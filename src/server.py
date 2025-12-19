import os
import tempfile
from pathlib import Path
from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from src.routes.deploy import register_deploy_routes
from src.routes.health import register_health_routes
from src.routes.processing import register_processing_routes

# =================================================
# PATH SETUP
# =================================================
BASE_DIR = Path(__file__).resolve().parent        # backend/src
ROOT_DIR = BASE_DIR.parent                       # backend

ENV_PATH = ROOT_DIR / ".env"
DIST_ZIP_PATH = ROOT_DIR / "dist.zip"
UPLOAD_DIR = Path(
    os.getenv(
        "UPLOAD_DIR",
        str(Path(tempfile.gettempdir()) / "alphawave_uploads"),
    )
)

UPLOAD_DIR.mkdir(exist_ok=True, parents=True)

# =================================================
# LOAD ENV
# =================================================
print(f"[startup] loading dotenv: {ENV_PATH}", flush=True)
load_dotenv(ENV_PATH)
print(f"[startup] PORT env: {os.getenv('PORT')}", flush=True)
print(f"[startup] upload dir: {UPLOAD_DIR}", flush=True)

NETLIFY_TOKEN = os.getenv("NETLIFY_TOKEN")
PORT = int(os.getenv("PORT", 5001))

if not NETLIFY_TOKEN:
    print("❌ NETLIFY_TOKEN chưa được set")

# =================================================
# CHECK IMAGEMAGICK (STARTUP)
# =================================================
import subprocess

try:
    from src.imagemagick import resolve_imagemagick
except ModuleNotFoundError:
    from imagemagick import resolve_imagemagick

print("[startup] checking ImageMagick...", flush=True)
try:
    magick = resolve_imagemagick(BASE_DIR)
    if magick is None:
        raise RuntimeError(
            "ImageMagick not found (set IMAGEMAGICK_BIN/MAGICK_BIN or install ImageMagick)"
        )
    out = subprocess.check_output(
        magick.command + ["-version"],
        env=magick.env,
        stderr=subprocess.STDOUT,
    )
    print(f"[startup] ImageMagick OK ({magick.source}):", flush=True)
    print(f"[startup] ImageMagick cmd: {magick.command[0]}", flush=True)
    print(out.decode(errors="replace"), flush=True)
except Exception as e:
    print("❌ [startup] ImageMagick NOT FOUND or ERROR", flush=True)
    print(str(e), flush=True)


# =================================================
# FASTAPI APP
# =================================================
app = FastAPI(title="AlphaWave Deploy API")

cors_allow_all = (os.getenv("CORS_ALLOW_ALL") or "").strip().lower() in {
    "1",
    "true",
    "yes",
}
cors_origins_raw = os.getenv("CORS_ORIGINS")
cors_origins = (
    [origin.strip() for origin in cors_origins_raw.split(",") if origin.strip()]
    if cors_origins_raw
    else [
        "http://localhost:5173",
        "https://alphawaveprep.netlify.app",
    ]
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"] if cors_allow_all else cors_origins,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

register_deploy_routes(
    app,
    netlify_token=NETLIFY_TOKEN,
    dist_zip_path=DIST_ZIP_PATH,
    upload_dir=UPLOAD_DIR,
)
register_processing_routes(app, upload_dir=UPLOAD_DIR)
register_health_routes(app)

# uvicorn src.server:app --host 0.0.0.0 --port 5001 --reload
