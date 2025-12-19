import os
import time
import requests
from pathlib import Path
from dotenv import load_dotenv

# =================================================
# LOAD ENV
# =================================================
BASE_DIR = Path(__file__).resolve().parent.parent.parent
ENV_PATH = BASE_DIR / ".env"

load_dotenv(ENV_PATH)

NETLIFY_TOKEN = os.getenv("NETLIFY_TOKEN")

if not NETLIFY_TOKEN:
    raise RuntimeError("❌ NETLIFY_TOKEN chưa được cấu hình")


# =================================================
# NETLIFY CREATE & DEPLOY
# =================================================
def create_and_deploy_site(zip_path: Path, site_name: str) -> dict:
    """
    Tạo site mới và deploy zip lên Netlify
    """
    if not zip_path.exists():
        raise FileNotFoundError(f"Zip không tồn tại: {zip_path}")

    print("\n=== NETLIFY DEPLOY ===")
    print("Site name:", site_name)
    print("Zip:", zip_path)

    # 1️⃣ Create site
    create_res = requests.post(
        "https://api.netlify.com/api/v1/sites",
        headers={
            "Authorization": f"Bearer {NETLIFY_TOKEN}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "alphawave-backend/1.0",
        },
        json={"name": site_name},
    )

    if not create_res.ok:
        raise Exception(f"Tạo site thất bại: {create_res.text}")

    site_json = create_res.json()
    site_id = site_json.get("id")

    if not site_id:
        raise Exception("Response không có site_id")

    # 2️⃣ Deploy zip
    with open(zip_path, "rb") as f:
        deploy_res = requests.post(
            f"https://api.netlify.com/api/v1/sites/{site_id}/deploys",
            headers={
                "Authorization": f"Bearer {NETLIFY_TOKEN}",
                "Content-Type": "application/zip",
            },
            data=f,
        )

    if not deploy_res.ok:
        raise Exception(f"Deploy thất bại: {deploy_res.text}")

    deploy_json = deploy_res.json()

    return {
        "site_id": site_id,
        "site_name": site_name,
        "deploy_id": deploy_json.get("id"),
        "liveUrl": deploy_json.get("ssl_url") or site_json.get("ssl_url"),
    }


def wait_for_deploy_ready(
    deploy_id: str | None,
    *,
    timeout_seconds: int = 120,
    poll_interval_seconds: float = 2.0,
) -> dict | None:
    """
    Poll Netlify deploy state until it's ready (or timeout).
    """
    if not deploy_id:
        return None

    deadline = time.time() + max(1, timeout_seconds)
    last: dict | None = None

    while time.time() < deadline:
        res = requests.get(
            f"https://api.netlify.com/api/v1/deploys/{deploy_id}",
            headers={
                "Authorization": f"Bearer {NETLIFY_TOKEN}",
                "Accept": "application/json",
                "User-Agent": "alphawave-backend/1.0",
            },
        )

        if not res.ok:
            raise Exception(f"Không thể kiểm tra deploy: {res.text}")

        last = res.json()
        state = (last.get("state") or "").lower()

        if state in {"ready", "current"}:
            return last

        if state in {"error", "failed"}:
            raise Exception(f"Deploy lỗi: state={state}")

        time.sleep(max(0.2, poll_interval_seconds))

    return last
