from __future__ import annotations

import json
import shutil
import time
import traceback
import zipfile
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from src.docx_parser.parser import parse_docx


def register_processing_routes(app: FastAPI, *, upload_dir: Path):
    @app.post("/api/processing")
    async def processing(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
        filename = file.filename or ""
        if not filename.lower().endswith(".docx"):
            raise HTTPException(400, "Chỉ chấp nhận file DOCX (.docx)")

        timestamp = int(time.time() * 1000)
        work_dir = upload_dir / f"processing_{timestamp}"
        docx_path = work_dir / "input.docx"
        image_dir = work_dir / "extracted_images"

        work_dir.mkdir(parents=True, exist_ok=True)

        with open(docx_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        try:
            result = parse_docx(
                str(docx_path),
                image_output_dir=image_dir,
                image_public_dir="extracted_images",
            )

            zip_path = work_dir / f"processing_{timestamp}.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
                zip_out.writestr(
                    "data.json",
                    json.dumps(result, ensure_ascii=False, indent=2),
                )

                if image_dir.exists():
                    for path in image_dir.rglob("*"):
                        if path.is_file():
                            if path.suffix.lower() == ".wmf":
                                continue
                            zip_out.write(
                                path,
                                arcname=path.relative_to(work_dir).as_posix(),
                            )

            background_tasks.add_task(shutil.rmtree, work_dir, ignore_errors=True)
            return FileResponse(
                zip_path,
                media_type="application/zip",
                headers={
                    "Content-Disposition": f'attachment; filename="processing_{timestamp}.zip"'
                },
                background=background_tasks,
            )
        except Exception as e:
            shutil.rmtree(work_dir, ignore_errors=True)
            print("❌ Lỗi processing DOCX:\n", traceback.format_exc())
            return JSONResponse(
                status_code=500,
                content={
                    "error": "processing_failed",
                    "detail": str(e),
                },
            )

