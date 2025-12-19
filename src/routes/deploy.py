from __future__ import annotations

import shutil
import time
import zipfile
from pathlib import Path

import requests
from fastapi import FastAPI, File, HTTPException, UploadFile


def deploy_to_netlify(*, zip_path: Path, site_name: str, netlify_token: str):
    if not netlify_token:
        raise Exception("NETLIFY_TOKEN chưa cấu hình")

    print("\n=== BẮT ĐẦU DEPLOY ===")
    print("Site name:", site_name)
    print("Zip path:", zip_path)

    # 1️⃣ Create site
    print("[1/3] Tạo site Netlify...")
    create_res = requests.post(
        "https://api.netlify.com/api/v1/sites",
        headers={
            "Authorization": f"Bearer {netlify_token}",
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
        raise Exception("Response thiếu site ID")

    print("Site ID:", site_id)

    # 2️⃣ Deploy zip
    print("[2/3] Upload ZIP...")
    with open(zip_path, "rb") as f:
        deploy_res = requests.post(
            f"https://api.netlify.com/api/v1/sites/{site_id}/deploys",
            headers={
                "Authorization": f"Bearer {netlify_token}",
                "Content-Type": "application/zip",
            },
            data=f,
        )

    if not deploy_res.ok:
        raise Exception(f"Deploy thất bại: {deploy_res.text}")

    deploy_json = deploy_res.json()
    print("Deploy ID:", deploy_json.get("id"))

    # 3️⃣ Done
    print("[3/3] Deploy hoàn tất")

    return {
        "site": site_json,
        "deploy": deploy_json,
        "liveUrl": deploy_json.get("ssl_url") or site_json.get("ssl_url"),
    }


def register_deploy_routes(
    app: FastAPI,
    *,
    netlify_token: str | None,
    dist_zip_path: Path,
    upload_dir: Path,
):
    @app.post("/api/deploy")
    @app.post("/api/create")
    async def deploy(file: UploadFile = File(...)):
        if not netlify_token:
            raise HTTPException(500, "Server chưa cấu hình NETLIFY_TOKEN")

        if not (file.filename or "").lower().endswith(".zip"):
            raise HTTPException(400, "Chỉ chấp nhận file ZIP")

        if not dist_zip_path.exists():
            raise HTTPException(400, f"dist.zip không tồn tại tại {dist_zip_path}")

        timestamp = int(time.time() * 1000)

        data_zip_path = upload_dir / f"data_{timestamp}.zip"
        merged_zip_path = upload_dir / f"dist_with_data_{timestamp}.zip"

        # Save uploaded zip
        with open(data_zip_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        try:
            # Merge zip
            with zipfile.ZipFile(
                merged_zip_path,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as merged_zip:
                # 1️⃣ Add dist.zip
                with zipfile.ZipFile(dist_zip_path, "r") as dist_zip:
                    for entry in dist_zip.infolist():
                        if not entry.is_dir():
                            merged_zip.writestr(
                                entry.filename,
                                dist_zip.read(entry.filename),
                            )

                # 2️⃣ Add _redirects
                merged_zip.writestr("_redirects", "/* /index.html 200")
                print("-> Đã thêm _redirects")

                # 3️⃣ Add data.zip
                with zipfile.ZipFile(data_zip_path, "r") as data_zip:
                    for entry in data_zip.infolist():
                        if entry.is_dir():
                            merged_zip.writestr(entry.filename, b"")
                        else:
                            merged_zip.writestr(
                                entry.filename,
                                data_zip.read(entry.filename),
                            )
                        print(f"-> Merge data entry: {entry.filename}")

            # Deploy to Netlify
            site_name = f"alphawave-quiz-{timestamp}"
            result = deploy_to_netlify(
                zip_path=merged_zip_path,
                site_name=site_name,
                netlify_token=netlify_token,
            )

            return {
                "message": "Deploy thành công",
                "url": result["liveUrl"],
            }
        except Exception as e:
            print("❌ Lỗi deploy:", str(e))
            raise HTTPException(500, str(e))
        finally:
            for p in [data_zip_path, merged_zip_path]:
                try:
                    p.unlink()
                except Exception:
                    pass

