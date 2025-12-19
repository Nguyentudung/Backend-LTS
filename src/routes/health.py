from __future__ import annotations

from fastapi import FastAPI


def register_health_routes(app: FastAPI):
    @app.get("/")
    def health():
        return {"status": "ok"}

    @app.get("/healthz")
    def healthz():
        return {"status": "ok"}

