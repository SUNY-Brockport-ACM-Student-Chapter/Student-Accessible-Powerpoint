import importlib.util
from pathlib import Path

from fastapi import FastAPI, status

from .jobs import cancel_job, commit_job, get_job_status, start_job

APP_ROOT = Path(__file__).resolve().parents[1]

app = FastAPI(
    title="Student Accessible PowerPoint Processing Service",
    description="Job orchestration and Chroma wrapper API for the Next.js migration.",
    version="0.1.0",
)


def include_chroma_routes() -> None:
    chroma_app_path = APP_ROOT / "chroma-api" / "app.py"
    spec = importlib.util.spec_from_file_location("sap_chroma_api", chroma_app_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load Chroma API app")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    for route in module.app.routes:
        if route.path in {"/openapi.json", "/docs", "/docs/oauth2-redirect", "/redoc"}:
            continue
        app.router.routes.append(route)


include_chroma_routes()


app.add_api_route(
    "/jobs/{job_id}/start",
    start_job,
    methods=["POST"],
    status_code=status.HTTP_202_ACCEPTED,
)
app.add_api_route(
    "/jobs/{job_id}/commit",
    commit_job,
    methods=["POST"],
    status_code=status.HTTP_202_ACCEPTED,
)
app.add_api_route("/jobs/{job_id}/cancel", cancel_job, methods=["POST"])
app.add_api_route("/jobs/{job_id}/status", get_job_status, methods=["GET"])
