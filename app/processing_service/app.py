from fastapi import FastAPI, status

from .jobs import cancel_job, commit_job, get_job_status, start_job

app = FastAPI(
    title="Student Accessible PowerPoint Processing Service",
    description="Job orchestration API for the Next.js migration.",
    version="0.1.0",
)


@app.get("/health")
def health_check():
    return {"status": "healthy"}


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
