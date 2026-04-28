import io
import os
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from fastapi import BackgroundTasks, Depends, Header, HTTPException, status
from pydantic import BaseModel, Field

from .storage import download_upload_object, upload_output_object
from .webhooks import post_processor_webhook

APP_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP_ROOT))
sys.path.insert(0, str(APP_ROOT / "pptx_rag_quizzer"))

from models.models import Type  # noqa: E402
from pptx_rag_quizzer.image import Image as ImageProcessor  # noqa: E402
from pptx_rag_quizzer.pptx import (  # noqa: E402
    parse_powerpoint,
    rebuild_presentation_with_accessible_features,
)
from pptx_rag_quizzer.rag_core import RAGCore  # noqa: E402


class StartJobRequest(BaseModel):
    storage_object: str
    presentation_name: str


class CommitDescription(BaseModel):
    order_number: int
    alt_text: str
    slide_number: int | None = None


class CommitJobRequest(BaseModel):
    descriptions: list[CommitDescription] = Field(default_factory=list)
    storage_object: str | None = None
    presentation_name: str | None = None


@dataclass
class JobRuntimeState:
    status: str
    phase: str | None = None
    storage_object: str | None = None
    presentation_name: str | None = None
    collection_id: str | None = None
    output_object_path: str | None = None
    error: str | None = None
    cancelled: bool = False
    descriptions: list[dict[str, Any]] = field(default_factory=list)


JOB_STATES: dict[str, JobRuntimeState] = {}


def require_processor_secret(
    x_sap_processor_secret: str | None = Header(default=None),
) -> None:
    expected = os.getenv("PY_SERVICE_SHARED_SECRET")
    if not expected or x_sap_processor_secret != expected:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid processor shared secret",
        )


def api_ok(data: dict[str, Any]):
    return {"ok": True, "data": data}


def _webhook(job_id: str, **payload: Any) -> None:
    post_processor_webhook({"jobId": job_id, **payload})


def _image_items(presentation):
    for slide in presentation.slides:
        for item in slide.items:
            if item.type == Type.image:
                yield slide, item


def _apply_commit_descriptions(presentation, descriptions: list[CommitDescription]) -> None:
    by_slide_and_order = {
        (description.slide_number, description.order_number): description.alt_text
        for description in descriptions
        if description.slide_number is not None
    }
    by_order = {
        description.order_number: description.alt_text
        for description in descriptions
        if description.slide_number is None
    }

    for slide, item in _image_items(presentation):
        alt_text = by_slide_and_order.get((slide.slide_number, item.order_number))
        if alt_text is None:
            alt_text = by_order.get(item.order_number)
        if alt_text:
            item.content = alt_text


def _describe_job(job_id: str, request: StartJobRequest) -> None:
    state = JOB_STATES[job_id]

    try:
        _webhook(
            job_id,
            status="parsing",
            phase="Downloading uploaded deck",
            progressCurrent=0,
            progressTotal=0,
        )
        file_bytes = download_upload_object(request.storage_object)

        if state.cancelled:
            return

        presentation = parse_powerpoint(io.BytesIO(file_bytes), request.presentation_name)
        image_items = list(_image_items(presentation))

        _webhook(
            job_id,
            status="describing",
            phase="Building presentation context",
            progressCurrent=0,
            progressTotal=len(image_items),
        )

        rag_core = RAGCore()
        collection_id = rag_core.create_collection(presentation)
        image_processor = ImageProcessor(rag_core)
        descriptions: list[dict[str, Any]] = []

        for index, (slide, item) in enumerate(image_items, start=1):
            if state.cancelled:
                return

            description = image_processor.describe_image(
                image_bytes=item.image_bytes,
                image_format=item.extension,
                slide_number=slide.slide_number,
                collection_id=collection_id,
            )
            item.content = description
            descriptions.append(
                {
                    "slide_number": slide.slide_number,
                    "order_number": item.order_number,
                    "item_type": Type.image.value,
                    "alt_text": description,
                }
            )
            _webhook(
                job_id,
                status="describing",
                phase="Describing images",
                progressCurrent=index,
                progressTotal=len(image_items),
                collectionId=collection_id,
            )

        state.status = "awaiting_review"
        state.phase = "Awaiting review"
        state.collection_id = collection_id
        state.descriptions = descriptions
        _webhook(
            job_id,
            status="awaiting_review",
            phase="Awaiting review",
            progressCurrent=len(image_items),
            progressTotal=len(image_items),
            collectionId=collection_id,
            descriptions=descriptions,
        )
    except Exception as exc:
        state.status = "error"
        state.error = str(exc)
        _webhook(
            job_id,
            status="error",
            phase="Processing failed",
            errorCode="PARSE_FAILED",
            errorMessage=str(exc),
        )
        print(traceback.format_exc())


def _commit_job(job_id: str, request: CommitJobRequest) -> None:
    state = JOB_STATES[job_id]

    try:
        storage_object = request.storage_object or state.storage_object
        presentation_name = request.presentation_name or state.presentation_name
        if not storage_object or not presentation_name:
            raise ValueError("Missing original upload reference for commit")

        _webhook(job_id, status="rebuilding", phase="Rebuilding deck")
        file_bytes = download_upload_object(storage_object)
        presentation = parse_powerpoint(io.BytesIO(file_bytes), presentation_name)
        _apply_commit_descriptions(presentation, request.descriptions)

        rebuilt = rebuild_presentation_with_accessible_features(
            presentation,
            io.BytesIO(file_bytes),
        )
        output_buffer = io.BytesIO()
        rebuilt.save(output_buffer)
        output_object_path = f"{job_id}/{Path(presentation_name).stem}-accessible.pptx"
        upload_output_object(output_object_path, output_buffer.getvalue())

        state.status = "ready"
        state.output_object_path = output_object_path
        _webhook(
            job_id,
            status="ready",
            phase="Ready for download",
            outputObjectPath=output_object_path,
        )
    except Exception as exc:
        state.status = "error"
        state.error = str(exc)
        _webhook(
            job_id,
            status="error",
            phase="Rebuild failed",
            errorCode="REBUILD_FAILED",
            errorMessage=str(exc),
        )
        print(traceback.format_exc())


def start_job(
    job_id: str,
    request: StartJobRequest,
    background_tasks: BackgroundTasks,
    _auth: None = Depends(require_processor_secret),
):
    JOB_STATES[job_id] = JobRuntimeState(
        status="queued",
        storage_object=request.storage_object,
        presentation_name=request.presentation_name,
    )
    background_tasks.add_task(_describe_job, job_id, request)
    return api_ok({"jobId": job_id})


def commit_job(
    job_id: str,
    request: CommitJobRequest,
    background_tasks: BackgroundTasks,
    _auth: None = Depends(require_processor_secret),
):
    if job_id not in JOB_STATES:
        JOB_STATES[job_id] = JobRuntimeState(status="rebuilding")
    background_tasks.add_task(_commit_job, job_id, request)
    return api_ok({"jobId": job_id})


def cancel_job(job_id: str, _auth: None = Depends(require_processor_secret)):
    state = JOB_STATES.setdefault(job_id, JobRuntimeState(status="cancelled"))
    state.cancelled = True
    state.status = "cancelled"
    _webhook(job_id, status="cancelled", phase="Cancelled")
    return api_ok({"jobId": job_id})


def get_job_status(job_id: str, _auth: None = Depends(require_processor_secret)):
    state = JOB_STATES.get(job_id)
    if not state:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Job not found")

    return api_ok(
        {
            "jobId": job_id,
            "status": state.status,
            "phase": state.phase,
            "collectionId": state.collection_id,
            "outputObjectPath": state.output_object_path,
            "error": state.error,
        }
    )
