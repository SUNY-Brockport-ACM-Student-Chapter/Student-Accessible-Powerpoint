import os
from urllib.parse import quote

import requests

PPTX_MIME_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)


def _required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def _storage_headers(content_type: str | None = None) -> dict[str, str]:
    headers = {
        "apikey": _required_env("SUPABASE_SERVICE_ROLE_KEY"),
        "authorization": f"Bearer {_required_env('SUPABASE_SERVICE_ROLE_KEY')}",
    }
    if content_type:
        headers["content-type"] = content_type
    return headers


def _object_url(bucket: str, object_path: str) -> str:
    encoded_path = "/".join(quote(part, safe="") for part in object_path.split("/"))
    return (
        f"{_required_env('SUPABASE_URL').rstrip('/')}"
        f"/storage/v1/object/{bucket}/{encoded_path}"
    )


def download_upload_object(object_path: str) -> bytes:
    response = requests.get(
        _object_url(_required_env("SUPABASE_UPLOADS_BUCKET"), object_path),
        headers=_storage_headers(),
        timeout=60,
    )
    response.raise_for_status()
    return response.content


def upload_output_object(object_path: str, content: bytes) -> str:
    response = requests.post(
        _object_url(_required_env("SUPABASE_OUTPUTS_BUCKET"), object_path),
        headers={
            **_storage_headers(PPTX_MIME_TYPE),
            "x-upsert": "false",
        },
        data=content,
        timeout=60,
    )
    response.raise_for_status()
    return object_path
