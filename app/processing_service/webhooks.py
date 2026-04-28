import os
from typing import Any

import requests


def _required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def post_processor_webhook(payload: dict[str, Any]) -> None:
    app_url = _required_env("NEXT_PUBLIC_APP_URL").rstrip("/")
    response = requests.post(
        f"{app_url}/api/webhooks/processor",
        headers={
            "content-type": "application/json",
            "x-sap-processor-secret": _required_env("PY_SERVICE_SHARED_SECRET"),
        },
        json=payload,
        timeout=30,
    )
    response.raise_for_status()
