from __future__ import annotations

import threading
from typing import Any

import requests
import streamlit as st


def backend_settings() -> tuple[str | None, str | None]:
    """Return the Apps Script endpoint and shared secret, if configured."""
    try:
        url = st.secrets.get("apps_script_url")
        secret = st.secrets.get("apps_script_secret")
    except Exception:
        return None, None
    return str(url).strip() if url else None, str(secret).strip() if secret else None


def backend_enabled() -> bool:
    url, secret = backend_settings()
    return bool(url and secret)


def _post_to_endpoint(
    url: str,
    secret: str,
    action: str,
    payload: dict[str, Any],
    timeout: int,
) -> tuple[bool, str]:
    body = {"secret": secret, "action": action, "payload": payload}
    try:
        response = requests.post(url, json=body, timeout=timeout)
        response.raise_for_status()
        result = response.json()
    except Exception as exc:
        return False, f"Backend request failed: {exc}"

    if not result.get("ok"):
        return False, str(result.get("error", "The backend rejected the request."))
    return True, str(result.get("message", "Done."))


def post_backend(
    action: str,
    payload: dict[str, Any],
    timeout: int = 30,
) -> tuple[bool, str]:
    """Send one action and wait for the Apps Script response."""
    url, secret = backend_settings()
    if not url or not secret:
        return False, "Google Sheets and email delivery are not configured yet."
    return _post_to_endpoint(url, secret, action, payload, timeout)


def post_backend_background(action: str, payload: dict[str, Any]) -> bool:
    """Dispatch usage logging without blocking the Streamlit interaction."""
    url, secret = backend_settings()
    if not url or not secret:
        return False

    def worker() -> None:
        _post_to_endpoint(url, secret, action, payload, timeout=8)

    threading.Thread(target=worker, daemon=True).start()
    return True
