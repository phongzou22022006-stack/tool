"""
Google Sheets DB — gọi Google Apps Script Web App.
Thay thế hoàn toàn Supabase, không cần Service Account hay gspread.

Cần 2 biến môi trường:
  GOOGLE_WEBAPP_URL   — URL của Apps Script Web App sau khi deploy
  WEBAPP_SECRET       — Token bảo mật (tùy chọn, khớp với Script Properties)
"""
import logging
import os
import httpx

logger = logging.getLogger(__name__)

_WEBAPP_URL = os.getenv("GOOGLE_WEBAPP_URL", "")
_SECRET     = os.getenv("WEBAPP_SECRET", "")

# Timeout cho mỗi lần gọi Web App (Sheets đôi khi chậm)
_TIMEOUT = 30


def _get(action: str, **params) -> dict:
    """Gọi GET request tới Web App."""
    if not _WEBAPP_URL:
        raise RuntimeError("GOOGLE_WEBAPP_URL chưa được cấu hình trong .env")
    resp = httpx.get(
        _WEBAPP_URL,
        params={"action": action, "token": _SECRET, **params},
        timeout=_TIMEOUT,
        follow_redirects=True,
    )
    resp.raise_for_status()
    return resp.json()


def _post(action: str, **body) -> dict:
    """Gọi POST request tới Web App."""
    if not _WEBAPP_URL:
        raise RuntimeError("GOOGLE_WEBAPP_URL chưa được cấu hình trong .env")
    resp = httpx.post(
        _WEBAPP_URL,
        params={"token": _SECRET},
        json={"action": action, **body},
        timeout=_TIMEOUT,
        follow_redirects=True,
    )
    resp.raise_for_status()
    if not resp.content or not resp.text.strip():
        return {}
    try:
        return resp.json()
    except Exception:
        logger.warning(f"_post({action}): response không phải JSON: {resp.text[:100]}")
        return {}


# ── Source Pages ───────────────────────────────────────────

def get_active_source_pages() -> list[dict]:
    return _get("get_active_sources")


def update_source_page_scraped_at(page_id: str, scraped_at: str):
    """page_id ở đây là fb_page_url."""
    _post("update_source_scraped_at", page_url=page_id, scraped_at=scraped_at)


# ── Destination Pages ──────────────────────────────────────

def get_destination_pages_by_group(group_id: str) -> list[dict]:
    return _get("get_destinations", group_id=group_id)


# ── Dedup ──────────────────────────────────────────────────

def is_post_exists(fb_post_id: str) -> bool:
    result = _get("is_dedup", fb_post_id=fb_post_id)
    return result.get("exists", False)


def save_dedup(fb_post_id: str, source_page_id: str,
               destination_page_id: str):
    _post("save_dedup",
          fb_post_id=fb_post_id,
          source_page_id=source_page_id,
          destination_page_id=destination_page_id)


# ── Schedule tracking ──────────────────────────────────────

def get_dest_last_scheduled(destination_page_id: str) -> str | None:
    result = _get("get_last_scheduled", dest_page_id=destination_page_id)
    return result.get("last_scheduled_at")


def update_dest_last_scheduled(destination_page_id: str,
                                scheduled_at: str):
    _post("update_last_scheduled",
          dest_page_id=destination_page_id,
          last_scheduled_at=scheduled_at)


# ── Logs ───────────────────────────────────────────────────

def save_log(scheduled_post_id: str | None, fb_post_id: str,
             destination_page_id: str, result: str,
             error_message: str | None = None,
             source_page_url: str | None = None,
             post_url: str | None = None):
    _post("save_log",
          fb_post_id=fb_post_id,
          destination_page_id=destination_page_id,
          result=result,
          error_message=error_message or "",
          source_page_url=source_page_url or "",
          post_url=post_url or "")


# ── Pending Review ─────────────────────────────────────────

def save_pending_review(rows: list[dict]):
    if not rows:
        return
    _post("save_pending_review", rows=rows)


def get_approved_reviews(limit: int = 20) -> list[dict]:
    result = _get("get_approved_reviews", limit=str(limit))
    return result.get("rows", [])


def update_review_status(row_id: str, status: str, fields: dict | None = None):
    _post("update_review_status", id=row_id, status=status, fields=fields or {})


# ── Apify Keys ─────────────────────────────────────────────

def get_active_apify_key(_safe_limit: int = 450) -> str | None:
    result = _get("get_apify_key")
    return result.get("api_key")


def increment_apify_key_usage(api_key: str, count: int = 1):
    _post("increment_apify_usage", api_key=api_key, count=count)


def mark_apify_key_exhausted(api_key: str):
    """Đánh dấu key lỗi/hết quota bằng cách đẩy usage lên cao để không được chọn nữa."""
    _post("increment_apify_usage", api_key=api_key, count=500)
