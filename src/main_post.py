"""
Entry point cho JOB POST APPROVED.
Chạy bởi GitHub Actions theo lịch riêng.

Flow:
  1. Lấy các dòng pending_review có status=approved
  2. Đăng/hẹn giờ từng bài lên Facebook Page tương ứng
  3. Cập nhật trạng thái scheduled hoặc failed
"""
import logging
import sys
from datetime import datetime, timedelta, timezone

from src.config import DB_BACKEND

if DB_BACKEND == "sheets":
    from src.sheets_db import get_approved_reviews, update_review_status, save_log
else:
    raise RuntimeError("Semi-auto review MVP hiện chỉ hỗ trợ DB_BACKEND=sheets")

from src.fb_poster import FacebookPoster

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("main_post")


def compute_schedule_time(index: int) -> datetime:
    now = datetime.now(timezone.utc)
    candidate = now + timedelta(minutes=15 + index * 10)
    bangkok_hour = (candidate.hour + 7) % 24
    if bangkok_hour >= 23 or bangkok_hour < 7:
        target_date = candidate.date()
        if bangkok_hour >= 23:
            target_date = (candidate + timedelta(days=1)).date()
        candidate = datetime(target_date.year, target_date.month, target_date.day, 0, 5, tzinfo=timezone.utc)
    return candidate


def run():
    logger.info("=" * 60)
    logger.info("BẮT ĐẦU JOB POST APPROVED")
    logger.info("=" * 60)

    rows = get_approved_reviews(limit=20)
    if not rows:
        logger.info("Không có bài approved nào. Thoát.")
        return

    total_ok = 0
    total_errors = 0

    for idx, row in enumerate(rows):
        try:
            schedule_time = compute_schedule_time(idx)
            poster = FacebookPoster(
                page_id=str(row["destination_page_id"]),
                access_token=str(_require(row, "fb_access_token")),
            )
            fb_created_id = poster.post(
                content=str(row.get("rewritten_caption") or row.get("original_text") or ""),
                image_urls=_split_lines(row.get("media_urls", "")) if str(row.get("media_type", "")) == "image" else [],
                video_url=_pick_video_url(row),
                reel_url=_pick_reel_url(row),
                scheduled_at=schedule_time,
            )
            update_review_status(
                str(row["id"]),
                "scheduled",
                {
                    "scheduled_time": schedule_time.isoformat(),
                    "facebook_post_id": fb_created_id,
                    "updated_at": datetime.now(timezone.utc).isoformat(),
                },
            )
            save_log(
                scheduled_post_id=None,
                fb_post_id=str(row.get("fb_post_id") or ""),
                destination_page_id=str(row["destination_page_id"]),
                result="scheduled",
                source_page_url=str(row.get("source_url") or ""),
                post_url=str(row.get("source_url") or ""),
            )
            total_ok += 1
        except Exception as e:
            total_errors += 1
            logger.error(f"Lỗi post approved id={row.get('id')}: {e}")
            update_review_status(
                str(row.get("id")),
                "failed",
                {
                    "review_note": f"schedule failed: {type(e).__name__}",
                    "updated_at": datetime.now(timezone.utc).isoformat(),
                },
            )
            save_log(
                scheduled_post_id=None,
                fb_post_id=str(row.get("fb_post_id") or ""),
                destination_page_id=str(row.get("destination_page_id") or ""),
                result="failed",
                error_message=str(e),
                source_page_url=str(row.get("source_url") or ""),
                post_url=str(row.get("source_url") or ""),
            )

    logger.info(f"KẾT THÚC JOB POST APPROVED: ok={total_ok} errors={total_errors}")
    if total_errors > 0 and total_ok == 0:
        sys.exit(1)


def _split_lines(value: str) -> list[str]:
    return [x.strip() for x in str(value or "").splitlines() if x.strip()]


def _pick_video_url(row: dict) -> str | None:
    if str(row.get("media_type", "")) != "video":
        return None
    lines = _split_lines(row.get("media_urls", ""))
    return lines[0] if lines else None


def _pick_reel_url(row: dict) -> str | None:
    if str(row.get("media_type", "")) != "reel":
        return None
    lines = _split_lines(row.get("media_urls", ""))
    return lines[0] if lines else None


def _require(row: dict, key: str) -> str:
    value = row.get(key)
    if not value:
        raise RuntimeError(f"Thiếu field bắt buộc: {key}")
    return str(value)


if __name__ == "__main__":
    run()
