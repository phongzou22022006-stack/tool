"""
Entry point cho JOB SCRAPE.
Chạy bởi GitHub Actions mỗi 8 tiếng.

Flow semi-auto review:
  1. Lấy tất cả source pages đang hoạt động (gộp theo group)
  2. Với mỗi group: scrape tất cả source pages qua Apify
  3. Với mỗi bài mới (chưa có trong dedup):
     → Chạy safety filter
     → Tạo AI summary / rewritten caption / hashtag / CTA suggestion
     → Ghi vào tab pending_review để duyệt thủ công
  4. Chỉ job scheduler riêng mới đăng các bài đã được approved
"""
import logging
import sys
from datetime import datetime, timezone

from src.config import DB_BACKEND

if DB_BACKEND == "sheets":
    from src.sheets_db import (
        get_active_source_pages,
        get_destination_pages_by_group,
        is_post_exists,
        save_dedup,
        save_log,
        save_pending_review,
        update_source_page_scraped_at,
    )
else:
    from src.database import (
        get_active_source_pages,
        get_destination_pages_by_group,
        is_post_exists,
        save_dedup,
        save_log,
        update_source_page_scraped_at,
    )
from src.apify_scraper import scrape_pages
from src.post_scheduler import max_posts_for_dest
from src.review_builder import build_candidate_from_post

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("main_scrape")


def run():
    logger.info("=" * 60)
    logger.info("BẮT ĐẦU JOB SCRAPE + PENDING REVIEW")
    logger.info("=" * 60)

    all_pages = get_active_source_pages()
    if not all_pages:
        logger.info("Không có source page nào. Thoát.")
        return

    # Gộp pages theo group_id
    groups: dict[str, list[dict]] = {}
    for page in all_pages:
        groups.setdefault(page["group_id"], []).append(page)

    logger.info(f"Tìm thấy {len(all_pages)} trang nguồn trong {len(groups)} nhóm")

    total_new     = 0
    total_errors  = 0

    for group_id, pages in groups.items():
        group_name = pages[0].get("page_groups", {}).get("name", group_id[:8])
        logger.info(f"\n--- Nhóm: {group_name} ({len(pages)} trang) ---")

        page_urls   = [p["fb_page_url"] for p in pages]
        url_to_page = {p["fb_page_url"]: p for p in pages}

        # Scrape
        try:
            scraped = scrape_pages(page_urls)
        except Exception as e:
            logger.error(f"Lỗi scrape group {group_name}: {e}")
            total_errors += 1
            continue

        # Gom tất cả bài mới (chưa tồn tại trong dedup) của group này
        new_posts: list[tuple[dict, dict]] = []  # [(post_data, page), ...]
        for page_url, posts in scraped.items():
            page = url_to_page.get(page_url)
            if not page:
                continue
            for post in posts:
                if not is_post_exists(post["fb_post_id"]):
                    new_posts.append((post, page))

            update_source_page_scraped_at(
                page["id"],
                datetime.now(timezone.utc).isoformat(),
            )

        if not new_posts:
            logger.info("  → Không có bài mới trong 8 tiếng qua")
            continue

        # Sắp xếp theo tương tác (cao nhất lên đầu)
        new_posts.sort(key=lambda x: x[0].get("engagement_score", 0), reverse=True)

        logger.info(f"  → {len(new_posts)} bài mới (đã sắp xếp theo tương tác):")
        for i, (p, _) in enumerate(new_posts):
            logger.info(
                f"    Bài {i+1}: id={p['fb_post_id'][:20]} | "
                f"imgs={len(p.get('image_urls') or [])} | "
                f"video={'có' if p.get('video_url') else 'không'} | "
                f"❤️{p.get('likes',0)} 💬{p.get('comments',0)} 🔁{p.get('shares',0)}"
            )

        # Lấy danh sách trang đích
        destinations = get_destination_pages_by_group(group_id)
        if not destinations:
            logger.warning(f"  → Group {group_name} không có trang đích nào")
            continue

        # Với mỗi trang đích → giới hạn số bài + tạo candidate
        for dest in destinations:
            dest_id   = dest["id"]
            dest_name = dest.get("fb_page_name", dest_id[:8])
            limit     = max_posts_for_dest(dest)

            # Chỉ lấy top-N bài chưa dedup cho trang đích này
            dest_posts = new_posts[:limit]
            candidates = []

            logger.info(
                f"  {dest_name}: tạo candidate cho {len(dest_posts)}/{len(new_posts)} bài "
                f"(max={limit})"
            )

            for idx, (post, page) in enumerate(dest_posts):
                fb_post_id = post["fb_post_id"]

                try:
                    candidate = build_candidate_from_post(
                        post=post,
                        page=page,
                        dest=dest,
                        group_id=group_id,
                    )
                    if candidate is None:
                        logger.info(f"    Bài {idx+1} bị filter block, bỏ qua")
                        continue

                    candidates.append(candidate)
                    save_dedup(fb_post_id, page["id"], dest_id)
                    save_log(
                        scheduled_post_id   = None,
                        fb_post_id          = fb_post_id,
                        destination_page_id = dest_id,
                        result              = "pending_review",
                        source_page_url     = page.get("fb_page_url", ""),
                        post_url            = post.get("post_url", ""),
                    )
                    logger.info(f"    Bài {idx+1}: candidate OK, risk={candidate.filter_result.risk_level}")

                except Exception as e:
                    logger.error(
                        f"  Lỗi tạo candidate bài {fb_post_id[:20]}... → {dest_name}: {e}"
                    )
                    save_log(
                        scheduled_post_id   = None,
                        fb_post_id          = fb_post_id,
                        destination_page_id = dest_id,
                        result              = "failed",
                        error_message       = str(e),
                        source_page_url     = page.get("fb_page_url", ""),
                        post_url            = post.get("post_url", ""),
                    )
                    total_errors += 1

            if candidates:
                save_pending_review([c.to_pending_review_row() for c in candidates])
                logger.info(f"  {dest_name}: đã ghi {len(candidates)} candidate vào pending_review")

        total_new += len(new_posts)

    logger.info("\n" + "=" * 60)
    logger.info(f"KẾT THÚC: +{total_new} bài mới | {total_errors} lỗi")
    logger.info("=" * 60)

    if total_errors > 0 and total_new == 0:
        sys.exit(1)


if __name__ == "__main__":
    run()
