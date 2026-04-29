from __future__ import annotations
from datetime import datetime, timezone
from uuid import uuid5, NAMESPACE_URL

from src.ai_enrichment import enrich_post_with_ai
from src.models import CandidatePost
from src.safety_filter import evaluate_post_safety


def build_candidate_from_post(*, post: dict, page: dict, dest: dict, group_id: str) -> CandidatePost | None:
    content = post.get("content") or ""
    image_urls = post.get("image_urls") or []
    video_url = post.get("video_url")
    reel_url = post.get("reel_url")
    media_urls = list(image_urls)
    if video_url:
        media_urls.append(video_url)
    if reel_url:
        media_urls.append(reel_url)
    media_type = "reel" if reel_url else "video" if video_url else "image" if image_urls else "text"

    filter_result = evaluate_post_safety(
        text=content,
        media_type=media_type,
        media_urls=media_urls,
        likes=int(post.get("likes") or 0),
        comments=int(post.get("comments") or 0),
        shares=int(post.get("shares") or 0),
    )
    if filter_result.result == "block":
        return None

    ai = enrich_post_with_ai(
        original_text=content,
        page_tone=dest.get("tone") or "ngắn, tự nhiên, tiếng Việt",
    )
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    fb_post_id = str(post["fb_post_id"])
    dest_id = str(dest["id"])
    candidate_id = str(uuid5(NAMESPACE_URL, f"{dest_id}:{fb_post_id}"))

    return CandidatePost(
        id=candidate_id,
        group_id=group_id,
        destination_page_id=dest_id,
        source_page=page.get("fb_page_name") or page.get("fb_page_url", ""),
        source_url=post.get("post_url") or page.get("fb_page_url", ""),
        fb_post_id=fb_post_id,
        original_text=content,
        media_type=media_type,
        media_urls=media_urls,
        likes=int(post.get("likes") or 0),
        comments=int(post.get("comments") or 0),
        shares=int(post.get("shares") or 0),
        filter_result=filter_result,
        ai=ai,
        created_at=now,
        updated_at=now,
    )
