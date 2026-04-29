from __future__ import annotations
import re
from src.models import FilterResult

SPAM_KEYWORDS = ["kiếm tiền nhanh", "cam kết lợi nhuận", "inbox nhận giá", "slot cuối", "lãi khủng"]
SENSITIVE_KEYWORDS = ["chính trị", "tôn giáo", "bạo lực", "tự tử", "khiêu dâm", "cờ bạc"]
URL_RE = re.compile(r"https?://\S+", re.I)

def evaluate_post_safety(*, text: str, media_type: str, media_urls: list[str], likes: int, comments: int, shares: int) -> FilterResult:
    normalized = " ".join((text or "").lower().split())
    warnings: list[str] = []
    blocks: list[str] = []
    if len(normalized) < 20 and not media_urls:
        blocks.append("too_short")
    if any(k in normalized for k in SPAM_KEYWORDS):
        blocks.append("spam_keyword")
    links = URL_RE.findall(text or "")
    if len(links) >= 3:
        blocks.append("too_many_links")
    elif links:
        warnings.append("external_link")
    if any(k in normalized for k in SENSITIVE_KEYWORDS):
        warnings.append("sensitive_keyword")
    if media_type in {"video", "reel"}:
        warnings.append("video_or_reel")
    if max(0, likes) + max(0, comments) + max(0, shares) > 100000:
        warnings.append("engagement_anomaly")
    if blocks:
        return FilterResult("high", "block", blocks + warnings)
    if warnings:
        return FilterResult("medium", "warning", warnings)
    return FilterResult("low", "pass", [])
