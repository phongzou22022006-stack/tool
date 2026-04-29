from __future__ import annotations
from dataclasses import dataclass

@dataclass(frozen=True)
class FilterResult:
    risk_level: str
    result: str
    reasons: list[str]

@dataclass(frozen=True)
class AIEnrichment:
    ai_summary: str
    rewritten_caption: str
    suggested_hashtags: list[str]
    suggested_cta: str
    confidence: float = 0.0

@dataclass(frozen=True)
class CandidatePost:
    id: str
    group_id: str
    destination_page_id: str
    source_page: str
    source_url: str
    fb_post_id: str
    original_text: str
    media_type: str
    media_urls: list[str]
    likes: int
    comments: int
    shares: int
    filter_result: FilterResult
    ai: AIEnrichment
    created_at: str
    updated_at: str

    def to_pending_review_row(self) -> dict:
        return {
            "id": self.id,
            "group_id": self.group_id,
            "destination_page_id": self.destination_page_id,
            "source_page": self.source_page,
            "source_url": self.source_url,
            "fb_post_id": self.fb_post_id,
            "original_text": self.original_text,
            "ai_summary": self.ai.ai_summary,
            "rewritten_caption": self.ai.rewritten_caption,
            "suggested_hashtags": " ".join(self.ai.suggested_hashtags),
            "suggested_cta": self.ai.suggested_cta,
            "media_type": self.media_type,
            "media_urls": "\n".join(self.media_urls),
            "likes": self.likes,
            "comments": self.comments,
            "shares": self.shares,
            "risk_level": self.filter_result.risk_level,
            "filter_result": self.filter_result.result,
            "filter_reasons": ",".join(self.filter_result.reasons),
            "status": "pending_review",
            "reviewed_by": "",
            "reviewed_at": "",
            "review_note": "",
            "scheduled_time": "",
            "facebook_post_id": "",
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }
