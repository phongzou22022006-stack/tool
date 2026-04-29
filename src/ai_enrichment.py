from __future__ import annotations
import json
import os
import httpx
from src.models import AIEnrichment

class AIEnrichmentError(RuntimeError):
    pass

def fallback_enrich_post(*, original_text: str, page_tone: str = "") -> AIEnrichment:
    compact = " ".join((original_text or "").split())
    summary = compact[:140].rstrip() + ("..." if len(compact) > 140 else "")
    return AIEnrichment(
        ai_summary=summary or "Không có đủ nội dung để tóm tắt.",
        rewritten_caption=compact or "Nội dung này cần được biên tập thêm trước khi đăng.",
        suggested_hashtags=["#capnhat"],
        suggested_cta="Bạn nghĩ sao về nội dung này?",
        confidence=0.35,
    )

def parse_ai_enrichment_json(raw: str) -> AIEnrichment:
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise AIEnrichmentError(str(exc)) from exc
    tags = data.get("suggested_hashtags", [])
    if isinstance(tags, str):
        tags = [t for t in tags.split() if t]
    if not isinstance(tags, list):
        tags = []
    return AIEnrichment(
        ai_summary=str(data.get("ai_summary", "")).strip(),
        rewritten_caption=str(data.get("rewritten_caption", "")).strip(),
        suggested_hashtags=[str(t).strip() for t in tags if str(t).strip()],
        suggested_cta=str(data.get("suggested_cta", "")).strip(),
        confidence=float(data.get("confidence", 0.0)),
    )

def enrich_post_with_ai(*, original_text: str, page_tone: str = "") -> AIEnrichment:
    api_key = os.getenv("AI_API_KEY")
    base_url = os.getenv("AI_BASE_URL")
    model = os.getenv("AI_MODEL", "gpt-4o-mini")
    if not api_key or not base_url:
        return fallback_enrich_post(original_text=original_text, page_tone=page_tone)
    prompt = build_prompt(original_text, page_tone)
    r = httpx.post(
        base_url.rstrip("/") + "/chat/completions",
        headers={"Authorization": f"Bearer {api_key}"},
        json={"model": model, "messages": [{"role": "system", "content": "Return strict JSON only."}, {"role": "user", "content": prompt}], "temperature": 0.4},
        timeout=60,
        follow_redirects=True,
    )
    r.raise_for_status()
    content = r.json()["choices"][0]["message"]["content"]
    return parse_ai_enrichment_json(content)

def build_prompt(original_text: str, page_tone: str) -> str:
    return f"""Viết lại nội dung Facebook cho chế độ duyệt thủ công.\n\nNội dung gốc:\n{original_text}\n\nGiọng page: {page_tone or 'ngắn, tự nhiên, tiếng Việt'}\n\nTrả JSON với keys: ai_summary, rewritten_caption, suggested_hashtags, suggested_cta, confidence. Không bịa, không copy nguyên văn."""
