from src.ai_enrichment import fallback_enrich_post, parse_ai_enrichment_json

def test_fallback_has_required_fields():
    r = fallback_enrich_post(original_text="Một quán cà phê mới ở Hà Nội có không gian yên tĩnh và đồ uống ổn.")
    assert r.ai_summary
    assert r.rewritten_caption
    assert r.suggested_hashtags
    assert r.suggested_cta

def test_parse_ai_json():
    raw='{"ai_summary":"Tóm tắt","rewritten_caption":"Caption","suggested_hashtags":["#a"],"suggested_cta":"CTA","confidence":0.8}'
    r=parse_ai_enrichment_json(raw)
    assert r.ai_summary == "Tóm tắt"
    assert r.suggested_hashtags == ["#a"]
