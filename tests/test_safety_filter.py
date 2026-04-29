from src.safety_filter import evaluate_post_safety

def test_blocks_short_text_without_media():
    r = evaluate_post_safety(text="sale", media_type="text", media_urls=[], likes=0, comments=0, shares=0)
    assert r.result == "block"
    assert "too_short" in r.reasons

def test_warns_sensitive_keyword():
    r = evaluate_post_safety(text="Tin chính trị đang gây tranh cãi trong cộng đồng", media_type="text", media_urls=[], likes=1, comments=0, shares=0)
    assert r.result == "warning"
    assert "sensitive_keyword" in r.reasons

def test_passes_normal_content():
    r = evaluate_post_safety(text="Một quán cà phê mới ở Hà Nội có không gian yên tĩnh và đồ uống ổn.", media_type="image", media_urls=["x.jpg"], likes=1, comments=0, shares=0)
    assert r.result == "pass"
