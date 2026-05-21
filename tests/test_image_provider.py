"""image_provider 测试 — 真实 API 用 mock 替代,避免 CI 成本。"""

import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch, MagicMock

sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.pipeline.image_provider import (
    ManualImageProvider, LocalAssetProvider,
    OpenAIImageProvider, get_provider,
    ASPECT_TO_API_SIZE, parse_aspect, crop_to_aspect,
)
from scripts.pipeline.image_styles import (
    STYLES, build_prompt, list_styles,
)

# 用项目内 tests/_tmp/ 避免 sandbox /tmp 权限问题
TMP = Path(__file__).parent / "_tmp"
TMP.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════════
# Aspect 工具测试
# ══════════════════════════════════════════════════════════
def test_parse_aspect():
    assert abs(parse_aspect("16:9") - 1.7777) < 0.001
    assert abs(parse_aspect("1:1") - 1.0) < 0.001
    assert abs(parse_aspect("9:16") - 0.5625) < 0.001


def test_aspect_map_has_common_ratios():
    assert "16:9" in ASPECT_TO_API_SIZE
    assert "1:1" in ASPECT_TO_API_SIZE
    assert "9:16" in ASPECT_TO_API_SIZE
    # 16:9 应该用 1536x1024 + 后处理裁
    api_size, crop = ASPECT_TO_API_SIZE["16:9"]
    assert api_size == "1536x1024"
    assert crop == "16:9"


# ══════════════════════════════════════════════════════════
# ManualImageProvider 测试
# ══════════════════════════════════════════════════════════
def test_manual_generate_returns_path():
    p = ManualImageProvider(work_dir=str(TMP / "manual"))
    path = p.generate("test prompt", aspect="16:9")
    assert isinstance(path, Path)
    stub = path.with_suffix(".prompt.txt")
    assert stub.exists()
    stub.unlink()


def test_manual_generate_accepts_aspect():
    """新签名:aspect= 取代 size=。"""
    p = ManualImageProvider(work_dir=str(TMP / "manual"))
    # 应该接受 aspect 参数(不抛 TypeError)
    p.generate("test", aspect="1:1")
    p.generate("test", aspect="9:16")
    p.generate("test", aspect="16:9", quality="high")


def test_manual_edit_returns_path():
    p = ManualImageProvider(work_dir=str(TMP / "manual"))
    path = p.edit(Path("/tmp/fake.png"), "remove background")
    assert isinstance(path, Path)
    stub = path.with_suffix(".prompt.txt")
    if stub.exists():
        stub.unlink()


# ══════════════════════════════════════════════════════════
# image_styles 测试(v0.5.0 — 返回 tuple)
# ══════════════════════════════════════════════════════════
def test_styles_dict_not_empty():
    assert len(STYLES) >= 10
    # v0.5.0 必须含核心风格
    assert "hero_editorial" in STYLES
    assert "hero_keynote_dark" in STYLES
    assert "concept_growth_metric" in STYLES
    assert "audience_lifestyle_cn" in STYLES


def test_styles_have_aspect():
    """每个风格必须含 aspect 字段。"""
    for k, v in STYLES.items():
        assert "prompt" in v, f"{k} missing prompt"
        assert "aspect" in v, f"{k} missing aspect"
        assert "description" in v, f"{k} missing description"


def test_build_prompt_returns_tuple():
    """v0.5.0:build_prompt 返回 (prompt, aspect) 元组。"""
    result = build_prompt(
        subject="AcmeCorp brand launch",
        style="hero_editorial",
        accent_hex="#7B61FF",
    )
    assert isinstance(result, tuple)
    assert len(result) == 2
    prompt, aspect = result
    assert isinstance(prompt, str)
    assert isinstance(aspect, str)
    # hex 强约束应在 prompt 里
    assert "#7B61FF" in prompt
    # aspect 应该是预设的 16:9
    assert aspect == "16:9"


def test_build_prompt_override_aspect():
    """override_aspect 应覆盖默认。"""
    _, asp = build_prompt(
        subject="x", style="hero_editorial", override_aspect="1:1"
    )
    assert asp == "1:1"


def test_build_prompt_unknown_style_raises():
    try:
        build_prompt("test", style="nonexistent_xxx")
    except ValueError as e:
        assert "Unknown style" in str(e)
        return
    raise AssertionError("应该 raise ValueError")


def test_list_styles():
    """list_styles 返回每个风格的 aspect + description。"""
    styles = list_styles()
    assert "hero_editorial" in styles
    assert "aspect" in styles["hero_editorial"]
    assert "description" in styles["hero_editorial"]


# ══════════════════════════════════════════════════════════
# OpenAIImageProvider 测试(用 mock 避免真调 API)
# ══════════════════════════════════════════════════════════
def test_openai_provider_requires_api_key():
    """无 api_key 时应当 raise。"""
    # 清除环境变量
    saved = {}
    for k in ("PPTX_CJK_IMAGE_API_KEY", "OPENAI_API_KEY"):
        saved[k] = os.environ.pop(k, None)
    try:
        try:
            OpenAIImageProvider()
        except RuntimeError as e:
            assert "API key" in str(e)
            return
        raise AssertionError("应该 raise RuntimeError")
    finally:
        for k, v in saved.items():
            if v is not None:
                os.environ[k] = v


def test_openai_provider_init_with_explicit_args():
    p = OpenAIImageProvider(
        api_base="https://example.com/v1",
        api_key="sk-test",
        model="gpt-image-2",
        work_dir=str(TMP / "openai"),
    )
    assert p.api_base == "https://example.com/v1"
    assert p.api_key == "sk-test"
    assert p.model == "gpt-image-2"


def test_openai_provider_proxy_localhost_to_127():
    """sandbox 兼容性:HTTPS_PROXY 含 localhost 时自动改 127.0.0.1。"""
    saved = os.environ.get("HTTPS_PROXY")
    os.environ["HTTPS_PROXY"] = "http://localhost:8080"
    try:
        p = OpenAIImageProvider(api_base="https://x/v1", api_key="sk-test")
        assert p._proxies is not None
        assert "127.0.0.1" in p._proxies["https"]
        assert "localhost" not in p._proxies["https"]
    finally:
        if saved:
            os.environ["HTTPS_PROXY"] = saved
        else:
            os.environ.pop("HTTPS_PROXY", None)


@patch("requests.post")
def test_openai_generate_mocked(mock_post):
    """generate 用 mock 模拟 API 返回。"""
    fake_b64 = (
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
    )
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"data": [{"b64_json": fake_b64}]}
    mock_post.return_value = mock_resp

    p = OpenAIImageProvider(
        api_base="https://example.com/v1",
        api_key="sk-test",
        model="gpt-image-2",
        work_dir=str(TMP / "mock"),
    )
    img_path = p.generate("test prompt", aspect="1:1")
    assert img_path.exists()
    assert img_path.stat().st_size > 0
    img_path.unlink()


@patch("requests.post")
def test_openai_generate_with_aspect_16_9(mock_post):
    """aspect='16:9' 应当请求 1536x1024。"""
    fake_b64 = (
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
    )
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"data": [{"b64_json": fake_b64}]}
    mock_post.return_value = mock_resp

    p = OpenAIImageProvider(
        api_base="https://example.com/v1",
        api_key="sk-test",
        model="gpt-image-2",
        work_dir=str(TMP / "mock"),
    )
    img_path = p.generate("test", aspect="16:9")
    # 检查传给 API 的 size 参数
    call_payload = mock_post.call_args.kwargs.get("json", {})
    assert call_payload.get("size") == "1536x1024"
    img_path.unlink()


@patch("requests.post")
def test_openai_generate_raises_on_error(mock_post):
    """非 200 状态码应 raise。"""
    mock_resp = MagicMock()
    mock_resp.status_code = 401
    mock_resp.text = '{"error": "Invalid API key"}'
    mock_post.return_value = mock_resp

    p = OpenAIImageProvider(
        api_base="https://example.com/v1",
        api_key="sk-bad",
        model="gpt-image-2",
        work_dir=str(TMP / "mock"),
    )
    try:
        p.generate("test", aspect="1:1")
    except RuntimeError as e:
        assert "401" in str(e) or "API 调用失败" in str(e)
        return
    raise AssertionError("应该 raise")


# ══════════════════════════════════════════════════════════
# get_provider 工厂方法测试
# ══════════════════════════════════════════════════════════
def test_get_provider_manual():
    p = get_provider("manual", work_dir=str(TMP / "factory"))
    assert isinstance(p, ManualImageProvider)


def test_get_provider_openai():
    p = get_provider("openai",
                      api_base="https://x/v1", api_key="sk-x",
                      work_dir=str(TMP / "factory"))
    assert isinstance(p, OpenAIImageProvider)


def test_get_provider_unknown_raises():
    try:
        get_provider("alien_mode")
    except ValueError as e:
        assert "Unknown" in str(e)
        return
    raise AssertionError("应该 raise")


if __name__ == "__main__":
    passed = 0
    failed = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"✅ {name}")
                passed += 1
            except Exception as e:
                print(f"❌ {name}: {type(e).__name__}: {e}")
                failed += 1
    print(f"\n{passed} passed, {failed} failed")
