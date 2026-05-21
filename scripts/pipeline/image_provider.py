"""B3 Image Provider — 生图抽象层(支持手动/本地素材/OpenAI API)。

设计目标:让 skill 在不同模式下都可用,调用方代码不变。

四种 provider:
- ManualImageProvider: 默认,交互式让用户提供图片路径(适合手动产图后回填)
- LocalAssetProvider: 从本地素材库索引现成图片
- OpenAIImageProvider: ⭐ 调用 OpenAI Image API(或兼容服务)
- MCPImageProvider: Phase 2 占位

⭐ Aspect Ratio 处理(v0.5.0):
所有 provider 的 generate() 接受 aspect 参数(如 "16:9"),内部:
1. 选最接近的 API 支持尺寸调用 API
2. 用 PIL 在生成后精确裁切到目标比例(center-crop 模式)
3. 调用方拿到的就是精确比例,不需要 slide 端拉伸

OpenAI Image API 支持的 size:
- 1024x1024 (1:1) — 方图
- 1024x1536 (2:3) — 竖图
- 1536x1024 (3:2) — 横图,接近 16:9
- auto

切换方式:
    # 16:9 hero 图(自动用 1536x1024 + 裁到精确 16:9)
    img = provider.generate(prompt, aspect="16:9")

    # 1:1 方图
    img = provider.generate(prompt, aspect="1:1")

⚠️ 安全:
- API key 从环境变量读取(PPTX_CJK_IMAGE_API_KEY),不写在代码里
- 用 .env + python-dotenv 管理本地配置
- .env 文件已在 .gitignore 中排除
"""

import base64
import os
from pathlib import Path
from typing import Protocol


# ══════════════════════════════════════════════════════════
# Aspect Ratio 工具
# ══════════════════════════════════════════════════════════
# OpenAI Image API 实际支持的 size(其他全部走最接近的 + PIL 后处理)
SUPPORTED_API_SIZES = {
    "1:1":  "1024x1024",
    "2:3":  "1024x1536",   # 竖图
    "3:2":  "1536x1024",   # 横图(最接近 16:9)
}

# 业务常用比例 → API size 映射 + 是否需要后处理裁剪
ASPECT_TO_API_SIZE = {
    "1:1":  ("1024x1024", None),       # 直接用,无需裁
    "16:9": ("1536x1024", "16:9"),     # 生 3:2,裁到 16:9
    "9:16": ("1024x1536", "9:16"),     # 生 2:3,裁到 9:16
    "4:3":  ("1536x1024", "4:3"),      # 生 3:2,裁到 4:3
    "3:4":  ("1024x1536", "3:4"),      # 生 2:3,裁到 3:4
    "2:3":  ("1024x1536", None),       # 直接
    "3:2":  ("1536x1024", None),       # 直接
    "21:9": ("1536x1024", "21:9"),     # 超宽
}


def parse_aspect(aspect: str) -> float:
    """'16:9' → 1.778"""
    a, b = aspect.split(":")
    return float(a) / float(b)


def crop_to_aspect(image_path: Path, target_aspect: str) -> Path:
    """用 PIL 把图 center-crop 到精确比例,覆盖原文件。"""
    try:
        from PIL import Image
    except ImportError:
        print(f"  ⚠️ Pillow 未安装,跳过 aspect 后处理")
        return image_path

    target_ratio = parse_aspect(target_aspect)
    img = Image.open(image_path)
    iw, ih = img.size
    cur_ratio = iw / ih

    if abs(cur_ratio - target_ratio) < 0.01:
        return image_path  # 已经接近目标比例

    if cur_ratio > target_ratio:
        # 图太宽,裁左右
        new_w = max(1, int(ih * target_ratio))
        x0 = (iw - new_w) // 2
        cropped = img.crop((x0, 0, x0 + new_w, ih))
    else:
        # 图太高,裁上下
        new_h = max(1, int(iw / target_ratio))
        y0 = (ih - new_h) // 2
        cropped = img.crop((0, y0, iw, y0 + new_h))

    # 边界保护:裁后尺寸 < 2px 视为无效,直接保留原图
    if cropped.size[0] < 2 or cropped.size[1] < 2:
        img.close()
        return image_path

    cropped.save(image_path, format="PNG", optimize=True)
    print(f"  ✂️ 已裁切到 {target_aspect}: {iw}×{ih} → {cropped.size[0]}×{cropped.size[1]}")
    img.close()
    return image_path


# ══════════════════════════════════════════════════════════
# 配置加载(从 .env 或环境变量)
# ══════════════════════════════════════════════════════════
def _load_env():
    """尝试加载项目根目录下的 .env 文件。"""
    try:
        from dotenv import load_dotenv
    except ImportError:
        return  # python-dotenv 未安装,跳过

    # 从当前文件向上找 .env(到 skill 根目录)
    cur = Path(__file__).resolve()
    for parent in [cur.parent, *cur.parents]:
        env_file = parent / ".env"
        if env_file.exists():
            load_dotenv(env_file)
            return


_load_env()


# ══════════════════════════════════════════════════════════
# 接口
# ══════════════════════════════════════════════════════════
class ImageProvider(Protocol):
    """图像生成 / 编辑能力的统一接口(v0.5.0 — aspect 优先)。"""

    def generate(self, prompt: str,
                 aspect: str = "16:9",
                 quality: str = "medium",
                 **kwargs) -> Path:
        """文生图,返回本地路径(精确比例)。

        Args:
            prompt: 提示词
            aspect: 业务比例(16:9 / 1:1 / 9:16 / 4:3 / 3:2 / 2:3 / 21:9)
            quality: low / medium / high / auto
        """
        ...

    def edit(self, image: Path, prompt: str,
             mask: Path | None = None, **kwargs) -> Path:
        """图编辑(局部精修,客户改稿场景)。"""
        ...

    def variation(self, image: Path, n: int = 1, **kwargs) -> list[Path]:
        """生成同主题变体。"""
        ...


# ══════════════════════════════════════════════════════════
# 实现 1: 手动(默认)
# ══════════════════════════════════════════════════════════
class ManualImageProvider:
    """交互式 — 打印 prompt,等待用户提供图片路径。

    使用场景:
    - 用户用其他工具(Midjourney / DALL-E web / 自建工具)产图后回填
    - 不希望脚本里硬编码 OpenAI key
    """

    def __init__(self, work_dir: str = "work/images"):
        self.work_dir = Path(work_dir)
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self._counter = 0

    def _placeholder(self, prompt: str, kind: str) -> Path:
        self._counter += 1
        stub = self.work_dir / f"{kind}_{self._counter:03d}.prompt.txt"
        stub.write_text(
            f"[ManualImageProvider {kind}]\nprompt: {prompt}\n",
            encoding="utf-8",
        )
        return self.work_dir / f"{kind}_{self._counter:03d}.png"

    def generate(self, prompt: str,
                 aspect: str = "16:9",
                 quality: str = "medium",
                 size: str | None = None,
                 **kwargs) -> Path:
        actual_size = size or ASPECT_TO_API_SIZE.get(aspect, ("1536x1024", None))[0]
        print(f"\n📸 [Manual] 需要生成一张图")
        print(f"   prompt: {prompt[:80]}{'...' if len(prompt) > 80 else ''}")
        print(f"   aspect: {aspect} (≈ {actual_size}), quality: {quality}")
        print(f"   请用其他工具产图后,放到 {self.work_dir}/ 下")
        return self._placeholder(prompt, "gen")

    def edit(self, image: Path, prompt: str,
             mask: Path | None = None, **kwargs) -> Path:
        print(f"\n✏️ [Manual] 需要编辑图片")
        print(f"   原图: {image}")
        print(f"   修改: {prompt}")
        if mask:
            print(f"   蒙版: {mask}")
        return self._placeholder(prompt, "edit")

    def variation(self, image: Path, n: int = 1, **kwargs) -> list[Path]:
        print(f"\n🎲 [Manual] 需要 {n} 个变体")
        print(f"   原图: {image}")
        return [self._placeholder("variation", "var") for _ in range(n)]


# ══════════════════════════════════════════════════════════
# 实现 2: 本地素材库
# ══════════════════════════════════════════════════════════
class LocalAssetProvider:
    """从本地素材库匹配图片。"""

    def __init__(self, assets_dir: str | None = None):
        if assets_dir is None:
            assets_dir = os.getenv("PPTX_CJK_ASSET_DIR", "assets")
        self.assets_dir = Path(assets_dir).expanduser()
        if not self.assets_dir.exists():
            raise FileNotFoundError(f"Assets dir not found: {self.assets_dir}")

    def _search(self, prompt: str) -> Path | None:
        keywords = [w.lower() for w in prompt.split() if len(w) >= 3]
        candidates = []
        for f in self.assets_dir.rglob("*.png"):
            name = f.name.lower()
            score = sum(1 for kw in keywords if kw in name)
            if score > 0:
                candidates.append((score, f))
        candidates.sort(reverse=True)
        return candidates[0][1] if candidates else None

    def generate(self, prompt: str,
                 aspect: str = "16:9",
                 quality: str = "medium",
                 size: str | None = None,
                 **kwargs) -> Path:
        match = self._search(prompt)
        if match:
            print(f"📦 [Local] 命中素材: {match.name}")
            # 若 aspect 不匹配,后处理裁切
            try:
                from PIL import Image
                with Image.open(match) as im:
                    iw, ih = im.size
                target_ratio = parse_aspect(aspect)
                if abs(iw / ih - target_ratio) > 0.02:
                    # 拷贝到 work/images/ 然后裁切,避免污染原素材
                    import shutil
                    out = self.assets_dir.parent / "_cropped" / match.name
                    out.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy(match, out)
                    crop_to_aspect(out, aspect)
                    return out
            except ImportError:
                pass
            return match
        print(f"⚠️ [Local] 未在 {self.assets_dir} 找到匹配 '{prompt}'")
        return self.assets_dir / f"PLACEHOLDER_{prompt[:30]}.png"

    def edit(self, image: Path, prompt: str,
             mask: Path | None = None, **kwargs) -> Path:
        print(f"⚠️ [Local] 不支持图编辑,返回原图")
        return image

    def variation(self, image: Path, n: int = 1, **kwargs) -> list[Path]:
        return [image] * n


# ══════════════════════════════════════════════════════════
# 实现 3: OpenAI Image API(⭐ 新增)
# ══════════════════════════════════════════════════════════
class OpenAIImageProvider:
    """调用 OpenAI Image API 或兼容服务(如国内 relay)生成配图。

    实现说明:
    - 直接用 requests 调 REST API,不走 openai SDK
    - 原因:openai SDK 的 httpx 在某些代理环境下有兼容问题
    - requests 自动尊重 HTTPS_PROXY / HTTP_PROXY 环境变量

    ⚠️ API Key 从环境变量读取,不写在代码里。
    支持的环境变量(优先级降序):
    1. 构造参数 api_key
    2. PPTX_CJK_IMAGE_API_KEY
    3. OPENAI_API_KEY

    Args:
        api_base: API 根地址,需含 /v1。例如:
                  - 官方: https://api.openai.com/v1
                  - 国内 relay: https://your-relay.com/v1
        api_key: API key
        model: 模型名,如 gpt-image-1, gpt-image-2, dall-e-3
        work_dir: 下载图片的本地目录
        timeout: 请求超时(秒)
    """

    def __init__(self,
                 api_base: str | None = None,
                 api_key: str | None = None,
                 model: str | None = None,
                 work_dir: str = "work/images",
                 timeout: int = 180):
        try:
            import requests  # noqa: F401
        except ImportError:
            raise RuntimeError(
                "OpenAIImageProvider 需要 requests 包,运行:\n"
                "  pip install requests"
            )

        self.api_base = (
            api_base
            or os.getenv("PPTX_CJK_IMAGE_API_BASE")
            or os.getenv("OPENAI_BASE_URL")
            or "https://api.openai.com/v1"
        ).rstrip("/")
        self.api_key = (
            api_key
            or os.getenv("PPTX_CJK_IMAGE_API_KEY")
            or os.getenv("OPENAI_API_KEY")
        )
        if not self.api_key:
            raise RuntimeError(
                "未找到 API key。请通过以下方式之一配置:\n"
                "  1. 环境变量 PPTX_CJK_IMAGE_API_KEY=sk-...\n"
                "  2. .env 文件(参考 .env.example)\n"
                "  3. OpenAIImageProvider(api_key='sk-...')"
            )
        self.model = (
            model
            or os.getenv("PPTX_CJK_IMAGE_MODEL")
            or "gpt-image-1"
        )

        self.timeout = timeout
        self.work_dir = Path(work_dir)
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self._counter = 0

        # 探测系统代理,sandbox 兼容性处理
        # 一些受限环境的 HTTPS_PROXY 含 'localhost',Python socket 可能解析失败,
        # 改用 127.0.0.1 替代
        self._proxies = self._detect_proxies()

    def _detect_proxies(self) -> dict | None:
        """从环境变量探测代理,localhost → 127.0.0.1。"""
        https = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
        http = os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy")
        if not https and not http:
            return None
        proxies = {}
        if https:
            proxies["https"] = https.replace("//localhost", "//127.0.0.1")
        if http:
            proxies["http"] = http.replace("//localhost", "//127.0.0.1")
        return proxies

    # ── HTTP 请求(用 requests,自动走系统代理) ───────────
    def _post_json(self, endpoint: str, payload: dict) -> dict:
        import requests
        url = f"{self.api_base}{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        resp = requests.post(url, headers=headers, json=payload,
                              timeout=self.timeout, proxies=self._proxies)
        if resp.status_code != 200:
            raise RuntimeError(
                f"API 调用失败 ({resp.status_code}): "
                f"{resp.text[:500]}"
            )
        return resp.json()

    def _post_multipart(self, endpoint: str, files: dict, data: dict) -> dict:
        import requests
        url = f"{self.api_base}{endpoint}"
        headers = {"Authorization": f"Bearer {self.api_key}"}
        resp = requests.post(url, headers=headers, files=files, data=data,
                              timeout=self.timeout, proxies=self._proxies)
        if resp.status_code != 200:
            raise RuntimeError(
                f"API 调用失败 ({resp.status_code}): "
                f"{resp.text[:500]}"
            )
        return resp.json()

    # ── 内部:把 API 返回的图片数据写入本地文件 ────────────
    def _save_from_response(self, image_obj: dict, kind: str = "gen") -> Path:
        """处理返回的单个图片对象(b64_json 或 url),写入本地。"""
        self._counter += 1
        out_path = self.work_dir / f"{kind}_{self._counter:03d}.png"

        # 优先 b64_json
        if image_obj.get("b64_json"):
            data = base64.b64decode(image_obj["b64_json"])
            out_path.write_bytes(data)
            return out_path

        # URL 模式
        if image_obj.get("url"):
            import requests
            r = requests.get(image_obj["url"], timeout=self.timeout,
                              proxies=self._proxies)
            r.raise_for_status()
            out_path.write_bytes(r.content)
            return out_path

        raise RuntimeError(f"API 返回既无 b64_json 也无 url: {image_obj}")

    # ── 接口实现 ──────────────────────────────────────────
    def generate(self, prompt: str,
                 aspect: str = "16:9",
                 size: str | None = None,
                 quality: str = "medium", n: int = 1, **kwargs) -> Path:
        """文生图。

        Args:
            prompt: 提示词
            aspect: 业务比例,如 "16:9" / "1:1" / "9:16" / "4:3"
                    自动映射到 API 支持的 size + PIL 后处理裁切
            size: 直接指定 API size(覆盖 aspect),如 "1536x1024"
            quality: low / medium / high / auto
            n: 生成数量(返回首张)
        """
        # 选 API size + 是否需要裁切
        if size:
            api_size = size
            crop_target = None
        elif aspect in ASPECT_TO_API_SIZE:
            api_size, crop_target = ASPECT_TO_API_SIZE[aspect]
        else:
            api_size, crop_target = "1536x1024", "16:9"

        print(f"\n🎨 [OpenAI/{self.model}] 生成图片")
        print(f"   prompt: {prompt[:80]}{'...' if len(prompt) > 80 else ''}")
        print(f"   aspect: {aspect} → API size {api_size}"
              f"{' + crop ' + crop_target if crop_target else ' (直出)'}")
        print(f"   quality: {quality}")

        payload = {
            "model": self.model,
            "prompt": prompt,
            "size": api_size,
            "n": n,
        }
        # gpt-image-1/2 默认就是 b64_json,不接受 response_format 参数
        # 其他模型(qwen / dalle-3 等)默认返 URL,需要显式要 b64_json
        if not self.model.startswith("gpt-image"):
            payload["response_format"] = "b64_json"
        if quality in ("low", "medium", "high", "auto"):
            payload["quality"] = quality
        for k, v in kwargs.items():
            payload[k] = v

        try:
            result = self._post_json("/images/generations", payload)
        except Exception as e:
            print(f"   ❌ 调用失败: {type(e).__name__}: {e}")
            raise

        if not result.get("data"):
            raise RuntimeError(f"API 返回无 data 字段: {result}")

        path = self._save_from_response(result["data"][0], kind="gen")

        # PIL 精确裁切到目标比例(避免 slide 端拉伸失真)
        if crop_target:
            crop_to_aspect(path, crop_target)

        print(f"   ✅ 已保存: {path}")
        return path

    def edit(self, image: Path, prompt: str,
             mask: Path | None = None, size: str = "1024x1024",
             **kwargs) -> Path:
        """图编辑(局部精修)。"""
        print(f"\n✏️ [OpenAI/{self.model}] 编辑图片")
        print(f"   原图: {image}")
        print(f"   修改: {prompt[:80]}")

        files = {"image": open(image, "rb")}
        data = {
            "model": self.model,
            "prompt": prompt,
            "size": size,
        }
        if mask:
            files["mask"] = open(mask, "rb")
            print(f"   蒙版: {mask}")

        try:
            result = self._post_multipart("/images/edits", files, data)
        except Exception as e:
            print(f"   ❌ 调用失败: {type(e).__name__}: {e}")
            raise
        finally:
            for f in files.values():
                if hasattr(f, "close"):
                    f.close()

        path = self._save_from_response(result["data"][0], kind="edit")
        print(f"   ✅ 已保存: {path}")
        return path

    def variation(self, image: Path, n: int = 1,
                   size: str = "1024x1024", **kwargs) -> list[Path]:
        """同主题变体。"""
        print(f"\n🎲 [OpenAI/{self.model}] 生成 {n} 个变体")

        files = {"image": open(image, "rb")}
        data = {
            "model": self.model,
            "n": str(n),
            "size": size,
        }
        try:
            result = self._post_multipart("/images/variations", files, data)
        finally:
            files["image"].close()

        paths = []
        for img_obj in result["data"]:
            paths.append(self._save_from_response(img_obj, kind="var"))
        return paths


# ══════════════════════════════════════════════════════════
# 实现 4: MCP(Phase 2 占位)
# ══════════════════════════════════════════════════════════
class MCPImageProvider:
    """通过 mcp-image-gen server 调用。Phase 2 实现。"""

    def __init__(self):
        raise NotImplementedError(
            "MCPImageProvider 是 Phase 2 实现。\n"
            "目前请用 get_provider('openai')。"
        )

    def generate(self, prompt: str, **kw) -> Path:
        raise NotImplementedError

    def edit(self, image: Path, prompt: str, mask: Path | None = None,
             **kw) -> Path:
        raise NotImplementedError

    def variation(self, image: Path, n: int = 1, **kw) -> list[Path]:
        raise NotImplementedError


# ══════════════════════════════════════════════════════════
# 工厂方法
# ══════════════════════════════════════════════════════════
def get_provider(mode: str = "manual", **kwargs) -> ImageProvider:
    """返回指定模式的 provider。

    Args:
        mode: "manual" / "local" / "openai" / "mcp"
        kwargs: 传给具体 provider 的参数

    Examples:
        # 默认手动
        provider = get_provider()

        # OpenAI(从 .env / 环境变量自动读)
        provider = get_provider("openai")

        # OpenAI + 显式参数
        provider = get_provider("openai",
                                api_base="https://...",
                                api_key="sk-...",
                                model="gpt-image-2")

        # 本地素材库
        provider = get_provider("local", assets_dir="~/assets/")
    """
    if mode == "manual":
        return ManualImageProvider(**kwargs)
    if mode == "local":
        return LocalAssetProvider(**kwargs)
    if mode == "openai":
        return OpenAIImageProvider(**kwargs)
    if mode == "mcp":
        return MCPImageProvider(**kwargs)
    raise ValueError(f"Unknown provider mode: {mode}")
