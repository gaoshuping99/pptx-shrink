"""mediacodecs.py — 媒体编码封装（subprocess 调用系统工具，跨平台）。

约定：每个 optimize_* 返回 (new_bytes|None, action_key, reason_key)。
action_key/reason_key 是稳定的英文短语 key，由 report.py 按 --lang 翻译展示；
带参数的用 "key|param" 形式（如 "resize|1280"），report 层解析。
new_bytes 为 None 表示未产出更小结果（调用方保留原件）。
所有工具走临时文件；主脚本负责阈值门与"仅当更小才替换"。

硬依赖：magick(IM7) / pngquant / ffmpeg / ffprobe（缺则主脚本报错退出）。
可选：jpegtran（缺则 magick 兜底）/ pyftsubset（仅字体子集化）。
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile

IS_MAC = platform.system() == "Darwin"

# 光栅图片扩展名
RASTER_EXTS = {"jpg", "jpeg", "png", "gif", "bmp", "tiff", "tif"}
VECTOR_EXTS = {"emf", "wmf"}
VIDEO_EXTS = {"mp4", "mov", "m4v", "avi", "mkv"}
AUDIO_EXTS = {"wav", "mp3", "m4a", "aac"}

# 工具解析缓存：name -> 绝对路径 or None（跨平台，靠 PATH；Windows 自动匹配 .exe）
_TOOL_CACHE: dict[str, str | None] = {}


def _tool(name: str) -> str | None:
    """解析外部工具的绝对路径（shutil.which 跨平台，自动处理 Windows .exe/.cmd）。"""
    if name not in _TOOL_CACHE:
        _TOOL_CACHE[name] = shutil.which(name)
    return _TOOL_CACHE[name]


def has_tool(name: str) -> bool:
    return _tool(name) is not None


def _run(cmd: list[str], timeout: int = 300) -> tuple[int, bytes, bytes]:
    """cmd[0] 是工具名，会先经 _tool 解析成绝对路径。"""
    resolved = _tool(cmd[0]) or cmd[0]
    try:
        p = subprocess.run([resolved, *cmd[1:]], capture_output=True, timeout=timeout)
        return p.returncode, p.stdout, p.stderr
    except (FileNotFoundError, OSError) as e:
        return 127, b"", str(e).encode()


def _tmp(suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    return path


def _write_tmp(data: bytes, suffix: str) -> str:
    path = _tmp(suffix)
    with open(path, "wb") as f:
        f.write(data)
    return path


def _read(path: str) -> bytes | None:
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError:
        return None


# ---------------- 图片：alpha 探测 ----------------
def png_has_alpha(data: bytes) -> bool:
    """PNG 是否含有效（非全不透明）alpha。"""
    from PIL import Image
    import io
    try:
        im = Image.open(io.BytesIO(data))
        if "A" not in im.getbands() and im.mode not in ("RGBA", "LA", "PA"):
            return False
        alpha = im.convert("RGBA").getchannel("A")
        return alpha.getextrema()[0] < 255
    except Exception:
        return True  # 探测失败保守当有 alpha（不转 JPEG）


def image_pixel_size(data: bytes) -> tuple[int, int] | None:
    from PIL import Image
    import io
    try:
        return Image.open(io.BytesIO(data)).size
    except Exception:
        return None


# ---------------- JPEG ----------------
def optimize_jpeg(data: bytes, target_px: tuple[int, int] | None, quality: int = 88):
    src = _write_tmp(data, ".jpg")
    out = _tmp(".jpg")
    try:
        resize = []
        if target_px:
            resize = ["-resize", f"{target_px[0]}x{target_px[1]}>"]
        rc, _, err = _run(["magick", src, *resize, "-strip",
                           "-sampling-factor", "4:2:0", "-interlace", "Plane",
                           "-quality", str(quality), out])
        cand = _read(out) if rc == 0 else None
        act = f"jpeg-recompress|{quality}" + (f"|{target_px[0]}" if target_px else "")
        # 无损兜底：jpegtran（若装了），否则 magick 无损重存
        if cand is None or len(cand) >= len(data):
            jtb = _jpeg_lossless(src, data)
            if jtb is not None:
                return jtb, "jpeg-lossless", ""
            return None, "", "not-smaller"
        return cand, act, ""
    finally:
        for p in (src, out):
            if os.path.exists(p):
                os.unlink(p)


def _jpeg_lossless(src_path: str, orig: bytes) -> bytes | None:
    """JPEG 无损优化：优先 jpegtran，缺失则 magick -strip 兜底。返回更小的 bytes 或 None。"""
    if has_tool("jpegtran"):
        jt = _tmp(".jpg")
        rc, _, _ = _run(["jpegtran", "-copy", "none", "-optimize", "-progressive",
                         "-outfile", jt, src_path])
        b = _read(jt) if rc == 0 else None
        if os.path.exists(jt):
            os.unlink(jt)
        if b and len(b) < len(orig):
            return b
    # magick 兜底
    o = _tmp(".jpg")
    rc, _, _ = _run(["magick", src_path, "-strip", "-interlace", "Plane", o])
    b = _read(o) if rc == 0 else None
    if os.path.exists(o):
        os.unlink(o)
    if b and len(b) < len(orig):
        return b
    return None


# ---------------- PNG ----------------
def optimize_png(data: bytes, target_px: tuple[int, int] | None,
                 allow_to_jpeg: bool = True, quality: int = 88):
    """返回 (bytes, action_key, reason_key, new_ext)；new_ext 指示跨格式（'jpeg' or None）。"""
    has_alpha = png_has_alpha(data)

    # 无有效 alpha 且允许跨格式 → 转 JPEG（收益最大）
    if not has_alpha and allow_to_jpeg:
        src = _write_tmp(data, ".png")
        out = _tmp(".jpg")
        try:
            resize = ["-resize", f"{target_px[0]}x{target_px[1]}>"] if target_px else []
            rc, _, _ = _run(["magick", src, *resize, "-background", "white",
                             "-flatten", "-strip", "-sampling-factor", "4:2:0",
                             "-quality", str(quality), out])
            cand = _read(out) if rc == 0 else None
            if cand and len(cand) < len(data):
                act = f"png-to-jpeg|{quality}" + (f"|{target_px[0]}" if target_px else "")
                return cand, act, "", "jpeg"
        finally:
            for p in (src, out):
                if os.path.exists(p):
                    os.unlink(p)
        # 转 JPEG 没变小则落到同格式分支

    # 有 alpha 或转 JPEG 不划算 → pngquant 量化（同格式）
    src = _write_tmp(data, ".png")
    resized = src
    try:
        if target_px:
            r = _tmp(".png")
            rc, _, _ = _run(["magick", src, "-resize",
                             f"{target_px[0]}x{target_px[1]}>", "-strip", r])
            if rc == 0 and os.path.exists(r):
                resized = r
        out = _tmp(".png")
        # pngquant rc=99 表示达不到质量下限；退而用无下限
        rc, _, _ = _run(["pngquant", "--quality=65-90", "--strip", "--force",
                         "--output", out, resized])
        if rc == 99 or not os.path.exists(out) or os.path.getsize(out) == 0:
            rc, _, _ = _run(["pngquant", "--strip", "--force", "--output", out, resized])
        cand = _read(out) if os.path.exists(out) else None
        if cand and len(cand) < len(data):
            act = "png-quantize" + (f"|{target_px[0]}" if target_px else "")
            return cand, act, "", None
        # 无损兜底：magick strip
        o2 = _tmp(".png")
        rc, _, _ = _run(["magick", resized, "-strip", o2])
        b2 = _read(o2) if rc == 0 else None
        if os.path.exists(o2):
            os.unlink(o2)
        if b2 and len(b2) < len(data):
            return b2, "png-lossless", "", None
        return None, "", "not-smaller", None
    finally:
        for p in {src, resized}:
            if os.path.exists(p):
                os.unlink(p)


# ---------------- 视频/音频 ----------------
def ffprobe_info(path: str) -> dict:
    rc, out, _ = _run(["ffprobe", "-v", "quiet", "-print_format", "json",
                       "-show_streams", "-show_format", path])
    import json
    if rc != 0:
        return {}
    try:
        return json.loads(out)
    except Exception:
        return {}


def optimize_video(data: bytes, ext: str, av_codec: str = "auto",
                   crf: int = 23, max_width: int = 1920, max_height: int = 1080,
                   fps: float | None = None, audio_bitrate: str = "128k"):
    """同容器重编码 mp4/mov→同扩展名。返回 (bytes, action_key, reason_key)。

    av_codec:
      "auto"  — macOS 用 hevc_videotoolbox（硬件、快），其它平台用 libx264。
      "hevc"  — 强制 hevc_videotoolbox（仅 macOS 有意义）。
      "x264"  — 强制 libx264（全平台）。
    crf/max_width/max_height/fps/audio_bitrate — 可选覆盖；默认= 视觉无损档
    （CRF23 / ≤1080p / 保持帧率 / 音轨 128k）。
    """
    # 解析实际使用的编码器
    if av_codec == "auto":
        use_hevc = IS_MAC
    else:
        use_hevc = (av_codec == "hevc")

    src = _write_tmp(data, f".{ext}")
    out = _tmp(f".{ext}")
    try:
        # 缩放：不超过 max_width×max_height，且绝不放大；-2 保持偶数尺寸
        vf = (f"scale='min({int(max_width)},iw)':'min({int(max_height)},ih)'"
              f":force_original_aspect_ratio=decrease,scale=trunc(iw/2)*2:trunc(ih/2)*2")
        fps_cmd = ["-r", str(fps)] if fps else []
        if use_hevc:
            # videotoolbox 用 q:v（无 crf）；把 crf 粗略映射到 q:v（crf 越大质量越低）
            qv = max(1, min(100, 100 - int(crf) * 2))
            vcmd = ["-c:v", "hevc_videotoolbox", "-q:v", str(qv), "-tag:v", "hvc1"]
            act = "video-hevc"
        else:
            vcmd = ["-c:v", "libx264", "-crf", str(int(crf)), "-preset", "medium"]
            act = "video-h264"
        rc, _, err = _run(["ffmpeg", "-y", "-i", src, "-vf", vf, *fps_cmd, *vcmd,
                           "-c:a", "aac", "-b:a", str(audio_bitrate),
                           "-movflags", "+faststart", out],
                          timeout=1800)
        cand = _read(out) if rc == 0 else None
        if cand and len(cand) < len(data):
            return cand, act, ""
        # HEVC 失败或更大 → 回退 x264（仅当刚才试的是 hevc）
        if use_hevc:
            return optimize_video(data, ext, av_codec="x264", crf=crf,
                                  max_width=max_width, max_height=max_height,
                                  fps=fps, audio_bitrate=audio_bitrate)
        return None, "", "not-smaller"
    finally:
        for p in (src, out):
            if os.path.exists(p):
                os.unlink(p)


def optimize_audio(data: bytes, ext: str, bitrate: str = "128k"):
    """同容器：mp3/m4a/aac 重编码到指定码率；wav 不装 aac → 跳过。"""
    if ext == "wav":
        return None, "", "wav-same-container"
    src = _write_tmp(data, f".{ext}")
    out = _tmp(f".{ext}")
    try:
        codec = "aac" if ext in ("m4a", "aac") else "libmp3lame"
        rc, _, _ = _run(["ffmpeg", "-y", "-i", src, "-c:a", codec, "-b:a", str(bitrate), out])
        cand = _read(out) if rc == 0 else None
        if cand and len(cand) < len(data):
            return cand, "audio-recompress", ""
        return None, "", "not-smaller"
    finally:
        for p in (src, out):
            if os.path.exists(p):
                os.unlink(p)


# ---------------- 字体子集化 ----------------
def subset_font(data: bytes, unicodes: str) -> tuple[bytes | None, str, str]:
    """pyftsubset 裁剪（从 PATH 解析）。不处理 OOXML 混淆——由调用方判断 fntdata。"""
    if not has_tool("pyftsubset"):
        return None, "", "pyftsubset-unavailable"
    src = _write_tmp(data, ".ttf")
    out = _tmp(".ttf")
    try:
        rc, _, err = _run(["pyftsubset", src, f"--unicodes={unicodes}",
                           "--desubroutinize", f"--output-file={out}"])
        cand = _read(out) if rc == 0 else None
        if cand and len(cand) < len(data):
            return cand, "font-subset", ""
        return None, "", "not-smaller"
    finally:
        for p in (src, out):
            if os.path.exists(p):
                os.unlink(p)
