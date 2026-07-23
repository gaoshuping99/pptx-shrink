#!/usr/bin/env python3
"""compress_pptx.py — PPT 压缩编排入口（跨平台：macOS / Windows / Linux）。

用法:
  python compress_pptx.py <input.pptx> [--out DIR]
      [--no-repack-formats]   # 默认开: PNG→JPEG 跨格式(会成对改 Content_Types+rels)
      [--subset-fonts]        # 默认关: 字体子集化(破坏性)
      [--apply-crop]          # 默认关: 裁剪像素丢弃(改 slide XML)
      [--av-codec auto|hevc|x264] [--jpeg-quality 88] [--retina 2.0]
      [--min-save-kb 64] [--lang en|zh]
  python compress_pptx.py <input.pptx> --analyze-only   # 只出报告, 不写新 pptx

核心不动 XML；唯一例外是 PNG→JPEG 时成对改 [Content_Types].xml 与对应 rels 的扩展名。
默认输出到输入文件同目录（可 --out 覆盖）。
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import os
import platform
import re
import shutil
import sys
import zipfile

# Windows 控制台默认 GBK，无法编码 emoji/中文，会让"成功的压缩"因 print 崩溃而退出失败。
# 统一把 stdout/stderr 切到 UTF-8 且对无法编码的字符降级替换（不崩溃）。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # Py3.7+
    except (AttributeError, ValueError):
        pass


def _safe_print(text: str) -> None:
    """在任何终端编码下安全打印（兜底：按当前编码替换不可编码字符）。"""
    try:
        print(text)
    except UnicodeEncodeError:
        enc = (getattr(sys.stdout, "encoding", None) or "utf-8")
        sys.stdout.write(text.encode(enc, "replace").decode(enc, "replace") + "\n")


sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import media as M
import mediacodecs as C
import cleaners as CL
import report as R

MARKER = "docProps/pptxshrink.json"
PRECOMPRESSED_EXT = {"jpg", "jpeg", "png", "gif", "mp4", "mov", "m4v", "m4a",
                     "mp3", "aac", "heic", "webp"}

# 硬依赖：三平台都有现成包（键=命令名，值=各平台包名）
REQUIRED_TOOLS = {
    "ffmpeg":   {"brew": "ffmpeg",      "choco": "ffmpeg",      "apt": "ffmpeg"},
    "ffprobe":  {"brew": "ffmpeg",      "choco": "ffmpeg",      "apt": "ffmpeg"},
    "magick":   {"brew": "imagemagick", "choco": "imagemagick", "apt": "imagemagick"},
    "pngquant": {"brew": "pngquant",    "choco": "pngquant",    "apt": "pngquant"},
}


def check_deps() -> None:
    """启动时校验硬依赖；缺失则打印各平台安装命令并退出。"""
    missing = [t for t in REQUIRED_TOOLS if shutil.which(t) is None]
    pil_missing = False
    try:
        import PIL  # noqa: F401
    except ImportError:
        pil_missing = True

    if not missing and not pil_missing:
        return

    sysname = platform.system()
    lines = ["Error: missing required dependencies; pptx-shrink cannot run.",
             "错误：缺少必需依赖工具，无法运行。", ""]
    if missing:
        lines.append(f"Missing CLI tools / 缺少工具: {', '.join(missing)}")
        if sysname == "Windows":
            # choco 与 winget 包名不同，分两行各自可直接复制执行
            choco_pkgs = sorted({REQUIRED_TOOLS[t]["choco"] for t in missing})
            winget_map = {"ffmpeg": "Gyan.FFmpeg", "imagemagick": "ImageMagick.ImageMagick",
                          "pngquant": "pngquant.pngquant"}
            winget_pkgs = sorted({winget_map.get(REQUIRED_TOOLS[t]["choco"],
                                                  REQUIRED_TOOLS[t]["choco"]) for t in missing})
            lines.append(f"  choco:  choco install {' '.join(choco_pkgs)}")
            lines.append(f"  winget: winget install {' '.join(winget_pkgs)}")
        elif sysname == "Darwin":
            pkgs = sorted({REQUIRED_TOOLS[t]["brew"] for t in missing})
            lines.append(f"  brew install {' '.join(pkgs)}")
        else:
            pkgs = sorted({REQUIRED_TOOLS[t]["apt"] for t in missing})
            lines.append(f"  sudo apt install {' '.join(pkgs)}")
    if pil_missing:
        lines.append("Missing Python lib Pillow / 缺少 Pillow")
        lines.append("  pip install Pillow")
    lines += ["",
              "Optional (missing only disables that feature) / 可选增强：",
              "  jpegtran  (JPEG lossless; magick used as fallback)",
              "  fonttools/pyftsubset  (--subset-fonts)",
              "  LibreOffice  (extra render validation)"]
    print("\n".join(lines), file=sys.stderr)
    sys.exit(1)


def _now_iso() -> str:
    return datetime.datetime.now().isoformat(timespec="seconds")


def _is_encrypted(path: str) -> bool:
    """OLE 复合文档头 (加密 OOXML) → D0CF11E0。"""
    try:
        with open(path, "rb") as f:
            return f.read(8) == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    except OSError:
        return False



def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _find_soffice() -> str | None:
    # 先查 PATH（跨平台，含 Windows soffice.exe）
    for name in ("soffice", "soffice.exe", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    # 再查各平台常见安装位置
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",   # macOS
        "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",       # Windows
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice", "/usr/bin/libreoffice",                # Linux
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def _structural_issues(path: str) -> set[str]:
    """返回一个"结构问题"集合（悬空引用/孤儿部件/未声明扩展名）。
    用集合是为了和原文件做差集——只有压缩产物"新增"的问题才算真问题，
    原文件本就存在的（如 PowerPoint 残留的 [trash]/*.dat、Target=NULL 外链）一律放行。
    """
    import posixpath
    issues: set[str] = set()
    with zipfile.ZipFile(path, "r") as z:
        names = set(z.namelist())
        referenced: set[str] = set()
        for n in names:
            if not n.endswith(".rels"):
                continue
            data = z.read(n)
            host_dir = posixpath.dirname(posixpath.dirname(n))
            for m in re.finditer(rb'<Relationship\b[^>]*/?>', data):
                tag = m.group(0)
                if b'TargetMode="External"' in tag:
                    continue
                tm = re.search(rb'Target="([^"]+)"', tag)
                if not tm:
                    continue
                tgt = tm.group(1).decode("utf-8", "replace")
                if tgt.startswith("http") or tgt.startswith("mailto:") or tgt == "NULL":
                    continue
                rp = posixpath.normpath(posixpath.join(host_dir, tgt))
                referenced.add(rp)
                if rp not in names:
                    issues.add(f"dangling:{rp}")
        # Content_Types 覆盖
        ct = z.read("[Content_Types].xml")
        declared_ext = {e.decode().lower() for e in
                        re.findall(rb'<Default\s+Extension="([^"]+)"', ct)}
        overridden = {p.decode() for p in
                      re.findall(rb'<Override\s+PartName="/([^"]+)"', ct)}
        for n in names:
            if n.endswith("/") or n == "[Content_Types].xml":
                continue
            if n not in overridden:
                ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
                if ext and ext not in declared_ext:
                    issues.add(f"undeclared-ext:{ext}")
        # 孤儿部件
        EXEMPT = {"[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"}
        for n in names:
            if n.endswith("/") or n.endswith(".rels") or n in EXEMPT:
                continue
            if n not in referenced:
                issues.add(f"orphan:{n}")
        # 断裂的关系引用：XML 里每个 r:id/r:embed/r:link 都必须在该部件的 .rels 有对应 Relationship
        # （PowerPoint 严格检查此项；超链接关系被误删会导致 hlinkClick 引用悬空、文件损坏）
        for n in names:
            if not (n.endswith(".xml") and n.startswith("ppt/")):
                continue
            rp = f"{posixpath.dirname(n)}/_rels/{posixpath.basename(n)}.rels"
            rels_data = z.read(rp) if rp in names else b""
            rel_ids = set(re.findall(rb'<Relationship\b[^>]*\bId="([^"]+)"', rels_data))
            used = set(re.findall(rb'r:(?:id|embed|link)="([^"]+)"', z.read(n)))
            for rid in used - rel_ids:
                issues.add(f"broken-relref:{n}:{rid.decode()}")
    return issues


def validate_pptx(path: str, baseline_issues: set[str] | None = None) -> tuple[bool, str]:
    """校验产物能否正常打开（尽量复刻 PowerPoint 的严格判定）：
    1) zip 完整性 + 必需部件；
    2) 结构问题（悬空引用/孤儿部件/未声明扩展名）——只拦"新增"的，
       对比原文件基线放行其本就存在的（[trash]/*.dat、Target=NULL 等 PowerPoint 残留）；
    3) 若有 LibreOffice，再加一道渲染校验。
    """
    try:
        with zipfile.ZipFile(path, "r") as z:
            bad = z.testzip()
            if bad is not None:
                return False, f"corrupt zip: {bad}"
            names = set(z.namelist())
            if "[Content_Types].xml" not in names or "ppt/presentation.xml" not in names:
                return False, "missing core part (Content_Types/presentation.xml)"
        issues = _structural_issues(path)
        new_issues = issues - (baseline_issues or set())
        if new_issues:
            sample = "; ".join(sorted(new_issues)[:5])
            return False, f"新增结构问题（PowerPoint 可能判损坏）：{sample}"
    except zipfile.BadZipFile:
        return False, "not a valid zip/pptx"
    except Exception as e:
        return False, f"结构自检异常：{e}"

    # ---- 二级：LibreOffice 渲染（可选，更强）----
    soffice = _find_soffice()
    if soffice:
        import subprocess, tempfile
        with tempfile.TemporaryDirectory() as td:
            try:
                p = subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                                    "--outdir", td, path],
                                   capture_output=True, timeout=180)
                pdfs = [f for f in os.listdir(td) if f.endswith(".pdf")]
                if p.returncode == 0 and pdfs and os.path.getsize(os.path.join(td, pdfs[0])) > 0:
                    return True, "structure OK + LibreOffice render OK"
                return False, "LibreOffice failed to render (possibly corrupt)"
            except subprocess.TimeoutExpired:
                return True, "structure OK (LibreOffice render timed out, skipped)"
            except Exception:
                return True, "structure OK (LibreOffice check errored, skipped)"
    return True, "structure OK (LibreOffice not found, render check skipped)"


def _rewrite_content_types(ct: bytes) -> bytes:
    """确保 [Content_Types].xml 声明了 jpeg 的 Default（跨格式转换统一用 .jpeg 扩展名）。
    注意：必须精确检查 jpeg（不能用 jpe?g，否则文件只声明了 jpg 时会误判为已声明）。"""
    if re.search(rb'<Default\s+Extension="jpeg"', ct, re.I):
        return ct
    ins = b'<Default Extension="jpeg" ContentType="image/jpeg"/>'
    m = re.search(rb'<Default\b', ct)
    if m:
        return ct[:m.start()] + ins + ct[m.start():]
    m = re.search(rb'(<Types\b[^>]*>)', ct)
    if m:
        return ct[:m.end()] + ins + ct[m.end():]
    return ct


def _rewrite_rels_target(rels: bytes, old_base: str, new_base: str) -> bytes:
    """把 rels 里指向 old_base(如 image2.png) 的 Target 改成 new_base(image2.jpeg)。"""
    return rels.replace(old_base.encode(), new_base.encode())


def compress(input_pptx: str, out_dir: str, *, analyze_only: bool,
             repack_formats: bool, subset_fonts: bool, apply_crop: bool,
             drop_unused_layouts: bool, strip_fast_save: bool,
             av_codec: str, jpeg_quality: int, retina: float, min_save_kb: int,
             still_large_mb: float = 10.0, lang: str = "en") -> dict:
    if _is_encrypted(input_pptx):
        raise SystemExit("Error: encrypted PPTX is not supported. "
                         "Please remove the password in PowerPoint first. "
                         "/ 加密的 PPTX 暂不支持，请先去除密码。")
    try:
        infos, raw = M.read_pptx(input_pptx)
    except zipfile.BadZipFile:
        raise SystemExit(f"Error: {input_pptx} is not a valid pptx file.")

    index = M.build_index(raw, retina=retina)

    # 已处理标记（幂等）——存为输出目录旁的 sidecar，绝不塞进 pptx 包内，
    # 否则会成为"无 rels 引用的孤儿部件"，PowerPoint 判为损坏。
    prev_marker = {}
    sidecar = os.path.join(out_dir, f".{os.path.splitext(os.path.basename(input_pptx))[0]}.pptxshrink.json")
    if os.path.exists(sidecar):
        import json
        try:
            with open(sidecar, encoding="utf-8") as f:
                prev_marker = json.load(f).get("media_sha", {})
        except Exception:
            prev_marker = {}

    min_save = min_save_kb * 1024
    results: list[dict] = []
    optimized: dict[str, bytes] = {}      # name -> new bytes
    renames: dict[str, str] = {}          # old media name -> new media name (跨格式)
    xml_touched: list[str] = []
    warnings: list[str] = []

    # ---- analyze-only：仅出报告 ----
    if analyze_only:
        for name, data in raw.items():
            if not name.startswith(M.MEDIA_PREFIX) or name.endswith("/"):
                continue
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            kind = _kind_of(ext)
            results.append({"name": name, "kind": kind, "orig": len(data),
                            "new": len(data), "accepted": False, "action": "",
                            "reason": "analyze-only"})
        rep = R.build_report(input_pptx, None, os.path.getsize(input_pptx), None,
                             index, results, [], warnings, _now_iso())
        text = R.render_text(rep, index, analyze_only=True, lang=lang)
        stem = os.path.splitext(os.path.basename(input_pptx))[0]
        R.write_reports(out_dir, stem, rep, text)
        _safe_print(text)
        return rep

    # ---- 逐媒体压缩 ----
    for name, data in list(raw.items()):
        if not name.startswith(M.MEDIA_PREFIX) or name.endswith("/"):
            continue
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        kind = _kind_of(ext)
        entry = index.get(name)
        target = entry.max_display_px if entry else None

        # 幂等：sha 命中上次标记 → 跳过
        if prev_marker.get(name) == _sha(data):
            results.append({"name": name, "kind": kind, "orig": len(data),
                            "new": len(data), "accepted": False, "action": "",
                            "reason": "already-compressed"})
            continue

        new_bytes = None; action = ""; reason = ""; new_ext = None

        if ext in ("jpg", "jpeg"):
            new_bytes, action, reason = C.optimize_jpeg(data, target, jpeg_quality)
        elif ext == "png":
            new_bytes, action, reason, new_ext = C.optimize_png(
                data, target, allow_to_jpeg=repack_formats, quality=jpeg_quality)
        elif ext in ("tiff", "tif", "bmp"):
            # 当作 PNG 分支处理（转 JPEG 或量化）
            new_bytes, action, reason, new_ext = C.optimize_png(
                data, target, allow_to_jpeg=repack_formats, quality=jpeg_quality)
        elif ext == "gif":
            reason = "gif-kept"
        elif ext in C.VECTOR_EXTS:
            reason = "vector-kept"
        elif ext in C.VIDEO_EXTS and ext in ("mp4", "mov", "m4v"):
            new_bytes, action, reason = C.optimize_video(data, ext, av_codec)
        elif ext in C.AUDIO_EXTS:
            new_bytes, action, reason = C.optimize_audio(data, ext)
        else:
            reason = f"unknown-type|{ext}"

        # 阈值门：仅当更小且省得够多才接受
        if new_bytes is not None and len(new_bytes) < len(data):
            saved = len(data) - len(new_bytes)
            if saved >= min_save or len(new_bytes) / len(data) <= 0.90:
                optimized[name] = new_bytes
                # 跨格式改名
                if new_ext and not name.lower().endswith("." + new_ext):
                    base_noext = name.rsplit(".", 1)[0]
                    renames[name] = f"{base_noext}.{new_ext}"
                results.append({"name": name, "kind": kind, "orig": len(data),
                                "new": len(new_bytes), "accepted": True,
                                "action": action, "reason": ""})
                continue
            else:
                reason = f"below-threshold|{saved//1024}"
        results.append({"name": name, "kind": kind, "orig": len(data),
                        "new": len(data), "accepted": False, "action": "",
                        "reason": reason or "not-smaller"})

    # ---- 字体子集化（可选） ----
    if subset_fonts:
        _do_font_subset(raw, optimized, results, warnings)

    # ---- iSlide 式冗余清理（可选，破坏性，改 XML）----
    xml_overrides: dict[str, bytes] = {}
    drops: set[str] = set()
    if apply_crop:
        CL.apply_crop(raw, index, optimized, xml_overrides, results, warnings)
    if drop_unused_layouts:
        CL.drop_unused_layouts(raw, drops, xml_overrides, results, warnings)
    # 注：曾有 drop_hidden（删隐藏/画布外对象），因会误删 think-cell 等插件的
    # hidden 数据对象导致 PowerPoint 报损坏，已移除（体积收益极小、风险极高）。
    if strip_fast_save:
        CL.strip_fast_save(raw, drops, xml_overrides, results, warnings)
    for n in xml_overrides:
        if n not in xml_touched:
            xml_touched.append(n)

    # ---- 跨格式改名：成对改 Content_Types + rels ----
    if renames:
        ct = xml_overrides.get("[Content_Types].xml", raw.get("[Content_Types].xml", b""))
        ct2 = _rewrite_content_types(ct)
        if ct2 != ct:
            xml_overrides["[Content_Types].xml"] = ct2
            if "[Content_Types].xml" not in xml_touched:
                xml_touched.append("[Content_Types].xml")
        # 改所有引用了被改名 media 的 rels（media 反查里 host 的 rels）
        for old, new in renames.items():
            old_base = old.split("/")[-1]
            new_base = new.split("/")[-1]
            for relname, reldata in raw.items():
                if relname.endswith(".rels") and old_base.encode() in reldata:
                    cur = xml_overrides.get(relname, reldata)
                    upd = _rewrite_rels_target(cur, old_base, new_base)
                    if upd != cur:
                        xml_overrides[relname] = upd
                        if relname not in xml_touched:
                            xml_touched.append(relname)

    # ---- 重打包（原子写：先写临时文件，校验通过后再提交为正式产物）----
    stem = os.path.splitext(os.path.basename(input_pptx))[0]
    out_pptx = os.path.join(out_dir, f"{stem}.compressed.pptx")
    tmp_pptx = os.path.join(out_dir, f".{stem}.compressed.tmp.pptx")  # 临时备份产物

    import json
    final_media_sha = {}

    with zipfile.ZipFile(tmp_pptx, "w") as zo:
        written = set()
        for i in infos:
            n = i.filename
            if n == MARKER:
                continue  # 旧版本可能把标记塞进了包内，一律剔除
            if n in drops:
                continue  # 冗余清理删除的 part 不写回
            # 决定最终名与内容
            final_name = renames.get(n, n)
            payload = optimized.get(n, xml_overrides.get(n, raw[n]))
            if n in xml_overrides:
                payload = xml_overrides[n]
            if n in optimized:
                payload = optimized[n]
            ext = final_name.rsplit(".", 1)[-1].lower() if "." in final_name else ""
            comp = (zipfile.ZIP_STORED if ext in PRECOMPRESSED_EXT
                    else zipfile.ZIP_DEFLATED)
            zi = zipfile.ZipInfo(final_name, date_time=i.date_time)
            zi.external_attr = i.external_attr
            zi.compress_type = comp
            zo.writestr(zi, payload)
            written.add(final_name)
            if final_name.startswith(M.MEDIA_PREFIX):
                final_media_sha[final_name] = _sha(payload)
        # 注意：幂等标记不再写进包内（会成为无引用孤儿部件导致 PowerPoint 报损坏）。
        # 改在提交后写到输出目录旁的 sidecar 文件。

    # ---- 校验临时产物能否正常打开（对比原文件基线，只拦新增问题）----
    try:
        baseline = _structural_issues(input_pptx)
    except Exception:
        baseline = set()
    ok, detail = validate_pptx(tmp_pptx, baseline_issues=baseline)
    if not ok:
        # 回滚：丢弃临时产物，原文件从未被动过
        try:
            os.remove(tmp_pptx)
        except OSError:
            pass
        raise SystemExit(
            f"Error: post-compression validation failed ({detail}). "
            f"Discarded the temp output; the original file was NOT modified.\n"
            f"Try: --av-codec x264, or --no-clean.\n"
            f"（压缩后校验未通过，已丢弃临时产物，原文件未改动。）")

    # 校验通过：提交（临时产物 → 正式产物，原子 rename）
    if os.path.exists(out_pptx):
        os.remove(out_pptx)
    os.replace(tmp_pptx, out_pptx)
    warnings.append(f"post-check|{detail}")
    # 幂等命中提示：若有 media 因 sha 命中上次标记而跳过，明确告知（消除"第二次仍报同样节省"的困惑）
    _skipped = sum(1 for r in results if r.get("reason") == "already-compressed")
    if _skipped:
        warnings.append(f"idempotent-skip|{_skipped}")

    # 写幂等标记 sidecar（在包外，不影响 pptx 结构）
    try:
        with open(sidecar, "w", encoding="utf-8") as f:
            json.dump({"tool": "pptx-shrink", "version": "1.0.1",
                       "media_sha": final_media_sha}, f, ensure_ascii=False)
    except OSError:
        pass

    out_bytes = os.path.getsize(out_pptx)
    rep = R.build_report(input_pptx, out_pptx, os.path.getsize(input_pptx),
                         out_bytes, index, results, xml_touched, warnings, _now_iso())

    # ---- 压缩后仍很大 → 自动触发残余大头详细分析 ----
    rep["still_large"] = None
    if out_bytes > still_large_mb * 1024 * 1024:
        rep["still_large"] = _residual_analysis(rep, index, out_bytes, still_large_mb)

    text = R.render_text(rep, index, analyze_only=False, lang=lang)
    R.write_reports(out_dir, stem, rep, text)
    _safe_print(text)
    return rep


def _residual_analysis(rep, index, out_bytes, threshold_mb):
    """压缩后仍超阈值时，产出残余大头的详细定位（页+元素+体积+建议）。"""
    residual = []
    for it in rep["top_items"]:
        cur = it["new"] if it["accepted"] else it["orig"]
        if cur <= 0:
            continue
        e = index.get(it["media"])
        pages = e.pages if e else []
        # 给建议（key，由 report 层按 lang 翻译）
        kind = it.get("kind", "")
        if kind == "video":
            # 已经是 x264 就别再建议 x264；改建议降分辨率/码率/外链
            hint = "hint-video-x264" if it.get("action") == "video-h264" else "hint-video"
        elif kind in ("png",):
            hint = "hint-png"
        elif kind in ("jpeg", "jpg"):
            hint = "hint-jpeg"
        elif kind == "vector":
            hint = "hint-vector"
        else:
            hint = "hint-done"
        residual.append({
            "media": it["media"], "kind": kind, "pages": pages,
            "bytes": cur, "pct_of_file": round(cur / out_bytes * 100, 1),
            "hint": hint,
        })
    return {"threshold_mb": threshold_mb, "output_bytes": out_bytes,
            "items": residual[:10]}



def _kind_of(ext: str) -> str:
    if ext in ("jpg", "jpeg"):
        return "jpeg"
    if ext == "png":
        return "png"
    if ext == "gif":
        return "gif"
    if ext in C.VECTOR_EXTS:
        return "vector"
    if ext in C.VIDEO_EXTS:
        return "video"
    if ext in C.AUDIO_EXTS:
        return "audio"
    if ext in ("ttf", "otf", "fntdata"):
        return "font"
    return ext or "unknown"


def _do_font_subset(raw, optimized, results, warnings):
    fonts = [n for n in raw if re.match(r'ppt/fonts/.*\.(fntdata|ttf|otf)$', n)]
    if not fonts:
        return
    used = set()
    for n, d in raw.items():
        if n.endswith(".xml") and re.search(r'/(slides|slideMasters|slideLayouts|notesSlides)/', n):
            for t in re.findall(rb'<a:t>(.*?)</a:t>', d, re.S):
                used.update(t.decode("utf-8", "replace"))
    if not used:
        warnings.append("font-subset-no-chars")
        return
    unicodes = ",".join(f"U+{ord(c):04X}" for c in sorted(used) if ord(c) > 31)
    for f in fonts:
        data = raw[f]
        # OOXML 字体混淆：fntdata 常被前 32 字节 XOR 混淆。若首字节非常见字体魔数，标注跳过。
        magic = data[:4]
        if not (magic in (b"\x00\x01\x00\x00", b"OTTO", b"true", b"ttcf") or f.endswith((".ttf", ".otf"))):
            warnings.append(f"font-subset-obfuscated|{f.split('/')[-1]}")
            results.append({"name": f, "kind": "font", "orig": len(data),
                            "new": len(data), "accepted": False, "action": "",
                            "reason": "obfuscated-font"})
            continue
        nb, action, reason = C.subset_font(data, unicodes)
        if nb and len(nb) < len(data):
            optimized[f] = nb
            warnings.append("font-subset-done")
            results.append({"name": f, "kind": "font", "orig": len(data),
                            "new": len(nb), "accepted": True, "action": action, "reason": ""})
        else:
            results.append({"name": f, "kind": "font", "orig": len(data),
                            "new": len(data), "accepted": False, "action": "",
                            "reason": reason or "not-smaller"})


def main():
    ap = argparse.ArgumentParser(description="视觉无损最大级别压缩 PPTX")
    ap.add_argument("input", help="输入 .pptx")
    ap.add_argument("--out", default=None,
                    help="输出目录(默认: 输入文件所在目录)")
    ap.add_argument("--analyze-only", action="store_true", help="只出体积分析报告")
    ap.add_argument("--no-repack-formats", action="store_true", help="关闭 PNG→JPEG 跨格式转换")
    ap.add_argument("--subset-fonts", action="store_true",
                    help="字体子集化(破坏性,默认关:会锁死文字编辑)")
    # 以下 3 项冗余清理默认开启，用 --no-* 关闭
    ap.add_argument("--no-crop", action="store_true", help="不做裁剪像素丢弃")
    ap.add_argument("--no-drop-layouts", action="store_true", help="不删无用版式")
    ap.add_argument("--no-strip-fast-save", action="store_true", help="不去冗余缩略图")
    ap.add_argument("--no-clean", action="store_true",
                    help="一键关闭全部冗余清理(仅做媒体压缩)")
    ap.add_argument("--av-codec", choices=["auto", "hevc", "x264"], default="auto",
                    help="视频编码器: auto(macOS用hevc硬件,其它用x264) / hevc / x264")
    ap.add_argument("--jpeg-quality", type=int, default=88)
    ap.add_argument("--retina", type=float, default=2.0)
    ap.add_argument("--min-save-kb", type=int, default=64)
    ap.add_argument("--still-large-mb", type=float, default=10.0,
                    help="压缩后超过此体积(MB)则附残余大头详细分析(默认10)")
    ap.add_argument("--lang", choices=["en", "zh"], default="en",
                    help="报告语言 (default: en)")
    args = ap.parse_args()

    check_deps()  # 硬依赖校验，缺则报错退出

    if not os.path.isfile(args.input):
        raise SystemExit(f"Error: file not found: {args.input}")
    # 默认输出到输入文件同目录（~/Desktop 不通用）
    out_dir = args.out or os.path.dirname(os.path.abspath(args.input))
    os.makedirs(out_dir, exist_ok=True)

    clean = not args.no_clean
    compress(args.input, out_dir,
             analyze_only=args.analyze_only,
             repack_formats=not args.no_repack_formats,
             subset_fonts=args.subset_fonts,
             apply_crop=clean and not args.no_crop,
             drop_unused_layouts=clean and not args.no_drop_layouts,
             strip_fast_save=clean and not args.no_strip_fast_save,
             av_codec=args.av_codec,
             jpeg_quality=args.jpeg_quality,
             retina=args.retina,
             min_save_kb=args.min_save_kb,
             still_large_mb=args.still_large_mb,
             lang=args.lang)


if __name__ == "__main__":
    main()
