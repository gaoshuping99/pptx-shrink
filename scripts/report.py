"""report.py — 压缩报告（JSON 机器可读 + 双语文本面向用户）。

数据来源：media.build_index 的 MediaEntry + 压缩阶段每项的 result dict。
每个 result: {name, kind, orig, new, accepted, action, reason}

i18n：action/reason/warning/hint 都是稳定英文 key（部分带 "|参数"），
本模块的 STRINGS/_t() 按 lang(en|zh) 翻译展示。JSON 报告保留英文 key（机读稳定）。
"""

from __future__ import annotations

import json


def _fmt(n: int) -> str:
    """字节→人类可读。"""
    x = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if x < 1024 or unit == "GB":
            if unit == "B":
                return f"{int(x)}B"
            return f"{x:.1f}{unit}"
        x /= 1024


# ---------------- i18n 词表 ----------------
# 静态 UI 标签
UI = {
    "en": {
        "analyze_title": "PPT size analysis 📊",
        "file_size": "File size",
        "done_title": "PPT compression done ✅",
        "before_after": "Before: {a}  →  After: {b}   (saved {p}%, -{d})",
        "by_type": "By type:",
        "top": "Top items (located by page / element):",
        "not_compressed": "Not further compressed ({n}):",
        "still_large": "⚠️ Still large after compression ({sz} > {th}MB). Biggest remaining:",
        "pct_of_file": "of file",
        "notes": "Notes:",
        "output": "Output:",
        "count_unit": "",
        "global": "global", "unref": "unreferenced", "dash": "—",
        "masters": "master/layout", "page": "p{n}", "pages_more": " +{n} more",
    },
    "zh": {
        "analyze_title": "PPT 体积分析 📊",
        "file_size": "文件大小",
        "done_title": "PPT 压缩完成 ✅",
        "before_after": "原始：{a}  →  压缩后：{b}   （省 {p}%，减 {d}）",
        "by_type": "按类型：",
        "top": "体积 Top（定位到页/元素）：",
        "not_compressed": "未进一步压缩（{n} 项）：",
        "still_large": "⚠️ 压缩后仍较大（{sz} > {th}MB），残余体积大头：",
        "pct_of_file": "占文件",
        "notes": "提示：",
        "output": "产出：",
        "count_unit": "个",
        "global": "全局", "unref": "未引用", "dash": "—",
        "masters": "母版/版式", "page": "第{n}页", "pages_more": " 等{n}页",
    },
}

# 媒体/元素类型
KIND = {
    "en": {"jpeg": "Image(JPEG)", "jpg": "Image(JPEG)", "png": "Image(PNG)",
           "video": "Video", "audio": "Audio", "font": "Font", "gif": "Image(GIF)",
           "vector": "Vector", "layout": "Layout", "thumbnail": "Thumbnail",
           "orphan-media": "Unref media", "orphan-embed": "Unref embed",
           "orphan-tag": "Unref tag", "orphan-aux": "Unref part"},
    "zh": {"jpeg": "图片(JPEG)", "jpg": "图片(JPEG)", "png": "图片(PNG)",
           "video": "视频", "audio": "音频", "font": "字体", "gif": "图片(GIF)",
           "vector": "矢量图", "layout": "版式", "thumbnail": "缩略图",
           "orphan-media": "无引用媒体", "orphan-embed": "无引用嵌入",
           "orphan-tag": "无引用标签", "orphan-aux": "无引用部件"},
}
ELEM = {
    "en": {"picture": "picture", "shape-fill": "shape fill", "background": "background",
           "video": "video", "audio": "audio", "unknown": "media"},
    "zh": {"picture": "图片", "shape-fill": "形状填充图", "background": "背景图",
           "video": "视频", "audio": "音频", "unknown": "媒体"},
}

# action / reason / warning / hint 的翻译（{0} 为可选参数）
MSG = {
    "en": {
        "jpeg-recompress": "JPEG recompress q{0}", "jpeg-lossless": "JPEG lossless",
        "png-to-jpeg": "PNG→JPEG q{0}", "png-quantize": "PNG quantize",
        "png-lossless": "PNG lossless",
        "video-hevc": "video HEVC re-encode", "video-h264": "video H.264 re-encode",
        "audio-recompress": "audio re-encode 128k", "font-subset": "font subset",
        "drop-unused-layout": "removed unused layout", "drop-thumbnail": "removed thumbnail",
        "drop-orphan-media": "removed unref media", "drop-orphan-embed": "removed unref embed",
        "drop-orphan-tag": "removed unref tag", "drop-orphan-aux": "removed unref part",
        "resize": " resize→{0}px", "crop-discard": "+crop discard",
        # reasons
        "not-smaller": "not smaller after re-encode",
        "gif-kept": "GIF (possibly animated), kept",
        "vector-kept": "vector (EMF/WMF), kept by design",
        "unknown-type": "unknown media type .{0}, kept",
        "below-threshold": "saved only {0}KB, below threshold",
        "wav-same-container": "WAV kept (no cross-container re-encode)",
        "analyze-only": "analyze-only mode", "already-compressed": "already compressed (idempotent skip)",
        "obfuscated-font": "obfuscated font, skipped",
        "pyftsubset-unavailable": "pyftsubset unavailable",
        # warnings
        "crop-discard|": "Crop-pixel discard: processed {0} (cropped areas removed, not restorable).",
        "drop-unused-layouts|": "Removed {0} unused layout(s) (not referenced by any slide).",
        "strip-fast-save|": "Removed {0} redundant thumbnail(s) (PowerPoint regenerates on save).",
        "font-subset-done": "Fonts subset: deck text can no longer be edited (new chars show as tofu).",
        "font-subset-no-chars": "Font subset skipped: no characters collected.",
        "font-subset-obfuscated|": "Font subset skipped for {0}: looks OOXML-obfuscated.",
        "post-check|": "Post-compression validation: {0}",
        # residual hints
        "hint-video": "video is the biggest chunk: try --av-codec x264, or external link / shorter clip",
        "hint-png": "large transparent PNG: verify transparency is needed, else convert to JPEG",
        "hint-jpeg": "image still large: lower --jpeg-quality (e.g. 82), or its displayed size is large",
        "hint-vector": "vector image, kept by design",
        "hint-done": "already optimized as much as possible",
    },
    "zh": {
        "jpeg-recompress": "JPEG重编码 q{0}", "jpeg-lossless": "JPEG无损优化",
        "png-to-jpeg": "PNG→JPEG q{0}", "png-quantize": "PNG量化",
        "png-lossless": "PNG无损优化",
        "video-hevc": "视频HEVC重编码", "video-h264": "视频H.264重编码",
        "audio-recompress": "音频重编码 128k", "font-subset": "字体子集化",
        "drop-unused-layout": "删除无用版式", "drop-thumbnail": "删除缩略图",
        "drop-orphan-media": "删除无引用媒体", "drop-orphan-embed": "删除无引用嵌入",
        "drop-orphan-tag": "删除无引用标签", "drop-orphan-aux": "删除无引用部件",
        "resize": " resize→{0}px", "crop-discard": "+裁剪丢弃",
        "not-smaller": "重编码后未变小",
        "gif-kept": "GIF(可能动图)，保留",
        "vector-kept": "矢量图(EMF/WMF)，按设计保留",
        "unknown-type": "未识别媒体类型 .{0}，保留",
        "below-threshold": "节省仅 {0}KB，低于阈值",
        "wav-same-container": "WAV 保留（不做跨容器重编码）",
        "analyze-only": "分析模式未压缩", "already-compressed": "已压缩过(幂等跳过)",
        "obfuscated-font": "疑似混淆字体，跳过",
        "pyftsubset-unavailable": "pyftsubset 不可用",
        "crop-discard|": "裁剪像素丢弃：处理了 {0} 处（被裁区域已删除，不可再拖回）。",
        "drop-unused-layouts|": "删除无用版式 {0} 个（未被任何幻灯片引用）。",
        "strip-fast-save|": "去除冗余缩略图 {0} 个（PowerPoint 保存时会自动重建）。",
        "font-subset-done": "已子集化字体：deck 文字将无法再增改（新字符会显示为缺字）。",
        "font-subset-no-chars": "字体子集化跳过：未采集到用字。",
        "font-subset-obfuscated|": "字体子集化跳过 {0}：疑似 OOXML 混淆字体。",
        "post-check|": "完成后校验：{0}",
        "hint-video": "视频仍是大头：可 --av-codec x264 或改用外部链接/降低时长分辨率",
        "hint-png": "透明PNG较大：确认是否真需透明，否则可转JPEG",
        "hint-jpeg": "图片仍大：可下调 --jpeg-quality（如 82）或该图显示尺寸本身很大",
        "hint-vector": "矢量图未压缩（按设计保留）",
        "hint-done": "该媒体已尽力压缩",
    },
}

_GLOBAL_KINDS = {"layout", "thumbnail", "orphan-media", "orphan-embed", "orphan-tag", "orphan-aux"}


def _t(key: str, lang: str) -> str:
    """翻译一个 'key' 或 'key|arg1|arg2' 形式的消息。未知 key 原样返回。"""
    if not key:
        return ""
    parts = key.split("|")
    base = parts[0]
    args = parts[1:]
    tbl = MSG.get(lang, MSG["en"])
    # action 里 resize 是拼接后缀（如 "png-to-jpeg|88|1280"）——特殊处理
    if base in ("jpeg-recompress", "png-to-jpeg") and args:
        s = tbl.get(base, base).format(args[0])
        if len(args) > 1:
            s += tbl["resize"].format(args[1])
        return s
    if base in ("png-quantize", "font-subset") and args:
        return tbl.get(base, base) + tbl["resize"].format(args[0])
    # 带参数的 warning/reason：key 存成 "base|"
    if args and (base + "|") in tbl:
        try:
            return tbl[base + "|"].format(args[0])
        except Exception:
            return tbl[base + "|"]
    return tbl.get(base, MSG["en"].get(base, base))


def _pages_str(entry, lang: str) -> str:
    u = UI[lang]
    scopes = entry.scopes
    pages = entry.pages
    if pages:
        head = "、".join(u["page"].format(n=p) for p in pages[:4]) if lang == "zh" \
            else ", ".join(u["page"].format(n=p) for p in pages[:4])
        if len(pages) > 4:
            head += u["pages_more"].format(n=len(pages))
        if scopes & {"slideMaster", "slideLayout"}:
            head += f" ({u['masters']})"
        return head
    if scopes & {"slideMaster", "slideLayout"}:
        return u["masters"]
    return u["unref"]


def build_report(input_path: str, output_path: str | None,
                 input_bytes: int, output_bytes: int | None,
                 index: dict, results: list[dict],
                 xml_touched: list[str], warnings: list[str],
                 now_iso: str) -> dict:
    by_type: dict[str, dict] = {}
    for r in results:
        k = r.get("kind", "unknown")
        b = by_type.setdefault(k, {"count": 0, "orig": 0, "new": 0})
        b["count"] += 1
        b["orig"] += r["orig"]
        b["new"] += r["new"] if r.get("accepted") else r["orig"]

    accepted = [r for r in results if r.get("accepted")]
    skipped = [r for r in results if not r.get("accepted")]

    def saved(r):
        return r["orig"] - (r["new"] if r.get("accepted") else r["orig"])
    ranked = sorted(results, key=lambda r: (saved(r), r["orig"]), reverse=True)

    def item_dict(r):
        e = index.get(r["name"])
        return {
            "media": r["name"],
            "kind": r.get("kind"),
            "elem": sorted(e.elems)[0] if e and e.elems else r.get("kind"),
            "pages": e.pages if e else [],
            "scope": sorted(e.scopes)[0] if e and e.scopes else "slide",
            "orig": r["orig"],
            "new": r["new"] if r.get("accepted") else r["orig"],
            "saved_pct": round(saved(r) / r["orig"] * 100, 1) if r["orig"] else 0.0,
            "action": r.get("action", ""),
            "crop_discarded": bool(r.get("crop_discarded")),
            "accepted": bool(r.get("accepted")),
        }

    top_items = [item_dict(r) for r in ranked if r["orig"] > 0][:10]
    not_compressed = [
        {"media": r["name"], "kind": r.get("kind"),
         "pages": index[r["name"]].pages if r["name"] in index else [],
         "reason": r.get("reason", "not-smaller")}
        for r in skipped
    ]

    saved_bytes = (input_bytes - output_bytes) if output_bytes is not None else 0
    return {
        "tool": "pptx-shrink", "version": "1.1.0", "generated_at": now_iso,
        "input": {"path": input_path, "bytes": input_bytes},
        "output": {"path": output_path, "bytes": output_bytes} if output_path else None,
        "summary": {
            "saved_bytes": saved_bytes,
            "saved_pct": round(saved_bytes / input_bytes * 100, 1) if input_bytes and output_bytes else 0.0,
            "media_count": len(results),
            "media_optimized": len(accepted),
            "media_skipped": len(skipped),
        },
        "by_type": by_type,
        "top_items": top_items,
        "not_compressed": not_compressed,
        "xml_touched": xml_touched,
        "warnings": warnings,
    }


def render_text(report: dict, index: dict, analyze_only: bool = False, lang: str = "en") -> str:
    if lang not in UI:
        lang = "en"
    u = UI[lang]
    L = []
    inp = report["input"]
    out = report.get("output")
    if analyze_only or not out:
        L.append(u["analyze_title"])
        L.append(f"{u['file_size']}: {_fmt(inp['bytes'])}")
    else:
        s = report["summary"]
        L.append(u["done_title"])
        L.append(u["before_after"].format(a=_fmt(inp['bytes']), b=_fmt(out['bytes']),
                                           p=s['saved_pct'], d=_fmt(s['saved_bytes'])))
    L.append("")

    # 按类型
    if report["by_type"]:
        L.append(u["by_type"])
        for k, b in sorted(report["by_type"].items(), key=lambda kv: -kv[1]["orig"]):
            name = KIND[lang].get(k, k)
            cu = u["count_unit"]
            if analyze_only:
                L.append(f"  {name} {b['count']}{cu}   {_fmt(b['orig'])}")
            else:
                L.append(f"  {name} {b['count']}{cu}   {_fmt(b['orig'])} → {_fmt(b['new'])}")
        L.append("")

    # Top-N
    if report["top_items"]:
        L.append(u["top"])
        for i, it in enumerate(report["top_items"], 1):
            e = index.get(it["media"])
            if e:
                loc = _pages_str(e, lang)
            elif it.get("kind") in _GLOBAL_KINDS:
                loc = u["global"]
            else:
                loc = u["dash"]
            elem = ELEM[lang].get(it["elem"], KIND[lang].get(it.get("kind"), it["elem"]))
            base = it["media"].split("/")[-1]
            if analyze_only or not it["accepted"]:
                L.append(f"  {i}. {loc} · {elem}  {base}  {_fmt(it['orig'])}")
            else:
                act = _t(it["action"], lang)
                if it.get("crop_discarded"):
                    act = (act + " " + _t("crop-discard", lang)).strip()
                tail = f"  [{act}]" if act else ""
                L.append(f"  {i}. {loc} · {elem}  {base}  "
                         f"{_fmt(it['orig'])} → {_fmt(it['new'])}  (-{it['saved_pct']:.0f}%){tail}")
        L.append("")

    # 未压缩项
    if report["not_compressed"] and not analyze_only:
        L.append(u["not_compressed"].format(n=len(report['not_compressed'])))
        for nc in report["not_compressed"][:8]:
            e = index.get(nc["media"])
            loc = _pages_str(e, lang) if e else u["dash"]
            L.append(f"  · {loc} · {nc['media'].split('/')[-1]} —— {_t(nc['reason'], lang)}")
        L.append("")

    # 残余大头
    sl = report.get("still_large")
    if sl and sl.get("items"):
        L.append(u["still_large"].format(sz=_fmt(sl['output_bytes']), th=sl['threshold_mb']))
        for j, r in enumerate(sl["items"], 1):
            pages = r.get("pages") or []
            if pages:
                loc = ("、".join(u["page"].format(n=p) for p in pages[:3]) if lang == "zh"
                       else ", ".join(u["page"].format(n=p) for p in pages[:3]))
            else:
                loc = u["global"]
            base = r["media"].split("/")[-1]
            L.append(f"  {j}. {loc} · {base}  {_fmt(r['bytes'])}"
                     f" ({u['pct_of_file']} {r['pct_of_file']}%) — {_t(r['hint'], lang)}")
        L.append("")

    if report["warnings"]:
        L.append(u["notes"])
        for w in report["warnings"]:
            L.append(f"  ⚠️ {_t(w, lang)}")
        L.append("")

    if out:
        L.append(u["output"])
        L.append(f"  {out['path']}")
    return "\n".join(L)


def write_reports(out_dir: str, stem: str, report: dict, text: str) -> tuple[str, str]:
    import os
    jp = os.path.join(out_dir, f"{stem}.report.json")
    tp = os.path.join(out_dir, f"{stem}.report.txt")
    with open(jp, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    with open(tp, "w", encoding="utf-8") as f:
        f.write(text)
    return jp, tp
