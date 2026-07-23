"""media.py — pptx 媒体反查索引 + 显示尺寸计算（只读解析，从不回写 XML）。

职责：
- 用 zipfile 读取 pptx，建立 media_path -> {pages, scope, elems, refs, display_px} 索引。
- 幻灯片真实页序：presentation.xml <p:sldIdLst> 的 r:id 顺序 -> presentation.xml.rels。
- media->页：各 slide/master/layout/notes 的 _rels 里 image/video/audio 关系。
- 显示尺寸：slide XML 里该 blip 所在 <p:spPr><a:ext cx cy/> (EMU) -> px。

所有解析基于正则/ElementTree 只读，绝不修改 XML bytes。
"""

from __future__ import annotations

import re
import posixpath
from dataclasses import dataclass, field

EMU_PER_INCH = 914400
DISPLAY_DPI = 96
EMU_PER_PX = EMU_PER_INCH / DISPLAY_DPI  # 9525.0

MEDIA_PREFIX = "ppt/media/"

# OOXML 里稳定的命名空间前缀；正则用字面前缀，对非标准前缀在 _resolve_prefixes 兜底。
_RE_SLDID = re.compile(rb'<p:sldId\b[^>]*\br:id="([^"]+)"')
_RE_PRES_REL = re.compile(
    rb'<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="(slides/slide\d+\.xml)"'
)
# rels 里 Id / Type / Target / TargetMode 顺序不定，分别抓取后再组合。
_RE_REL = re.compile(rb'<Relationship\b[^>]*/?>')
_RE_ATTR_ID = re.compile(rb'\bId="([^"]+)"')
_RE_ATTR_TYPE = re.compile(rb'\bType="([^"]+)"')
_RE_ATTR_TARGET = re.compile(rb'\bTarget="([^"]+)"')
_RE_ATTR_MODE = re.compile(rb'\bTargetMode="([^"]+)"')

# 媒体类关系类型（结尾）
_MEDIA_REL_KINDS = ("image", "video", "audio", "media")


@dataclass
class MediaRef:
    """一次媒体引用（某页某元素引用了某个 media）。"""
    page: int | None            # 幻灯片页码（1-based）；母版/版式引用为 None
    scope: str                  # slide / slideLayout / slideMaster / notesSlide
    host: str                   # 宿主 xml 路径
    elem: str                   # picture / shape-fill / background / video / audio / unknown
    display_px: tuple[int, int] | None  # 该引用处的显示像素 (w,h)


@dataclass
class MediaEntry:
    path: str                          # ppt/media/xxx
    size: int                          # 原始字节数
    ext: str                           # 小写扩展名，无点
    refs: list[MediaRef] = field(default_factory=list)

    @property
    def pages(self) -> list[int]:
        return sorted({r.page for r in self.refs if r.page is not None})

    @property
    def scopes(self) -> set[str]:
        return {r.scope for r in self.refs}

    @property
    def elems(self) -> set[str]:
        return {r.elem for r in self.refs}

    @property
    def max_display_px(self) -> tuple[int, int] | None:
        """一图多引用时取最大显示尺寸，保证任一处都不糊。"""
        dims = [r.display_px for r in self.refs if r.display_px]
        if not dims:
            return None
        return (max(d[0] for d in dims), max(d[1] for d in dims))


def emu_to_px(emu: int, retina: float = 1.0) -> int:
    return max(1, round(emu / EMU_PER_PX * retina))


def _norm_target(rels_path: str, target: str) -> str:
    """把 rels 里的相对 Target 归一到包内绝对路径。
    rels_path 如 ppt/slides/_rels/slide1.xml.rels，其宿主目录是 ppt/slides。
    """
    host_dir = posixpath.dirname(posixpath.dirname(rels_path))  # 去掉 _rels
    return posixpath.normpath(posixpath.join(host_dir, target))


def _rels_path_for(xml_path: str) -> str:
    d = posixpath.dirname(xml_path)
    b = posixpath.basename(xml_path)
    return f"{d}/_rels/{b}.rels"


def _scope_of(xml_path: str) -> str:
    if "/slideLayouts/" in xml_path:
        return "slideLayout"
    if "/slideMasters/" in xml_path:
        return "slideMaster"
    if "/notesSlides/" in xml_path:
        return "notesSlide"
    return "slide"


def _parse_rels(data: bytes, rels_path: str) -> dict[str, dict]:
    """返回 rid -> {'target': 包内路径 or 原始, 'kind': image/video/..., 'external': bool}"""
    out: dict[str, dict] = {}
    for m in _RE_REL.finditer(data):
        tag = m.group(0)
        rid = _RE_ATTR_ID.search(tag)
        typ = _RE_ATTR_TYPE.search(tag)
        tgt = _RE_ATTR_TARGET.search(tag)
        if not (rid and typ and tgt):
            continue
        typ_s = typ.group(1).decode("utf-8", "replace")
        kind = typ_s.rsplit("/", 1)[-1].lower()
        if kind not in _MEDIA_REL_KINDS:
            continue
        mode = _RE_ATTR_MODE.search(tag)
        external = bool(mode and mode.group(1) == b"External")
        tgt_s = tgt.group(1).decode("utf-8", "replace")
        target = tgt_s if external else _norm_target(rels_path, tgt_s)
        out[rid.group(1).decode("utf-8", "replace")] = {
            "target": target, "kind": kind, "external": external,
        }
    return out


# 元素类型判定：在 slide XML 里根据 r:embed/r:link 所处上下文推断，并抓取同一 spPr 的 a:ext。
def _classify_refs(slide_xml: bytes, rid2rel: dict[str, dict]):
    """产出 [(rid, elem, display_px|None)]。用轻量正则扫描 <p:pic>/<p:bg>/videoFile 等块。"""
    results = []
    text = slide_xml

    # 1) 图片/形状填充：<a:blip r:embed="rIdX"/> ... 其后可能有 <a:ext cx cy/>
    for m in re.finditer(rb'<(?:a|p):blip\b[^>]*\br:(?:embed|link)="([^"]+)"', text):
        rid = m.group(1).decode()
        # 找该 blip 往回最近的容器判断 elem
        pre = text[max(0, m.start() - 400):m.start()]
        if b"<p:bg" in pre or b"<a:bgPr" in pre:
            elem = "background"
        elif b"<p:spPr" in pre and b"<p:pic" not in pre:
            elem = "shape-fill"
        else:
            elem = "picture"
        # 找该引用之后最近的 <a:ext cx cy/>（同一 xfrm 块内）
        post = text[m.end():m.end() + 1200]
        ext = re.search(rb'<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', post)
        display = None
        if ext:
            display = (int(ext.group(1)), int(ext.group(2)))
        results.append((rid, elem, display))

    # 2) 视频/音频。PowerPoint 的视频对象引用较杂，需覆盖多种：
    #    <a:videoFile r:link|r:embed>、<a:audioFile ...>、<p:videoFile ...>、
    #    以及 <p14:media r:embed> / <*:media r:embed|r:link>（真正的 .../media 关系载体）。
    for m in re.finditer(rb'<(?:a|p):videoFile\b[^>]*\br:(?:link|embed)="([^"]+)"', text):
        results.append((m.group(1).decode(), "video", None))
    for m in re.finditer(rb'<(?:a|p):audioFile\b[^>]*\br:(?:link|embed)="([^"]+)"', text):
        results.append((m.group(1).decode(), "audio", None))
    # p14:media / a14:media / 任意前缀 :media —— 载 .../media 关系的 r:embed(或 r:link)
    for m in re.finditer(rb'<[a-z0-9]+:media\b[^>]*\br:(?:embed|link)="([^"]+)"', text):
        results.append((m.group(1).decode(), "video", None))

    return results


def build_index(raw: dict[str, bytes], retina: float = 1.0) -> dict[str, MediaEntry]:
    """核心：建 media_path -> MediaEntry 索引。raw 为 {zip内路径: bytes}。"""
    # 媒体清单
    entries: dict[str, MediaEntry] = {}
    for name, data in raw.items():
        if name.startswith(MEDIA_PREFIX) and "/" == name[len(MEDIA_PREFIX)-1]:
            pass
        if name.startswith(MEDIA_PREFIX) and not name.endswith("/"):
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            entries[name] = MediaEntry(path=name, size=len(data), ext=ext)

    # 页序：sldIdLst r:id 顺序 -> presentation.xml.rels
    pres = raw.get("ppt/presentation.xml", b"")
    prels = raw.get("ppt/_rels/presentation.xml.rels", b"")
    rid2slide = {rid.decode(): tgt.decode() for rid, tgt in _RE_PRES_REL.findall(prels)}
    slide_seq: list[tuple[int, str]] = []
    for pageno, rid in enumerate(_RE_SLDID.findall(pres), 1):
        tgt = rid2slide.get(rid.decode())
        if tgt:
            slide_seq.append((pageno, "ppt/" + tgt))

    def add_ref(target: str, ref: MediaRef):
        e = entries.get(target)
        if e is not None:
            e.refs.append(ref)

    # slide -> layout 引用映射（用于母版/版式媒体的页归属回填）
    slide_layout: dict[str, list[int]] = {}   # layout_path -> [页码]
    layout_master: dict[str, str] = {}         # layout_path -> master_path

    # 扫 slides
    for pageno, spath in slide_seq:
        sdata = raw.get(spath, b"")
        rid2rel = _parse_rels(raw.get(_rels_path_for(spath), b""), _rels_path_for(spath))
        # 记录该 slide 用的 layout
        for rid, rel in rid2rel.items():
            pass
        # slide->layout 在 slide 的 rels 里，Type 以 slideLayout 结尾（非媒体类，单独扫）
        srels = raw.get(_rels_path_for(spath), b"")
        for m in _RE_REL.finditer(srels):
            tag = m.group(0)
            typ = _RE_ATTR_TYPE.search(tag); tgt = _RE_ATTR_TARGET.search(tag)
            if typ and tgt and typ.group(1).decode().rsplit("/", 1)[-1] == "slideLayout":
                lp = _norm_target(_rels_path_for(spath), tgt.group(1).decode())
                slide_layout.setdefault(lp, []).append(pageno)
        # 分类媒体引用
        for rid, elem, display in _classify_refs(sdata, rid2rel):
            rel = rid2rel.get(rid)
            if not rel or rel["external"]:
                continue
            dpx = None
            if display:
                dpx = (emu_to_px(display[0], retina), emu_to_px(display[1], retina))
            add_ref(rel["target"], MediaRef(page=pageno, scope="slide", host=spath,
                                            elem=elem, display_px=dpx))

    # layout -> master 映射
    for lp in list(slide_layout.keys()):
        lrels = raw.get(_rels_path_for(lp), b"")
        for m in _RE_REL.finditer(lrels):
            tag = m.group(0)
            typ = _RE_ATTR_TYPE.search(tag); tgt = _RE_ATTR_TARGET.search(tag)
            if typ and tgt and typ.group(1).decode().rsplit("/", 1)[-1] == "slideMaster":
                layout_master[lp] = _norm_target(_rels_path_for(lp), tgt.group(1).decode())

    # 扫 masters & layouts & notes 的媒体（页归属靠引用链回填）
    for name in raw:
        if not re.match(r'ppt/(slideMasters|slideLayouts|notesSlides)/[^/]+\.xml$', name):
            continue
        scope = _scope_of(name)
        rid2rel = _parse_rels(raw.get(_rels_path_for(name), b""), _rels_path_for(name))
        # 该宿主影响的页码集合
        if scope == "slideLayout":
            host_pages = slide_layout.get(name, [])
        elif scope == "slideMaster":
            # 该 master 下所有 layout 的页
            host_pages = []
            for lp, mp in layout_master.items():
                if mp == name:
                    host_pages += slide_layout.get(lp, [])
        else:
            host_pages = []
        for rid, elem, display in _classify_refs(raw.get(name, b""), rid2rel):
            rel = rid2rel.get(rid)
            if not rel or rel["external"]:
                continue
            dpx = None
            if display:
                dpx = (emu_to_px(display[0], retina), emu_to_px(display[1], retina))
            if host_pages:
                for pg in sorted(set(host_pages)):
                    add_ref(rel["target"], MediaRef(page=pg, scope=scope, host=name,
                                                    elem=elem, display_px=dpx))
            else:
                add_ref(rel["target"], MediaRef(page=None, scope=scope, host=name,
                                                elem=elem, display_px=dpx))

    return entries


def read_pptx(path: str) -> tuple[list, dict[str, bytes]]:
    """读 pptx 到内存。返回 (infolist, {name: bytes})。加密/坏包抛异常由调用方处理。"""
    import zipfile
    with zipfile.ZipFile(path, "r") as z:
        infos = z.infolist()
        raw = {i.filename: z.read(i.filename) for i in infos}
    return infos, raw
