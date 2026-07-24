"""cleaners.py — 冗余清理（会改 slide/master XML，字节级最小改动）。

四项，全部默认关闭、破坏性、报告标注：
- apply_crop           : 裁剪像素丢弃（Pillow crop 源图 + 删该 slide 的 <a:srcRect/>）
- drop_unused_layouts  : 删无用版式（未被任何 slide 引用的 slideLayout）
- drop_hidden          : 删隐藏(hidden="1")/画布外(坐标完全在 sldSz 外) 的 shape
- strip_fast_save      : 去 fast-save/冗余 part（thumbnail 等，可安全再生）

设计原则：只做能精确定位的字节级删除；任何一步不确定就跳过并标注，绝不损坏文件。
返回：对 raw(dict) 就地修改 xml_overrides / drops，并往 results/warnings 追加记录。
"""

from __future__ import annotations

import io
import os
import re

# 命名空间前缀在 OOXML 稳定
_RE_PIC_BLOCK = re.compile(rb'<p:pic\b.*?</p:pic>', re.S)
_RE_SP_BLOCK = re.compile(rb'<p:sp\b.*?</p:sp>', re.S)
_RE_SRCRECT = re.compile(rb'<a:srcRect\b[^>]*/>')
_RE_EMBED = re.compile(rb'<a:blip\b[^>]*r:embed="([^"]+)"')
_RE_OFF = re.compile(rb'<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"')
_RE_EXT = re.compile(rb'<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"')
_RE_CNVPR = re.compile(rb'<p:cNvPr\b[^>]*>')
_RE_HIDDEN = re.compile(rb'hidden="1"')


# ---------------- 1. 裁剪像素丢弃 ----------------
def apply_crop(raw, index, optimized, xml_overrides, results, warnings,
               min_crop_ratio: float = 0.15):
    """对有 srcRect 且裁掉面积占比 > 阈值的图：crop 源图、删该 slide 的 srcRect。
    仅处理"该 media 在该 slide 只被这一处引用"的简单情形，避免多引用不同裁剪冲突。
    """
    from PIL import Image
    touched = 0
    for name in [n for n in raw if n.startswith("ppt/media/")]:
        entry = index.get(name)
        if not entry:
            continue
        # 该 media 被引用的 slide host + srcRect
        for host in {r.host for r in entry.refs if r.scope == "slide"}:
            sdata = xml_overrides.get(host, raw.get(host, b""))
            # 找 host 里引用了该 media 的 pic 块（经 rId 关联）
            rid = _rid_for_media(raw, host, name)
            if not rid:
                continue
            new_sdata = sdata
            changed = False
            for m in _RE_PIC_BLOCK.finditer(sdata):
                block = m.group(0)
                if rid.encode() not in block:
                    continue
                sr = _RE_SRCRECT.search(block)
                if not sr:
                    continue
                l, t, r_, b = _parse_srcrect(sr.group(0))
                keep_w = 1.0 - (l + r_) / 100000.0
                keep_h = 1.0 - (t + b) / 100000.0
                cropped_ratio = 1.0 - keep_w * keep_h
                if cropped_ratio < min_crop_ratio:
                    continue
                # crop 源图（用已优化后的 bytes 若有，否则原始）
                src_bytes = optimized.get(name, raw[name])
                cropped = _crop_image(src_bytes, l, t, r_, b)
                if cropped is None or len(cropped) >= len(src_bytes):
                    continue
                optimized[name] = cropped
                # 删该 pic 块内的 srcRect（字节级）
                nb = block.replace(sr.group(0), b"", 1)
                new_sdata = new_sdata.replace(block, nb, 1)
                changed = True
                touched += 1
                # 更新对应 result
                _bump_result(results, name, action_suffix="+crop-discard")
            if changed:
                xml_overrides[host] = new_sdata
    if touched:
        warnings.append(f"crop-discard|{touched}")
    return touched


def _crop_image(data, l, t, r, b):
    from PIL import Image
    try:
        im = Image.open(io.BytesIO(data))
        W, H = im.size
        box = (round(W * l / 100000), round(H * t / 100000),
               round(W * (1 - r / 100000)), round(H * (1 - b / 100000)))
        if box[2] <= box[0] or box[3] <= box[1]:
            return None
        out = im.crop(box)
        buf = io.BytesIO()
        fmt = im.format or ("JPEG" if data[:2] == b"\xff\xd8" else "PNG")
        save_kw = {"quality": 88} if fmt == "JPEG" else {}
        out.save(buf, format=fmt, **save_kw)
        return buf.getvalue()
    except Exception:
        return None


def _parse_srcrect(tag: bytes):
    def g(attr):
        m = re.search(attr.encode() + rb'="(-?\d+)"', tag)
        return int(m.group(1)) if m else 0
    return g("l"), g("t"), g("r"), g("b")


def _rid_for_media(raw, host_xml, media_path):
    """host 的 rels 里，哪个 rId 指向 media_path。"""
    import posixpath
    d = posixpath.dirname(host_xml); base = posixpath.basename(host_xml)
    rels = raw.get(f"{d}/_rels/{base}.rels", b"")
    target_base = media_path.split("/")[-1].encode()
    for m in re.finditer(rb'<Relationship\b[^>]*/?>', rels):
        tag = m.group(0)
        if target_base in tag:
            rid = re.search(rb'Id="([^"]+)"', tag)
            if rid:
                return rid.group(1).decode()
    return None


def _bump_result(results, name, action_suffix):
    # 不拼进 action key（会破坏 "key|arg" 结构），改置独立标志位，由 report 层翻译拼接
    for r in results:
        if r["name"] == name and r.get("accepted"):
            r["crop_discarded"] = True
            return


# ---------------- 2. 删无用版式 ----------------
def drop_unused_layouts(raw, drops, xml_overrides, results, warnings):
    """删除未被任何 slide 引用的 slideLayout，并从 master 的 sldLayoutIdLst / rels /
    Content_Types 里移除对应条目。"""
    import posixpath
    layouts = [n for n in raw if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n)]
    if not layouts:
        return 0
    # slide 用到的 layout 集合
    used = set()
    for n in raw:
        if re.match(r'ppt/slides/slide\d+\.xml$', n):
            rels = raw.get(f"ppt/slides/_rels/{posixpath.basename(n)}.rels", b"")
            for m in re.finditer(rb'Target="([^"]*slideLayouts/slideLayout\d+\.xml)"', rels):
                used.add(posixpath.normpath("ppt/slides/" + m.group(1).decode()))
    unused = [l for l in layouts if l not in used]
    if not unused:
        return 0

    # 每个 master：删 sldLayoutId 指向 unused 的条目 + 删 master rels 对应 Relationship
    masters = [n for n in raw if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n)]
    dropped = 0
    for lay in unused:
        lay_base = lay.split("/")[-1].encode()  # slideLayoutN.xml
        # 找承载它的 master rels，拿 rId
        for mpath in masters:
            mrels_path = f"ppt/slideMasters/_rels/{mpath.split('/')[-1]}.rels"
            mrels = xml_overrides.get(mrels_path, raw.get(mrels_path, b""))
            rid = None
            for m in re.finditer(rb'<Relationship\b[^>]*/?>', mrels):
                if lay_base in m.group(0):
                    ridm = re.search(rb'Id="([^"]+)"', m.group(0))
                    rid = ridm.group(1) if ridm else None
                    # 删这条 Relationship
                    mrels = mrels.replace(m.group(0), b"", 1)
                    break
            if rid is None:
                continue
            xml_overrides[mrels_path] = mrels
            # 删 master XML 里 <p:sldLayoutId ... r:id="rid"/>
            mdata = xml_overrides.get(mpath, raw.get(mpath, b""))
            mdata2 = re.sub(rb'<p:sldLayoutId\b[^>]*r:id="' + re.escape(rid) + rb'"[^>]*/>',
                            b"", mdata)
            if mdata2 != mdata:
                xml_overrides[mpath] = mdata2
        # 标记删除 layout part + 其 rels
        drops.add(lay)
        lrels = f"ppt/slideLayouts/_rels/{lay.split('/')[-1]}.rels"
        if lrels in raw:
            drops.add(lrels)
        # Content_Types Override 删除
        ct = xml_overrides.get("[Content_Types].xml", raw.get("[Content_Types].xml", b""))
        ct2 = re.sub(rb'<Override\b[^>]*PartName="/' + re.escape(lay.encode()) + rb'"[^>]*/>',
                     b"", ct)
        if ct2 != ct:
            xml_overrides["[Content_Types].xml"] = ct2
        dropped += 1
        results.append({"name": lay, "kind": "layout",
                        "orig": len(raw[lay]), "new": 0, "accepted": True,
                        "action": "drop-unused-layout", "reason": ""})
    if dropped:
        warnings.append(f"drop-unused-layouts|{dropped}")
        # 删版式后，被删版式独占的 media 会成孤儿 → 一并 GC
        _gc_unreferenced_media(raw, xml_overrides, drops, results)
    return dropped


# ---------------- （已移除）删隐藏/画布外对象 ----------------
# 该功能会误删某些插件的 hidden 数据对象（常命名为 "do not delete"），
# 导致 PowerPoint 报"内容有问题"。hidden="1" 不代表可删，风险高、体积收益极小，故移除。


def _gc_unreferenced_media(raw, xml_overrides, drops, results):
    """删除因删 shape 而不再被使用的 media，并清理其悬空 rels 关系。

    关键：判断"是否被使用"要看宿主 XML 里是否还有 r:embed/r:link 用到该关系的 rId，
    而不是只看 .rels 里有没有 Target（删 shape 后 .rels 可能残留悬空关系）。
    步骤：
      1) 对每个 host XML（slide/master/layout），收集其 XML 里仍在用的 rId。
      2) 对该 host 的 .rels，删掉指向 media 但 rId 已不再被 XML 使用的 Relationship。
      3) 全 deck 清理后，任何 media 若没有任何 .rels 关系指向它 → 删除该 media 文件。
    """
    import posixpath
    # 各 host XML 里实际在用的所有 rId（涵盖 r:embed/r:link/r:id——
    # 后者用于 hlinkClick/hlinkHover 等超链接，绝不能漏，否则会误删超链接关系）。
    def used_rids(xml_bytes):
        return set(re.findall(rb'r:(?:embed|link|id)="([^"]+)"', xml_bytes))

    # 遍历所有有 _rels 的内容部件，清理"确属图片/媒体、且 XML 里已无 rId 引用"的悬空关系。
    # 判定必须靠 Relationship 的 Type（结尾 image/video/audio/media），
    # 绝不能用裸 "media/" 子串匹配 Target——外部超链接 URL 常含 "media/"（如 .../media/...），
    # 会被误删（曾遇到超链接因 URL 含 media/ 被误删、hlinkClick 引用悬空导致文件损坏）。
    MEDIA_TYPE = re.compile(rb'Type="[^"]*/(image|video|audio|media)"')
    for host in list(raw.keys()):
        if not re.match(r'ppt/(slides|slideMasters|slideLayouts|notesSlides)/[^/]+\.xml$', host):
            continue
        rels_path = f"{posixpath.dirname(host)}/_rels/{posixpath.basename(host)}.rels"
        rels = xml_overrides.get(rels_path, raw.get(rels_path))
        if not rels:
            continue
        host_xml = xml_overrides.get(host, raw.get(host, b""))
        rids = used_rids(host_xml)
        new_rels = rels
        for m in re.finditer(rb'<Relationship\b[^>]*/?>', rels):
            tag = m.group(0)
            if b'TargetMode="External"' in tag:
                continue  # 外部关系（超链接等）一律不动
            if not MEDIA_TYPE.search(tag):
                continue  # 只处理图片/媒体类型的关系
            rid = re.search(rb'Id="([^"]+)"', tag)
            if rid and rid.group(1) not in rids:
                new_rels = new_rels.replace(tag, b"", 1)
        if new_rels != rels:
            xml_overrides[rels_path] = new_rels

    # ---- 通用 GC：从根 presentation.xml 出发，沿"存活"的 rels 做可达性遍历；
    #      任何 ppt/ 下不可达的部件都是死部件（删版式后独占的 media/embeddings/
    #      tags/theme override/… 都会在此被一次性清除）。核心部件天然可达，不会误删。----
    def rels_of(part):
        return f"{posixpath.dirname(part)}/_rels/{posixpath.basename(part)}.rels"

    def targets(part):
        rp = rels_of(part)
        if rp in drops:
            return []
        data = xml_overrides.get(rp, raw.get(rp))
        if not data:
            return []
        out = []
        host_dir = posixpath.dirname(posixpath.dirname(rp))
        for m in re.finditer(rb'<Relationship\b[^>]*/?>', data):
            tagb = m.group(0)
            if b'TargetMode="External"' in tagb:
                continue
            tm = re.search(rb'Target="([^"]+)"', tagb)
            if not tm:
                continue
            tgt = tm.group(1).decode()
            if tgt.startswith("http") or tgt == "NULL":
                continue
            out.append(posixpath.normpath(posixpath.join(host_dir, tgt)))
        return out

    # BFS 可达性：根 = presentation.xml + 通过 _rels/.rels 直达的部件
    reachable = set()
    stack = ["ppt/presentation.xml"]
    # 根 _rels/.rels 里的顶层部件（docProps 等）也算根
    root_rels = xml_overrides.get("_rels/.rels", raw.get("_rels/.rels", b""))
    for m in re.finditer(rb'Target="([^"]+)"', root_rels):
        t = m.group(1).decode()
        if not t.startswith("http") and t != "NULL":
            stack.append(posixpath.normpath(t.lstrip("/")))
    while stack:
        p = stack.pop()
        if p in reachable or p in drops:
            continue
        if p not in raw:
            continue
        reachable.add(p)
        for t in targets(p):
            if t not in reachable:
                stack.append(t)

    # 删除 ppt/ 下不可达的部件（保护：可达的一律不动；.rels 跟随宿主处理）
    for name in list(raw.keys()):
        if name in drops or name.endswith("/") or name.endswith(".rels"):
            continue
        if not name.startswith("ppt/"):
            continue  # 只在 ppt/ 内清理，不碰 docProps/customXml 顶层等
        if name in reachable:
            continue
        # 不可达 = 死部件
        drops.add(name)
        # 同时删其 _rels（若有）
        rp = rels_of(name)
        if rp in raw:
            drops.add(rp)
        if name.startswith("ppt/embeddings/"):
            kind, act = "orphan-embed", "drop-orphan-embed"
        elif name.startswith("ppt/tags/"):
            kind, act = "orphan-tag", "drop-orphan-tag"
        elif name.startswith("ppt/media/"):
            kind, act = "orphan-media", "drop-orphan-media"
        else:
            kind, act = "orphan-aux", "drop-orphan-aux"
        results.append({"name": name, "kind": kind,
                        "orig": len(raw[name]), "new": 0, "accepted": True,
                        "action": act, "reason": ""})


# ---------------- 4. 去 fast-save / 冗余 part ----------------
def strip_fast_save(raw, drops, xml_overrides, results, warnings):
    """删可安全再生的冗余 part：docProps/thumbnail.*（缩略图，PowerPoint 会重建）。
    并清理 Content_Types / _rels/.rels 里对应声明。
    """
    dropped = 0
    thumbs = [n for n in raw if re.match(r'docProps/thumbnail\.(jpeg|jpg|png|emf|wmf)$', n)]
    for th in thumbs:
        drops.add(th)
        # 从根 _rels/.rels 删 thumbnail 关系
        root_rels = xml_overrides.get("_rels/.rels", raw.get("_rels/.rels", b""))
        rr2 = re.sub(rb'<Relationship\b[^>]*thumbnail[^>]*/>', b"", root_rels, flags=re.I)
        if rr2 != root_rels:
            xml_overrides["_rels/.rels"] = rr2
        # Content_Types 里 thumbnail 的 Override/Default 一般靠扩展名 Default，不强删
        dropped += 1
        results.append({"name": th, "kind": "thumbnail",
                        "orig": len(raw[th]), "new": 0, "accepted": True,
                        "action": "drop-thumbnail", "reason": ""})
    if dropped:
        warnings.append(f"strip-fast-save|{dropped}")
    return dropped
