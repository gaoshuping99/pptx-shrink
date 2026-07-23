# pptx-shrink 参考

## 架构：为什么不用 pptx skill 的 unpack.py/pack.py
那两个脚本会 pretty-print XML（minidom 重序列化）、转义智能引号、pack 时 schema 校验+重序列化——对"编辑文本"有用，对"纯压缩媒体"多余且有风险（改 XML bytes、拖慢、校验噪音）。

本 skill 用标准库 `zipfile` 自解包/重打包：
- 非媒体部件（XML/rels/`[Content_Types].xml`）**原始 bytes 原样透传**。
- 只替换 `ppt/media/` 下二进制。
- 保持 `infolist()` 原始顺序、`date_time`、`external_attr`。
- 媒体是已压缩格式 → 重打包 `ZIP_STORED`（不再二次 deflate）；XML/rels → `ZIP_DEFLATED`。
- **唯一动 XML 的地方**：PNG→JPEG 时成对改 `[Content_Types].xml`（补 jpeg Default）+ 引用它的 rels 的 Target 扩展名。任一步失败则该图整体回退原件。

## 文件分工
- `media.py` — 读 pptx、建 `media_path → MediaEntry{refs:[MediaRef], size, ext}` 反查索引。
  - 页序：`presentation.xml` 的 `<p:sldId r:id>` 顺序 → `presentation.xml.rels`。
  - media→页：各 slide/master/layout/notes 的 `_rels` 里 image/video/audio 关系。
  - 元素类型：slide XML 里 `<a:blip r:embed>` 的上下文（picture/shape-fill/background）、`<a:videoFile>`/`<a:audioFile>`。
  - 显示尺寸：blip 之后最近的 `<a:ext cx cy>`(EMU) → px。母版/版式媒体的页归属靠 slide→layout→master 引用链回填。
  - EMU→px：`px = EMU / 914400 * 96`（=EMU/9525）。
- `mediacodecs.py` — 编码封装（**注意：不能叫 codecs.py，会撞 Python 标准库**）。
- `report.py` — JSON + 中文文本报告。
- `compress_pptx.py` — 编排：analyze/compress、阈值门、成对改 XML、重打包、幂等标记。

## 图片决策树
```
JPEG → magick resize(WxH>) -strip -sampling-factor 4:2:0 -quality 88
       没变小 → jpegtran -optimize -progressive 无损 → 仍没变小 → 保留原件
PNG
  ├─ 无有效 alpha 且允许跨格式 → 转 JPEG(先 -flatten 白底)   [跨格式, 改 XML]
  └─ 有 alpha / 转 JPEG 不划算 → 先按 target resize → pngquant --quality=65-90
        pngquant 退出码 99 = 达不到质量下限 → 用无下限重试
        仍没变小 → magick -strip 无损 → 保留原件
TIFF/BMP → 按 PNG 分支（转 JPEG 或量化）
GIF → 保留（本机无 gifsicle）；EMF/WMF → 跳过（不光栅化）
```
- 降采样目标 = `min(源图实际像素, 显示px × retina)`，**绝不放大**；一图多引用取最大显示尺寸。

## 视频/音频（同容器红线）
- 视频 mp4/mov/m4v → 同扩展名。编码器按 `--av-codec` 与平台选择：
  `auto`（默认）在 **macOS** 用 `hevc_videotoolbox -q:v 60 -tag:v hvc1`（硬件加速、体积小），
  **非 macOS** 直接用 `libx264 -crf 23 -preset medium`（避免尝试不存在的编码器）；
  `hevc` 失败自动回退 `libx264`。统一 `scale='min(1920,iw)':-2`；音轨 `aac 128k`；`+faststart`。
- 音频：mp3/m4a/aac 同容器重编码 128k；**wav 装不下 aac → 跳过**（改容器会导致播放失败）。
- 仅当输出更小才替换。

## 字体子集化（--subset-fonts, 默认关）
- 收集全 deck `<a:t>` 用字 → `pyftsubset --unicodes=... --desubroutinize`。
- OOXML 嵌入字体 `.fntdata` 常被前 32 字节 XOR presentation GUID 混淆。**本 skill 保守判断魔数**（`\x00\x01\x00\x00`/`OTTO`/`true`/`ttcf`），非常见魔数即判为混淆并跳过并在报告标注，绝不损坏字体。要支持混淆字体需另加 deobfuscate/reobfuscate。

## 幂等与阈值
- 阈值门：`orig-new >= min_save_kb*1024` 或 `new/orig <= 0.90` 才接受。
- 输出写 `docProps/pptxshrink.json` = `{tool, version, media_sha:{path:sha256}}`；再压时命中 sha 直接跳过。

## 已知边界
| 场景 | 行为 |
|---|---|
| 加密 pptx | 检测 OLE 头 `D0CF11E0` → 报错，不猜密码 |
| 坏 zip | BadZipFile → 报错 |
| 外部链接媒体 | rels `TargetMode=External`/`r:link` 不在包内 → 跳过 |
| 非标准命名空间前缀 | 正则用字面 a/r/p；异常前缀可能漏认（可后续从根 xmlns 反解增强） |
| 超大文件 | v1 整包读内存；>数百 MB 的极端大文件后续可加流式 |
| `--apply-crop` | v1 预留未实现，报告标注 |

## 测试基线（随机噪声图，收益偏极端）
36.1MB（4000px JPEG + 3000px 无透明 PNG + 1200px 透明 PNG，3 页）→ 571KB（-98.5%），
LibreOffice 可渲染成 PDF（= PowerPoint 可打开），幂等二次压缩 0 变动。
真实照片 deck 收益通常 50–85%。
