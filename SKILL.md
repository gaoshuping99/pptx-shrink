---
name: pptx-shrink
description: >
  Compress / shrink / 压缩 a .pptx presentation to the smallest size with
  visually-lossless quality, and report where the size comes from (per page,
  per element, per byte). Trigger on: 压缩PPT / 压缩pptx / PPT太大 / PPT瘦身 /
  reduce or shrink presentation file size / make a deck smaller / optimize
  images in a deck / deck too big to email. Downsamples & re-encodes images,
  converts opaque PNG→JPEG, quantizes transparent PNGs, re-encodes embedded
  video/audio, optional font subsetting & crop-pixel discard, strips redundancy.
  Media-only mode (--no-clean) leaves XML untouched (except a paired
  Content_Types/rels rename for PNG→JPEG); the default cleanup mode also edits some
  slide/master XML to discard cropped pixels and drop unused layouts. Output stays a
  standard editable .pptx. Cross-platform (macOS / Windows / Linux). Bilingual report (en/zh).
license: MIT
---

# pptx-shrink — visually-lossless maximum PPTX compression

Compress a .pptx to the smallest size a human can't tell apart, and report where
the size comes from (**page + element + bytes**). Architecture follows NXPowerLite
(per-media re-encode, stays editable); the size report follows iSlide's file analysis.

## When to use
When the user says: compress/shrink PPT, PPT too big, deck too large to send,
"why is this deck so big", 压缩PPT / PPT瘦身 / 发不出去.

## Requirements (hard dependencies)
Must be on `PATH`. The tool checks at startup and prints install commands if missing.

| Tool | macOS | Windows | Linux |
|---|---|---|---|
| ffmpeg (+ffprobe) | `brew install ffmpeg` | `choco install ffmpeg` | `apt install ffmpeg` |
| ImageMagick 7 (`magick`) | `brew install imagemagick` | `choco install imagemagick` | `apt install imagemagick` |
| pngquant | `brew install pngquant` | `choco install pngquant` | `apt install pngquant` |
| Python **Pillow** | `pip install Pillow` | same | same |

**Optional** (missing only disables that feature, never errors):
- `jpegtran` — JPEG lossless optimization (falls back to `magick`)
- `fonttools`/`pyftsubset` — `--subset-fonts`
- LibreOffice — extra render validation pass

Python 3.9+. Uses only stdlib `zipfile` for (un)packing.

## One-shot compress (max, safe defaults)
```bash
python3 scripts/compress_pptx.py <input.pptx>
```
- Output → **input file's directory** as `<name>.compressed.pptx` + `.report.json`/`.report.txt`.
- On by default: image downsample (2× retina) + JPEG re-encode, opaque **PNG→JPEG**,
  transparent PNG quantize, video/audio re-encode, and redundancy cleanup
  (crop-pixel discard, drop unused layouts, drop redundant thumbnail).
- **What gets modified**: media bytes under `ppt/media` are always rewritten.
  In **default (cleanup) mode**, some XML is also edited — `[Content_Types].xml` and rels
  (PNG→JPEG rename), slide XML (crop-pixel discard removes `<a:srcRect>`), master XML +
  rels (drop unused layouts), and `_rels/.rels` (drop thumbnail). Use **`--no-clean`** for
  media-only compression that leaves XML untouched (bar the PNG→JPEG rename). Either way the
  output is validated to open correctly before it's written.
- **Safety**: writes a temp file → validates it opens (zip integrity + OOXML structure
  + no dangling relationship refs, diffed against the original so pre-existing PowerPoint
  quirks don't cause false rollbacks; plus a LibreOffice render pass if available) →
  only then commits. **If validation fails it discards the temp and errors; the original
  is never touched.**
- If still large (>10MB) it appends a "biggest remaining" breakdown with page + element + hints.

## Analyze only
```bash
python3 scripts/compress_pptx.py <input.pptx> --analyze-only
```

## Options
| Flag | Default | Meaning |
|---|---|---|
| `--out DIR` | input file's dir | output directory |
| `--lang en\|zh` | en | report language |
| `--av-codec auto\|hevc\|x264` | auto | video codec (auto: hevc_videotoolbox on macOS, libx264 elsewhere) |
| `--no-repack-formats` | off | disable PNG→JPEG cross-format |
| `--subset-fonts` | off | font subsetting (**destructive**: locks text editing) |
| `--no-crop` / `--no-drop-layouts` / `--no-strip-fast-save` | (on) | disable a cleanup |
| `--no-clean` | off | disable all cleanup (media compression only) |
| `--jpeg-quality N` | 88 | JPEG quality (85–90 = visually lossless) |
| `--retina F` | 2.0 | pixel headroom kept when downsampling |
| `--min-save-kb N` | 64 | min saving per file to accept |
| `--still-large-mb N` | 10 | threshold for the residual breakdown |

## Notes / caveats
- `--subset-fonts` locks text editing (new chars → tofu); default off.
- PNG→JPEG only for **opaque** PNGs; transparent ones go through pngquant.
- Encrypted pptx → clear error (remove password first).
- External-link media / EMF·WMF vector / animated GIF are kept by design (noted in report).
- **Idempotent**: a sidecar `<name>.pptxshrink.json` records media sha256; re-running skips
  already-processed media (no double-lossy degradation).

See `reference.md` for architecture and the pitfalls hardened against.
