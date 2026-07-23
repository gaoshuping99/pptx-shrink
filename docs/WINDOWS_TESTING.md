# Windows Acceptance Test / Windows 验收清单

This skill was developed and validated on macOS. The logic is cross-platform, but
Windows-specific behavior (subprocess resolution, `.exe` tools, the HEVC→x264 codec
path, LibreOffice detection, output paths) must be verified on a real Windows machine
before publishing to a skill hub.

本 skill 在 macOS 上开发验证。逻辑已跨平台，但以下 **Windows 专属行为**需在真机验证后
再上架：subprocess 解析、`.exe` 工具、HEVC→x264 编码器路径、LibreOffice 探测、输出路径。

Run each step, compare against **Expected**. Note anything that differs.
逐项运行，对照 **Expected**；记录任何不符。

---

## 0. Setup / 环境准备

```powershell
# Install hard dependencies (choco; or use winget/scoop)
choco install ffmpeg imagemagick pngquant -y
pip install Pillow

# (optional) LibreOffice for the extra render-validation pass
choco install libreoffice-fresh -y

# Verify tools are on PATH
where ffmpeg
where ffprobe
where magick
where pngquant
python -c "import PIL; print(PIL.__version__)"
```

**Expected / 预期**: every `where` prints a path ending in `.exe`; Pillow prints a version.

---

## 1. Dependency check fires correctly / 依赖探测

Temporarily rename or remove one tool from PATH (e.g. rename `pngquant.exe`), then:

```powershell
python scripts\compress_pptx.py some.pptx
```

**Expected / 预期**: exits non-zero (`echo %ERRORLEVEL%` ≠ 0), prints
`Missing CLI tools ...` and a `choco install ...` line naming the missing package.
Restore the tool afterward.

---

## 2. Basic compress + output location / 基本压缩与输出位置

Use a real deck with images (ideally also an embedded video). Put it at e.g.
`C:\Users\<you>\Documents\deck.pptx`, then:

```powershell
python scripts\compress_pptx.py "C:\Users\<you>\Documents\deck.pptx"
```

**Expected / 预期**:
- Prints `PPT compression done ✅` with a Before→After size line.
- Output `deck.compressed.pptx` + `deck.report.txt` + `deck.report.json` land
  **in the same folder as the input** (`C:\Users\<you>\Documents\`), NOT the Desktop.
- Original `deck.pptx` is unchanged (same size/mtime).
- Report contains `Post-compression validation: structure OK ...`.

---

## 3. **Open the output in real PowerPoint (Windows)** / 用真 PowerPoint 打开 ⭐ 最关键

Double-click `deck.compressed.pptx` in **Microsoft PowerPoint on Windows**.

**Expected / 预期**: opens **without** the "PowerPoint found a problem with content"
repair dialog. Spot-check a few slides — images look fine, embedded video plays,
hyperlinks work. (This is the definitive check that macOS/LibreOffice can't fully replicate.)

---

## 4. Video codec path (no macOS-only encoder) / 视频编码器路径

`hevc_videotoolbox` is macOS-only. On Windows, `--av-codec auto` must use `libx264`.
Use a deck with an embedded video:

```powershell
python scripts\compress_pptx.py "deck_with_video.pptx"
```

**Expected / 预期**: video shrinks; report action shows `video H.264 re-encode`
(NOT `video HEVC re-encode`); no ffmpeg "Unknown encoder 'hevc_videotoolbox'" error.
Also try `--av-codec x264` explicitly — same result.

---

## 5. LibreOffice detection / LibreOffice 探测

If LibreOffice is installed (step 0), the report note should read
`structure OK + LibreOffice render OK`. If NOT installed, it should read
`structure OK (LibreOffice not found, render check skipped)` — and still succeed.

**Expected / 预期**: both cases succeed; the note reflects whether LibreOffice was found.
(Verifies the Windows path `C:\Program Files\LibreOffice\program\soffice.exe` /
`where soffice` detection works.)

---

## 6. Bilingual report / 双语报告

```powershell
python scripts\compress_pptx.py deck.pptx --lang zh
```

**Expected / 预期**: report is in Chinese (`PPT 压缩完成 ✅`, `按类型：` …).
Default (no `--lang`) is English. No garbled characters in the Windows console
(if you see mojibake, run `chcp 65001` first — that's a console encoding issue, not a bug).

---

## 7. Analyze-only / 仅分析

```powershell
python scripts\compress_pptx.py deck.pptx --analyze-only
```

**Expected / 预期**: prints a size breakdown (per page/element), writes NO
`.compressed.pptx` (only the report), exits 0.

---

## 8. Idempotency / 幂等

Run step 2 twice on the same input.

**Expected / 预期**: second run saves ~0% (media already processed, skipped via the
`.pptxshrink.json` sidecar). No degradation, no error.

---

## 9. Paths with spaces / 含空格路径

Windows decks often live under paths with spaces. Test one:

```powershell
python scripts\compress_pptx.py "C:\Users\<you>\OneDrive - Company\My Deck (final).pptx"
```

**Expected / 预期**: works; output lands beside the input; no path-parsing error.

---

## Pass criteria / 通过标准

All of the above behave as Expected, and **step 3 (real PowerPoint opens the output
without a repair prompt) passes on at least 2–3 different real decks** (mixed images,
video, hyperlinks, many layouts). Then it's safe to publish to a skill hub.

If step 3 ever shows the repair dialog, capture the deck and run the built-in strict
check to see what structural issue was introduced (compare against the original):
the validator in `scripts/compress_pptx.py` (`_structural_issues`) already flags
dangling refs / orphan parts / undeclared extensions / broken relationship refs.
