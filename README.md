# pptx-shrink

Visually-lossless maximum compression for PowerPoint `.pptx` files, with a
size-breakdown report (per page, per element, per byte). Cross-platform
(macOS / Windows / Linux). Bilingual report (English / 中文).

Compresses images (downsample + re-encode, opaque PNG→JPEG, PNG quantize),
re-encodes embedded video/audio, optionally subsets fonts and discards cropped
pixels, and strips redundancy — while keeping the output a standard **editable**
`.pptx`. Only `ppt/media` bytes are rewritten via `zipfile`; XML is left intact
(the one exception is a paired `[Content_Types].xml`/rels rename for PNG→JPEG).

## Install dependencies

Hard dependencies (must be on `PATH`):

```bash
# macOS
brew install ffmpeg imagemagick pngquant
pip install Pillow

# Windows
choco install ffmpeg imagemagick pngquant   # or winget
pip install Pillow

# Linux
sudo apt install ffmpeg imagemagick pngquant
pip install Pillow
```

Optional (missing only disables that feature): `jpegtran` (JPEG lossless; falls
back to ImageMagick), `fonttools`/`pyftsubset` (`--subset-fonts`), LibreOffice
(extra render-validation pass).

## Usage

```bash
# Compress to smallest visually-lossless size (safe defaults)
python3 scripts/compress_pptx.py deck.pptx

# Analyze only — show what's bloating the deck
python3 scripts/compress_pptx.py deck.pptx --analyze-only

# Chinese report
python3 scripts/compress_pptx.py deck.pptx --lang zh
```

Output lands next to the input as `deck.compressed.pptx` plus `.report.txt` /
`.report.json`. The original file is never modified. See `SKILL.md` for all flags.

## Safety

Writes to a temp file, validates it (zip integrity, OOXML structure, no dangling
relationship references — diffed against the original so pre-existing quirks don't
trip it, plus a LibreOffice render pass when available), and only then commits.
If validation fails, the temp is discarded and the original is untouched.

## License

MIT — see `LICENSE`.
