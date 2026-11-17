# pptx2tex

## Purpose

pptx2tex converts PowerPoint presentations (.pptx) to LaTeX Beamer slides using a custom theme of the Institute for Modelling Hydraulic and Environmental Systems (University of Stuttgart). The converter extracts images, videos, and text content while preserving slide structure, animations, and positioned elements.

## Requirements

- Python 3.8+
- [python-pptx](https://python-pptx.readthedocs.io/) library

Install dependencies:

```bash
pip install python-pptx
```

Optional tools, used when available:

- [Inkscape](https://inkscape.org/): converts EMF/WMF pictures and SVG-only Office icons, which Pillow cannot render.
- LibreOffice (`soffice`) and `pdftoppm`: render slides made of many freeform shapes to images (and EMF/WMF fallback).

For compiling the generated LaTeX, you need a LaTeX distribution with the beamer package (e.g., TeX Live or MiKTeX, also on Windows). Both pdfLaTeX and LuaLaTeX work and give the same Helvetica-like font.

## Usage

1. Place your PowerPoint presentation(s) in the `pptx-input/` directory.

2. Run the converter:

   ```bash
   python pptx2tex.py
   ```

   Optional: `python pptx2tex.py presentation.pptx --output-dir /some/folder` writes the output to `/some/folder/<presentation-name>/` instead of `tex-output/` (keeps this repository clean).

3. Find the output in `tex-output/<presentation-name>/` containing:
   - `beamer-main.tex` - Main LaTeX file
   - `section-*.tex` - Per-section content files
   - `beamertheme.sty` - Beamer theme style file
   - `beamerProject-TexStudio.txss2` - Project file for [TeXstudio](https://www.texstudio.org/)
   - `fig/` - Extracted images
   - `videos/` - Extracted videos with their poster frames (generic thumbnail if a video has none)
   - `theme/` - Theme assets (logos, backgrounds)

4. Compile the presentation:

   ```bash
   cd tex-output/<presentation-name>
   pdflatex beamer-main.tex   # or lualatex
   ```

   Or open `beamerProject-TexStudio.txss2` in [TeXstudio](https://www.texstudio.org/) for an integrated editing experience. Note that two runs are required to resolve references (for the Table of Contents).

### Videos

Videos are not embedded in the PDF. Each poster frame is a relative link to `videos/<clip>` (macro `\pptxvideo` of the theme), so keep the `videos/` folder next to the PDF:

- Adobe Acrobat Reader (Windows, macOS), SumatraPDF, Okular, Evince: a click opens the clip in the default player (Acrobat asks for confirmation).
- [pdfpc](https://pdfpc.github.io/): plays the clip inside the slide. With `\pptxvideoautoplaytrue` in the preamble, pdfpc starts and loops it automatically (pdfpc only: Acrobat would then look for a file named `clip.mp4?autostart&loop`).
- Browser PDF viewers block such links; the poster and the file name stay visible.

### What the converter handles

- Pictures (also EMF/WMF and SVG-only icons via Inkscape), with their rotation and flips, capped at 2000 px and saved as JPG.
- Videos, including videos in content placeholders and several videos per slide, with their poster frames.
- Tables (as `tabular`), bullet levels, symbol-font characters (Wingdings/Symbol arrows, check marks), soft line breaks.
- Section-header slides start a new `\section`; section-header slides that also carry pictures, videos, tables or charts are converted as normal frames.
- Not converted: PowerPoint equations (OMML). Slides containing them get a `% TODO` comment.

### Template Customization

The `template/` directory contains the source files that are copied to each output:
- `beamer-main-template.tex` - Template with placeholders for title, author, sections
- `beamertheme.sty` - Beamer theme definition
- `theme/` - Logos and background images
- `videos/generic-thumbnail.jpg` - Video placeholder for clips without a poster frame

## Disclaimer

This code was vibecoded with [Claude Code](https://claude.ai/code) by Anthropic.

## License

BSD 3-Clause License - see [LICENSE](LICENSE) file for details.

Copyright (c) 2026, Sebastian Schwindt
