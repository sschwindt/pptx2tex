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

For compiling the generated LaTeX, you need a LaTeX distribution with LuaLaTeX and the beamer package (e.g., TeX Live or MiKTeX).

## Usage

1. Place your PowerPoint presentation(s) in the `pptx-input/` directory.

2. Run the converter:

   ```bash
   python pptx2tex.py
   ```

3. Find the output in `tex-output/<presentation-name>/` containing:
   - `beamer-main.tex` - Main LaTeX file
   - `section-*.tex` - Per-section content files
   - `beamertheme.sty` - Beamer theme style file
   - `beamerProject-TexStudio.txss2` - Project file for [TeXstudio](https://www.texstudio.org/)
   - `fig/` - Extracted images
   - `videos/` - Extracted videos with placeholder thumbnail
   - `theme/` - Theme assets (logos, backgrounds)

4. Compile the presentation:

   ```bash
   cd tex-output/<presentation-name>
   lualatex beamer-main.tex
   ```

   Or open `beamerProject-TexStudio.txss2` in [TeXstudio](https://www.texstudio.org/) for an integrated editing experience. Note that multiple runs of `lualatex` may be required to resolve references (for the Table of Contents) and generate the PDF file.

### Template Customization

The `template/` directory contains the source files that are copied to each output:
- `beamer-main-template.tex` - Template with placeholders for title, author, sections
- `beamertheme.sty` - Beamer theme definition
- `theme/` - Logos and background images
- `videos/generic-thumbnail.jpg` - Default video placeholder

## Disclaimer

This code was vibecoded with [Claude Code](https://claude.ai/code) by Anthropic.

## License

BSD 3-Clause License - see [LICENSE](LICENSE) file for details.

Copyright (c) 2026, Sebastian Schwindt
