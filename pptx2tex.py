#!/usr/bin/env python3
"""
PPTX to LaTeX Beamer Converter

Converts PowerPoint presentations (.pptx) to LaTeX Beamer format using
the beamertheme.sty style file. Output defaults to beamer-slides.tex
for compatibility with beamerProject-TexStudio.txss2.

Usage:
    python pptx2tex.py [input.pptx] [output.tex]

If no arguments provided, processes all .pptx files in pptx-input/
"""

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Optional

import zipfile
import shutil

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
    from pptx.enum.dml import MSO_THEME_COLOR
except ImportError:
    print("Error: python-pptx is required. Install with: pip install python-pptx")
    sys.exit(1)

try:
    from PIL import Image
    import io
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    print("Warning: Pillow not installed. Images will not be converted to JPG. Install with: pip install Pillow")


def patch_pptx_content_types(pptx_path: str) -> str:
    """
    Patch PPTX file to add missing content types for webp and other formats.
    Returns path to patched file (may be a temp copy).
    """
    import tempfile
    import xml.etree.ElementTree as ET

    # Create a temp copy
    temp_dir = tempfile.mkdtemp()
    temp_pptx = os.path.join(temp_dir, "patched.pptx")
    shutil.copy2(pptx_path, temp_pptx)

    # Add missing content types
    content_types_to_add = {
        '.webp': 'image/webp',
        '.svg': 'image/svg+xml',
        '.heic': 'image/heic',
        '.heif': 'image/heif',
        '.avif': 'image/avif',
        '.mp4': 'video/mp4',
        '.webm': 'video/webm',
        '.m4v': 'video/x-m4v',
        '.mov': 'video/quicktime',
    }

    try:
        with zipfile.ZipFile(temp_pptx, 'a') as zf:
            # Read existing content types
            ct_xml = zf.read('[Content_Types].xml').decode('utf-8')
            root = ET.fromstring(ct_xml)

            ns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
            ET.register_namespace('', ns['ct'])

            # Get existing extensions
            existing = set()
            for default in root.findall('.//ct:Default', ns):
                ext = default.get('Extension', '').lower()
                existing.add('.' + ext if not ext.startswith('.') else ext)

            # Add missing types
            modified = False
            for ext, content_type in content_types_to_add.items():
                if ext.lower() not in existing:
                    elem = ET.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/content-types}Default')
                    elem.set('Extension', ext.lstrip('.'))
                    elem.set('ContentType', content_type)
                    modified = True

            if modified:
                # Write back
                new_ct = ET.tostring(root, encoding='unicode')
                # Remove old and add new
                with zipfile.ZipFile(pptx_path, 'r') as original:
                    with zipfile.ZipFile(temp_pptx, 'w') as new_zip:
                        for item in original.infolist():
                            if item.filename == '[Content_Types].xml':
                                new_zip.writestr(item, new_ct)
                            else:
                                new_zip.writestr(item, original.read(item.filename))

        return temp_pptx
    except Exception as e:
        print(f"Warning: Could not patch content types: {e}")
        return pptx_path


class PPTXToLatexConverter:
    """Converts PPTX presentations to LaTeX Beamer format."""

    def __init__(self, input_path: str, output_path: str,
                 fig_dir: str = "fig", video_dir: str = "videos"):
        self.input_path = Path(input_path)
        self.output_path = Path(output_path)
        self.fig_dir = Path(fig_dir)
        self.video_dir = Path(video_dir)
        self.image_counter = 0
        self.video_counter = 0

        # Ensure output directories exist
        self.fig_dir.mkdir(parents=True, exist_ok=True)
        self.video_dir.mkdir(parents=True, exist_ok=True)

    def escape_latex(self, text: str) -> str:
        """Escape special LaTeX characters."""
        if not text:
            return ""

        # Order matters - escape backslash first
        replacements = [
            ('\\', r'\textbackslash{}'),
            ('&', r'\&'),
            ('%', r'\%'),
            ('$', r'\$'),
            ('#', r'\#'),
            ('_', r'\_'),
            ('{', r'\{'),
            ('}', r'\}'),
            ('~', r'\textasciitilde{}'),
            ('^', r'\textasciicircum{}'),
        ]

        for old, new in replacements:
            text = text.replace(old, new)

        return text

    def clean_text(self, text: str) -> str:
        """Clean and normalize text from PPTX."""
        if not text:
            return ""
        # Remove excessive whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    def rgb_to_latex_color(self, rgb_color) -> Optional[str]:
        """Convert python-pptx RGB color to LaTeX color definition."""
        if rgb_color is None:
            return None
        try:
            # rgb_color is an RGBColor object with red, green, blue attributes (0-255)
            r, g, b = rgb_color[0], rgb_color[1], rgb_color[2]
            # Return as RGB values for \textcolor[RGB]{r,g,b}{text}
            return f"{r},{g},{b}"
        except (TypeError, IndexError, AttributeError):
            return None

    def get_font_size_command(self, font_size_pt: float, base_size_pt: float = 11.0) -> Optional[str]:
        """Convert font size to LaTeX size command relative to base size."""
        if font_size_pt is None or font_size_pt <= 0:
            return None

        ratio = font_size_pt / base_size_pt

        # Map ratio to LaTeX size commands
        if ratio < 0.6:
            return "\\tiny"
        elif ratio < 0.75:
            return "\\scriptsize"
        elif ratio < 0.85:
            return "\\footnotesize"
        elif ratio < 0.95:
            return "\\small"
        elif ratio < 1.1:
            return None  # Normal size, no command needed
        elif ratio < 1.3:
            return "\\large"
        elif ratio < 1.6:
            return "\\Large"
        elif ratio < 2.0:
            return "\\LARGE"
        elif ratio < 2.5:
            return "\\huge"
        else:
            return "\\Huge"

    def format_run_to_latex(self, run, base_font_size: float = 11.0, para_font=None) -> str:
        """Convert a single text run to LaTeX with formatting.

        Args:
            run: The text run from python-pptx
            base_font_size: Base font size for relative sizing
            para_font: Paragraph-level font as fallback for inherited properties
        """
        text = run.text
        if not text:
            return ""

        # Escape LaTeX special characters
        escaped_text = self.escape_latex(text)

        # Track formatting wrappers
        prefix = ""
        suffix = ""

        # Check for font properties
        font = run.font

        # Bold - check run font, then paragraph font
        is_bold = font.bold
        if is_bold is None and para_font:
            is_bold = para_font.bold
        if is_bold:
            prefix += "\\textbf{"
            suffix = "}" + suffix

        # Italic - check run font, then paragraph font
        is_italic = font.italic
        if is_italic is None and para_font:
            is_italic = para_font.italic
        if is_italic:
            prefix += "\\textit{"
            suffix = "}" + suffix

        # Font size - check run font, then paragraph font
        font_size = None
        try:
            if font.size is not None:
                font_size = font.size.pt
            elif para_font and para_font.size is not None:
                font_size = para_font.size.pt

            if font_size is not None:
                size_cmd = self.get_font_size_command(font_size, base_font_size)
                if size_cmd:
                    prefix = "{" + size_cmd + " " + prefix
                    suffix = suffix + "}"
        except (AttributeError, TypeError):
            pass

        # Font color
        try:
            rgb = None
            # Check run-level color
            if font.color and font.color.type is not None:
                if font.color.rgb:
                    rgb = self.rgb_to_latex_color(font.color.rgb)
                elif font.color.theme_color is not None:
                    # Theme colors - map common ones
                    # This is a simplified mapping; actual theme colors depend on the template
                    pass

            if rgb and rgb != "0,0,0":
                prefix = f"\\textcolor[RGB]{{{rgb}}}" + "{" + prefix
                suffix = suffix + "}"
        except (AttributeError, TypeError):
            pass

        return prefix + escaped_text + suffix

    def needs_space_between(self, prev_text: str, curr_text: str) -> bool:
        """Determine if a space is needed between two text segments."""
        if not prev_text or not curr_text:
            return False

        prev_char = prev_text[-1]
        curr_char = curr_text[0]

        # Already has spacing
        if prev_char in ' \n\t' or curr_char in ' \n\t':
            return False

        # Don't add space after opening brackets or before closing ones
        if prev_char in '([{' or curr_char in ')]}':
            return False

        # Don't add space around colons that are part of times or ratios
        if prev_char == ':' and curr_char.isdigit():
            return False

        # Add space between: word ending and capital letter starting (new word)
        if prev_char.isalpha() and curr_char.isupper():
            return True

        # Add space between: lowercase/digit and uppercase
        if (prev_char.islower() or prev_char.isdigit()) and curr_char.isupper():
            return True

        # Add space between: letter and digit (except subscripts like H2O)
        if prev_char.isalpha() and curr_char.isdigit():
            # Check if it looks like a unit (m³, m², etc.) - no space needed
            if prev_text.endswith(('m', 'cm', 'km', 's', 'h', 'kg', 'g', 'l', 'L')):
                return False
            return True

        # Add space between: digit and letter (like "77.9m" -> "77.9 m")
        if prev_char.isdigit() and curr_char.isalpha():
            return True

        # Add space between: punctuation (except specific ones) and alphanumeric
        if prev_char in '.!?;' and curr_char.isalnum():
            return True

        return False

    def process_paragraph_with_formatting(self, para, base_font_size: float = 11.0) -> dict:
        """Process a paragraph preserving run-level formatting."""
        formatted_parts = []
        raw_text_parts = []

        # Get paragraph-level font as fallback for inherited properties
        para_font = None
        try:
            para_font = para.font
        except AttributeError:
            pass

        # Check if paragraph has runs
        runs = list(para.runs)

        if runs:
            for run in runs:
                if run.text:
                    formatted_parts.append(self.format_run_to_latex(run, base_font_size, para_font))
                    raw_text_parts.append(run.text)
        else:
            # Fallback: if no runs, use paragraph text directly
            if para.text:
                escaped = self.escape_latex(para.text)
                formatted_parts.append(escaped)
                raw_text_parts.append(para.text)

        # Join with space where needed
        formatted_text = ""
        for i, part in enumerate(formatted_parts):
            if i > 0:
                prev_raw = raw_text_parts[i-1]
                curr_raw = raw_text_parts[i]
                if self.needs_space_between(prev_raw, curr_raw):
                    formatted_text += " "
            formatted_text += part

        level = para.level if para.level is not None else 0

        return {
            'text': formatted_text,
            'raw_text': ''.join(raw_text_parts),
            'level': level,
            'is_bullet': para.level is not None and para.level >= 0
        }

    def extract_image(self, shape, slide_num: int, slide_width=None, slide_height=None) -> Optional[dict]:
        """Extract image from shape, convert to high-resolution JPG (220 dpi), and save to fig directory.

        Returns dict with 'filename' and 'width_ratio' (relative to slide width).
        """
        try:
            if hasattr(shape, 'image'):
                image = shape.image
                self.image_counter += 1
                filename = f"slide{slide_num:02d}_img{self.image_counter:03d}.jpg"
                filepath = self.fig_dir / filename

                if HAS_PIL:
                    # Convert to JPG at 220 dpi
                    img = Image.open(io.BytesIO(image.blob))
                    # Convert to RGB if necessary (e.g., PNG with transparency)
                    if img.mode in ('RGBA', 'LA', 'P'):
                        # Create white background for transparent images
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'P':
                            img = img.convert('RGBA')
                        background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                        img = background
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                    # Save as JPG with 220 dpi
                    img.save(filepath, 'JPEG', quality=95, dpi=(220, 220))
                else:
                    # Fallback: save original format if PIL not available
                    ext = image.ext
                    filename = f"slide{slide_num:02d}_img{self.image_counter:03d}.{ext}"
                    filepath = self.fig_dir / filename
                    with open(filepath, 'wb') as f:
                        f.write(image.blob)

                # Calculate width ratio relative to slide
                width_ratio = 0.8  # Default
                if slide_width and hasattr(shape, 'width') and shape.width:
                    width_ratio = min(0.95, shape.width / slide_width)
                    # Round to reasonable precision
                    width_ratio = round(width_ratio, 2)

                return {
                    'filename': filename,
                    'width_ratio': width_ratio
                }
        except Exception as e:
            print(f"Warning: Could not extract image: {e}")
        return None

    def extract_video(self, shape, slide_num: int) -> Optional[str]:
        """Extract video from shape and save to videos directory."""
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.MEDIA:
                # Access the media part
                if hasattr(shape, '_element'):
                    # Try to get video from the shape's media
                    for rel in shape.part.rels.values():
                        if "video" in rel.reltype.lower():
                            self.video_counter += 1
                            # Determine extension from content type
                            content_type = getattr(rel.target_part, 'content_type', 'video/mp4')
                            ext = content_type.split('/')[-1] if '/' in content_type else 'mp4'
                            filename = f"slide{slide_num:02d}_vid{self.video_counter:03d}.{ext}"
                            filepath = self.video_dir / filename

                            with open(filepath, 'wb') as f:
                                f.write(rel.target_part.blob)

                            return filename
        except Exception as e:
            print(f"Warning: Could not extract video: {e}")
        return None

    def process_text_frame(self, text_frame, base_font_size: float = 11.0) -> list:
        """Process text frame and return list of paragraphs with formatting info."""
        paragraphs = []
        for para in text_frame.paragraphs:
            # Use the new formatting-aware processor
            para_info = self.process_paragraph_with_formatting(para, base_font_size)
            # Only include if there's actual content
            if para_info['raw_text'].strip():
                paragraphs.append(para_info)
        return paragraphs

    def emu_to_cm(self, emu) -> float:
        """Convert EMU (English Metric Units) to centimeters."""
        if emu is None:
            return 0.0
        # 1 inch = 914400 EMU, 1 inch = 2.54 cm
        return emu / 914400 * 2.54

    def emu_to_textwidth(self, emu, slide_width) -> float:
        """Convert EMU to fraction of textwidth."""
        if emu is None or slide_width is None or slide_width == 0:
            return 0.8
        return min(0.95, emu / slide_width)

    def extract_all_images_from_shape(self, shape, slide_num: int, slide_width=None, slide_height=None, depth=0) -> list:
        """Recursively extract images from any shape type.

        Handles: Pictures, grouped shapes, placeholder images, fill images.
        Returns list of dicts with 'filename', 'width_ratio', 'left', 'top'.
        """
        images = []
        max_depth = 5  # Prevent infinite recursion

        if depth > max_depth:
            return images

        try:
            # Handle grouped shapes - recurse into children
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                for child_shape in shape.shapes:
                    images.extend(self.extract_all_images_from_shape(
                        child_shape, slide_num, slide_width, slide_height, depth + 1
                    ))
                return images

            # Handle regular pictures
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                img_info = self.extract_image(shape, slide_num, slide_width, slide_height)
                if img_info:
                    # Add position info
                    img_info['left'] = self.emu_to_textwidth(shape.left, slide_width) if slide_width else 0
                    img_info['top'] = self.emu_to_cm(shape.top)
                    images.append(img_info)
                return images

            # Handle placeholder shapes that might contain images
            if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                # Check if placeholder has an image
                if hasattr(shape, 'image'):
                    img_info = self.extract_image(shape, slide_num, slide_width, slide_height)
                    if img_info:
                        img_info['left'] = self.emu_to_textwidth(shape.left, slide_width) if slide_width else 0
                        img_info['top'] = self.emu_to_cm(shape.top)
                        images.append(img_info)
                return images

            # Handle shapes with fill images (background images on shapes)
            if hasattr(shape, 'fill'):
                try:
                    fill = shape.fill
                    if fill.type is not None and hasattr(fill, 'picture') and fill.picture:
                        # Extract the fill image
                        self.image_counter += 1
                        filename = f"slide{slide_num:02d}_img{self.image_counter:03d}.jpg"
                        filepath = self.fig_dir / filename

                        if HAS_PIL and hasattr(fill, '_fill'):
                            # Try to get the image blob from the fill
                            blob = fill._fill.blip.blob
                            if blob:
                                img = Image.open(io.BytesIO(blob))
                                if img.mode in ('RGBA', 'LA', 'P'):
                                    background = Image.new('RGB', img.size, (255, 255, 255))
                                    if img.mode == 'P':
                                        img = img.convert('RGBA')
                                    background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                                    img = background
                                elif img.mode != 'RGB':
                                    img = img.convert('RGB')
                                img.save(filepath, 'JPEG', quality=95, dpi=(220, 220))

                                width_ratio = self.emu_to_textwidth(shape.width, slide_width)
                                images.append({
                                    'filename': filename,
                                    'width_ratio': round(width_ratio, 2),
                                    'left': self.emu_to_textwidth(shape.left, slide_width) if slide_width else 0,
                                    'top': self.emu_to_cm(shape.top)
                                })
                except Exception:
                    pass  # Fill image extraction is optional

        except Exception as e:
            print(f"Warning: Could not extract image from shape: {e}")

        return images

    def shape_to_tikz(self, shape, slide_width, slide_height) -> Optional[str]:
        """Convert an AutoShape to TikZ code.

        Returns TikZ code string or None if not an AutoShape with content.
        """
        try:
            # Only handle AutoShapes
            if shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
                return None

            # Get shape text
            shape_text = ""
            if hasattr(shape, 'text_frame') and shape.text_frame:
                for para in shape.text_frame.paragraphs:
                    if para.text.strip():
                        shape_text += self.escape_latex(para.text.strip()) + " "
                shape_text = shape_text.strip()

            if not shape_text:
                return None  # Skip shapes without text

            # Get position and size in cm
            left_cm = self.emu_to_cm(shape.left)
            top_cm = self.emu_to_cm(shape.top)
            width_cm = self.emu_to_cm(shape.width)
            height_cm = self.emu_to_cm(shape.height)

            # Convert slide coordinates to TikZ coordinates
            # Slide coordinates: origin top-left, y increases downward
            # TikZ coordinates: we'll use origin top-left too for simplicity
            slide_width_cm = self.emu_to_cm(slide_width) if slide_width else 25.4  # ~10 inches default
            slide_height_cm = self.emu_to_cm(slide_height) if slide_height else 14.29  # ~5.6 inches default

            # Normalize to textwidth-based coordinates
            x_pos = left_cm / slide_width_cm
            y_pos = 1.0 - (top_cm / slide_height_cm)  # Flip y for TikZ

            # Determine shape style based on AutoShape type
            shape_style = "draw, rounded corners, fill=yellow!20"

            # Check if it's a callout type
            try:
                auto_shape_type = shape.auto_shape_type
                if auto_shape_type is not None:
                    type_name = str(auto_shape_type).lower()
                    if 'callout' in type_name:
                        shape_style = "draw, fill=yellow!30, rounded corners, drop shadow"
                    elif 'arrow' in type_name:
                        shape_style = "draw, ->, thick"
                    elif 'rectangle' in type_name or 'rect' in type_name:
                        shape_style = "draw, fill=blue!10"
                    elif 'oval' in type_name or 'ellipse' in type_name:
                        shape_style = "draw, ellipse, fill=green!10"
            except Exception:
                pass

            # Generate TikZ code
            # Position as fraction of textwidth, convert to actual position
            tikz_code = f"""\\begin{{tikzpicture}}[overlay, remember picture]
  \\node[{shape_style}, text width={width_cm:.1f}cm, align=center] at ([xshift={x_pos:.2f}\\textwidth, yshift={y_pos * 5:.1f}cm]current page.south west) {{{shape_text}}};
\\end{{tikzpicture}}"""

            return tikz_code

        except Exception as e:
            print(f"Warning: Could not convert shape to TikZ: {e}")
            return None

    def get_all_shape_content(self, slide, slide_num: int, slide_width, slide_height) -> dict:
        """Extract all content from a slide: title, text items, images, shapes.

        Returns dict with 'title', 'subtitle', 'content_items', 'images', 'videos', 'tikz_shapes'.
        """
        result = {
            'title': '',
            'subtitle': '',
            'content_items': [],
            'images': [],
            'videos': [],
            'tikz_shapes': []
        }

        # Track if we've found title/subtitle
        title_found = False
        subtitle_found = False

        for shape in slide.shapes:
            try:
                shape_type = shape.shape_type
                shape_processed = False

                # Debug: Print shape info for troubleshooting
                # print(f"  Shape: type={shape_type}, has_image={hasattr(shape, 'image')}, text='{shape.text[:30] if hasattr(shape, 'text') and shape.text else ''}'")

                # Handle grouped shapes - extract images and text recursively
                if shape_type == MSO_SHAPE_TYPE.GROUP:
                    result['images'].extend(
                        self.extract_all_images_from_shape(shape, slide_num, slide_width, slide_height)
                    )
                    # Also check for text in grouped shapes
                    for child_shape in shape.shapes:
                        if hasattr(child_shape, 'text_frame') and child_shape.text:
                            paragraphs = self.process_text_frame(child_shape.text_frame)
                            result['content_items'].extend(paragraphs)
                    continue

                # Check for image in ANY shape that has an 'image' attribute
                if hasattr(shape, 'image') and shape.image is not None:
                    img_info = self.extract_image(shape, slide_num, slide_width, slide_height)
                    if img_info:
                        result['images'].append(img_info)
                        shape_processed = True

                # Handle pictures specifically
                if shape_type == MSO_SHAPE_TYPE.PICTURE:
                    if not shape_processed:
                        img_info = self.extract_image(shape, slide_num, slide_width, slide_height)
                        if img_info:
                            result['images'].append(img_info)
                    continue

                # Handle videos/media
                if shape_type == MSO_SHAPE_TYPE.MEDIA:
                    vid_file = self.extract_video(shape, slide_num)
                    if vid_file:
                        result['videos'].append(vid_file)
                    continue

                # Handle AutoShapes (callouts, etc.) - convert to TikZ
                # But don't skip - the text content should also be processed
                if shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    tikz = self.shape_to_tikz(shape, slide_width, slide_height)
                    if tikz:
                        result['tikz_shapes'].append(tikz)
                    # Don't continue - fall through to process text content too
                    # But skip adding to content_items since TikZ already has the text
                    if tikz:
                        continue

                # Handle text frames (including remaining AutoShapes without TikZ)
                if hasattr(shape, 'text_frame') and shape.text and shape.text.strip():
                    # Check if this is a placeholder
                    is_placeholder_handled = False
                    try:
                        if shape.is_placeholder:
                            ph_type = shape.placeholder_format.type
                            # Title placeholder
                            if ph_type in [1, 3] and not title_found:  # TITLE or CENTER_TITLE
                                result['title'] = self.clean_text(shape.text)
                                title_found = True
                                is_placeholder_handled = True
                            # Subtitle placeholder (type 2) - this is often the frame subtitle
                            elif ph_type == 2 and not subtitle_found:
                                result['subtitle'] = self.clean_text(shape.text)
                                subtitle_found = True
                                is_placeholder_handled = True
                            # Body/Content placeholder - process as content
                            elif ph_type in [6, 7]:  # BODY, OBJECT
                                paragraphs = self.process_text_frame(shape.text_frame)
                                result['content_items'].extend(paragraphs)
                                is_placeholder_handled = True
                    except (ValueError, AttributeError):
                        pass

                    # Process as regular content if not a special placeholder
                    if not is_placeholder_handled:
                        paragraphs = self.process_text_frame(shape.text_frame)
                        result['content_items'].extend(paragraphs)

            except Exception as e:
                print(f"Warning: Could not process shape: {e}")
                continue

        return result

    def is_title_slide(self, slide, slide_index: int = 0) -> bool:
        """Detect if slide is likely a title slide."""
        layout_name = slide.slide_layout.name.lower() if slide.slide_layout.name else ""

        # Check layout name for title indicators
        if 'title' in layout_name and 'content' not in layout_name:
            return True

        # First slide with certain characteristics is likely a title slide
        if slide_index == 0:
            # Count text shapes and check for typical title slide patterns
            text_shapes = [s for s in slide.shapes if hasattr(s, 'text') and s.text.strip()]
            # Title slides typically have few text elements (title, subtitle, author)
            if len(text_shapes) <= 4:
                return True

        return False

    def is_thank_you_slide(self, slide) -> bool:
        """Detect if slide is likely a thank-you slide."""
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                text_lower = shape.text.lower()
                if any(phrase in text_lower for phrase in ['thank you', 'thanks', 'danke', 'vielen dank']):
                    return True
        return False

    def extract_title_info(self, slide) -> dict:
        """Extract title, subtitle, and author from title slide.

        Note: In PPTX files, the SUBTITLE placeholder typically contains the
        author name, while the BODY placeholder contains the chapter/subtitle text.
        We map these to LaTeX fields accordingly:
        - PPTX SUBTITLE placeholder -> LaTeX \\author{}
        - PPTX BODY placeholder -> LaTeX \\subtitle{}
        """
        info = {'title': '', 'subtitle': '', 'author': '', 'date': ''}
        other_texts = []  # Collect non-placeholder text

        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                text = self.clean_text(shape.text)
                if not text:
                    continue

                # Try to identify by placeholder type
                try:
                    if shape.is_placeholder:
                        ph_type = shape.placeholder_format.type
                        # Title placeholder types
                        if ph_type in [1, 3]:  # TITLE or CENTER_TITLE
                            info['title'] = text
                        elif ph_type == 2:  # PPTX SUBTITLE contains author name
                            info['author'] = text
                        elif ph_type == 6:  # BODY contains chapter/subtitle
                            if not info['subtitle']:
                                info['subtitle'] = text
                        elif ph_type == 13:  # DATE
                            info['date'] = text
                        continue
                except (ValueError, AttributeError):
                    pass

                # Collect other text shapes for potential subtitle detection
                other_texts.append(text)

        # If no subtitle found in placeholders, use first non-title/author text
        if not info['subtitle'] and other_texts:
            for text in other_texts:
                if text != info['title'] and text != info['author']:
                    info['subtitle'] = text
                    break

        if not info['title'] and other_texts:
            # First non-empty text as title if no title found yet
            info['title'] = other_texts[0]

        return info

    def slide_to_latex(self, slide, slide_num: int, slide_width=None, slide_height=None) -> str:
        """Convert a single slide to LaTeX frame using comprehensive shape extraction."""
        lines = []

        # Use the new comprehensive content extraction
        content = self.get_all_shape_content(slide, slide_num, slide_width, slide_height)

        title = content['title']
        subtitle = content['subtitle']
        content_items = content['content_items']
        images = content['images']
        videos = content['videos']
        tikz_shapes = content['tikz_shapes']

        # If no title found, try to use first content as title
        if not title and content_items:
            title = content_items[0].get('raw_text', content_items[0]['text'])
            content_items = content_items[1:]

        # Build the frame
        escaped_title = self.escape_latex(title) if title else "Slide"
        lines.append(f"\\begin{{frame}}{{{escaped_title}}}")

        # Add subtitle if present (as framesubtitle)
        if subtitle:
            escaped_subtitle = self.escape_latex(subtitle)
            lines.append(f"\\framesubtitle{{{escaped_subtitle}}}")

        # Add content as itemize if there are bullet points
        if content_items:
            has_bullets = any(item.get('is_bullet', False) for item in content_items)

            if has_bullets:
                lines.append("\\begin{itemize}")
                current_level = 0
                for item in content_items:
                    level = item.get('level', 0)
                    # Text is already formatted and escaped
                    text = item['text']

                    # Handle nesting
                    while level > current_level:
                        lines.append("  " * (current_level + 1) + "\\begin{itemize}")
                        current_level += 1
                    while level < current_level:
                        lines.append("  " * current_level + "\\end{itemize}")
                        current_level -= 1

                    lines.append("  " * (level + 1) + f"\\item {text}")

                # Close any open itemize environments
                while current_level > 0:
                    lines.append("  " * current_level + "\\end{itemize}")
                    current_level -= 1
                lines.append("\\end{itemize}")
            else:
                # Just paragraphs - text is already formatted
                for item in content_items:
                    text = item['text']
                    lines.append(f"{text}")
                    lines.append("")

        # Add images with appropriate sizing
        for img_info in images:
            filename = img_info['filename']
            width_ratio = img_info.get('width_ratio', 0.8)
            lines.append(f"\\begin{{center}}")
            lines.append(f"  \\includegraphics[width={width_ratio}\\textwidth]{{{filename}}}")
            lines.append(f"\\end{{center}}")

        # Add TikZ shapes (callouts, arrows, etc.)
        for tikz in tikz_shapes:
            lines.append("")
            lines.append("% AutoShape converted to TikZ")
            lines.append(tikz)

        # Add videos (using movie15 package from style)
        for vid in videos:
            lines.append(f"% Video: {vid}")
            lines.append(f"\\includemovie[poster, autoplay]{{0.8\\textwidth}}{{0.6\\textwidth}}{{videos/{vid}}}")

        lines.append("\\end{frame}")
        lines.append("")

        return '\n'.join(lines)

    def convert(self) -> str:
        """Convert the entire presentation to LaTeX."""
        # Patch the PPTX to handle missing content types
        patched_path = patch_pptx_content_types(str(self.input_path))
        temp_dir_to_cleanup = None
        if patched_path != str(self.input_path):
            temp_dir_to_cleanup = os.path.dirname(patched_path)

        try:
            prs = Presentation(patched_path)
        except Exception as e:
            if temp_dir_to_cleanup:
                shutil.rmtree(temp_dir_to_cleanup, ignore_errors=True)
            raise

        # Get slide dimensions for image sizing
        slide_width = prs.slide_width
        slide_height = prs.slide_height

        # Build LaTeX document
        latex_lines = [
            "\\documentclass[aspectratio=169]{beamer}",
            "\\usepackage{beamertheme}",
            "",
            "% Presentation metadata",
        ]

        # Extract title slide info from first slide
        title_info = {'title': 'Presentation', 'subtitle': '', 'author': '', 'date': '\\today'}
        if prs.slides:
            first_slide = prs.slides[0]
            if self.is_title_slide(first_slide, 0):
                title_info.update(self.extract_title_info(first_slide))

        # Extract subtitle and author separately
        subtitle = title_info['subtitle'] if title_info['subtitle'] else ''
        author_name = title_info['author'] if title_info['author'] else ''

        latex_lines.extend([
            f"\\title{{{self.escape_latex(title_info['title'])}}}",
            f"\\subtitle{{{self.escape_latex(subtitle)}}}",
            f"\\author{{{self.escape_latex(author_name)}}}",
            f"\\date{{{title_info['date'] if title_info['date'] else '\\\\today'}}}",
            "",
            "\\begin{document}",
            "",
            "% Title slide",
            "\\maketitle",
            "",
        ])

        # Process each slide
        thank_you_slide_content = None

        for i, slide in enumerate(prs.slides):
            slide_num = i + 1

            # Skip title slide (already handled with \maketitle)
            if i == 0 and self.is_title_slide(slide, i):
                continue

            # Handle thank you slide specially
            if self.is_thank_you_slide(slide):
                # Store for later - we'll add it at the end using \thankyou
                thank_you_slide_content = slide
                continue

            latex_lines.append(f"% Slide {slide_num}")
            latex_lines.append(self.slide_to_latex(slide, slide_num, slide_width, slide_height))

        # Add thank you slide at the end, using the author name from the title slide
        escaped_author = self.escape_latex(author_name) if author_name else "Author Name"
        if thank_you_slide_content:
            latex_lines.append("% Thank you slide")
            # Extract any text from the thank you slide for the message
            thank_text = "Thank you for your attention!"
            for shape in thank_you_slide_content.shapes:
                if hasattr(shape, 'text') and shape.text:
                    text = self.clean_text(shape.text)
                    if text and ('thank' in text.lower() or 'danke' in text.lower()):
                        thank_text = text
                        break

            latex_lines.append(f"\\thankyou{{{self.escape_latex(thank_text)}}}{{{escaped_author}}}{{Position}}{{email@example.com}}{{theme/logos/drop.png}}")
        else:
            # Add a default thank you slide
            latex_lines.append("% Thank you slide")
            latex_lines.append(f"\\thankyou{{Thank you for your attention!}}{{{escaped_author}}}{{Position}}{{email@example.com}}{{theme/logos/drop.png}}")

        latex_lines.extend([
            "",
            "\\end{document}",
        ])

        # Clean up temporary patched file
        if temp_dir_to_cleanup:
            shutil.rmtree(temp_dir_to_cleanup, ignore_errors=True)

        return '\n'.join(latex_lines)

    def save(self):
        """Convert and save the LaTeX output."""
        latex_content = self.convert()

        with open(self.output_path, 'w', encoding='utf-8') as f:
            f.write(latex_content)

        print(f"Converted: {self.input_path} -> {self.output_path}")
        print(f"  Images saved to: {self.fig_dir}/")
        print(f"  Videos saved to: {self.video_dir}/")


def main():
    parser = argparse.ArgumentParser(
        description="Convert PowerPoint presentations to LaTeX Beamer format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s                           # Convert all .pptx in pptx-input/ to beamer-slides.tex
  %(prog)s presentation.pptx         # Convert specific file to beamer-slides.tex
  %(prog)s input.pptx output.tex     # Specify custom output filename
        """
    )
    parser.add_argument(
        'input',
        nargs='?',
        help='Input PPTX file (default: process all in pptx-input/)'
    )
    parser.add_argument(
        'output',
        nargs='?',
        help='Output TEX file (default: beamer-slides.tex)'
    )
    parser.add_argument(
        '--fig-dir',
        default='fig',
        help='Directory for extracted images (default: fig)'
    )
    parser.add_argument(
        '--video-dir',
        default='videos',
        help='Directory for extracted videos (default: videos)'
    )

    args = parser.parse_args()

    # Determine input files
    if args.input:
        input_files = [Path(args.input)]
    else:
        # Process all PPTX files in pptx-input/
        input_dir = Path('pptx-input')
        if not input_dir.exists():
            print(f"Error: Input directory '{input_dir}' does not exist")
            sys.exit(1)

        input_files = list(input_dir.glob('*.pptx'))
        if not input_files:
            print(f"No .pptx files found in '{input_dir}'")
            sys.exit(1)

    # Process each file
    for input_path in input_files:
        if not input_path.exists():
            print(f"Error: File '{input_path}' does not exist")
            continue

        # Determine output path - always use beamer-slides.tex for TexStudio compatibility
        if args.output and len(input_files) == 1:
            output_path = Path(args.output)
        else:
            output_path = Path('beamer-slides.tex')

        try:
            converter = PPTXToLatexConverter(
                str(input_path),
                str(output_path),
                fig_dir=args.fig_dir,
                video_dir=args.video_dir
            )
            converter.save()
        except Exception as e:
            print(f"Error converting {input_path}: {e}")
            raise


if __name__ == '__main__':
    main()