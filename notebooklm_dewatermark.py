#!/usr/bin/env python3
"""NotebookLM PPT Watermark Remover

Remove the NotebookLM watermark (logo + text) from the bottom-right corner
of slides in exported PPTX files.

Usage:
    python3 notebooklm_dewatermark.py input.pptx
    python3 notebooklm_dewatermark.py input.pptx -o output.pptx
    python3 notebooklm_dewatermark.py *.pptx
"""

import argparse
import io
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Watermark region parameters (offset from bottom-right corner)
WATERMARK_WIDTH = 175
WATERMARK_HEIGHT = 30
MARGIN_RIGHT = 3
MARGIN_BOTTOM = 3
# Extra height for sampling strip above watermark
SAMPLE_PADDING = 5


def remove_watermark(img: Image.Image) -> Image.Image:
    """Remove the NotebookLM watermark from the bottom-right corner.

    Strategy: copy a strip of pixels from directly above the watermark area
    and paste it over the watermark. This naturally continues any background
    texture or gradient.
    """
    result = img.copy()
    w, h = result.size

    wm_left = w - WATERMARK_WIDTH - MARGIN_RIGHT
    wm_top = h - WATERMARK_HEIGHT - MARGIN_BOTTOM
    wm_right = w - MARGIN_RIGHT
    wm_bottom = h - MARGIN_BOTTOM

    sample_top = max(0, wm_top - WATERMARK_HEIGHT - SAMPLE_PADDING)
    sample_bottom = wm_top

    sample_strip = result.crop((wm_left, sample_top, wm_right, sample_bottom))

    target_height = wm_bottom - wm_top
    sample_strip = sample_strip.resize(
        (wm_right - wm_left, target_height), Image.LANCZOS
    )
    result.paste(sample_strip, (wm_left, wm_top))

    return result


def process_pptx(input_path: Path, output_path: Path) -> int:
    """Process a single PPTX file. Returns the number of slides processed."""
    prs = Presentation(str(input_path))
    processed = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue

            img_blob = shape.image.blob
            content_type = shape.image.content_type
            img = Image.open(io.BytesIO(img_blob))

            cleaned = remove_watermark(img)

            buf = io.BytesIO()
            fmt = "PNG" if "png" in content_type else "JPEG"
            cleaned.save(buf, format=fmt)
            buf.seek(0)

            # Replace image data in the PPTX package
            pic = shape._element
            ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            ns_r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
            blip = pic.find(f".//{ns_a}blip")
            r_id = blip.get(f"{ns_r}embed")
            image_part = slide.part.rels[r_id].target_part
            image_part._blob = buf.read()

            processed += 1

    prs.save(str(output_path))
    return processed


def main():
    parser = argparse.ArgumentParser(
        description="Remove NotebookLM watermark from PPTX slides"
    )
    parser.add_argument(
        "input", nargs="+", type=Path, help="Input PPTX file(s)"
    )
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Output file path (only for single file input)"
    )
    args = parser.parse_args()

    for input_path in args.input:
        if not input_path.exists():
            print(f"Error: file not found: {input_path}", file=sys.stderr)
            continue

        if input_path.suffix.lower() != ".pptx":
            print(f"Skipping non-PPTX file: {input_path}", file=sys.stderr)
            continue

        if args.output and len(args.input) == 1:
            output_path = args.output
        else:
            output_path = input_path.with_stem(input_path.stem + "_clean")

        count = process_pptx(input_path, output_path)
        print(f"Done: {input_path.name} -> {output_path.name} ({count} slides)")


if __name__ == "__main__":
    main()
