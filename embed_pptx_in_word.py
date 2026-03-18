"""
embed_pptx_in_word.py
---------------------
Embeds a PPTX file as an OLE object into an existing Word document,
replacing every paragraph that contains only the text "embedded_pptx".

Dependencies:
    pip install spire.doc Pillow

Usage:
    from embed_pptx_in_word import embed_pptx_in_word

    # From file paths
    result_bytes = embed_pptx_in_word(
        word_input="path/to/document.docx",
        pptx_input="path/to/presentation.pptx",
    )

    # From raw bytes
    with open("doc.docx", "rb") as f:
        word_bytes = f.read()
    with open("pres.pptx", "rb") as f:
        pptx_bytes = f.read()

    result_bytes = embed_pptx_in_word(
        word_input=word_bytes,
        pptx_input=pptx_bytes,
        pptx_display_name="MyPresentation.pptx",  # shown under icon
    )

    with open("output.docx", "wb") as f:
        f.write(result_bytes)
"""

import os
import io
import tempfile
import shutil
from typing import Union

from PIL import Image, ImageDraw, ImageFont
from spire.doc import Document, DocPicture, OleObjectType, FileFormat


# ---------------------------------------------------------------------------
# Icon generator
# ---------------------------------------------------------------------------

def _generate_pptx_icon(filename: str, output_path: str, width: int = 120, height: int = 140) -> None:
    """
    Generates a PowerPoint-style icon PNG with the filename label below it.

    Args:
        filename:    Display name shown under the icon (truncated if > 16 chars).
        output_path: Where to save the PNG file.
        width:       Icon width in pixels  (default 120).
        height:      Icon height in pixels (default 140).
    """
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    body_margin = 8
    fold_size   = 22
    body_x0, body_y0 = body_margin, 10
    body_x1, body_y1 = width - body_margin, height - 30

    # File body (white with light border)
    draw.polygon(
        [
            (body_x0, body_y0),
            (body_x1 - fold_size, body_y0),
            (body_x1, body_y0 + fold_size),
            (body_x1, body_y1),
            (body_x0, body_y1),
        ],
        fill=(255, 255, 255),
        outline=(180, 180, 180),
    )

    # Folded corner
    draw.polygon(
        [
            (body_x1 - fold_size, body_y0),
            (body_x1, body_y0 + fold_size),
            (body_x1 - fold_size, body_y0 + fold_size),
        ],
        fill=(210, 210, 210),
        outline=(180, 180, 180),
    )

    # PowerPoint red/orange banner
    banner_y0 = body_y0 + 20
    banner_y1 = banner_y0 + 32
    draw.rectangle([body_x0, banner_y0, body_x1, banner_y1], fill=(209, 52, 28))

    # Try to load a real font; fall back to default
    try:
        font_big   = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
        font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 9)
    except OSError:
        font_big   = ImageFont.load_default()
        font_small = ImageFont.load_default()

    # "PPT" label inside banner
    draw.text((body_x0 + 8, banner_y0 + 5), "PPT", fill=(255, 255, 255), font=font_big)

    # Filename below the icon body (truncate if needed)
    label = filename if len(filename) <= 16 else filename[:13] + "..."
    bbox  = draw.textbbox((0, 0), label, font=font_small)
    text_w = bbox[2] - bbox[0]
    draw.text(((width - text_w) // 2, body_y1 + 4), label, fill=(50, 50, 50), font=font_small)

    img.save(output_path, "PNG")


# ---------------------------------------------------------------------------
# Main function
# ---------------------------------------------------------------------------

def embed_pptx_in_word(
    word_input: Union[str, bytes],
    pptx_input: Union[str, bytes],
    pptx_display_name: str = "presentation.pptx",
    placeholder: str = "embedded_pptx",
    icon_width: int = 50,
    icon_height: int = 58,
) -> bytes:
    """
    Embeds a PPTX file as an OLE object into an existing Word document,
    replacing every paragraph whose text is exactly ``placeholder``.

    Args:
        word_input:        Path (str) or raw bytes of the input .docx file.
        pptx_input:        Path (str) or raw bytes of the .pptx file to embed.
        pptx_display_name: Filename shown under the icon inside the Word doc.
                           Ignored when pptx_input is a file path (the real
                           filename is used instead).
        placeholder:       The exact paragraph text to search for and replace.
                           Default: ``"embedded_pptx"``.
        icon_width:        Width  of the OLE icon in the document (points).
        icon_height:       Height of the OLE icon in the document (points).

    Returns:
        bytes: The modified .docx file content.

    Raises:
        ValueError: If no placeholder paragraph is found in the document.
        FileNotFoundError: If a supplied file path does not exist.
    """
    tmp_dir = tempfile.mkdtemp(prefix="embed_pptx_")
    try:
        # ------------------------------------------------------------------ #
        # 1. Resolve word file → temp path                                    #
        # ------------------------------------------------------------------ #
        if isinstance(word_input, (bytes, bytearray)):
            word_path = os.path.join(tmp_dir, "input.docx")
            with open(word_path, "wb") as fh:
                fh.write(word_input)
        else:
            word_path = str(word_input)
            if not os.path.isfile(word_path):
                raise FileNotFoundError(f"Word file not found: {word_path}")

        # ------------------------------------------------------------------ #
        # 2. Resolve pptx file → temp path + display name                    #
        # ------------------------------------------------------------------ #
        if isinstance(pptx_input, (bytes, bytearray)):
            # Use supplied display name (or fallback)
            display_name = pptx_display_name or "presentation.pptx"
            pptx_path = os.path.join(tmp_dir, display_name)
            with open(pptx_path, "wb") as fh:
                fh.write(pptx_input)
        else:
            pptx_path = str(pptx_input)
            if not os.path.isfile(pptx_path):
                raise FileNotFoundError(f"PPTX file not found: {pptx_path}")
            display_name = os.path.basename(pptx_path)

        # ------------------------------------------------------------------ #
        # 3. Generate the icon PNG with the filename label                    #
        # ------------------------------------------------------------------ #
        icon_path = os.path.join(tmp_dir, "pptx_icon.png")
        _generate_pptx_icon(display_name, icon_path, width=120, height=140)

        # ------------------------------------------------------------------ #
        # 4. Open the Word document                                           #
        # ------------------------------------------------------------------ #
        doc = Document()
        doc.LoadFromFile(word_path)

        # ------------------------------------------------------------------ #
        # 5. Find and replace all placeholder paragraphs                      #
        # ------------------------------------------------------------------ #

        replacements = 0
        
        for sec_idx in range(doc.Sections.Count):
            section = doc.Sections.get_Item(sec_idx)
            for para_idx in range(section.Paragraphs.Count):
                para = section.Paragraphs.get_Item(para_idx)
        
                if placeholder not in para.Text:
                    continue
        
                # Walk through child objects and remove only the placeholder text run
                children = para.ChildObjects
                for i in range(children.Count - 1, -1, -1):
                    child = children.get_Item(i)
                    child_text = getattr(child, 'Text', '')
                    if placeholder in child_text:
                        # Replace this child with OLE object
                        children.RemoveAt(i)
                        picture = DocPicture(doc)
                        picture.LoadImage(icon_path)
                        picture.Width  = icon_width
                        picture.Height = icon_height
                        para.AppendOleObject(
                            pptx_path,
                            picture,
                            OleObjectType.PowerPointPresentation,
                        )
                        replacements += 1
                        break

        if replacements == 0:
            raise ValueError(
                f'No paragraph with text "{placeholder}" found in the document. '
                "Make sure the placeholder is on its own paragraph."
            )

        # ------------------------------------------------------------------ #
        # 6. Save to bytes and return                                         #
        # ------------------------------------------------------------------ #
        output_path = os.path.join(tmp_dir, "output.docx")
        doc.SaveToFile(output_path, FileFormat.Docx2013)
        doc.Close()

        with open(output_path, "rb") as fh:
            return fh.read()

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# CLI / quick test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python embed_pptx_in_word.py <input.docx> <file.pptx> [output.docx]")
        sys.exit(1)

    word_arg  = sys.argv[1]
    pptx_arg  = sys.argv[2]
    out_arg   = sys.argv[3] if len(sys.argv) > 3 else "output_with_pptx.docx"

    result = embed_pptx_in_word(word_arg, pptx_arg)
    with open(out_arg, "wb") as f:
        f.write(result)

    print(f"Done → {out_arg}")
