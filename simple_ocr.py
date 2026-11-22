"""Extract plain text from an image or PDF using local Ollama.

Install: pip install PyMuPDF Pillow
Usage:   python simple_ocr.py scan.pdf
         python simple_ocr.py photo.jpg -o output.txt
"""

import argparse
import base64
import io
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

import pymupdf
from PIL import Image, ImageOps


MODEL = "AuditAid/PaddleOCR-VL-1.6-0.9B:latest"
URL = "http://localhost:11434/api/chat"


def page_images(path):
    """Yield PNG bytes, rendering PDFs one page at a time."""
    if path.suffix.lower() == ".pdf":
        with pymupdf.open(path) as document:
            for page in document:
                yield page.get_pixmap(dpi=150, alpha=False).tobytes("png")
    else:
        with Image.open(path) as image:
            image = ImageOps.exif_transpose(image).convert("RGB")
            buffer = io.BytesIO()
            image.save(buffer, format="PNG")
            yield buffer.getvalue()


def extract_text(path):
    """Return the OCR text for all pages as a single string."""
    pages = []
    for png in page_images(Path(path)):
        payload = {
            "model": MODEL,
            "messages": [{
                "role": "user",
                "content": (
                    "OCR: Extract all visible text in reading order. "
                    "Return only plain text in the original language. "
                    "No Markdown, formatting, commentary, or summary."
                ),
                "images": [base64.b64encode(png).decode("ascii")],
            }],
            "stream": False,
            "options": {
                "temperature": 0,
                "num_ctx": 16384,
                "num_predict": 8192,
            },
        }
        request = urllib.request.Request(
            URL,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(request, timeout=3600) as response:
            result = json.load(response)
        if result.get("error"):
            raise RuntimeError(result["error"])
        if result.get("done_reason") == "length":
            raise RuntimeError(
                "OCR reached the 8192-token output limit. "
                "Try cropping the page into smaller sections."
            )
        pages.append(result["message"]["content"].strip())
    return "\n\n".join(pages)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Image or PDF to read")
    parser.add_argument("-o", "--output", type=Path, help="Save UTF-8 text to this file")
    args = parser.parse_args()
    if args.output and args.input.resolve() == args.output.resolve():
        parser.error("Output must be different from the input file.")
    try:
        text = extract_text(args.input)
        if args.output:
            args.output.write_text(text, encoding="utf-8")
        else:
            sys.stdout.reconfigure(encoding="utf-8")
            print(text)
    except urllib.error.HTTPError as error:
        parser.exit(1, f"Ollama error: {error.read().decode('utf-8', errors='replace')}\n")
    except (OSError, ValueError, RuntimeError, KeyError) as error:
        parser.exit(1, f"Error: {error}\n")


if __name__ == "__main__":
    main()
