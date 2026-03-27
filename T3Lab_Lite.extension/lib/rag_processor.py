# -*- coding: utf-8 -*-
"""
RAG Processor for T3Lab Assistant

Handles PDF text extraction and image encoding so they can be sent
to the Claude API as context (Retrieval-Augmented Generation).

Supported inputs:
  - PDF  (.pdf)  → extracted plain text via .NET or fallback byte-scan
  - Image (.png, .jpg, .jpeg, .bmp, .gif, .webp) → base64-encoded bytes

Author: Tran Tien Thanh
"""

from __future__ import unicode_literals

import os
import base64

# ── .NET / IronPython imports ─────────────────────────────────────────────────
try:
    import clr
    clr.AddReference('System')
    clr.AddReference('System.IO')
    from System.IO import File, FileInfo
    HAS_DOTNET = True
except Exception:
    HAS_DOTNET = False

# ─── Supported extensions ─────────────────────────────────────────────────────

IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'}
PDF_EXT    = '.pdf'

SUPPORTED_EXTS = IMAGE_EXTS | {PDF_EXT}

# MIME map for Claude vision content blocks
_MIME_MAP = {
    '.jpg':  'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.png':  'image/png',
    '.gif':  'image/gif',
    '.webp': 'image/webp',
    '.bmp':  'image/png',   # convert via re-encode or just send as png
}

# Max sizes
MAX_IMAGE_BYTES = 5 * 1024 * 1024   # 5 MB
MAX_PDF_BYTES   = 20 * 1024 * 1024  # 20 MB
MAX_PDF_CHARS   = 12000             # truncate extracted text to this length


# ─── File validation ──────────────────────────────────────────────────────────

def is_supported(file_path):
    """Return True if file_path has a supported extension."""
    ext = os.path.splitext(file_path)[1].lower()
    return ext in SUPPORTED_EXTS


def is_image(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    return ext in IMAGE_EXTS


def is_pdf(file_path):
    return os.path.splitext(file_path)[1].lower() == PDF_EXT


# ─── PDF text extraction ──────────────────────────────────────────────────────

def extract_pdf_text(pdf_path):
    """Extract readable text from a PDF file.

    Tries (in order):
      1. iTextSharp (if available in Revit environment)
      2. Byte-level ASCII scan (fallback — catches plain-text streams)

    Returns:
        str: Extracted text (may be empty if PDF is scanned/image-only).
    """
    # Guard: file exists and not too large
    if not os.path.isfile(pdf_path):
        return u''
    try:
        size = os.path.getsize(pdf_path)
    except Exception:
        size = 0
    if size == 0 or size > MAX_PDF_BYTES:
        return u'[PDF quá lớn để xử lý — tối đa {} MB]'.format(MAX_PDF_BYTES // (1024 * 1024))

    text = _try_itextsharp(pdf_path)
    if not text:
        text = _fallback_byte_scan(pdf_path)

    if text:
        text = text.strip()
        if len(text) > MAX_PDF_CHARS:
            text = text[:MAX_PDF_CHARS] + u'\n[... nội dung bị cắt ngắn ...]'
    return text or u''


def _try_itextsharp(pdf_path):
    """Attempt to use iTextSharp for PDF text extraction (Revit may have it)."""
    try:
        clr.AddReference('itextsharp')
        from iTextSharp.text.pdf import PdfReader
        from iTextSharp.text.pdf.parser import PdfTextExtractor

        reader = PdfReader(pdf_path)
        pages  = []
        n = reader.NumberOfPages
        for i in range(1, min(n + 1, 51)):   # max 50 pages
            try:
                page_text = PdfTextExtractor.GetTextFromPage(reader, i)
                if page_text:
                    pages.append(page_text.strip())
            except Exception:
                pass
        reader.Close()
        return u'\n\n'.join(pages)
    except Exception:
        return u''


def _fallback_byte_scan(pdf_path):
    """Extract printable ASCII text streams from raw PDF bytes (simple heuristic)."""
    try:
        with open(pdf_path, 'rb') as f:
            raw = f.read(500000)  # read first 500 KB only

        # Decode with errors='ignore' to get printable chars
        text = raw.decode('latin-1', errors='ignore')

        # PDF text streams live between BT ... ET markers
        import re
        chunks = []
        for m in re.finditer(r'BT\s*(.*?)\s*ET', text, re.DOTALL):
            block = m.group(1)
            # Extract content inside parentheses: (some text)
            for t in re.findall(r'\(([^)]{1,500})\)', block):
                # Filter non-printable
                clean = ''.join(c for c in t if 32 <= ord(c) < 127 or c in '\n\r\t')
                if len(clean) > 2:
                    chunks.append(clean)

        if chunks:
            return u' '.join(chunks)

        # Last resort: extract any long printable runs
        runs = re.findall(r'[A-Za-z0-9\s.,;:!?/()\-]{20,}', text)
        return u' '.join(r.strip() for r in runs[:200])
    except Exception:
        return u''


# ─── Image encoding ───────────────────────────────────────────────────────────

def encode_image_base64(image_path):
    """Read an image file and return (base64_string, media_type).

    Returns:
        tuple(str, str): (base64_data, media_type) or (None, None) on failure.
    """
    if not os.path.isfile(image_path):
        return None, None
    try:
        size = os.path.getsize(image_path)
    except Exception:
        size = 0
    if size == 0 or size > MAX_IMAGE_BYTES:
        return None, None

    ext = os.path.splitext(image_path)[1].lower()
    media_type = _MIME_MAP.get(ext, 'image/jpeg')

    try:
        with open(image_path, 'rb') as f:
            data = f.read()
        b64 = base64.b64encode(data).decode('ascii')
        return b64, media_type
    except Exception:
        return None, None


# ─── Build context / message blocks for Claude API ───────────────────────────

def build_text_context(attached_files):
    """Build a plain-text context string from a list of file paths.

    Used for text-only Claude API calls (no vision).
    PDFs → extracted text; images → placeholder note.

    Returns:
        str: Context block to prepend to user message, or '' if nothing.
    """
    parts = []
    for path in attached_files:
        name = os.path.basename(path)
        if is_pdf(path):
            text = extract_pdf_text(path)
            if text:
                parts.append(
                    u'=== Nội dung PDF: {} ===\n{}\n=== Hết PDF ==='.format(name, text)
                )
            else:
                parts.append(u'[PDF "{}" không trích xuất được văn bản]'.format(name))
        elif is_image(path):
            parts.append(u'[Hình ảnh đính kèm: {} — xem phần vision bên dưới]'.format(name))
    return u'\n\n'.join(parts)


def build_vision_content_blocks(user_text, attached_files):
    """Build a Claude API `content` array supporting text + images.

    Structure follows Claude's multi-modal messages spec:
    [
      {"type": "text", "text": "..."},
      {"type": "image", "source": {"type": "base64", "media_type": "...", "data": "..."}},
      ...
    ]

    PDF content is injected as text blocks preceding the user text.

    Returns:
        list: Content blocks for the `messages[].content` field.
    """
    blocks = []

    # 1. PDF text blocks
    for path in attached_files:
        if is_pdf(path):
            name = os.path.basename(path)
            text = extract_pdf_text(path)
            if text:
                blocks.append({
                    "type": "text",
                    "text": u'=== Nội dung PDF: {} ===\n{}\n=== Hết PDF ==='.format(name, text)
                })
            else:
                blocks.append({
                    "type": "text",
                    "text": u'[PDF "{}" không trích xuất được văn bản — có thể là PDF scan]'.format(name)
                })

    # 2. Image blocks
    for path in attached_files:
        if is_image(path):
            b64, media_type = encode_image_base64(path)
            if b64:
                blocks.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type,
                        "data": b64
                    }
                })
            else:
                blocks.append({
                    "type": "text",
                    "text": u'[Hình ảnh "{}" quá lớn hoặc không đọc được]'.format(
                        os.path.basename(path))
                })

    # 3. User text (always last)
    blocks.append({"type": "text", "text": user_text})

    return blocks


def has_images(attached_files):
    """Return True if any file in the list is an image."""
    return any(is_image(p) for p in attached_files)


def summarize_attachments(attached_files):
    """Return a short human-readable list of attached file names."""
    return u', '.join(os.path.basename(p) for p in attached_files)
