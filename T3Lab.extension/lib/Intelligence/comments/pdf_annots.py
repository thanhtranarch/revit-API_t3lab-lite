# -*- coding: utf-8 -*-
"""
PDF annotation extractor — pure Python primary path, iTextSharp as an
opportunistic accelerator (the DLL is NOT bundled with the extension).

Handles the shapes Bluebeam/Acrobat actually write:
  - literal strings with escapes and nested parens
  - <hex> strings, UTF-16BE with BOM FEFF (Bluebeam default for /Contents)
  - /Annots arrays inline on the page object or as an indirect object
  - objects packed in /ObjStm streams (FlateDecode via zlib, with a .NET
    DeflateStream fallback; failure marks the result "partial", never
    crashes)

Returned record schema (see plan §3.3):
    {"id","page","subtype","author","subject","content","rect",
     "modified","color","source"}

CPython-3 testable; IronPython 2.7 safe.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import os
import re

MARKUP_SUBTYPES = frozenset([
    'Text', 'FreeText', 'Square', 'Circle', 'Highlight', 'StrikeOut',
    'Underline', 'Squiggly', 'Ink', 'Polygon', 'PolyLine', 'Caret',
    'Stamp', 'Line',
])

_MAX_PDF_BYTES = 20 * 1024 * 1024

_OBJ_RE = re.compile(r'(\d+)\s+(\d+)\s+obj\b(.*?)endobj', re.DOTALL)
_SUBTYPE_RE = re.compile(r'/Subtype\s*/(\w+)')
_RECT_RE = re.compile(r'/Rect\s*\[([^\]]*)\]')
_COLOR_RE = re.compile(r'/C\s*\[([^\]]*)\]')
_ANNOTS_INLINE_RE = re.compile(r'/Annots\s*\[([^\]]*)\]')
_ANNOTS_REF_RE = re.compile(r'/Annots\s+(\d+)\s+\d+\s+R')
_REF_RE = re.compile(r'(\d+)\s+\d+\s+R')

_QUICK_MARKUP_RE = re.compile(
    br'/Subtype\s*/(Text|FreeText|Square|Circle|Highlight|StrikeOut'
    br'|Underline|Squiggly|Ink|Polygon|PolyLine|Caret|Stamp|Line)\b')


# ─── public API ───────────────────────────────────────────────────────────────

def extract_annotations(pdf_path):
    """Extract markup annotations. Returns (records, partial_flag)."""
    recs = _via_itextsharp(pdf_path)
    if recs:
        return recs, False
    return _via_rawscan(pdf_path)


def has_annotations(pdf_path):
    """Cheap byte-level probe: does this PDF carry visible markup
    annotations? (Objects hidden in compressed streams may be missed —
    a keyword-routed request still gets the full parse.)"""
    try:
        size = os.path.getsize(pdf_path)
        if size == 0 or size > _MAX_PDF_BYTES:
            return False
        with open(pdf_path, 'rb') as f:
            raw = f.read()
        return _QUICK_MARKUP_RE.search(raw) is not None
    except Exception:
        return False


def extract_page_texts(pdf_path, max_pages=20):
    """[(page_no, text)] — thin wrapper over rag_processor."""
    try:
        from Intelligence import rag_processor
        return rag_processor.extract_pdf_pages(pdf_path, max_pages=max_pages)
    except Exception:
        return []


# ─── iTextSharp accelerator (in-Revit only, best effort) ─────────────────────

def _via_itextsharp(pdf_path):
    try:
        import clr
        clr.AddReference('itextsharp')
        from iTextSharp.text.pdf import PdfReader, PdfName

        reader = PdfReader(pdf_path)
        recs = []
        n = 0
        for i in range(1, reader.NumberOfPages + 1):
            page = reader.GetPageN(i)
            annots = page.GetAsArray(PdfName.ANNOTS)
            if annots is None:
                continue
            for j in range(annots.Size):
                annot = annots.GetAsDict(j)
                if annot is None:
                    continue
                subtype = annot.GetAsName(PdfName.SUBTYPE)
                sub = str(subtype)[1:] if subtype else ''
                if sub not in MARKUP_SUBTYPES:
                    continue
                n += 1
                recs.append({
                    'id': 'a{}'.format(n),
                    'page': i,
                    'subtype': sub,
                    'author': _its_str(annot, PdfName.T),
                    'subject': _its_str(annot, PdfName.SUBJECT),
                    'content': _its_str(annot, PdfName.CONTENTS),
                    'rect': _its_rect(annot),
                    'modified': _its_str(annot, PdfName.M),
                    'color': [],
                    'source': 'itextsharp',
                })
        reader.Close()
        return recs
    except Exception:
        return []


def _its_str(annot, key):
    try:
        s = annot.GetAsString(key)
        return s.ToUnicodeString() if s is not None else ''
    except Exception:
        return ''


def _its_rect(annot):
    try:
        from iTextSharp.text.pdf import PdfName
        arr = annot.GetAsArray(PdfName.RECT)
        if arr is None:
            return []
        return [float(str(arr.GetAsNumber(i))) for i in range(arr.Size)]
    except Exception:
        return []


# ─── pure-Python raw parser ───────────────────────────────────────────────────

def _via_rawscan(pdf_path):
    """Parse the raw PDF object graph. Returns (records, partial_flag)."""
    try:
        size = os.path.getsize(pdf_path)
        if size == 0 or size > _MAX_PDF_BYTES:
            return [], True
        with open(pdf_path, 'rb') as f:
            raw = f.read()
    except Exception:
        return [], True

    src = raw.decode('latin-1')
    obj_map = {}      # objnum -> source string
    obj_order = []    # appearance order
    partial = False

    for m in _OBJ_RE.finditer(src):
        num = int(m.group(1))
        body = m.group(3)
        if num not in obj_map:
            obj_map[num] = body
            obj_order.append(num)
        # objects packed inside object streams
        if re.search(r'/Type\s*/ObjStm\b', body):
            ok = _explode_objstm(body, obj_map, obj_order)
            if not ok:
                partial = True

    # page objects in document order (page-tree walk with fallback)
    page_nums = _ordered_pages(obj_map, obj_order)

    recs = []
    n = 0
    for page_no, pnum in enumerate(page_nums, start=1):
        body = obj_map.get(pnum, '')
        annot_refs = []
        m_inline = _ANNOTS_INLINE_RE.search(body)
        if m_inline:
            annot_refs = [int(r) for r in
                          _REF_RE.findall(m_inline.group(1))]
        else:
            m_ref = _ANNOTS_REF_RE.search(body)
            if m_ref:
                arr_body = obj_map.get(int(m_ref.group(1)), '')
                annot_refs = [int(r) for r in _REF_RE.findall(arr_body)]
        for ref in annot_refs:
            rec = _parse_annot(obj_map.get(ref, ''), page_no)
            if rec is not None:
                n += 1
                rec['id'] = 'a{}'.format(n)
                recs.append(rec)
    return recs, partial


def _explode_objstm(body, obj_map, obj_order):
    """Inflate a /Type/ObjStm stream and add its packed objects."""
    try:
        m_n = re.search(r'/N\s+(\d+)', body)
        m_first = re.search(r'/First\s+(\d+)', body)
        m_stream = re.search(r'stream\r?\n', body)
        if not (m_n and m_first and m_stream):
            return False
        if not re.search(r'/Filter\s*/FlateDecode\b', body):
            return False
        start = m_stream.end()
        end = body.rfind('endstream')
        if end <= start:
            return False
        data = body[start:end].encode('latin-1')
        # trailing EOL before 'endstream' is not stream data
        data = data.rstrip(b'\r\n')
        inflated = _inflate(data)
        if inflated is None:
            return False
        text = inflated.decode('latin-1')
        count = int(m_n.group(1))
        first = int(m_first.group(1))
        header = text[:first].split()
        pairs = []
        for i in range(0, min(len(header), count * 2) - 1, 2):
            pairs.append((int(header[i]), int(header[i + 1])))
        for idx, (objnum, off) in enumerate(pairs):
            begin = first + off
            stop = first + pairs[idx + 1][1] if idx + 1 < len(pairs) else len(text)
            if objnum not in obj_map:
                obj_map[objnum] = text[begin:stop]
                obj_order.append(objnum)
        return True
    except Exception:
        return False


def _inflate(data):
    """zlib inflate with a .NET DeflateStream fallback. None on failure."""
    try:
        import zlib
        return zlib.decompress(data)
    except ImportError:
        pass
    except Exception:
        return None
    try:
        import clr
        clr.AddReference('System')
        from System.IO import MemoryStream
        from System.IO.Compression import DeflateStream, CompressionMode
        # skip the 2-byte zlib header for raw-deflate reading
        payload = data[2:] if data[:1] == b'\x78' else data
        ms = MemoryStream(bytearray(payload))
        ds = DeflateStream(ms, CompressionMode.Decompress)
        out = MemoryStream()
        buf = bytearray(4096)
        while True:
            read = ds.Read(buf, 0, 4096)
            if read <= 0:
                break
            out.Write(buf, 0, read)
        return bytes(bytearray(out.ToArray()))
    except Exception:
        return None


def _ordered_pages(obj_map, obj_order):
    """Page object numbers in reading order (tree walk, else file order)."""
    kids_of = {}
    root_pages = None
    for num in obj_order:
        body = obj_map[num]
        if re.search(r'/Type\s*/Catalog\b', body):
            m = re.search(r'/Pages\s+(\d+)\s+\d+\s+R', body)
            if m:
                root_pages = int(m.group(1))
        if re.search(r'/Type\s*/Pages\b', body):
            m = re.search(r'/Kids\s*\[([^\]]*)\]', body)
            if m:
                kids_of[num] = [int(r) for r in _REF_RE.findall(m.group(1))]

    def _is_page(num):
        body = obj_map.get(num, '')
        return re.search(r'/Type\s*/Page\b(?!s)', body) is not None

    if root_pages is not None:
        ordered = []
        stack = [root_pages]
        seen = set()
        while stack:
            cur = stack.pop(0)
            if cur in seen:
                continue
            seen.add(cur)
            if _is_page(cur):
                ordered.append(cur)
            else:
                stack = kids_of.get(cur, []) + stack
        if ordered:
            return ordered
    return [num for num in obj_order if _is_page(num)]


def _parse_annot(body, page_no):
    """One annotation object → record dict, or None if not a markup."""
    if not body:
        return None
    m = _SUBTYPE_RE.search(body)
    if not m or m.group(1) not in MARKUP_SUBTYPES:
        return None
    rect = []
    m_rect = _RECT_RE.search(body)
    if m_rect:
        try:
            rect = [float(x) for x in m_rect.group(1).split()]
        except Exception:
            rect = []
    color = []
    m_col = _COLOR_RE.search(body)
    if m_col:
        try:
            color = [float(x) for x in m_col.group(1).split()]
        except Exception:
            color = []
    return {
        'id': '',
        'page': page_no,
        'subtype': m.group(1),
        'author': _get_string(body, '/T'),
        'subject': _get_string(body, '/Subj'),
        'content': _get_string(body, '/Contents'),
        'rect': rect,
        'modified': _get_string(body, '/M'),
        'color': color,
        'source': 'rawscan',
    }


# ─── PDF string decoding ──────────────────────────────────────────────────────

_ESC_MAP = {'n': '\n', 'r': '\r', 't': '\t', 'b': '\b', 'f': '\f',
            '(': '(', ')': ')', '\\': '\\'}


def _get_string(body, key):
    """Value of `key` when it is a PDF string — literal or hex."""
    for m in re.finditer(re.escape(key) + r'(?![A-Za-z])\s*', body):
        pos = m.end()
        if pos >= len(body):
            continue
        ch = body[pos]
        if ch == '(':
            return _decode_pdf_text(_scan_literal(body, pos))
        if ch == '<' and body[pos:pos + 2] != '<<':
            m_hex = re.match(r'<([0-9A-Fa-f\s]*)>', body[pos:])
            if m_hex:
                hex_str = re.sub(r'\s+', '', m_hex.group(1))
                if len(hex_str) % 2:
                    hex_str += '0'
                try:
                    chars = ''.join(
                        chr(int(hex_str[i:i + 2], 16))
                        for i in range(0, len(hex_str), 2))
                except Exception:
                    return ''
                return _decode_pdf_text(chars)
    return ''


def _scan_literal(src, start):
    """Scan a (literal) PDF string starting at src[start] == '('. Returns
    the raw (unescaped) char string."""
    depth = 0
    i = start
    out = []
    ln = len(src)
    while i < ln:
        ch = src[i]
        if ch == '\\':
            nxt = src[i + 1] if i + 1 < ln else ''
            if nxt in _ESC_MAP:
                out.append(_ESC_MAP[nxt])
                i += 2
                continue
            m = re.match(r'[0-7]{1,3}', src[i + 1:i + 4])
            if m:
                out.append(chr(int(m.group(0), 8) % 256))
                i += 1 + len(m.group(0))
                continue
            i += 2   # unknown escape / line continuation
            continue
        if ch == '(':
            depth += 1
            if depth > 1:
                out.append(ch)
            i += 1
            continue
        if ch == ')':
            depth -= 1
            if depth == 0:
                return ''.join(out)
            out.append(ch)
            i += 1
            continue
        out.append(ch)
        i += 1
    return ''.join(out)


def _decode_pdf_text(chars):
    """PDF text-string decoding: UTF-16BE when BOM FEFF, else latin-1-ish."""
    if not chars:
        return ''
    if len(chars) >= 2 and ord(chars[0]) == 0xFE and ord(chars[1]) == 0xFF:
        try:
            raw = ''.join(chars[2:]).encode('latin-1')
            return raw.decode('utf-16-be', 'replace')
        except Exception:
            return ''.join(chars[2:])
    return chars.strip()
