# -*- coding: utf-8 -*-
"""
pdf_text — pure-Python PDF text extraction (no external packages).

Why this exists
---------------
rag_processor's old fallback scanned raw PDF bytes for `BT ... ET` and
parenthesised strings. In every modern PDF the content streams are
FlateDecode-COMPRESSED, so those markers are not present in the raw bytes and
the "last resort" regex matched random compressed binary instead — the RAG
index and the folder digests filled up with garbage like
`/Type/Catalog/Pages 2 0 R` and `|* Q9nG%6bHiw?`.

This module does the real thing:
  1. scans the indirect objects, following `stream`/`endstream` correctly;
  2. expands PDF 1.5+ object streams (`/Type/ObjStm`) so page objects inside
     them are visible;
  3. walks the page tree (/Root → /Pages → /Kids) for true page order;
  4. inflates each page's content stream(s);
  5. parses the text operators (Tj, TJ, ', ") and maps the bytes through the
     font's /ToUnicode CMap when present — required for subset-embedded fonts
     whose codes are glyph indices (\\x01\\x02…), not ASCII.

Falls back gracefully: any page that cannot be parsed is simply skipped, and
extract_pages() returns [] rather than raising, so callers keep their existing
behaviour when a PDF is scanned/image-only or malformed.

Python 2.7 (IronPython) and Python 3 compatible: all byte handling goes
through bytearray so indexing yields ints on both.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import re
import zlib

MAX_PAGES = 400
_MAX_STREAM = 12 * 1024 * 1024      # inflated-stream sanity cap

# `br''` (not `rb''`) — the only byte-literal prefix valid on BOTH py2.7 and py3
_RE_OBJ = re.compile(br'(\d+)\s+(\d+)\s+obj\b')
_RE_TYPE_PAGE = re.compile(br'/Type\s*/Page\b(?!s)')
_RE_KIDS = re.compile(br'/Kids\s*\[(.*?)\]', re.S)
_RE_REF = re.compile(br'(\d+)\s+(\d+)\s+R\b')
_RE_CONTENTS = re.compile(br'/Contents\s*(?:(\d+)\s+(\d+)\s+R|\[(.*?)\])', re.S)
# `/Length 6 0 R` is an INDIRECT reference, not a length of 6 — the negative
# lookahead makes those fall through to the endstream search instead of
# truncating the stream to a handful of bytes.
_RE_LENGTH = re.compile(br'/Length\s+(\d+)(?!\s+\d+\s+R)')
_RE_LENGTH_REF = re.compile(br'/Length\s+(\d+)\s+(\d+)\s+R')
_RE_FONT_BLOCK = re.compile(br'/Font\s*<<(.*?)>>', re.S)
_RE_FONT_REF = re.compile(br'/Font\s+(\d+)\s+(\d+)\s+R')
_RE_FONT_ENTRY = re.compile(br'/([A-Za-z0-9_.\-]+)\s+(\d+)\s+(\d+)\s+R')
_RE_TOUNICODE = re.compile(br'/ToUnicode\s+(\d+)\s+(\d+)\s+R')
_RE_RESOURCES_REF = re.compile(br'/Resources\s+(\d+)\s+(\d+)\s+R')


def _inflate(data):
    """zlib-inflate, tolerating raw-deflate and trailing junk. b'' on failure."""
    for wbits in (15, -15, 47):
        try:
            out = zlib.decompressobj(wbits).decompress(bytes(data), _MAX_STREAM)
            if out:
                return out
        except Exception:
            continue
    return b''


class _Pdf(object):
    """Minimal object store for one PDF file."""

    def __init__(self, raw):
        self.raw = bytearray(raw)
        self.objs = {}        # obj_num -> (body_bytes, stream_bytes_or_None)
        self._scan_objects()
        self._expand_object_streams()

    # ── object scanning ───────────────────────────────────────────────────

    def _scan_objects(self):
        raw = bytes(self.raw)
        starts = [(m.start(), int(m.group(1)), m.end())
                  for m in _RE_OBJ.finditer(raw)]
        for i, (pos, num, body_at) in enumerate(starts):
            end = starts[i + 1][0] if i + 1 < len(starts) else len(raw)
            chunk = raw[body_at:end]
            sm = re.search(br'\bstream\r?\n', chunk)
            if not sm:
                self.objs[num] = (chunk, None)
                continue
            body = chunk[:sm.start()]
            data_at = sm.end()
            # prefer the declared /Length, fall back to the endstream marker
            stream = None
            lm = _RE_LENGTH.search(body)
            if lm:
                try:
                    n = int(lm.group(1))
                    if 0 < n <= len(chunk) - data_at:
                        stream = chunk[data_at:data_at + n]
                except Exception:
                    stream = None
            if stream is None:
                # indirect /Length N 0 R → resolve it if we can
                n = self._resolve_len(body)
                if n and 0 < n <= len(chunk) - data_at:
                    stream = chunk[data_at:data_at + n]
            if stream is None:
                e = chunk.find(br'endstream', data_at)
                stream = chunk[data_at:e if e >= 0 else len(chunk)]
            self.objs[num] = (body, stream)

    def _expand_object_streams(self):
        """PDF 1.5+ /ObjStm: object definitions live inside a compressed
        stream. Without this, page objects are invisible."""
        for num in list(self.objs.keys()):
            body, stream = self.objs[num]
            if stream is None or br'/ObjStm' not in body:
                continue
            data = self._decode_stream(body, stream)
            if not data:
                continue
            try:
                n = int(re.search(br'/N\s+(\d+)', body).group(1))
                first = int(re.search(br'/First\s+(\d+)', body).group(1))
            except Exception:
                continue
            header = data[:first].split()
            for i in range(0, min(len(header) - 1, n * 2), 2):
                try:
                    onum = int(header[i])
                    off = int(header[i + 1])
                except Exception:
                    continue
                if onum in self.objs:
                    continue          # never shadow a real object
                nxt = (int(header[i + 3]) + first
                       if i + 3 < len(header) else len(data))
                self.objs[onum] = (data[first + off:nxt], None)

    def _decode_stream(self, body, stream):
        if not stream:
            return b''
        if br'/FlateDecode' in body:
            return _inflate(stream)
        return bytes(stream)

    def _resolve_len(self, body):
        """Indirect /Length N 0 R → the integer stored in object N, or None."""
        m = _RE_LENGTH_REF.search(body)
        if not m:
            return None
        try:
            b2, _s2 = self.get(int(m.group(1)))
            im = re.search(br'(\d+)', b2)
            return int(im.group(1)) if im else None
        except Exception:
            return None

    def get(self, num):
        return self.objs.get(num, (b'', None))

    def stream_of(self, num):
        body, stream = self.get(num)
        return self._decode_stream(body, stream)

    # ── page tree ─────────────────────────────────────────────────────────

    def page_objects(self):
        """Page object numbers in document order."""
        root = None
        for num, (body, _s) in self.objs.items():
            if br'/Type' in body and br'/Pages' in body and br'/Kids' in body:
                if root is None:
                    root = num
        ordered = []
        seen = set()

        def walk(num, depth=0):
            if depth > 64 or num in seen:
                return
            seen.add(num)
            body, _s = self.get(num)
            if _RE_TYPE_PAGE.search(body):
                ordered.append(num)
                return
            km = _RE_KIDS.search(body)
            if not km:
                return
            for m in _RE_REF.finditer(km.group(1)):
                walk(int(m.group(1)), depth + 1)

        # walk every /Pages node that has kids; roots first
        for num in ([root] if root else []) + sorted(self.objs.keys()):
            if num is None:
                continue
            body, _s = self.get(num)
            if br'/Kids' in body:
                walk(num)
        if not ordered:   # malformed tree → fall back to object order
            ordered = [n for n in sorted(self.objs)
                       if _RE_TYPE_PAGE.search(self.get(n)[0])]
        return ordered[:MAX_PAGES]

    # ── fonts / ToUnicode ─────────────────────────────────────────────────

    def font_maps(self, page_body):
        """{font_name: {code: unicode}} for a page's /Font resources."""
        body = page_body
        rm = _RE_RESOURCES_REF.search(body)
        if rm:
            body = body + self.get(int(rm.group(1)))[0]
        # /Font may itself be an indirect reference (/Font 13 0 R)
        fr = _RE_FONT_REF.search(body)
        if fr:
            body = body + b' /Font <<' + self.get(int(fr.group(1)))[0] + b'>>'
        out = {}
        for fb in _RE_FONT_BLOCK.finditer(body):
            for m in _RE_FONT_ENTRY.finditer(fb.group(1)):
                name = m.group(1)
                fbody = self.get(int(m.group(2)))[0]
                tm = _RE_TOUNICODE.search(fbody)
                if tm:
                    cmap = parse_cmap(self.stream_of(int(tm.group(1))))
                    if cmap:
                        out[name] = cmap
        return out

    def page_text(self, num):
        body, _s = self.get(num)
        cm = _RE_CONTENTS.search(body)
        if not cm:
            return ''
        refs = []
        if cm.group(1):
            refs.append(int(cm.group(1)))
        elif cm.group(3):
            refs.extend(int(r.group(1)) for r in _RE_REF.finditer(cm.group(3)))
        data = b''.join(self.stream_of(r) for r in refs)
        if not data:
            return ''
        return render_content(data, self.font_maps(body))


def parse_cmap(data):
    """Parse a /ToUnicode CMap into {code_int: unicode_str}."""
    if not data:
        return {}
    out = {}

    def _uni(h):
        try:
            raw = bytearray.fromhex(h.decode('ascii'))
        except Exception:
            return ''
        try:
            return bytes(raw).decode('utf-16-be', 'ignore')
        except Exception:
            return ''

    for blk in re.findall(br'beginbfchar(.*?)endbfchar', data, re.S):
        for src, dst in re.findall(br'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', blk):
            try:
                out[int(src, 16)] = _uni(dst)
            except Exception:
                pass
    for blk in re.findall(br'beginbfrange(.*?)endbfrange', data, re.S):
        # <lo> <hi> <dstStart>
        for lo, hi, dst in re.findall(
                br'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', blk):
            try:
                l, h = int(lo, 16), int(hi, 16)
                base = _uni(dst)
                if not base or h < l or h - l > 65535:
                    continue
                first = ord(base[0])
                for k in range(h - l + 1):
                    out[l + k] = unichr_(first + k) + base[1:]
            except Exception:
                pass
        # <lo> <hi> [<d1> <d2> ...]
        for lo, _hi, arr in re.findall(
                br'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[(.*?)\]', blk, re.S):
            try:
                l = int(lo, 16)
                for k, d in enumerate(re.findall(br'<([0-9A-Fa-f]+)>', arr)):
                    out[l + k] = _uni(d)
            except Exception:
                pass
    return out


def unichr_(i):
    try:
        return unichr(i)          # noqa: F821  (py2)
    except NameError:
        return chr(i)


def _decode_string(buf, cmap):
    """Map a PDF string's bytes to text through the font CMap (or cp1252)."""
    if cmap:
        # CMaps in practice are 1-byte (simple fonts) or 2-byte (Type0/CID)
        two = any(c > 0xFF for c in cmap) or len(buf) % 2 == 0 and \
            all((buf[i] == 0 or buf[i] < 0x20) for i in range(0, len(buf), 2))
        out = []
        if two and len(buf) >= 2:
            for i in range(0, len(buf) - 1, 2):
                code = (buf[i] << 8) | buf[i + 1]
                out.append(cmap.get(code, cmap.get(buf[i + 1], '')))
            if out and any(out):
                return ''.join(out)
        return ''.join(cmap.get(b, '') for b in buf)
    try:
        return bytes(buf).decode('cp1252', 'replace')
    except Exception:
        return bytes(buf).decode('latin-1', 'replace')


def _unescape(raw):
    """Resolve PDF literal-string escapes into a bytearray."""
    out = bytearray()
    i, n = 0, len(raw)
    simple = {0x6E: 0x0A, 0x72: 0x0D, 0x74: 0x09, 0x62: 0x08, 0x66: 0x0C}
    while i < n:
        c = raw[i]
        if c == 0x5C and i + 1 < n:      # backslash
            nxt = raw[i + 1]
            if nxt in simple:
                out.append(simple[nxt]); i += 2; continue
            if 0x30 <= nxt <= 0x37:      # octal \ddd
                j, val = i + 1, 0
                while j < n and j < i + 4 and 0x30 <= raw[j] <= 0x37:
                    val = val * 8 + (raw[j] - 0x30); j += 1
                out.append(val & 0xFF); i = j; continue
            if nxt == 0x0A:              # line continuation
                i += 2; continue
            out.append(nxt); i += 2; continue
        out.append(c); i += 1
    return out


def render_content(data, font_maps):
    """Extract visible text from a decompressed content stream."""
    buf = bytearray(data)
    out = []
    cur = {}          # active font cmap
    i, n = 0, len(buf)
    pending = []      # operands seen since the last operator
    lasty = None      # last text-baseline Y, for line-break detection

    def flush_string(raw_bytes, spacing_before=False):
        txt = _decode_string(_unescape(raw_bytes), cur)
        if txt:
            if spacing_before and out and not out[-1].endswith((' ', '\n')):
                out.append(' ')
            out.append(txt)

    while i < n:
        c = buf[i]
        # literal string ( ... )
        if c == 0x28:
            depth, j = 1, i + 1
            start = j
            while j < n and depth:
                if buf[j] == 0x5C:
                    j += 2; continue
                if buf[j] == 0x28:
                    depth += 1
                elif buf[j] == 0x29:
                    depth -= 1
                j += 1
            pending.append(('s', buf[start:j - 1]))
            i = j; continue
        # hex string < ... >  (skip dictionaries "<<")
        if c == 0x3C and i + 1 < n and buf[i + 1] != 0x3C:
            j = buf.find(b'>', i + 1)
            if j < 0:
                break
            hx = re.sub(br'[^0-9A-Fa-f]', b'', bytes(buf[i + 1:j]))
            if len(hx) % 2:
                hx += b'0'
            try:
                pending.append(('s', bytearray.fromhex(hx.decode('ascii'))))
            except Exception:
                pass
            i = j + 1; continue
        # name /Xxx
        if c == 0x2F:
            j = i + 1
            while j < n and buf[j] not in b' \t\r\n/[]<>(){}':
                j += 1
            pending.append(('n', bytes(buf[i + 1:j])))
            i = j; continue
        # number
        if (0x30 <= c <= 0x39) or c in (0x2D, 0x2B, 0x2E):
            j = i
            while j < n and ((0x30 <= buf[j] <= 0x39) or buf[j] in (0x2D, 0x2B, 0x2E)):
                j += 1
            try:
                pending.append(('#', float(bytes(buf[i:j]))))
            except Exception:
                pass
            i = j; continue
        # array delimiters are handled by the TJ operator itself
        if c in (0x5B, 0x5D):
            pending.append(('d', chr(c)))
            i += 1; continue
        # operator token
        if (0x41 <= c <= 0x5A) or (0x61 <= c <= 0x7A) or c in (0x27, 0x22):
            j = i
            while j < n and ((0x41 <= buf[j] <= 0x5A) or (0x61 <= buf[j] <= 0x7A)
                             or buf[j] in (0x27, 0x22, 0x2A, 0x30, 0x31)):
                j += 1
            op = bytes(buf[i:j])
            if op == b'Tf':
                for kind, val in reversed(pending):
                    if kind == 'n':
                        cur = font_maps.get(val, {})
                        break
            elif op in (b'Tj', b"'", b'"'):
                for kind, val in reversed(pending):
                    if kind == 's':
                        flush_string(val, op != b'Tj')
                        break
                if op in (b"'", b'"'):
                    out.append('\n')
            elif op == b'TJ':
                for kind, val in pending:
                    if kind == 's':
                        flush_string(val)
                    elif kind == '#' and val <= -180:
                        # large negative kern = inter-word gap
                        if out and not out[-1].endswith((' ', '\n')):
                            out.append(' ')
            elif op in (b'Td', b'TD', b'Tm', b'T*', b'ET'):
                # A newline only when the baseline actually MOVES DOWN.
                # Emitting one on every Td shatters a line into fragments
                # ("V" / "5" / "(202" / "5"), which is what made the first
                # digests look like word salad.
                nums = [v for k, v in pending if k == '#']
                newy = None
                if op == b'Tm' and len(nums) >= 6:
                    newy = nums[5]
                elif op in (b'Td', b'TD') and len(nums) >= 2:
                    newy = (lasty if lasty is not None else 0.0) + nums[1]
                elif op == b'T*':
                    newy = (lasty if lasty is not None else 0.0) - 1.0
                if op == b'ET':
                    pass                      # end of a text object: no break
                elif newy is None or lasty is None or abs(newy - lasty) > 1.5:
                    if out and not out[-1].endswith('\n'):
                        out.append('\n')
                elif out and not out[-1].endswith((' ', '\n')):
                    out.append(' ')           # same baseline → same line
                if newy is not None:
                    lasty = newy
            pending = []
            i = j if j > i else i + 1
            continue
        i += 1

    text = ''.join(out)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)
    return text.strip()


def extract_pages(path, max_pages=MAX_PAGES):
    """Return [(page_no, text)] for a PDF, or [] when nothing is extractable."""
    try:
        with open(path, 'rb') as f:
            raw = f.read()
    except Exception:
        return []
    if not raw[:1024].lstrip()[:5] == b'%PDF-' and b'%PDF-' not in raw[:1024]:
        return []
    try:
        pdf = _Pdf(raw)
    except Exception:
        return []
    pages = []
    try:
        nums = pdf.page_objects()
    except Exception:
        return []
    for idx, num in enumerate(nums[:max_pages], start=1):
        try:
            txt = pdf.page_text(num)
        except Exception:
            txt = ''
        if txt and txt.strip():
            pages.append((idx, txt.strip()))
    return pages
