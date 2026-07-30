# -*- coding: utf-8 -*-
"""
pdf_crypt — PDF standard security handler (decryption of "protected" PDFs).

Why this exists
---------------
Corporate BEP/manual PDFs are routinely published with an OWNER password only:
you can open and read them in any viewer, but printing/editing is restricted.
Every stream in such a file is still encrypted on disk, so a text extractor
that ignores encryption inflates garbage and reports "no text layer" — which
looks identical to a scanned document and sends you chasing OCR that cannot
help. (Observed on WHPL-A-INNO-MAN-ALL-01-0100-BIMManual.pdf: /V 4 /R 4,
/CFM /V2, 83 object streams, 0 pages recovered.)

The user password on these files is EMPTY, which is exactly the case the
standard handler lets us derive the file key for without any secret.

Scope
-----
  * /Filter /Standard, revisions 2, 3, 4 (RC4 40/128) — pure Python, works on
    IronPython 2.7 and CPython 3 alike.
  * revision 5/6 (/AESV3, AES-256) and /AESV2 (AES-128) — handled when an AES
    backend is reachable (.NET System.Security.Cryptography under Revit, or
    pycryptodome/cryptography under CPython); otherwise reported as
    unsupported rather than silently mis-decrypted.
  * Only the empty user password is attempted. A file with a real user
    password is reported as such so the digest can say so honestly.

Byte handling goes through bytearray so indexing yields ints on both
runtimes — see the _DELIMS note in pdf_text.py for why that matters.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import hashlib
import re

_PAD = bytearray([
    0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
    0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
    0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
    0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A])

_RE_ENCRYPT_REF = re.compile(br'/Encrypt\s+(\d+)\s+(\d+)\s+R')
_RE_ID = re.compile(br'/ID\s*\[\s*<([0-9A-Fa-f\s]*)>', re.S)
_RE_ID_LIT = re.compile(br'/ID\s*\[\s*\((.{0,64}?)\)', re.S)
_RE_FILTER = re.compile(br'/Filter\s*/([A-Za-z0-9]+)')
_RE_V = re.compile(br'/V\s+(\d+)')
_RE_R = re.compile(br'/R\s+(\d+)')
_RE_LEN = re.compile(br'/Length\s+(\d+)')
_RE_P = re.compile(br'/P\s+(-?\d+)')
_RE_CFM = re.compile(br'/CFM\s*/([A-Za-z0-9]+)')
_RE_ENCMETA = re.compile(br'/EncryptMetadata\s+(true|false)')
_RE_STMF = re.compile(br'/StmF\s*/([A-Za-z0-9]+)')

# RC4 ("V2") and the AES variants, as named by /CFM
_CFM_RC4 = ('V2',)
_CFM_AES = ('AESV2', 'AESV3')
_CFM_NONE = ('None', 'Identity')


def _le32(value):
    """4-byte little-endian two's-complement, as the /P entry is hashed.

    Done by hand rather than with struct.pack: the format string must be str
    on py3 but the module is bytes-oriented on py2, and struct is not always
    importable in a hosted IronPython engine.
    """
    v = int(value) & 0xFFFFFFFF
    return bytearray([v & 0xFF, (v >> 8) & 0xFF,
                      (v >> 16) & 0xFF, (v >> 24) & 0xFF])


def _md5(*parts):
    h = hashlib.md5()
    for p in parts:
        h.update(bytes(bytearray(p)))
    return bytearray(h.digest())


def rc4(key, data):
    """RC4 / ARC4 stream cipher. key and data may be bytes or bytearray."""
    key = bytearray(key)
    data = bytearray(data)
    if not key:
        return bytes(data)
    s = list(range(256))
    j = 0
    klen = len(key)
    for i in range(256):
        j = (j + s[i] + key[i % klen]) & 0xFF
        s[i], s[j] = s[j], s[i]
    out = bytearray(len(data))
    i = j = 0
    for n in range(len(data)):
        i = (i + 1) & 0xFF
        j = (j + s[i]) & 0xFF
        s[i], s[j] = s[j], s[i]
        out[n] = data[n] ^ s[(s[i] + s[j]) & 0xFF]
    return bytes(out)


# ─── AES-CBC backends (optional; RC4 files need none of this) ─────────────────

def _aes_cbc_decrypt(key, data):
    """AES-CBC decrypt with the IV prefixed to `data`. None when no backend."""
    data = bytes(bytearray(data))
    key = bytes(bytearray(key))
    if len(data) <= 16:
        return b''
    iv, body = data[:16], data[16:]
    body = body[:len(body) - (len(body) % 16)]
    if not body:
        return b''
    out = _aes_dotnet(key, iv, body)
    if out is None:
        out = _aes_pylib(key, iv, body)
    if out is None:
        return None
    # strip PKCS#7 padding when it looks valid
    if out:
        pad = bytearray(out)[-1]
        if 1 <= pad <= 16 and len(out) >= pad:
            out = out[:-pad]
    return out


def _aes_dotnet(key, iv, body):
    try:
        import clr                                   # noqa: F401
        clr.AddReference('System.Core')
        from System.Security.Cryptography import (Aes, CipherMode,
                                                  PaddingMode)
        from System import Array, Byte
        alg = Aes.Create()
        alg.Mode = CipherMode.CBC
        # `PaddingMode.None` is a syntax error — None is a keyword in Python
        alg.Padding = getattr(PaddingMode, 'None')    # we strip it ourselves
        alg.Key = Array[Byte](bytearray(key))
        alg.IV = Array[Byte](bytearray(iv))
        dec = alg.CreateDecryptor()
        buf = Array[Byte](bytearray(body))
        res = dec.TransformFinalBlock(buf, 0, buf.Length)
        return bytes(bytearray(res))
    except Exception:
        return None


def _aes_pylib(key, iv, body):
    try:
        from Crypto.Cipher import AES                 # pycryptodome
        return AES.new(key, AES.MODE_CBC, iv).decrypt(body)
    except Exception:
        pass
    try:
        from cryptography.hazmat.primitives.ciphers import (
            Cipher, algorithms, modes)
        dec = Cipher(algorithms.AES(key), modes.CBC(iv)).decryptor()
        return dec.update(body) + dec.finalize()
    except Exception:
        return None


def aes_available():
    """True when some AES backend can be reached (probe with a dummy block)."""
    probe = b'\x00' * 32
    return _aes_dotnet(b'\x00' * 16, b'\x00' * 16, probe) is not None or \
        _aes_pylib(b'\x00' * 16, b'\x00' * 16, probe) is not None


# ─── the handler ─────────────────────────────────────────────────────────────

class Decryptor(object):
    """Per-object decryption for one PDF. `reason` explains an unusable file."""

    def __init__(self, key, cfm, revision, reason=''):
        self.key = bytearray(key or b'')
        self.cfm = cfm
        self.revision = revision
        self.reason = reason

    @property
    def usable(self):
        return bool(self.key) and not self.reason

    def _object_key(self, num, gen):
        extra = bytearray([num & 0xFF, (num >> 8) & 0xFF, (num >> 16) & 0xFF,
                           gen & 0xFF, (gen >> 8) & 0xFF])
        parts = [self.key, extra]
        if self.cfm == 'AESV2':
            parts.append(bytearray(b'sAlT'))
        digest = _md5(*parts)
        return digest[:min(len(self.key) + 5, 16)]

    def decrypt(self, data, num, gen=0):
        """Decrypt one string/stream belonging to object (num, gen)."""
        if not data or not self.usable:
            return bytes(bytearray(data or b''))
        if self.cfm == 'AESV3':
            out = _aes_cbc_decrypt(self.key, data)      # key used directly
            return out if out is not None else b''
        if self.cfm == 'AESV2':
            out = _aes_cbc_decrypt(self._object_key(num, gen), data)
            return out if out is not None else b''
        if self.cfm in _CFM_NONE:
            return bytes(bytearray(data))
        return rc4(self._object_key(num, gen), data)


def _first_id(raw):
    """The first element of the trailer /ID array, as bytes."""
    m = _RE_ID.search(raw)
    if m:
        hx = re.sub(br'\s', b'', m.group(1))
        if len(hx) % 2:
            hx += b'0'
        try:
            return bytes(bytearray.fromhex(hx.decode('ascii')))
        except Exception:
            return b''
    m = _RE_ID_LIT.search(raw)
    return bytes(m.group(1)) if m else b''


def _int(rx, body, default=0):
    m = rx.search(body)
    try:
        return int(m.group(1)) if m else default
    except Exception:
        return default


def _split_cf(body):
    """(top_level_dict, crypt_filter_block) for an /Encrypt dictionary.

    The nested `/CF << /StdCF << /Length 16 ... >> >>` carries a /Length in
    BYTES while the encrypt dictionary's own /Length is in BITS. A plain
    search for /Length finds the nested one first and yields a 5-byte key
    instead of 16 — the file then "needs a password" even though the empty
    one is correct. Slice the sub-dictionary out before reading top-level keys.
    """
    m = re.search(br'/CF\s*<<', body)
    if not m:
        return body, b''
    buf = bytearray(body)
    i = m.end() - 2          # position of the opening '<<'
    depth = 0
    n = len(buf)
    while i < n - 1:
        pair = (buf[i], buf[i + 1])
        if pair == (0x3C, 0x3C):
            depth += 1
            i += 2
            continue
        if pair == (0x3E, 0x3E):
            depth -= 1
            i += 2
            if depth == 0:
                return body[:m.start()] + body[i:], body[m.start():i]
            continue
        i += 1
    return body, b''


def _pdf_string(body, key):
    """Value of `/key ( ... )` as raw bytes, honouring backslash escapes."""
    m = re.search(br'/' + key + br'\s*\(', body)
    if not m:
        hm = re.search(br'/' + key + br'\s*<([0-9A-Fa-f\s]+)>', body)
        if hm:
            hx = re.sub(br'\s', b'', hm.group(1))
            if len(hx) % 2:
                hx += b'0'
            try:
                return bytes(bytearray.fromhex(hx.decode('ascii')))
            except Exception:
                return b''
        return b''
    buf = bytearray(body)
    i = m.end()
    out = bytearray()
    depth = 1
    n = len(buf)
    # PDF literal-string escapes. Taking the raw next byte instead (i.e. 'r'
    # for \r, 'b' for \b) corrupts /O and /U — both of which feed the key
    # derivation, so the empty-password check fails and a perfectly readable
    # owner-locked file gets reported as needing a password.
    simple = {0x6E: 0x0A, 0x72: 0x0D, 0x74: 0x09, 0x62: 0x08, 0x66: 0x0C,
              0x28: 0x28, 0x29: 0x29, 0x5C: 0x5C}
    while i < n:
        c = buf[i]
        if c == 0x5C and i + 1 < n:          # backslash escape
            nxt = buf[i + 1]
            if nxt in simple:
                out.append(simple[nxt])
                i += 2
                continue
            if 0x30 <= nxt <= 0x37:          # octal \ddd
                j, val = i + 1, 0
                while j < n and j < i + 4 and 0x30 <= buf[j] <= 0x37:
                    val = val * 8 + (buf[j] - 0x30)
                    j += 1
                out.append(val & 0xFF)
                i = j
                continue
            if nxt in (0x0A, 0x0D):          # line continuation
                i += 2
                continue
            out.append(nxt)
            i += 2
            continue
        if c == 0x28:
            depth += 1
        elif c == 0x29:
            depth -= 1
            if depth == 0:
                break
        out.append(c)
        i += 1
    return bytes(out)


def build(raw, get_object, password=b''):
    """Return a Decryptor for `raw`, or None when the file is not encrypted.

    get_object(num) -> body_bytes, so the /Encrypt dictionary can be resolved
    through the caller's already-scanned object table (the encrypt dict itself
    is never encrypted).
    """
    raw = bytes(bytearray(raw))
    m = _RE_ENCRYPT_REF.search(raw)
    if not m:
        return None
    try:
        body = bytes(bytearray(get_object(int(m.group(1))) or b''))
    except Exception:
        body = b''
    if not body:
        return Decryptor(b'', '', 0, reason='encrypted (unreadable /Encrypt)')

    fm = _RE_FILTER.search(body)
    filt = fm.group(1).decode('ascii', 'ignore') if fm else ''
    if filt != 'Standard':
        return Decryptor(b'', '', 0,
                         reason='encrypted with a non-standard handler '
                                '(/Filter /%s)' % (filt or '?'))

    top, cf_block = _split_cf(body)
    v = _int(_RE_V, top, 0)
    r = _int(_RE_R, top, 0)
    p = _int(_RE_P, top, -1)
    o_val = _pdf_string(top, b'O')
    u_val = _pdf_string(top, b'U')
    em = _RE_ENCMETA.search(top)
    enc_meta = not (em and em.group(1) == b'false')

    length = _int(_RE_LEN, top, 0)
    if not length and cf_block:
        length = _int(_RE_LEN, cf_block, 0) * 8      # crypt filter is in bytes
    if not length:
        length = 40

    cfm = 'V2'
    if v >= 4:
        cm = _RE_CFM.search(cf_block or top)
        cfm = cm.group(1).decode('ascii', 'ignore') if cm else 'V2'
        sm = _RE_STMF.search(top)
        if sm and sm.group(1) in (b'Identity',):
            cfm = 'Identity'
    if v == 5 or r >= 5:
        cfm = 'AESV3'

    if cfm in _CFM_AES and not aes_available():
        return Decryptor(b'', cfm, r,
                         reason='encrypted with AES (%s) and no AES backend '
                                'is available' % cfm)

    if cfm == 'AESV3':
        key = _key_r5(u_val, o_val, body, password)
        if key is None:
            return Decryptor(b'', cfm, r,
                             reason='encrypted and needs a password')
        return Decryptor(key, cfm, r)

    # revisions 2-4: RC4/AES-128 file key from the padded password
    n = 5 if r == 2 else max(5, min(16, length // 8))
    parts = [bytearray(bytearray(password) + _PAD)[:32], bytearray(o_val[:32]),
             _le32(p), bytearray(_first_id(raw))]
    if r >= 4 and not enc_meta:
        parts.append(bytearray(b'\xff\xff\xff\xff'))
    key = _md5(*parts)
    if r >= 3:
        for _ in range(50):
            key = _md5(key[:n])
    key = key[:n]

    if u_val and not _check_user(key, u_val, raw, r):
        return Decryptor(b'', cfm, r,
                         reason='encrypted and needs a user password')
    return Decryptor(key, cfm, r)


def _check_user(key, u_val, raw, r):
    """Algorithm 4/5 — validate the empty-user-password key against /U."""
    try:
        if r == 2:
            return rc4(key, _PAD)[:32] == bytes(bytearray(u_val))[:32]
        val = _md5(_PAD, bytearray(_first_id(raw)))
        val = bytearray(rc4(key, val))
        for i in range(1, 20):
            k2 = bytearray((b ^ i) & 0xFF for b in key)
            val = bytearray(rc4(k2, val))
        return bytes(val)[:16] == bytes(bytearray(u_val))[:16]
    except Exception:
        return False


def _key_r5(u_val, o_val, body, password):
    """Revision 5/6 (AES-256): validate the empty user password and unwrap
    the file key from /U + /UE. None when the password is wrong."""
    try:
        if len(u_val) < 48:
            return None
        pwd = bytes(bytearray(password))
        salt = bytes(bytearray(u_val))[32:40]
        if hashlib.sha256(pwd + salt).digest() != bytes(bytearray(u_val))[:32]:
            return None
        ks = bytes(bytearray(u_val))[40:48]
        ikey = hashlib.sha256(pwd + ks).digest()
        ue = _pdf_string(body, b'UE')
        if not ue:
            return None
        out = _aes_dotnet(ikey, b'\x00' * 16, bytes(bytearray(ue))[:32])
        if out is None:
            out = _aes_pylib(ikey, b'\x00' * 16, bytes(bytearray(ue))[:32])
        return bytearray(out)[:32] if out else None
    except Exception:
        return None
