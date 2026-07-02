# -*- coding: utf-8 -*-
"""
LLM Provider

Abstract base class and shared HTTP helper for all LLM provider adapters.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "LLM Provider"

import json
import re

# ─── Shared HTTP backend ───────────────────────────────────────────────────────
# IronPython/pyRevit: use .NET WebClient.
# CPython (testing): fall back to urllib.

_USE_NET = False
try:
    import clr
    clr.AddReference('System.Net')
    from System.Net import WebClient
    from System.Text import Encoding as _NetEncoding
    _USE_NET = True
except Exception:
    pass

_HAS_URLLIB = False
if not _USE_NET:
    try:
        from urllib2 import urlopen, Request    # Python 2 / IronPython fallback
        _HAS_URLLIB = True
    except ImportError:
        try:
            from urllib.request import urlopen, Request  # Python 3
            _HAS_URLLIB = True
        except Exception:
            pass

HAS_HTTP = _USE_NET or _HAS_URLLIB


def http_get_auth(url, headers=None, timeout_ms=8000):
    """
    Authenticated GET request with optional headers.
    Returns response text string, or None on error.
    Mirrors http_post's dual .NET / urllib backend.
    """
    if _USE_NET:
        try:
            from System.Net import WebClient
            client = WebClient()
            try:
                client.Encoding = _NetEncoding.UTF8
                if headers:
                    for k, v in headers.items():
                        client.Headers.Add(k, v)
                return client.DownloadString(url)
            finally:
                try:
                    client.Dispose()
                except Exception:
                    pass
        except Exception:
            pass

    if _HAS_URLLIB:
        try:
            req = Request(url)
            if headers:
                for k, v in headers.items():
                    req.add_header(k, v)
            timeout_sec = float(timeout_ms) / 1000.0
            resp = urlopen(req, timeout=timeout_sec)
            raw = resp.read()
            return raw.decode("utf-8") if isinstance(raw, bytes) else raw
        except Exception:
            pass

    return None


def http_post(url, payload, headers=None, timeout_ms=60000):
    """
    POST a JSON-serialisable payload and return the response string.

    Args:
        url: target URL string.
        payload: dict to serialise as JSON.
        headers: optional dict of extra request headers.
        timeout_ms: request timeout in milliseconds. Default 60000 (60s) —
            fine for cloud APIs. Local providers (Ollama/LM Studio) doing
            CPU inference on a multi-billion-parameter model routinely need
            much longer; callers there should pass a larger value.
            IMPORTANT: plain System.Net.WebClient has NO Timeout property,
            so without explicitly building the request via HttpWebRequest
            (as done below), every local-model call silently inherited
            .NET's ~100s default and failed on any slower model/machine —
            indistinguishable from "the model answered badly", when in fact
            the request never completed at all.

    Returns:
        str: response body, or raises RuntimeError on failure.
    """
    body = json.dumps(payload, ensure_ascii=False)
    if _USE_NET:
        from System.Net import HttpWebRequest
        from System.IO import StreamReader
        body_bytes = _NetEncoding.UTF8.GetBytes(body)
        req = HttpWebRequest.Create(url)
        req.Method           = "POST"
        req.ContentType      = "application/json; charset=utf-8"
        req.Timeout          = timeout_ms
        req.ReadWriteTimeout  = timeout_ms
        if headers:
            for k, v in headers.items():
                req.Headers.Add(k, v)
        req.ContentLength = body_bytes.Length
        rs = req.GetRequestStream()
        try:
            rs.Write(body_bytes, 0, body_bytes.Length)
        finally:
            rs.Close()
        resp = req.GetResponse()
        try:
            reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
            try:
                return reader.ReadToEnd()
            finally:
                reader.Close()
        finally:
            resp.Close()

    if _HAS_URLLIB:
        if isinstance(body, type(u"")):
            body_bytes = body.encode("utf-8")
        else:
            body_bytes = body
        req_headers = {"Content-Type": "application/json; charset=utf-8"}
        if headers:
            req_headers.update(headers)
        req = Request(url, body_bytes, req_headers)
        resp = urlopen(req, timeout=float(timeout_ms) / 1000.0)
        raw = resp.read()
        return raw.decode("utf-8") if isinstance(raw, bytes) else raw

    raise RuntimeError("No HTTP client available")


def http_get(url, timeout_ms=4000):
    """GET url; return response string or None. Times out after timeout_ms."""
    try:
        if _USE_NET:
            from System.Net import HttpWebRequest
            from System.IO import StreamReader
            req = HttpWebRequest.Create(url)
            req.Method  = "GET"
            req.Timeout = timeout_ms
            resp = req.GetResponse()
            try:
                reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
                try:
                    return reader.ReadToEnd()
                finally:
                    reader.Close()
            finally:
                resp.Close()
        if _HAS_URLLIB:
            resp = urlopen(url, timeout=4)
            raw = resp.read()
            return raw.decode("utf-8") if isinstance(raw, bytes) else raw
    except Exception:
        return None


# ─── Streaming (Server-Sent Events) backend ────────────────────────────────────

def http_post_stream(url, payload, headers=None, on_line=None, timeout_ms=120000):
    """
    POST a JSON payload and stream the response back line-by-line.

    Each decoded text line of the response is passed to on_line(line) as it
    arrives. Used for SSE endpoints (request payloads carry "stream": true).

    Args:
        url: target URL string.
        payload: dict to serialise as JSON.
        headers: optional dict of extra request headers.
        on_line: callable(str) invoked once per response line.
        timeout_ms: socket timeout in milliseconds.

    Returns:
        bool: True when the stream completes. Raises on transport failure so the
        caller can fall back to a blocking request.
    """
    body = json.dumps(payload, ensure_ascii=False)

    if _USE_NET:
        # .NET HttpWebRequest streams the response without buffering it whole.
        from System.Net import HttpWebRequest
        from System.IO import StreamReader

        req = HttpWebRequest.Create(url)
        req.Method          = "POST"
        req.ContentType     = "application/json; charset=utf-8"
        req.Timeout         = timeout_ms
        req.ReadWriteTimeout = timeout_ms
        if headers:
            for k, v in headers.items():
                # Content-Type is set via the property above; everything else
                # (x-api-key, Authorization, anthropic-version, …) is unrestricted.
                req.Headers.Add(k, v)

        data = _NetEncoding.UTF8.GetBytes(body)
        req.ContentLength = data.Length
        rs = req.GetRequestStream()
        try:
            rs.Write(data, 0, data.Length)
        finally:
            rs.Close()

        resp = req.GetResponse()
        try:
            reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
            try:
                while True:
                    line = reader.ReadLine()
                    if line is None:
                        break
                    if on_line is not None:
                        on_line(line)
            finally:
                reader.Close()
        finally:
            resp.Close()
        return True

    if _HAS_URLLIB:
        body_bytes = body.encode("utf-8") if isinstance(body, type(u"")) else body
        req_headers = {"Content-Type": "application/json; charset=utf-8"}
        if headers:
            req_headers.update(headers)
        req  = Request(url, body_bytes, req_headers)
        resp = urlopen(req, timeout=120)
        for raw_line in resp:
            line = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else raw_line
            if on_line is not None:
                on_line(line.rstrip("\n"))
        return True

    raise RuntimeError("No HTTP client available")


def parse_anthropic_stream_line(line):
    """Return the text delta carried by one Anthropic SSE line, or None."""
    if not line:
        return None
    line = line.strip()
    if not line.startswith("data:"):
        return None
    data = line[5:].strip()
    if not data or data == "[DONE]":
        return None
    try:
        obj = json.loads(data)
    except Exception:
        return None
    if obj.get("type") == "content_block_delta":
        delta = obj.get("delta", {}) or {}
        if delta.get("type") in ("text_delta", None):
            return delta.get("text")
    return None


def parse_openai_stream_line(line):
    """Return the text delta carried by one OpenAI-format SSE line, or None.

    Shared by OpenAI and DeepSeek (both OpenAI-compatible). Reasoning models
    stream their chain-of-thought under `reasoning_content`, which is ignored —
    only the user-facing `content` delta is surfaced.
    """
    if not line:
        return None
    line = line.strip()
    if not line.startswith("data:"):
        return None
    data = line[5:].strip()
    if not data or data == "[DONE]":
        return None
    try:
        obj = json.loads(data)
    except Exception:
        return None
    try:
        choices = obj.get("choices") or []
        if not choices:
            return None
        delta = choices[0].get("delta") or {}
        return delta.get("content")
    except Exception:
        return None


class StreamingJSONExtractor(object):
    """
    Incrementally surface the human-readable `message` value out of a streaming
    JSON response of the form:

        {"intent": "...", "params": {...}, "message": "<reply>"}

    The T3Lab system prompt asks models to answer in JSON, but during streaming
    the user should only ever see the `message` text — never the raw braces.
    Feed the full accumulated raw text on each delta; ``display`` returns the
    best human-readable string so far (partial messages included). If the model
    replies in plain prose instead of JSON, the prose is returned verbatim.
    """

    _ESCAPES = {"n": "\n", "t": "\t", "r": "\r", '"': '"', "\\": "\\", "/": "/"}

    @staticmethod
    def _strip_fences(text):
        t = text.lstrip()
        if t.startswith("```"):
            nl = t.find("\n")
            if nl != -1:
                t = t[nl + 1:]
            stripped = t.rstrip()
            if stripped.endswith("```"):
                t = stripped[:-3]
        return t

    def display(self, raw):
        if not raw:
            return u""
        t = self._strip_fences(raw)
        head = t.lstrip()

        # Plain prose (model ignored the JSON instruction) → show as-is.
        if not head.startswith("{"):
            return raw.strip()

        key_idx = t.find('"message"')
        if key_idx == -1:
            # JSON object opened but the message field hasn't streamed yet.
            return u""

        colon = t.find(":", key_idx)
        if colon == -1:
            return u""
        q = t.find('"', colon)
        if q == -1:
            return u""

        out = []
        i = q + 1
        n = len(t)
        while i < n:
            c = t[i]
            if c == "\\" and i + 1 < n:
                out.append(self._ESCAPES.get(t[i + 1], t[i + 1]))
                i += 2
                continue
            if c == '"':          # unescaped closing quote → end of message
                break
            out.append(c)
            i += 1
        return u"".join(out)


# ─── Abstract base provider ────────────────────────────────────────────────────

class BaseLLMProvider(object):
    """
    Abstract base class for all LLM provider adapters.

    Subclasses must implement:
        chat(messages, system_prompt, user_content, max_tokens) → str | None
        check_health()                                          → bool
    """

    NAME         = "base"
    DISPLAY_NAME = "Base Provider"

    # True if this provider can handle image content blocks
    SUPPORTS_VISION = False

    def _debug_log(self, msg):
        """Best-effort debug log via pyRevit's logger; never raises.

        chat()/check_health() failures here are usually swallowed and
        returned as None/False, which looks identical to "not configured" —
        use this in the except-blocks that wrap the actual network call so a
        real API error (bad key, malformed response, rate limit) leaves a
        trace instead of vanishing silently.
        """
        try:
            from pyrevit import script
            script.get_logger().debug(u"{}: {}".format(self.NAME, msg))
        except Exception:
            pass

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Send a chat request and return the raw response text.

        Args:
            messages (list): prior [{role, content}] dicts — conversation history.
                             Content may be a string or a list of content blocks.
            system_prompt (str): system instruction string.
            user_content (str|list): current user input — plain string OR a list
                                     of Claude-format content blocks (text/image).
            max_tokens (int): maximum tokens in the response.

        Returns:
            str | None: raw response text, or None on failure.
        """
        raise NotImplementedError

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """
        Streaming variant of chat(). Calls on_delta(text_chunk) for each piece of
        text as it arrives and returns the full concatenated response.

        The base implementation has no real token streaming: it performs a normal
        blocking chat() and emits the whole reply as a single delta. Providers
        that support Server-Sent Events override this for true incremental output.

        Returns:
            str | None: full response text, or None on failure.
        """
        text = self.chat(messages, system_prompt, user_content, max_tokens, **kwargs)
        if text and on_delta:
            try:
                on_delta(text)
            except Exception:
                pass
        return text

    def check_health(self):
        """Return True if the provider is reachable and has credentials."""
        return False

    def supports_vision(self):
        return self.SUPPORTS_VISION

    def get_models(self):
        """Return a list of model name strings available for this provider."""
        return []

    def get_active_model(self):
        """Return the model name currently in use, or None."""
        return None

    def set_model(self, model_name):
        """
        Set the model to use for future requests.

        Returns:
            bool: True if the model was accepted.
        """
        return False

    # ── Shared utilities ───────────────────────────────────────────────────────

    @staticmethod
    def extract_json(text):
        """Extract the first JSON object from a response string."""
        text = text.strip()
        try:
            return json.loads(text)
        except Exception:
            pass
        m = re.search(r'\{[\s\S]*\}', text)
        if m:
            try:
                return json.loads(m.group())
            except Exception:
                pass
        return None

    @staticmethod
    def blocks_to_text(user_content):
        """
        Flatten a list of Claude-format content blocks to a plain text string.
        Used by providers that do not support vision.
        """
        if isinstance(user_content, list):
            parts = []
            for block in user_content:
                if isinstance(block, dict) and block.get("type") == "text":
                    parts.append(block.get("text", ""))
            return "\n".join(parts)
        return user_content or ""

    @staticmethod
    def has_image_blocks(user_content):
        """Return True if user_content contains at least one image block."""
        if not isinstance(user_content, list):
            return False
        for block in user_content:
            if isinstance(block, dict) and block.get("type") == "image":
                return True
        return False
