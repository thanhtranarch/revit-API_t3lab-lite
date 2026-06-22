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
            resp = urlopen(req, timeout=8)
            raw = resp.read()
            return raw.decode("utf-8") if isinstance(raw, bytes) else raw
        except Exception:
            pass

    return None


def http_post(url, payload, headers=None):
    """
    POST a JSON-serialisable payload and return the response string.

    Args:
        url: target URL string.
        payload: dict to serialise as JSON.
        headers: optional dict of extra request headers.

    Returns:
        str: response body, or raises RuntimeError on failure.
    """
    body = json.dumps(payload, ensure_ascii=False)
    if _USE_NET:
        body_bytes = _NetEncoding.UTF8.GetBytes(body)
        client = WebClient()
        try:
            client.Encoding = _NetEncoding.UTF8
            client.Headers.Add("Content-Type", "application/json; charset=utf-8")
            if headers:
                for k, v in headers.items():
                    client.Headers.Add(k, v)
            resp_bytes = client.UploadData(url, "POST", body_bytes)
            return _NetEncoding.UTF8.GetString(resp_bytes)
        finally:
            try:
                client.Dispose()
            except Exception:
                pass

    if _HAS_URLLIB:
        if isinstance(body, type(u"")):
            body_bytes = body.encode("utf-8")
        else:
            body_bytes = body
        req_headers = {"Content-Type": "application/json; charset=utf-8"}
        if headers:
            req_headers.update(headers)
        req = Request(url, body_bytes, req_headers)
        resp = urlopen(req, timeout=60)
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

    def chat(self, messages, system_prompt, user_content, max_tokens=400):
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
