# -*- coding: utf-8 -*-
"""
Optional semantic retrieval channel — Ollama embeddings + fusion helpers.

The knowledge stack stays fully functional without this (BM25-only);
when Ollama is reachable and the embedding model is installed, queries
gain a semantic re-rank fused with BM25 via reciprocal rank fusion.

Pure Python + injectable HTTP (CPython-3 testable). HTTP goes through
Intelligence.llm_provider's dual .NET/urllib backend.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import json
import math
import os


DEFAULT_MODEL = 'nomic-embed-text'
_PROBE_TIMEOUT_MS = 1500
_EMBED_TIMEOUT_MS = 60000
_PULL_TIMEOUT_MS = 600000     # first pull downloads ~270 MB
_VECTOR_DECIMALS = 4          # storage size: ~6 KB/chunk at 768 dims


def _candidate_hosts(configured=None):
    """Hosts to try, in order: configured/env → 127.0.0.1 → localhost."""
    env = os.environ.get('OLLAMA_HOST', '')
    out = []
    for h in (configured, env, 'http://127.0.0.1:11434', 'http://localhost:11434'):
        if h:
            h = h.rstrip('/')
            if not h.startswith('http'):
                h = 'http://' + h
            if h not in out:
                out.append(h)
    return out


class OllamaEmbedder(object):
    """Thin client for Ollama's embedding endpoints.

    Args:
        http_post_fn / http_get_fn: injectable transport (tests); default
            to Intelligence.llm_provider helpers.
        host: explicit Ollama host, else discovered.
        model: embedding model name (settings knowledge.embed_model).
    """

    MODEL = DEFAULT_MODEL

    def __init__(self, http_post_fn=None, http_get_fn=None, host=None, model=None):
        if model:
            self.MODEL = model
        self._host = host
        self._active_host = None
        self._available = None      # tri-state probe cache
        if http_post_fn is None or http_get_fn is None:
            from Intelligence.llm_provider import http_post, http_get
            http_post_fn = http_post_fn or http_post
            http_get_fn = http_get_fn or http_get
        self._post = http_post_fn
        self._get = http_get_fn

    # ── availability ──────────────────────────────────────────────────────

    def _probe(self):
        """Find a reachable host and whether MODEL is installed.
        Returns (host_or_None, model_installed)."""
        for host in _candidate_hosts(self._host):
            try:
                resp = self._get(host + '/api/tags', timeout_ms=_PROBE_TIMEOUT_MS)
                if not resp:
                    continue
                data = json.loads(resp)
                names = [m.get('name', '') for m in data.get('models', [])]
                base = self.MODEL.split(':')[0]
                installed = any(n == self.MODEL or n.split(':')[0] == base
                                for n in names)
                return host, installed
            except Exception:
                continue
        return None, False

    def is_available(self, recheck=False):
        """True when Ollama is reachable AND the embed model is installed."""
        if self._available is not None and not recheck:
            return self._available
        host, installed = self._probe()
        self._active_host = host
        self._available = bool(host and installed)
        return self._available

    def ensure_model(self):
        """Pull the embedding model if a host is up but the model is missing.
        BLOCKING (large download) — call from a background thread only.
        Returns True when the model is ready."""
        host, installed = self._probe()
        if not host:
            return False
        if installed:
            self._active_host = host
            self._available = True
            return True
        try:
            resp = self._post(host + '/api/pull',
                              {'name': self.MODEL, 'stream': False},
                              timeout_ms=_PULL_TIMEOUT_MS)
            ok = bool(resp) and '"error"' not in (resp or '')
            if ok:
                self._active_host = host
                self._available = True
            return ok
        except Exception:
            return False

    # ── embedding ─────────────────────────────────────────────────────────

    def embed(self, texts):
        """Embed a list of texts. Returns list of vectors (rounded floats),
        aligned with input; None on total failure."""
        if not texts:
            return []
        if not self.is_available():
            return None
        host = self._active_host
        # Preferred: batch /api/embed (Ollama >= 0.2.6)
        try:
            resp = self._post(host + '/api/embed',
                              {'model': self.MODEL, 'input': list(texts)},
                              timeout_ms=_EMBED_TIMEOUT_MS)
            if resp:
                data = json.loads(resp)
                vecs = data.get('embeddings')
                if vecs and len(vecs) == len(texts):
                    return [_round_vec(v) for v in vecs]
        except Exception:
            pass
        # Fallback: per-text /api/embeddings (older servers)
        out = []
        for text in texts:
            try:
                resp = self._post(host + '/api/embeddings',
                                  {'model': self.MODEL, 'prompt': text},
                                  timeout_ms=_EMBED_TIMEOUT_MS)
                data = json.loads(resp) if resp else {}
                vec = data.get('embedding')
                out.append(_round_vec(vec) if vec else None)
            except Exception:
                out.append(None)
        if any(v for v in out):
            return out
        return None


def _round_vec(vec):
    return [round(float(x), _VECTOR_DECIMALS) for x in vec]


# ─── math / fusion helpers ────────────────────────────────────────────────────

def cosine(a, b):
    """Cosine similarity of two equal-length vectors (pure Python)."""
    if not a or not b or len(a) != len(b):
        return 0.0
    dot = 0.0
    na = 0.0
    nb = 0.0
    for i in range(len(a)):
        x = a[i]
        y = b[i]
        dot += x * y
        na += x * x
        nb += y * y
    if na <= 0 or nb <= 0:
        return 0.0
    return dot / (math.sqrt(na) * math.sqrt(nb))


def rrf_fuse(ranked_lists, k=60, weights=None):
    """Reciprocal rank fusion over N ranked key lists.

    score(key) = sum_i weight_i / (k + rank_i(key))   (1-based ranks)

    Returns fused list of keys, best first.
    """
    if not ranked_lists:
        return []
    if weights is None:
        weights = [1.0] * len(ranked_lists)
    scores = {}
    for li, keys in enumerate(ranked_lists):
        w = weights[li] if li < len(weights) else 1.0
        for rank, key in enumerate(keys):
            scores[key] = scores.get(key, 0.0) + w / (k + rank + 1.0)
    return [key for key, _ in sorted(scores.items(), key=lambda kv: -kv[1])]


def get_default_embedder():
    """Embedder configured from settings — None when disabled.

    Never raises; safe to call from routing code.
    """
    try:
        from config.settings import get_settings
        s = get_settings()
        if not s.get_knowledge_option('embeddings_enabled', True):
            return None
        model = s.get_knowledge_option('embed_model', DEFAULT_MODEL)
        return OllamaEmbedder(model=model)
    except Exception:
        try:
            return OllamaEmbedder()
        except Exception:
            return None
