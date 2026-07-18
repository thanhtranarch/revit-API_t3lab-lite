# -*- coding: utf-8 -*-
"""
Knowledge package — RAG v2 for T3Lab Assistant.

Pure-Python (IronPython 2.7 + CPython 3 compatible) retrieval stack:
    vi_text.py         diacritic folding / tokenizing (VI + EN)
    chunker.py         page-aware text chunking
    bm25_index.py      BM25 inverted index (lexical channel)
    embeddings.py      Ollama embeddings + RRF fusion (semantic channel)
    knowledge_store.py persistent index manager (scan/index/search)
    knowledge_agent.py retrieval + single-chat-call answer with citations

Author: Tran Tien Thanh
"""
