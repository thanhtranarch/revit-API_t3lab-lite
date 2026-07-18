# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the T3Lab Assistant knowledge stack.

Run:  python3 dev/test_knowledge.py
Exit code 0 = all pass. No external test framework — plain asserts,
mirroring dev/audit_tools.py conventions.

Covers modules that must be importable OUTSIDE Revit (guarded clr imports):
    lib/Intelligence/knowledge/*        (vi_text, chunker, bm25, embeddings fusion)
    lib/Intelligence/agents/dispatcher  (keyword stage)
    lib/Intelligence/skills_engine      (frontmatter parsing)
    lib/Intelligence/comments/*         (pdf annots, sheet matcher)
"""
from __future__ import unicode_literals

import json
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─── vi_text ──────────────────────────────────────────────────────────────────

def test_vi_text():
    print('[vi_text]')
    from Intelligence.knowledge import vi_text

    check('fold basic', vi_text.fold_diacritics('tường') == 'tuong')
    check('fold dj', vi_text.fold_diacritics('đường Đông') == 'duong Dong')
    check('fold passthrough', vi_text.fold_diacritics('Wall-101') == 'Wall-101')
    check('fold empty', vi_text.fold_diacritics('') == '')

    toks = vi_text.tokenize('Có bao nhiêu bức tường trong dự án?')
    check('tokenize vi', 'tuong' in toks and 'bao' in toks and 'nhieu' in toks, repr(toks))
    check('tokenize stopword drop', 'trong' not in toks, repr(toks))
    toks2 = vi_text.tokenize('How many walls are in the project?')
    check('tokenize en', 'walls' in toks2 and 'many' in toks2, repr(toks2))
    check('tokenize stopword en', 'the' not in toks2 and 'in' not in toks2, repr(toks2))

    s = vi_text.word_match_score(['mat', 'bang', 'tang', '01'], ['mat', 'bang', 'tang', 'mai'])
    check('word_match_score partial', 0.7 < s <= 1.0, s)
    check('word_match_score empty', vi_text.word_match_score([], ['a']) == 0.0)


# ─── chunker ──────────────────────────────────────────────────────────────────

def test_chunker():
    print('[chunker]')
    from Intelligence.knowledge import chunker

    words_p1 = ' '.join('w{}'.format(i) for i in range(1000))
    words_p2 = ' '.join('v{}'.format(i) for i in range(100))
    chunks = chunker.chunk_pages([(1, words_p1), (2, words_p2)],
                                 target_words=300, overlap_words=50)

    check('chunks produced', len(chunks) >= 4, len(chunks))
    check('no chunk spans pages',
          all(c['page'] in (1, 2) for c in chunks))
    check('seq unique', len(set(c['seq'] for c in chunks)) == len(chunks))
    p1 = [c for c in chunks if c['page'] == 1]
    first_words = p1[0]['text'].split()
    second_words = p1[1]['text'].split()
    check('overlap present', first_words[-50:] == second_words[:50])
    all_p1_words = set()
    for c in p1:
        all_p1_words.update(c['text'].split())
    check('no words lost page1', len(all_p1_words) == 1000, len(all_p1_words))

    tiny = chunker.chunk_text('one two three')
    check('tiny text single chunk', len(tiny) == 1 and tiny[0]['page'] == 0)
    check('empty text no chunk', chunker.chunk_text('') == [])


# ─── bm25 ─────────────────────────────────────────────────────────────────────

CORPUS = [
    ('d_aaa', [
        {'page': 1, 'seq': 0,
         'text': 'Chiều cao lan can ban công tối thiểu 1100 mm theo tiêu chuẩn an toàn.'},
        {'page': 2, 'seq': 1,
         'text': 'Cửa thoát hiểm phải mở theo chiều thoát nạn, chiều rộng tối thiểu 800 mm.'},
    ]),
    ('d_bbb', [
        {'page': 1, 'seq': 0,
         'text': 'Fire rated walls shall achieve a two hour rating at stair cores.'},
        {'page': 1, 'seq': 1,
         'text': 'Handrail height for stairs is 900 mm measured from nosing.'},
    ]),
]


def test_bm25():
    print('[bm25]')
    from Intelligence.knowledge.bm25_index import BM25Index, make_chunk_key

    idx = BM25Index()
    for doc_id, chunks in CORPUS:
        idx.add_document(doc_id, chunks)

    check('size', idx.size == 4, idx.size)

    hits = idx.search('chiều cao lan can bao nhiêu?', top_k=3)
    check('vi query hits', len(hits) >= 1)
    check('vi query top is lan can chunk',
          hits[0][0] == make_chunk_key('d_aaa', 1, 0), hits and hits[0][0])

    hits2 = idx.search('fire rating of walls', top_k=3)
    check('en query top is fire chunk',
          hits2 and hits2[0][0] == make_chunk_key('d_bbb', 1, 0),
          hits2 and hits2[0][0])

    # diacritic-free query must still match diacritic content
    hits3 = idx.search('chieu cao lan can', top_k=3)
    check('folded query matches',
          hits3 and hits3[0][0] == make_chunk_key('d_aaa', 1, 0))

    # persistence round-trip
    idx2 = BM25Index.from_dict(json.loads(json.dumps(idx.to_dict())))
    hits4 = idx2.search('chiều cao lan can', top_k=1)
    check('round-trip search identical', hits4 and hits4[0][0] == hits[0][0])

    # allowed_docs filter restricts scoring to given documents
    hits5 = idx.search('height mm', top_k=5, allowed_docs=set(['d_bbb']))
    check('allowed_docs filter', hits5 and all(
        k.startswith('d_bbb#') for k, _ in hits5), hits5)

    idx.remove_document('d_aaa')
    check('remove_document size', idx.size == 2, idx.size)
    check('removed doc unfindable', not any(
        k.startswith('d_aaa#') for k, _ in idx.search('lan can', top_k=5)))


# ─── embeddings ───────────────────────────────────────────────────────────────

def test_embeddings():
    print('[embeddings]')
    from Intelligence.knowledge.embeddings import (
        OllamaEmbedder, cosine, rrf_fuse)

    check('cosine identical', abs(cosine([1.0, 2.0], [1.0, 2.0]) - 1.0) < 1e-9)
    check('cosine orthogonal', abs(cosine([1.0, 0.0], [0.0, 1.0])) < 1e-9)
    check('cosine mismatched len', cosine([1.0], [1.0, 2.0]) == 0.0)
    check('cosine zero vec', cosine([0.0, 0.0], [1.0, 1.0]) == 0.0)

    fused = rrf_fuse([['a', 'b', 'c'], ['b', 'a', 'd']])
    check('rrf both-listed first', fused[0] in ('a', 'b') and set(fused[:2]) == set(['a', 'b']), fused)
    check('rrf keeps all keys', set(fused) == set(['a', 'b', 'c', 'd']))
    fused_w = rrf_fuse([['a'], ['b']], weights=[2.0, 1.0])
    check('rrf weights', fused_w[0] == 'a', fused_w)

    # fake transport: /api/tags lists the model; /api/embed batches
    calls = {}

    def fake_get(url, timeout_ms=0):
        calls['get'] = url
        return json.dumps({'models': [{'name': 'nomic-embed-text:latest'}]})

    def fake_post(url, payload, headers=None, timeout_ms=0):
        calls['post'] = url
        if url.endswith('/api/embed'):
            return json.dumps({'embeddings': [[0.123456, 0.2]] * len(payload['input'])})
        return None

    emb = OllamaEmbedder(http_post_fn=fake_post, http_get_fn=fake_get,
                         host='http://x:11434')
    check('embedder available', emb.is_available() is True)
    vecs = emb.embed(['one', 'two'])
    check('batch embed', vecs is not None and len(vecs) == 2)
    check('vector rounding', vecs and vecs[0][0] == 0.1235, vecs and vecs[0][0])

    # fallback path: /api/embed missing → per-text /api/embeddings
    def fake_post2(url, payload, headers=None, timeout_ms=0):
        if url.endswith('/api/embed'):
            return None
        return json.dumps({'embedding': [0.5, 0.5]})

    emb2 = OllamaEmbedder(http_post_fn=fake_post2, http_get_fn=fake_get,
                          host='http://x:11434')
    vecs2 = emb2.embed(['one'])
    check('fallback per-text embed', vecs2 == [[0.5, 0.5]], vecs2)

    # unreachable host → unavailable, embed returns None
    def dead_get(url, timeout_ms=0):
        return None

    emb3 = OllamaEmbedder(http_post_fn=fake_post, http_get_fn=dead_get,
                          host='http://x:11434')
    check('unavailable when no host', emb3.is_available() is False)
    check('embed None when unavailable', emb3.embed(['x']) is None)


# ─── knowledge_store ──────────────────────────────────────────────────────────

def test_knowledge_store():
    print('[knowledge_store]')
    import tempfile, shutil, time as _time
    from Intelligence.knowledge.knowledge_store import KnowledgeStore

    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, 'docs')
    os.makedirs(src)
    try:
        with open(os.path.join(src, 'standard.md'), 'wb') as f:
            f.write('Chiều cao lan can ban công tối thiểu 1100 mm.\n'
                    'Cửa thoát hiểm rộng tối thiểu 800 mm.'.encode('utf-8'))
        with open(os.path.join(src, 'notes.txt'), 'wb') as f:
            f.write(b'Handrail height for stairs is 900 mm from nosing.')
        with open(os.path.join(src, 'skip.docx'), 'wb') as f:
            f.write(b'not indexable')

        store = KnowledgeStore(os.path.join(tmp, 'idx'), [src], 'test')
        r1 = store.scan()
        check('scan added 2', r1['added'] == 2, r1)

        r2 = store.scan()
        check('rescan unchanged', r2['unchanged'] == 2 and r2['added'] == 0, r2)

        hits = store.search('chiều cao lan can', top_k=3)
        check('store search hit', len(hits) >= 1)
        check('citation fields', hits and hits[0]['file'] == 'standard.md'
              and hits[0]['page'] == 0 and 'lan can' in hits[0]['text'])

        st = store.stats()
        check('stats', st['files'] == 2 and st['chunks'] == 2, st)

        # attachment indexing + reload from disk (fresh store instance)
        att = os.path.join(tmp, 'attach.txt')
        with open(att, 'wb') as f:
            f.write(b'Concrete cover for beams shall be 25 mm minimum.')
        entry = store.index_file(att)
        check('attachment indexed', entry and entry['chunks'] == 1, entry)

        store2 = KnowledgeStore(os.path.join(tmp, 'idx'), [src], 'test')
        hits2 = store2.search('concrete cover beams', top_k=2)
        check('persisted index reload', hits2 and hits2[0]['file'] == 'attach.txt')

        # file change → rescan reindexes; file removal → entry dropped
        _time.sleep(0.02)
        with open(os.path.join(src, 'standard.md'), 'ab') as f:
            f.write(b'\nExtra line about ramp slope 1:12 maximum.')
        os.utime(os.path.join(src, 'standard.md'), None)
        r3 = store2.scan()
        check('changed file reindexed', r3['updated'] == 1, r3)
        os.remove(os.path.join(src, 'notes.txt'))
        r4 = store2.scan()
        check('deleted file removed', r4['removed'] == 1, r4)

        # hybrid channel: fake embedder → embed_pending → fused search
        class FakeEmbedder(object):
            MODEL = 'fake'

            def is_available(self, recheck=False):
                return True

            def embed(self, texts):
                out = []
                for t in texts:
                    lowered = t.lower()
                    out.append([1.0, 0.0] if ('lan can' in lowered
                                              or 'chieu cao' in lowered)
                               else [0.0, 1.0])
                return out

        n_emb = store2.embed_pending(FakeEmbedder())
        check('embed_pending vectorized', n_emb >= 2, n_emb)
        st2 = store2.stats(include_vectors=True)
        check('stats vectors', st2['vectors'] == n_emb, st2)
        hits3 = store2.search('chiều cao lan can', top_k=2,
                              embedder=FakeEmbedder())
        check('hybrid search works', hits3 and hits3[0]['file'] == 'standard.md',
              hits3 and hits3[0]['file'])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ─── dispatcher (keyword stage) ───────────────────────────────────────────────

def test_dispatcher():
    print('[dispatcher]')
    from Intelligence.agents.dispatcher import AgentDispatcher

    d = AgentDispatcher()

    def label(text, **kw):
        return d.classify(text, allow_llm=False, **kw)['specialist']

    check('vi count → revit_data',
          label('Có bao nhiêu bức tường trong dự án?') == 'revit_data')
    check('en count → revit_data',
          label('how many walls are there') == 'revit_data')
    check('liet ke → revit_data',
          label('liệt kê các sheet') == 'revit_data')
    check('rename → revit_action',
          label('đổi tên sheet A-101 thành A-102') == 'revit_action')
    check('delete → revit_action',
          label('xóa các text note trong view này') == 'revit_action')
    check('export → revit_action',
          label('xuất pdf toàn bộ sheet') == 'revit_action')
    check('standard question → knowledge',
          label('tiêu chuẩn chiều cao lan can là bao nhiêu?') == 'knowledge')
    check('doc question → knowledge',
          label('trong tài liệu có nói về cấp chống cháy không') == 'knowledge')
    check('action beats knowledge',
          label('sửa chiều cao lan can theo tiêu chuẩn') == 'revit_action')
    check('cmt → comment',
          label('hoàn thiện các cmt trong bản vẽ') == 'comment')
    check('bluebeam → comment',
          label('xử lý markup bluebeam') == 'comment')
    check('annotated pdf attach → comment',
          d.classify('xem giúp file này', attached_pdf_annotated=True,
                     allow_llm=False)['specialist'] == 'comment')
    check('greeting → general', label('chào bạn, khỏe không?') == 'general')
    check('ban ve not draw-verb',
          label('bản vẽ này thuộc model nào') == 'general')
    check('empty → general', label('') == 'general')

    # LLM stage with a fake provider
    class FakeProvider(object):
        def chat(self, messages, system, user, max_tokens=60, **kw):
            return '{"label": "revit_data", "skill": null}'

    r = d.classify('mấy cái đó nằm đâu', provider=FakeProvider(),
                   allow_llm=True)
    check('llm stage used for ambiguous',
          r['specialist'] == 'revit_data' and r['source'] == 'llm', r)

    class BadProvider(object):
        def chat(self, *a, **kw):
            return 'not json at all'

    r2 = d.classify('mấy cái đó nằm đâu', provider=BadProvider(),
                    allow_llm=True)
    check('bad llm → general default',
          r2['specialist'] == 'general' and r2['source'] == 'default', r2)


# ─── knowledge_agent ──────────────────────────────────────────────────────────

def test_knowledge_agent():
    print('[knowledge_agent]')
    import tempfile, shutil
    from Intelligence.knowledge.knowledge_store import KnowledgeStore
    from Intelligence.knowledge.knowledge_agent import (
        KnowledgeAgent, format_citation_line)

    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, 'docs')
    os.makedirs(src)
    try:
        with open(os.path.join(src, 'fire.md'), 'wb') as f:
            f.write('Cửa thoát hiểm phải có chiều rộng tối thiểu 800 mm '
                    'và mở theo chiều thoát nạn.'.encode('utf-8'))
        store = KnowledgeStore(os.path.join(tmp, 'idx'), [src], 'test')
        store.scan()

        agent = KnowledgeAgent(store=store, fallback_store=None)
        seen = {}

        def chat_fn(system_prompt, query):
            seen['system'] = system_prompt
            seen['query'] = query
            return 'Chiều rộng tối thiểu là 800 mm [1].'

        res = agent.answer('cửa thoát hiểm rộng bao nhiêu?', [], chat_fn)
        check('agent done', res['status'] == 'done', res)
        check('citation built', res['citations'] and
              res['citations'][0]['file'] == 'fire.md', res.get('citations'))
        check('excerpt in query', '800 mm' in seen['query'])
        check('grounding rule in system', 'trích đoạn' in seen['system'])

        line = format_citation_line(res['citations'])
        check('citation line', 'fire.md' in line and 'Nguồn' in line, line)

        res2 = agent.answer('chủ đề hoàn toàn khác biệt xyz', [], chat_fn)
        check('no hits status', res2['status'] == 'no_hits', res2)

        def mute_fn(system_prompt, query):
            return None
        res3 = agent.answer('cửa thoát hiểm', [], mute_fn)
        check('llm_failed status', res3['status'] == 'llm_failed', res3)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ─── main ─────────────────────────────────────────────────────────────────────

def main():
    test_vi_text()
    test_chunker()
    test_bm25()
    test_embeddings()
    test_knowledge_store()
    test_dispatcher()
    test_knowledge_agent()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All knowledge-stack tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
