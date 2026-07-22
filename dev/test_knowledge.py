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

# Sandbox %APPDATA% BEFORE any config/settings import, so settings.json,
# projects/ and skills/ land in a throwaway dir instead of the real one.
import tempfile
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_test_')

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
    check('export → export',
          label('xuất pdf toàn bộ sheet') == 'export')
    check('standard question → knowledge',
          label('tiêu chuẩn chiều cao lan can là bao nhiêu?') == 'knowledge')
    check('doc question → knowledge',
          label('trong tài liệu có nói về cấp chống cháy không') == 'knowledge')
    check('action beats knowledge',
          label('sửa chiều cao lan can theo tiêu chuẩn') == 'revit_action')
    check('to do (color) → revit_action',
          label('tô đỏ tường') == 'revit_action')
    check('boi xanh → revit_action',
          label('bôi xanh các cột tầng 2') == 'revit_action')
    check('en to-do not action', label('tell me what to do') == 'general')
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


# ─── specialists ──────────────────────────────────────────────────────────────

def test_specialists():
    print('[specialists]')
    from Intelligence.agents.specialists import (
        get_spec, SPECIALISTS, READ_TOOLS, ACTION_TOOLS)

    check('registry names', set(SPECIALISTS.keys()) ==
          set(['general', 'revit_data', 'revit_action', 'knowledge',
               'comment', 'multi_doc', 'modeling', 'qa_check', 'export']))
    check('unknown → general', get_spec('nope').name == 'general')

    data = get_spec('revit_data')
    check('data read-only', not data.allows_writes and not data.use_launcher)
    check('data budget', data.max_iterations == 6)
    check('data subset all providers',
          data.tools_for(local=False) == READ_TOOLS and
          data.tools_for(local=True) == READ_TOOLS)
    check('data has no write tools',
          'delete_element' not in READ_TOOLS and 'set_parameter' not in READ_TOOLS)

    act = get_spec('revit_action')
    check('action local subset', act.tools_for(local=True) == ACTION_TOOLS)
    check('action cloud full', act.tools_for(local=False) is None)
    check('action has writes',
          'set_parameter' in ACTION_TOOLS and 'export_dwg' in ACTION_TOOLS)

    gen = get_spec('general')
    check('general default tools', gen.tools_for(local=True) is None)

    # prompt building must not need Revit — uses agent_loop's builder
    from Intelligence.agents.specialists import build_specialist_prompt
    p = build_specialist_prompt(data, 'CTX-HERE',
                                project_instructions='PROJ-RULE',
                                skills_block='## Active skill: X',
                                local=True)
    check('prompt has context', 'CTX-HERE' in p)
    check('prompt has role', 'DATA specialist' in p)
    check('prompt has few-shot (local)', 'Examples' in p)
    check('prompt has project rules', 'PROJ-RULE' in p)
    check('prompt has skills block', 'Active skill' in p)
    p2 = build_specialist_prompt(data, 'CTX', local=False)
    check('no few-shot on cloud', 'Examples' not in p2)


# ─── skills engine ────────────────────────────────────────────────────────────

def test_skills():
    print('[skills]')
    from Intelligence.skills_engine import (
        parse_frontmatter, SkillsEngine, build_skills_block)

    meta, body = parse_frontmatter(
        '---\n'
        'name: my-skill\n'
        'description: Mot skill thu nghiem\n'
        'triggers: ten sheet, doi ten sheet\n'
        'agents: [revit_action, general]\n'
        'tools: rename_element\n'
        '---\n'
        '# Body here\nRule 1.')
    check('fm name', meta.get('name') == 'my-skill')
    check('fm triggers list', meta.get('triggers') == ['ten sheet', 'doi ten sheet'])
    check('fm bracket list', meta.get('agents') == ['revit_action', 'general'])
    check('fm body', body.startswith('# Body here'))

    m2, b2 = parse_frontmatter('no frontmatter at all')
    check('no fence passthrough', m2 == {} and b2 == 'no frontmatter at all')

    # engine over the real built-in skills dir
    engine = SkillsEngine()
    n = engine.scan()
    check('builtin skills found', n >= 3, n)
    ids = [s['id'] for s in engine.get_catalog()]
    check('starter skills present',
          'sheet-naming-standard' in ids and 'qa-checklist' in ids
          and 'comment-resolution-playbook' in ids, ids)

    hits = engine.match('đổi tên sheet A-101 giúp mình')
    check('trigger match (folded)', 'sheet-naming-standard' in hits, hits)
    hits2 = engine.match('xử lý markup bluebeam')
    check('comment skill match', 'comment-resolution-playbook' in hits2, hits2)
    check('no match', engine.match('chào bạn') == [])

    body = engine.get_body('sheet-naming-standard')
    check('body loaded', 'Sheet Number' in body or 'sheet' in body.lower())

    check('filter by agent', engine.filter_for_specialist(
        ['sheet-naming-standard'], 'revit_action') == ['sheet-naming-standard'])
    check('filter blocks wrong agent', engine.filter_for_specialist(
        ['qa-checklist'], 'comment') == [])
    check('skill tools', 'rename_element' in
          engine.tools_for('sheet-naming-standard'))

    block = build_skills_block(['qa-checklist'])
    check('skills block built', 'Active skill' in block and 'QA' in block, block[:60])
    check('empty block', build_skills_block([]) == '')


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


# ─── pdf_annots / sheet_matcher / comment_agent ───────────────────────────────

def _make_pdf(objects):
    parts = [b'%PDF-1.4\n']
    for num, body in objects:
        parts.append('{} 0 obj\n'.format(num).encode('ascii'))
        parts.append(body if isinstance(body, bytes) else body.encode('latin-1'))
        parts.append(b'\nendobj\n')
    parts.append(b'trailer\n<< >>\n%%EOF\n')
    return b''.join(parts)


def _utf16be_hex(text):
    return 'FEFF' + ''.join('%04X' % ord(c) for c in text)


def test_pdf_annots():
    print('[pdf_annots]')
    import tempfile, shutil, zlib
    from Intelligence.comments import pdf_annots

    tmp = tempfile.mkdtemp()
    try:
        # fixture 1: literal + escaped parens + UTF-16BE hex + a Link (skip)
        pdf1 = os.path.join(tmp, 'A-101_MatBang.pdf')
        objs = [
            (1, '<< /Type /Catalog /Pages 2 0 R >>'),
            (2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'),
            (3, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] '
                '/Annots [4 0 R 5 0 R 6 0 R] >>'),
            (4, '<< /Type /Annot /Subtype /Text /Rect [100 700 120 720] '
                '/Contents (Move dimension \\(outside\\) gridline A) '
                '/T (Nguyen B) /Subj (Dim) /M (D:20260718) >>'),
            (5, '<< /Type /Annot /Subtype /FreeText /Rect [10 10 60 40] '
                '/Contents <{}> /T <{}> >>'.format(
                    _utf16be_hex('Sửa kích thước dầm'),
                    _utf16be_hex('Linh'))),
            (6, '<< /Type /Annot /Subtype /Link /Rect [0 0 1 1] >>'),
        ]
        with open(pdf1, 'wb') as f:
            f.write(_make_pdf(objs))

        check('has_annotations true', pdf_annots.has_annotations(pdf1))
        recs, partial = pdf_annots.extract_annotations(pdf1)
        check('two markup annots (Link skipped)', len(recs) == 2, len(recs))
        check('not partial', partial is False)
        r1 = [r for r in recs if r['subtype'] == 'Text'][0]
        check('literal content + escapes',
              r1['content'] == 'Move dimension (outside) gridline A',
              r1['content'])
        check('author + subject', r1['author'] == 'Nguyen B'
              and r1['subject'] == 'Dim')
        check('rect parsed', r1['rect'] == [100.0, 700.0, 120.0, 720.0])
        check('page number', r1['page'] == 1)
        r2 = [r for r in recs if r['subtype'] == 'FreeText'][0]
        check('utf16be content', r2['content'] == 'Sửa kích thước dầm',
              repr(r2['content']))
        check('utf16be author', r2['author'] == 'Linh', repr(r2['author']))

        # fixture 2: annotation packed inside a Flate ObjStm
        inner = ('<< /Type /Annot /Subtype /FreeText /Rect [10 10 50 30] '
                 '/Contents (Trong ObjStm) >>')
        header = '7 0 '
        stream = zlib.compress((header + inner).encode('latin-1'))
        body8 = (b'<< /Type /ObjStm /N 1 /First '
                 + str(len(header)).encode('ascii')
                 + b' /Filter /FlateDecode /Length '
                 + str(len(stream)).encode('ascii')
                 + b' >>\nstream\n' + stream + b'\nendstream')
        pdf2 = os.path.join(tmp, 'objstm.pdf')
        objs2 = [
            (1, '<< /Type /Catalog /Pages 2 0 R >>'),
            (2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'),
            (3, '<< /Type /Page /Parent 2 0 R /Annots [7 0 R] >>'),
            (8, body8),
        ]
        with open(pdf2, 'wb') as f:
            f.write(_make_pdf(objs2))
        recs2, partial2 = pdf_annots.extract_annotations(pdf2)
        check('objstm annot found', len(recs2) == 1 and
              recs2[0]['content'] == 'Trong ObjStm', recs2)
        check('objstm not partial', partial2 is False)

        # fixture 3: no annotations
        pdf3 = os.path.join(tmp, 'plain.pdf')
        with open(pdf3, 'wb') as f:
            f.write(_make_pdf([
                (1, '<< /Type /Catalog /Pages 2 0 R >>'),
                (2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'),
                (3, '<< /Type /Page /Parent 2 0 R >>')]))
        check('has_annotations false', not pdf_annots.has_annotations(pdf3))
        recs3, _p3 = pdf_annots.extract_annotations(pdf3)
        check('no annots empty', recs3 == [])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_sheet_matcher():
    print('[sheet_matcher]')
    from Intelligence.comments import sheet_matcher

    sheets = [
        {'name': 'MAT BANG TANG 1', 'number': 'A-101', 'id': 111},
        {'name': 'MAT DUNG TRUC A', 'number': 'A-201', 'id': 222},
    ]

    cands = sheet_matcher.extract_sheet_candidates(
        [(1, 'DU AN X ... SHEET NO: A-201 ... khung ten')],
        'A-101_MatBang.pdf')
    check('filename candidate first', cands and cands[0].startswith('A-101'),
          cands)
    check('keyword candidate found',
          any(c.replace('-', '') == 'A201' for c in cands), cands)

    m = sheet_matcher.match_sheets(['A-101'], sheets)
    check('exact match', m and m['id'] == 111 and m['score'] == 1.0, m)
    m2 = sheet_matcher.match_sheets(['A101'], sheets)
    check('separator-insensitive', m2 and m2['id'] == 111, m2)
    m3 = sheet_matcher.match_sheets(['Z-999'], sheets, 'khongkhop.pdf')
    check('no match none', m3 is None, m3)
    m4 = sheet_matcher.match_sheets([], sheets, 'Mat Bang Tang 1.pdf')
    check('fuzzy name match', m4 and m4['id'] == 111 and m4['score'] >= 0.55,
          m4)


def test_comment_agent():
    print('[comment_agent]')
    import tempfile, shutil
    from Intelligence.comments.comment_agent import CommentAgent

    tmp = tempfile.mkdtemp()
    try:
        pdf1 = os.path.join(tmp, 'A-101_MatBang.pdf')
        with open(pdf1, 'wb') as f:
            f.write(_make_pdf([
                (1, '<< /Type /Catalog /Pages 2 0 R >>'),
                (2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'),
                (3, '<< /Type /Page /Parent 2 0 R /Annots [4 0 R] >>'),
                (4, '<< /Type /Annot /Subtype /Text /Rect [1 2 3 4] '
                    '/Contents (Doi dim ra ngoai grid) /T (QA) >>')]))

        calls = []

        def fake_execute(name, args):
            calls.append(name)
            if name == 'revit_list_sheets':
                return {'count': 1, 'sheets': [
                    {'name': 'MAT BANG TANG 1', 'number': 'A-101', 'id': 111}]}
            if name == 'list_open_documents':
                return {'documents': [{'title': 'ModelA'}]}
            if name == 'revit_get_project_info':
                return {'name': 'Landmark'}
            return {}

        class FakeProvider(object):
            def chat(self, messages, system, user, max_tokens=0, **kw):
                assert 'a1' in user
                return ('{"items": [{"id": "a1", "action_type": '
                        '"fix_dimension", "description": "Doi dim ra ngoai",'
                        ' "instruction": "Move the dimension on sheet A-101 '
                        'outside gridline"}]}')

        agent = CommentAgent()
        report = agent.analyze(pdf1, fake_execute, FakeProvider(), None)
        check('sheet matched', report['sheet_match']
              and report['sheet_match']['number'] == 'A-101', report['sheet_match'])
        check('model open', report['model_open'] is True)
        check('mcp tools called', 'revit_list_sheets' in calls
              and 'list_open_documents' in calls)
        item = report['items'][0]
        check('proposal merged',
              item['proposal']['action_type'] == 'fix_dimension',
              item['proposal'])
        check('item sheet ref', item['matched_sheet']['id'] == 111)

        run_instr = agent.build_run_instruction(item, report)
        check('run instruction', 'A-101' in run_instr
              and 'Doi dim ra ngoai grid' in run_instr, run_instr[:80])
        note_instr = agent.build_note_instruction(item, report)
        check('note instruction', 'create_text_note' in note_instr
              and 'CMT-1' in note_instr, note_instr[:80])

        md = agent.report_to_markdown(report)
        check('markdown table', '| a1 |' in md and 'A-101' in md)

        # LLM mute → proposals stay manual, pipeline never crashes
        class MuteProvider(object):
            def chat(self, *a, **kw):
                return None
        report2 = agent.analyze(pdf1, fake_execute, MuteProvider(), None)
        check('mute → manual default',
              report2['items'][0]['proposal']['action_type'] == 'manual')

        # PDF without annotations → clean error
        pdf3 = os.path.join(tmp, 'plain.pdf')
        with open(pdf3, 'wb') as f:
            f.write(_make_pdf([
                (1, '<< /Type /Catalog /Pages 2 0 R >>'),
                (2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>'),
                (3, '<< /Type /Page /Parent 2 0 R >>')]))
        report3 = agent.analyze(pdf3, fake_execute, FakeProvider(), None)
        check('no annotations error', report3['error'] == 'no_annotations')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ─── project_store ────────────────────────────────────────────────────────────

def test_project_store():
    print('[project_store]')
    from config.project_store import ProjectStore
    from Intelligence.knowledge.knowledge_store import get_active_store

    ps = ProjectStore()
    meta = ps.create_project('Landmark Tower')
    check('created', meta['id'].startswith('p_'), meta)
    check('listed', any(p['id'] == meta['id'] for p in ps.list_projects()))

    ps.set_active_project(meta['id'])
    check('active id', ps.get_active_project_id() == meta['id'])

    ps.update_project(meta['id'], {'instructions': 'Follow standard ABC.'})
    check('prompt addendum', ps.get_active_prompt_addendum() ==
          'Follow standard ABC.')

    hp = ps.history_path(meta['id'], 'docA')
    check('history path scoped', meta['id'] in hp and hp.endswith('docA.json'))

    store = ps.knowledge_store_for(meta['id'])
    check('store scoped', store is not None and meta['id'] in store.index_dir)

    files_dir = os.path.join(ps.project_dir(meta['id']), 'files')
    with open(os.path.join(files_dir, 'rule.md'), 'wb') as f:
        f.write('Chiều cao lan can 1200 mm áp dụng riêng project này.'
                .encode('utf-8'))
    store.scan()

    astore = get_active_store()
    check('get_active_store → project store', astore is store)
    hits = astore.search('chiều cao lan can')
    check('project-scoped search', hits and hits[0]['file'] == 'rule.md',
          hits and hits[0]['file'])

    ps.delete_project(meta['id'])
    check('delete clears active', ps.get_active_project_id() is None)
    check('active store falls back to global',
          get_active_store() is not store)


# ─── main ─────────────────────────────────────────────────────────────────────

def main():
    test_vi_text()
    test_chunker()
    test_bm25()
    test_embeddings()
    test_knowledge_store()
    test_dispatcher()
    test_specialists()
    test_skills()
    test_knowledge_agent()
    test_project_store()
    test_pdf_annots()
    test_sheet_matcher()
    test_comment_agent()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All knowledge-stack tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
