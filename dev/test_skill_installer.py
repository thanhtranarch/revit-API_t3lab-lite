# -*- coding: utf-8 -*-
"""
CPython 3 test harness for Claude-skill compatibility and the GitHub
skill installer.

Run:  python3 dev/test_skill_installer.py
Exit code 0 = all pass. No external test framework — plain asserts, mirroring
dev/test_assistant_routing.py / dev/test_knowledge.py conventions.

Covers the two halves of "install Claude skills from a repo link":

    parsing     real Claude frontmatter (folded scalars, block lists,
                allowed-tools) must survive skills_engine's hand-rolled YAML
                subset, and a skill that declares no `triggers` must still
                get something for match() to fire on
    installing  a zipball is downloaded once, every SKILL.md inside it is
                found, normalised and written to the user skills folder with
                a provenance stamp that /skills update re-pulls

The network is never touched: fetchers are injected and archives are built
in memory with zipfile.
"""
from __future__ import unicode_literals

import io
import os
import sys
import tempfile
import zipfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

# Sandbox %APPDATA% BEFORE importing anything that resolves the user dirs.
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_skilltest_')

from Intelligence import skills_engine as se            # noqa: E402
from Intelligence import skill_installer as si          # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─────────────────────────────────────────────────────────────────────────────
# Fixtures
# ─────────────────────────────────────────────────────────────────────────────

CLAUDE_SKILL = """---
name: pdf-processing
description: >-
  Extracts text and tables from PDF files, fills forms and merges
  documents. Triggers: "extract pdf", "merge pdf", "fill pdf form".
allowed-tools:
  - Read
  - Bash
license: MIT
---

# PDF processing

Use pdfplumber for text, pypdf for page operations.
"""

MINIMAL_SKILL = """---
name: Brand Guidelines
description: Company colours and typography for slide decks.
---

Always use the corporate palette.
"""

T3LAB_SKILL = """---
name: sheet-naming
description: Sheet naming rules
triggers: ten sheet, sheet number
agents: revit_action
tools: rename_element
---

Body.
"""


def _zip(entries):
    """Build an in-memory zipball (GitHub wraps everything in one folder)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for path, content in entries:
            if isinstance(content, bytes):
                zf.writestr(path, content)
            else:
                zf.writestr(path, content.encode('utf-8'))
    return buf.getvalue()


REPO_ZIP = _zip([
    ('owner-repo-abc123/README.md', '# skills repo'),
    ('owner-repo-abc123/skills/pdf-processing/SKILL.md', CLAUDE_SKILL),
    ('owner-repo-abc123/skills/pdf-processing/scripts/extract.py',
     'print("hi")'),
    ('owner-repo-abc123/skills/pdf-processing/references/notes.md', 'notes'),
    ('owner-repo-abc123/skills/brand-guidelines/SKILL.md', MINIMAL_SKILL),
    ('owner-repo-abc123/skills/quick-tip.md', MINIMAL_SKILL),
    ('owner-repo-abc123/.git/config', 'nope'),
    ('owner-repo-abc123/node_modules/pkg/SKILL.md', MINIMAL_SKILL),
])


def _fake_fetcher(data=REPO_ZIP):
    calls = []

    def fetch(url, timeout=90):
        calls.append(url)
        return data
    fetch.calls = calls
    return fetch


def _tmpdir():
    return tempfile.mkdtemp(prefix='t3lab_skills_')


# ─────────────────────────────────────────────────────────────────────────────
# 1. Frontmatter — Claude YAML must parse
# ─────────────────────────────────────────────────────────────────────────────

def test_folded_description():
    meta, body = se.parse_frontmatter(CLAUDE_SKILL)
    check('folded description joins lines',
          meta.get('description', '').startswith('Extracts text and tables')
          and 'fills forms' in meta.get('description', ''),
          repr(meta.get('description')))
    check('folded description is one line',
          '\n' not in meta.get('description', ''))
    check('body survives frontmatter', body.startswith('# PDF processing'),
          repr(body[:40]))


def test_block_list_allowed_tools():
    meta, _ = se.parse_frontmatter(CLAUDE_SKILL)
    check('block list parsed', meta.get('allowed-tools') == ['Read', 'Bash'],
          repr(meta.get('allowed-tools')))
    check('allowed-tools NOT promoted to tools', not meta.get('tools'),
          repr(meta.get('tools')))


def test_legacy_frontmatter_unchanged():
    meta, body = se.parse_frontmatter(T3LAB_SKILL)
    check('legacy triggers still split',
          meta.get('triggers') == ['ten sheet', 'sheet number'],
          repr(meta.get('triggers')))
    check('legacy agents still split', meta.get('agents') == ['revit_action'])
    check('legacy tools still split', meta.get('tools') == ['rename_element'])
    check('legacy name/description', meta.get('name') == 'sheet-naming'
          and meta.get('description') == 'Sheet naming rules')
    check('legacy body', body == 'Body.', repr(body))


def test_builtin_skills_still_parse():
    """Every shipped skill must keep the same meta under the new parser."""
    d = os.path.join(LIB, 'Intelligence', 'skills')
    bad = []
    for entry in sorted(os.listdir(d)):
        if not entry.endswith('.md'):
            continue
        with io.open(os.path.join(d, entry), encoding='utf-8') as f:
            meta, body = se.parse_frontmatter(f.read())
        if not meta.get('name') or not meta.get('description'):
            bad.append(entry + ' (no name/description)')
        elif not meta.get('triggers'):
            bad.append(entry + ' (no triggers)')
        elif not body.strip():
            bad.append(entry + ' (empty body)')
    check('all built-in skills parse', not bad, ', '.join(bad))


def test_quoted_and_no_fence():
    meta, _ = se.parse_frontmatter('---\nname: "Quoted Name"\n---\nbody')
    check('quotes stripped', meta.get('name') == 'Quoted Name',
          repr(meta.get('name')))
    meta2, body2 = se.parse_frontmatter('no fence here')
    check('no fence yields raw body', meta2 == {} and body2 == 'no fence here')


# ─────────────────────────────────────────────────────────────────────────────
# 2. Trigger derivation
# ─────────────────────────────────────────────────────────────────────────────

def test_derive_from_trigger_line():
    meta, body = se.parse_frontmatter(CLAUDE_SKILL)
    trig = se.derive_triggers('pdf-processing', meta['name'],
                              meta['description'], body)
    low = [t.lower() for t in trig]
    check('quoted triggers lifted',
          'extract pdf' in low and 'merge pdf' in low, repr(trig))
    check('skill id added as phrase', 'pdf processing' in low, repr(trig))
    check('derivation is capped', len(trig) <= 12, len(trig))


def test_derive_falls_back_to_name():
    trig = se.derive_triggers('brand-guidelines', 'Brand Guidelines',
                              'Company colours and typography.', '')
    low = [t.lower() for t in trig]
    check('name/id used when no trigger line',
          'brand guidelines' in low, repr(trig))
    check('no prose swept in', all(len(t) <= 60 for t in trig), repr(trig))


def test_engine_matches_imported_skill():
    """End to end: an installed Claude skill activates on a real message."""
    dest = _tmpdir()
    si.install_from_zip_bytes(REPO_ZIP, si.parse_repo_url(
        'https://github.com/owner/repo'), dest_dir=dest)
    engine = se.get_skills_engine()
    engine.scan(extra_dirs=[dest])
    hits = engine.match('please extract pdf tables from this drawing set')
    check('imported skill matches a request', 'pdf-processing' in hits,
          repr(hits))
    check('imported skill is a reference playbook',
          engine.is_reference_skill('pdf-processing'))
    engine.scan()   # restore the shared singleton for later tests


def test_engine_reads_raw_claude_skill():
    """A Claude SKILL.md copied in by hand (no install) still activates."""
    dest = _tmpdir()
    folder = os.path.join(dest, 'pdf-processing')
    os.makedirs(folder)
    with io.open(os.path.join(folder, 'SKILL.md'), 'w',
                 encoding='utf-8') as f:
        f.write(CLAUDE_SKILL)
    engine = se.get_skills_engine()
    engine.scan(extra_dirs=[dest])
    metas = dict((m['id'], m) for m in engine.all_skills())
    meta = metas.get('pdf-processing', {})
    check('raw Claude skill registers', bool(meta), repr(sorted(metas)))
    check('derived triggers flagged', meta.get('triggers_derived') is True,
          repr(meta.get('triggers')))
    check('raw Claude skill matches a request',
          'pdf-processing' in engine.match('extract pdf from the drawings'),
          repr(meta.get('triggers')))
    check('allowed-tools surfaced for the UI',
          meta.get('allowed_tools') == ['Read', 'Bash'],
          repr(meta.get('allowed_tools')))
    engine.scan()   # restore the shared singleton for later tests


# ─────────────────────────────────────────────────────────────────────────────
# 3. Understanding the request
# ─────────────────────────────────────────────────────────────────────────────

def test_parse_repo_url():
    cases = [
        ('https://github.com/anthropics/skills', 'anthropics', 'skills',
         None, ''),
        ('see github.com/owner/repo.git please', 'owner', 'repo', None, ''),
        ('https://github.com/owner/repo/tree/main/document-skills',
         'owner', 'repo', 'main', 'document-skills'),
        ('git@github.com:owner/repo.git', 'owner', 'repo', None, ''),
    ]
    for text, owner, repo, ref, subdir in cases:
        got = si.parse_repo_url(text)
        check('parse {}'.format(text[:44]),
              got and got['owner'] == owner and got['repo'] == repo
              and got['ref'] == ref and got['subdir'] == subdir, repr(got))
    check('non-github text ignored',
          si.parse_repo_url('install skills please') is None)
    check('gitlab ignored',
          si.parse_repo_url('https://gitlab.com/owner/repo') is None)


def test_install_request_detection():
    yes = [
        'cài skill từ https://github.com/anthropics/skills',
        'cai dat skills tu github.com/owner/repo',
        'install the claude skills from https://github.com/owner/repo',
        'update skills from github.com/owner/repo',
        'thêm kỹ năng từ https://github.com/owner/repo',
    ]
    no = [
        'https://github.com/owner/repo',                      # link only
        'what does https://github.com/owner/repo do?',        # no verb
        'install revit 2025',                                 # no skill/link
        'cài đặt pyRevit giúp mình',
        'tóm tắt repo https://github.com/owner/repo cho tôi',  # discussing it
    ]
    for text in yes:
        check('install request: {}'.format(text[:40]),
              si.is_skill_install_request(text))
    for text in no:
        check('not an install request: {}'.format(text[:40]),
              not si.is_skill_install_request(text))


def test_parse_skills_command():
    cases = [
        ('/skills install https://github.com/owner/repo', 'install'),
        ('/skills https://github.com/owner/repo', 'install'),
        ('cài skill từ https://github.com/owner/repo', 'install'),
        ('/skills update', 'update'),
        ('cập nhật tất cả skill', 'update'),
        ('/skills', 'list'),
        ('/skills list', 'list'),
        ('/skills remove pdf-processing', 'remove'),
        ('/skills wat', 'help'),
    ]
    for text, action in cases:
        got = si.parse_skills_command(text)
        check('command {} → {}'.format(text[:36], action),
              got and got[0] == action, repr(got))
    check('plain chat is not a skills command',
          si.parse_skills_command('tô đỏ các tường ở view này') is None)
    check('remove carries the id',
          si.parse_skills_command('/skills remove pdf-processing')[1]
          == 'pdf-processing')


def test_zipball_url():
    check('default branch zipball',
          si.zipball_url('o', 'r') ==
          'https://api.github.com/repos/o/r/zipball')
    check('pinned ref zipball',
          si.zipball_url('o', 'r', 'v1.2') ==
          'https://api.github.com/repos/o/r/zipball/v1.2')


# ─────────────────────────────────────────────────────────────────────────────
# 4. Discovery
# ─────────────────────────────────────────────────────────────────────────────

def test_discover_skills():
    names = zipfile.ZipFile(io.BytesIO(REPO_ZIP)).namelist()
    found = si.discover_skills(names, default_id='repo')
    ids = [s['id'] for s in found]
    check('SKILL.md folders found',
          'pdf-processing' in ids and 'brand-guidelines' in ids, repr(ids))
    check('loose skills/*.md found', 'quick-tip' in ids, repr(ids))
    check('README is not a skill', 'readme' not in ids, repr(ids))
    check('junk folders skipped', 'pkg' not in ids, repr(ids))


def test_discover_respects_subdir():
    names = zipfile.ZipFile(io.BytesIO(REPO_ZIP)).namelist()
    found = si.discover_skills(names, subdir='skills/brand-guidelines')
    check('subdir narrows the install',
          [s['id'] for s in found] == ['brand-guidelines'], repr(found))


def test_strip_top_dir():
    check('zipball wrapper stripped',
          si.strip_top_dir(['a-b-c/x.md', 'a-b-c/y/z.md']) == 'a-b-c/')
    check('no common root → no prefix',
          si.strip_top_dir(['x.md', 'y/z.md']) == '')


# ─────────────────────────────────────────────────────────────────────────────
# 5. Installing
# ─────────────────────────────────────────────────────────────────────────────

def test_install_writes_skills():
    dest = _tmpdir()
    fetch = _fake_fetcher()
    report = si.install_from_github(
        'cài skill từ https://github.com/owner/repo', dest_dir=dest,
        fetcher=fetch)
    ids = [e['id'] for e in report['installed']]
    check('three skills installed', sorted(ids) ==
          ['brand-guidelines', 'pdf-processing', 'quick-tip'], repr(ids))
    check('one download only', len(fetch.calls) == 1, repr(fetch.calls))
    check('SKILL.md written', os.path.isfile(
        os.path.join(dest, 'pdf-processing', 'SKILL.md')))
    check('resources copied', os.path.isfile(
        os.path.join(dest, 'pdf-processing', 'scripts', 'extract.py')))
    check('provenance stamped', os.path.isfile(
        os.path.join(dest, 'pdf-processing', si.SOURCE_FILE)))
    with io.open(os.path.join(dest, 'pdf-processing', 'SKILL.md'),
                 encoding='utf-8') as f:
        meta, body = se.parse_frontmatter(f.read())
    check('installed skill has triggers', bool(meta.get('triggers')),
          repr(meta))
    check('installed skill keeps allowed-tools',
          meta.get('allowed-tools') == ['Read', 'Bash'], repr(meta))
    check('installed skill records its source',
          'github.com/owner/repo' in (meta.get('source') or ''),
          repr(meta.get('source')))
    check('installed body preserved', 'pdfplumber' in body, repr(body[:60]))


def test_reinstall_reports_update():
    dest = _tmpdir()
    si.install_from_github('https://github.com/owner/repo', dest_dir=dest,
                           fetcher=_fake_fetcher())
    report = si.install_from_github('https://github.com/owner/repo',
                                    dest_dir=dest, fetcher=_fake_fetcher())
    check('second run reports updates, not installs',
          not report['installed'] and len(report['updated']) == 3,
          repr(report['installed']) + repr(report['updated']))


def test_install_does_not_overwrite_a_hand_authored_skill():
    """remove_installed() refuses to delete a folder with no source stamp,
    treating that as the user's own skill. Installing must honour the same
    line: a hand-authored skill sharing an id with a repo skill must be left
    alone, not silently overwritten with the repo's content."""
    dest = _tmpdir()
    user_dir = os.path.join(dest, 'pdf-processing')
    os.makedirs(user_dir)
    with io.open(os.path.join(user_dir, 'SKILL.md'), 'w',
                encoding='utf-8') as f:
        f.write(u'---\nname: pdf-processing\ntriggers: ["my own thing"]\n'
                u'---\n\nHand-written body, not from any repo.\n')

    report = si.install_from_github('https://github.com/owner/repo',
                                    dest_dir=dest, fetcher=_fake_fetcher())

    skipped_ids = [sid for sid, _reason in report['skipped']]
    check('the conflicting skill is reported as skipped, not installed',
          'pdf-processing' in skipped_ids, repr(report['skipped']))
    check('the conflicting skill is not reported as installed or updated',
          'pdf-processing' not in [e['id'] for e in report['installed']]
          and 'pdf-processing' not in [e['id'] for e in report['updated']],
          repr(report))
    check('no source stamp was written into the user\'s folder',
          not os.path.isfile(os.path.join(user_dir, si.SOURCE_FILE)))
    with io.open(os.path.join(user_dir, 'SKILL.md'), encoding='utf-8') as f:
        check('the user\'s own SKILL.md content is untouched',
              'Hand-written body' in f.read())


def test_update_all_repulls():
    dest = _tmpdir()
    si.install_from_github('https://github.com/owner/repo', dest_dir=dest,
                           fetcher=_fake_fetcher())
    fetch = _fake_fetcher()
    report = si.update_all(dest_dir=dest, fetcher=fetch)
    check('update re-pulls the repo once', len(fetch.calls) == 1,
          repr(fetch.calls))
    check('update refreshes every installed skill',
          len(report['updated']) == 3, repr(report['updated']))
    check('inventory lists them', len(si.list_installed(dest)) == 3)


def test_update_with_nothing_installed():
    report = si.update_all(dest_dir=_tmpdir(), fetcher=_fake_fetcher())
    check('empty update is not an error',
          not report['errors'] and not report['updated'], repr(report))


def test_blocked_and_unsafe_files():
    data = _zip([
        ('o-r-1/skills/evil/SKILL.md', MINIMAL_SKILL),
        ('o-r-1/skills/evil/payload.exe', b'MZ\x00'),
        ('o-r-1/skills/evil/setup.bat', 'echo pwn'),
        ('o-r-1/skills/evil/notes.md', 'fine'),
    ])
    dest = _tmpdir()
    report = si.install_from_zip_bytes(
        data, si.parse_repo_url('https://github.com/o/r'), dest_dir=dest)
    check('skill still installs', [e['id'] for e in report['installed']]
          == ['evil'], repr(report['installed']))
    check('exe not written',
          not os.path.isfile(os.path.join(dest, 'evil', 'payload.exe')))
    check('bat not written',
          not os.path.isfile(os.path.join(dest, 'evil', 'setup.bat')))
    check('markdown resource written',
          os.path.isfile(os.path.join(dest, 'evil', 'notes.md')))
    check('blocks are reported', len(report['skipped']) == 2,
          repr(report['skipped']))


def test_path_traversal_blocked():
    data = _zip([
        ('o-r-1/skills/tricky/SKILL.md', MINIMAL_SKILL),
        ('o-r-1/skills/tricky/../../../escaped.md', 'nope'),
    ])
    dest = _tmpdir()
    si.install_from_zip_bytes(
        data, si.parse_repo_url('https://github.com/o/r'), dest_dir=dest)
    escaped = os.path.abspath(os.path.join(dest, '..', '..', '..',
                                           'escaped.md'))
    check('traversal did not escape the skills folder',
          not os.path.isfile(escaped), escaped)


def test_root_skill_does_not_swallow_nested_ones():
    """A repo whose root is itself a skill must not absorb skills/ below it."""
    data = _zip([
        ('o-r-1/SKILL.md', MINIMAL_SKILL),
        ('o-r-1/notes.md', 'root resource'),
        ('o-r-1/skills/nested/SKILL.md', CLAUDE_SKILL),
    ])
    dest = _tmpdir()
    report = si.install_from_zip_bytes(
        data, si.parse_repo_url('https://github.com/o/r'), dest_dir=dest)
    ids = sorted(e['id'] for e in report['installed'])
    check('root skill takes the repo name', 'r' in ids, repr(ids))
    check('root resource copied',
          os.path.isfile(os.path.join(dest, 'r', 'notes.md')))
    check('nested skill not copied into the root skill',
          not os.path.isdir(os.path.join(dest, 'r', 'skills')),
          repr(os.listdir(os.path.join(dest, 'r'))))


def test_update_notes_removed_upstream():
    """A skill deleted upstream is reported, not silently forgotten."""
    dest = _tmpdir()
    si.install_from_github('https://github.com/owner/repo', dest_dir=dest,
                           fetcher=_fake_fetcher())
    shrunk = _zip([
        ('owner-repo-def456/skills/pdf-processing/SKILL.md', CLAUDE_SKILL),
    ])
    report = si.update_all(dest_dir=dest, fetcher=_fake_fetcher(shrunk))
    check('surviving skill still updates',
          [e['id'] for e in report['updated']] == ['pdf-processing'],
          repr(report['updated']))
    check('removed skills are called out',
          any('no longer in the repo' in n for n in report['notes']),
          repr(report['notes']))
    check('removed skill left on disk',
          os.path.isdir(os.path.join(dest, 'brand-guidelines')))


def test_archive_without_skills():
    data = _zip([('o-r-1/README.md', '# nothing here')])
    report = si.install_from_zip_bytes(
        data, si.parse_repo_url('https://github.com/o/r'),
        dest_dir=_tmpdir())
    check('empty repo reports an error, installs nothing',
          report['errors'] and not report['installed'], repr(report))


def test_bad_download_is_reported():
    def boom(url, timeout=90):
        raise IOError('404 Not Found')
    report = si.install_from_github('https://github.com/o/r',
                                    dest_dir=_tmpdir(), fetcher=boom)
    check('download failure surfaces',
          any('404' in e for e in report['errors']), repr(report['errors']))
    check('failure report renders',
          'Could not install' in si.format_report(report),
          si.format_report(report)[:80])


def test_report_formatting():
    dest = _tmpdir()
    report = si.install_from_github('https://github.com/owner/repo',
                                    dest_dir=dest, fetcher=_fake_fetcher())
    text_en = si.format_report(report)
    text_vi = si.format_report(report, viet=True)
    check('report lists slash ids', '/pdf-processing' in text_en, text_en[:80])
    check('report is bilingual', 'Đã cài skill' in text_vi, text_vi[:80])
    inv = si.format_inventory(si.list_installed(dest))
    check('inventory renders', '/brand-guidelines' in inv, inv[:80])
    check('empty inventory tells the user what to do',
          'repo link' in si.format_inventory([]))


def test_remove_installed():
    dest = _tmpdir()
    si.install_from_github('https://github.com/owner/repo', dest_dir=dest,
                           fetcher=_fake_fetcher())
    ok, msg = si.remove_installed('pdf-processing', dest_dir=dest)
    check('remove deletes the folder',
          ok and not os.path.isdir(os.path.join(dest, 'pdf-processing')), msg)
    ok2, msg2 = si.remove_installed('not-there', dest_dir=dest)
    check('removing an unknown skill fails cleanly', not ok2, msg2)


TESTS = [
    ('Frontmatter (Claude compatibility)', [
        test_folded_description,
        test_block_list_allowed_tools,
        test_legacy_frontmatter_unchanged,
        test_builtin_skills_still_parse,
        test_quoted_and_no_fence,
    ]),
    ('Trigger derivation', [
        test_derive_from_trigger_line,
        test_derive_falls_back_to_name,
        test_engine_matches_imported_skill,
        test_engine_reads_raw_claude_skill,
    ]),
    ('Request understanding', [
        test_parse_repo_url,
        test_install_request_detection,
        test_parse_skills_command,
        test_zipball_url,
    ]),
    ('Discovery', [
        test_discover_skills,
        test_discover_respects_subdir,
        test_strip_top_dir,
    ]),
    ('Install / update', [
        test_install_writes_skills,
        test_reinstall_reports_update,
        test_install_does_not_overwrite_a_hand_authored_skill,
        test_update_all_repulls,
        test_update_with_nothing_installed,
        test_blocked_and_unsafe_files,
        test_path_traversal_blocked,
        test_root_skill_does_not_swallow_nested_ones,
        test_update_notes_removed_upstream,
        test_archive_without_skills,
        test_bad_download_is_reported,
        test_report_formatting,
        test_remove_installed,
    ]),
]


def main():
    for group, tests in TESTS:
        print('\n{}'.format(group))
        for t in tests:
            try:
                t()
            except Exception as ex:
                FAILURES.append(t.__name__)
                print('  ERROR {}  {}'.format(t.__name__, ex))

    print('')
    if FAILURES:
        print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('all passed')
    return 0


if __name__ == '__main__':
    sys.exit(main())
