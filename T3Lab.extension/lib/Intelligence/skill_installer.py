# -*- coding: utf-8 -*-
"""
SkillInstaller — install / update Claude-format skills from a GitHub repo.

The user pastes a repo link and asks for its skills:

    "cài skill từ https://github.com/anthropics/skills"
    "install the skills from github.com/owner/repo/tree/main/document-skills"
    "/skills install https://github.com/owner/repo"
    "/skills update"          → re-pull every skill installed from a repo

The repo is downloaded once as a zipball (no `git` on the machine required),
every Claude skill inside it is found, normalised to the frontmatter
SkillsEngine reads, and written to the user skills folder as
`<id>/SKILL.md` + its bundled resources.

Claude → T3Lab frontmatter compatibility
----------------------------------------
A Claude skill declares only `name` and `description` (plus optional
`allowed-tools`, `license`, `version`) and is activated by the model reading
that description. T3Lab's SkillsEngine activates on literal `triggers`, so an
imported skill with no triggers would never fire. `derive_triggers()` in
skills_engine bridges that: explicit "Triggers: …" lines in the description
or body are lifted out, and the skill name/id is added as a phrase.

`allowed-tools` names CLAUDE CODE tools (Read, Bash, …), not T3Lab Revit
tools — it is preserved verbatim for reference but never written into
`tools:`, so an imported skill stays a reference playbook unless it declares
real T3Lab tools itself.

Everything except `fetch_bytes()` and the two `install_*` writers is pure —
CPython-3 testable, see dev/test_skill_installer.py.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import io
import json
import os
import posixpath
import re
import time
import zipfile

from Intelligence.skills_engine import (parse_frontmatter, derive_triggers,
                                        _user_skills_dir)

__author__ = "Tran Tien Thanh"
__title__ = "Skill Installer"

# ── Limits (a skill pack is instructions + small helpers, never a payload) ──
MAX_ARCHIVE_BYTES = 40 * 1024 * 1024     # refuse a repo bigger than this
MAX_FILE_BYTES = 1 * 1024 * 1024         # per bundled file
MAX_SKILL_BYTES = 8 * 1024 * 1024        # per installed skill folder
MAX_SKILL_FILES = 120                    # per installed skill folder
MAX_SKILLS_PER_REPO = 60

# Never written to disk. These are downloaded from an arbitrary repo; the
# assistant reads skill bodies, it does not run them, and a binary that only
# ever sits in a folder is a liability with no upside.
BLOCKED_EXT = frozenset([
    '.exe', '.dll', '.msi', '.com', '.scr', '.pif', '.bat', '.cmd', '.vbs',
    '.vbe', '.wsf', '.wsh', '.ps1', '.psm1', '.jar', '.reg', '.lnk', '.sys',
    '.cpl', '.hta', '.apk', '.dmg', '.pkg', '.deb', '.rpm',
])

# Archive members that are never part of a skill.
JUNK_PARTS = frozenset([
    '.git', '.github', '.svn', '.hg', 'node_modules', '__macosx',
    '__pycache__', '.venv', 'venv', '.idea', '.vs', 'dist', 'build',
])

SOURCE_FILE = '.t3lab-source.json'

# Folders whose loose *.md files count as one-file skills (Claude also
# supports `skills/<name>.md` for short packs).
_SKILL_DIR_NAMES = ('skills', 'skill')
_NOT_A_SKILL_MD = frozenset([
    'readme.md', 'license.md', 'licence.md', 'changelog.md',
    'contributing.md', 'index.md', 'code_of_conduct.md', 'security.md',
])


# ═════════════════════════════════════════════════════════════════════════════
# 1. Understanding the request
# ═════════════════════════════════════════════════════════════════════════════

_REPO_RE = re.compile(
    r'(?:https?://)?(?:www\.)?github\.com[/:]'
    r'(?P<owner>[A-Za-z0-9][A-Za-z0-9._\-]*)/'
    r'(?P<repo>[A-Za-z0-9][A-Za-z0-9._\-]*?)(?:\.git)?'
    r'(?:/(?:tree|blob)/(?P<ref>[^/\s#?]+)(?P<subdir>(?:/[^\s#?]+)*))?'
    r'(?=[\s,;)\'"\]]|$)', re.I)

_SSH_RE = re.compile(
    r'git@github\.com:(?P<owner>[A-Za-z0-9][A-Za-z0-9._\-]*)/'
    r'(?P<repo>[A-Za-z0-9][A-Za-z0-9._\-]*?)(?:\.git)?(?=[\s,;)]|$)', re.I)

# Verbs that mean "put these on my machine", EN + VI (diacritics and folded).
_INSTALL_WORDS = (
    'install', 'add', 'import', 'load', 'download', 'fetch', 'pull', 'get',
    'update', 'upgrade', 'sync', 'refresh', 'reinstall',
    'cai', 'cài', 'cai dat', 'cài đặt', 'cai them', 'them', 'thêm',
    'tai', 'tải', 'tai ve', 'tải về', 'nap', 'nạp', 'nhap', 'nhập',
    'cap nhat', 'cập nhật', 'nang cap', 'nâng cấp', 'dong bo', 'đồng bộ',
)
_SKILL_WORDS = ('skill', 'skills', 'ky nang', 'kỹ năng', 'kynang')


def _fold(text):
    """Lowercase + strip Vietnamese diacritics (shared folding rules)."""
    try:
        from Intelligence.knowledge import vi_text
        folded = vi_text.fold_diacritics(text or '')
    except Exception:
        folded = text or ''
    return re.sub(r'\s+', ' ', folded.lower()).strip()


def parse_repo_url(text):
    """First GitHub repo reference in `text`, or None.

    Returns {'owner', 'repo', 'ref', 'subdir', 'url'}. `ref` is None for the
    default branch; `subdir` ('' when absent) limits the install to skills
    under that path, so a /tree/<branch>/<folder> link installs just that
    folder.
    """
    if not text:
        return None
    m = _REPO_RE.search(text) or _SSH_RE.search(text)
    if not m:
        return None
    groups = m.groupdict()
    owner = groups.get('owner') or ''
    repo = (groups.get('repo') or '').rstrip('.')
    if not owner or not repo:
        return None
    subdir = (groups.get('subdir') or '').strip('/')
    return {
        'owner': owner,
        'repo': repo,
        'ref': groups.get('ref') or None,
        'subdir': subdir,
        'url': 'https://github.com/{}/{}'.format(owner, repo),
    }


def is_skill_install_request(text):
    """True when the message asks to install/update skills from a repo.

    Needs all three signals — an install verb, the word "skill", and a GitHub
    link — so pasting a repo link while discussing it never triggers a write.
    """
    if not text:
        return False
    folded = _fold(text)
    if not any(w in folded for w in ('skill', 'ky nang', 'kynang')):
        return False
    if not any(_word_in(folded, w) for w in _INSTALL_WORDS):
        return False
    return parse_repo_url(text) is not None


def _word_in(folded, word):
    w = _fold(word)
    if not w:
        return False
    return re.search(r'(?:^|[^a-z0-9])' + re.escape(w) + r'(?:[^a-z0-9]|$)',
                     folded) is not None


def parse_skills_command(text):
    """Read a skills-management request. Returns (action, argument) or None.

    Actions: 'install' (argument = repo dict), 'update' (None), 'list' (None),
    'remove' (skill id), 'help' (None).

    Accepts the explicit command form (`/skills install <url>`) and the
    natural-language form ("cài skill từ <url>", "update my skills").
    """
    raw = (text or '').strip()
    if not raw:
        return None

    m = re.match(r'^/skills?\b\s*(.*)$', raw, re.S | re.I)
    if m:
        rest = (m.group(1) or '').strip()
        folded = _fold(rest)
        if not rest:
            return ('list', None)
        if parse_repo_url(rest):
            return ('install', parse_repo_url(rest))
        if folded.split(' ')[0] in ('update', 'upgrade', 'sync', 'refresh'):
            return ('update', None)
        if folded.split(' ')[0] in ('list', 'ls', 'show'):
            return ('list', None)
        if folded.split(' ')[0] in ('remove', 'delete', 'uninstall', 'rm'):
            arg = rest.split(None, 1)
            return ('remove', arg[1].strip() if len(arg) > 1 else None)
        return ('help', None)

    if is_skill_install_request(raw):
        return ('install', parse_repo_url(raw))

    folded = _fold(raw)
    if (any(w in folded for w in ('skill', 'ky nang'))
            and any(_word_in(folded, w) for w in
                    ('update', 'upgrade', 'cap nhat', 'nang cap', 'dong bo'))
            and any(_word_in(folded, w) for w in
                    ('all', 'tat ca', 'toan bo', 'moi', 'installed', 'da cai'))):
        return ('update', None)
    return None


# ═════════════════════════════════════════════════════════════════════════════
# 2. Downloading
# ═════════════════════════════════════════════════════════════════════════════

def zipball_url(owner, repo, ref=None):
    """GitHub zipball endpoint. Without `ref` this is the default branch,
    so a plain repo link works without knowing whether it uses main/master."""
    base = 'https://api.github.com/repos/{}/{}/zipball'.format(owner, repo)
    return base + ('/' + ref if ref else '')


def _fetch_dotnet(url, timeout):
    """.NET download — the reliable path inside Revit's IronPython."""
    import clr  # noqa: F401  (IronPython only)
    from System.Net import (HttpWebRequest, ServicePointManager,
                            SecurityProtocolType)
    from System.IO import MemoryStream
    try:
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    except Exception:
        pass
    req = HttpWebRequest.Create(url)
    req.Method = 'GET'
    req.UserAgent = 'T3Lab-Assistant'
    req.Accept = 'application/vnd.github+json'
    req.Timeout = int(timeout * 1000)
    resp = req.GetResponse()
    try:
        ms = MemoryStream()
        resp.GetResponseStream().CopyTo(ms)
        return bytes(bytearray(ms.ToArray()))
    finally:
        resp.Close()


def _fetch_urllib(url, timeout):
    try:
        from urllib.request import Request, urlopen
    except ImportError:                       # Python 2 / IronPython 2.7
        from urllib2 import Request, urlopen  # noqa
    req = Request(url, headers={'User-Agent': 'T3Lab-Assistant',
                                'Accept': 'application/vnd.github+json'})
    resp = urlopen(req, timeout=timeout)
    try:
        return resp.read()
    finally:
        try:
            resp.close()
        except Exception:
            pass


def fetch_bytes(url, timeout=90):
    """Download `url`. Tries .NET first (Revit), falls back to urllib."""
    errors = []
    for fn in (_fetch_dotnet, _fetch_urllib):
        try:
            data = fn(url, timeout)
            if data:
                return data
            errors.append('{}: empty response'.format(fn.__name__))
        except ImportError as ex:
            errors.append('{}: {}'.format(fn.__name__, ex))
        except Exception as ex:
            errors.append('{}: {}'.format(fn.__name__, ex))
    raise IOError('Download failed — ' + ' | '.join(errors))


# ═════════════════════════════════════════════════════════════════════════════
# 3. Finding skills inside the archive (pure)
# ═════════════════════════════════════════════════════════════════════════════

def _is_junk(name):
    parts = [p.lower() for p in name.split('/') if p]
    return any(p in JUNK_PARTS for p in parts)


def strip_top_dir(names):
    """GitHub zipballs wrap everything in one `owner-repo-sha/` folder.

    Returns that prefix ('' when the archive has no single common root).
    """
    tops = set()
    for n in names:
        head = n.split('/', 1)[0]
        if head:
            tops.add(head)
    if len(tops) == 1:
        only = tops.pop()
        # A single top-level FILE is not a wrapper directory.
        if any(n.startswith(only + '/') for n in names):
            return only + '/'
    return ''


def discover_skills(names, subdir='', default_id='skill'):
    """Locate every skill in a list of archive member paths.

    A skill is either a folder containing SKILL.md (the Claude layout) or a
    loose `<name>.md` inside a `skills/` folder. Returns a list of
    {'id', 'root', 'manifest', 'single_file'} ordered by path, ids deduped.
    """
    prefix = strip_top_dir(names)
    want = _norm_rel(subdir)
    manifests = []
    loose = []
    for name in names:
        if name.endswith('/') or _is_junk(name):
            continue
        rel = name[len(prefix):] if prefix and name.startswith(prefix) else name
        if not rel:
            continue
        if want and not (rel == want or rel.startswith(want + '/')):
            continue
        parts = rel.split('/')
        base = parts[-1].lower()
        if base == 'skill.md':
            manifests.append((name, rel, parts))
        elif (base.endswith('.md') and base not in _NOT_A_SKILL_MD
              and len(parts) >= 2
              and parts[-2].lower() in _SKILL_DIR_NAMES):
            loose.append((name, rel, parts))

    found = []
    seen_ids = set()
    manifest_roots = set()

    for name, rel, parts in sorted(manifests, key=lambda x: (len(x[2]), x[1])):
        root = name[:-len(parts[-1])]                 # keeps the zip prefix
        if any(_is_inside_skill(root, r) for r in manifest_roots):
            continue                                   # a skill's own resource
        manifest_roots.add(root)
        sid = parts[-2] if len(parts) >= 2 else default_id
        found.append({'id': _unique_id(_slug(sid) or default_id, seen_ids),
                      'root': root, 'manifest': name, 'single_file': False})

    for name, rel, parts in sorted(loose, key=lambda x: x[1]):
        root = name[:-len(parts[-1])]
        if root in manifest_roots or any(name.startswith(r)
                                         for r in manifest_roots):
            continue
        sid = parts[-1][:-3]
        found.append({'id': _unique_id(_slug(sid) or default_id, seen_ids),
                      'root': root, 'manifest': name, 'single_file': True})

    return found[:MAX_SKILLS_PER_REPO]


def _is_inside_skill(root, other_root):
    """True when `root` is a resource folder of the skill at `other_root`.

    A SKILL.md below another one is normally a bundled reference, not its own
    skill — except under a `skills/` segment, which is how a repo that is
    itself one skill also ships a library of others.
    """
    if root == other_root or not root.startswith(other_root):
        return False
    tail = root[len(other_root):].strip('/').split('/')
    return not any(p.lower() in _SKILL_DIR_NAMES for p in tail)


def _slug(text):
    """Folder-safe skill id: lowercase, dashes, no separators."""
    s = re.sub(r'[^A-Za-z0-9._\- ]+', '', (text or '')).strip()
    s = re.sub(r'[\s_]+', '-', s).strip('-.').lower()
    return s[:64]


def _unique_id(sid, seen):
    base = sid
    n = 2
    while sid in seen:
        sid = '{}-{}'.format(base, n)
        n += 1
    seen.add(sid)
    return sid


def _norm_rel(path):
    """Normalise an archive-relative path; '' for anything unsafe."""
    p = (path or '').replace('\\', '/').strip('/')
    if not p:
        return ''
    parts = []
    for part in p.split('/'):
        if part in ('', '.'):
            continue
        if part == '..' or ':' in part:
            return ''
        parts.append(part)
    return posixpath.join(*parts) if parts else ''


# ═════════════════════════════════════════════════════════════════════════════
# 4. Claude → T3Lab frontmatter (pure)
# ═════════════════════════════════════════════════════════════════════════════

# Written back in this order so an installed SKILL.md reads like a built-in.
_META_ORDER = ('name', 'description', 'triggers', 'agents', 'tools',
               'standard', 'allowed-tools', 'license', 'version', 'source')


def normalize_meta(meta, body, skill_id, source_url=None):
    """Claude frontmatter → the keys SkillsEngine reads. Returns (meta, notes).

    Only fills gaps: an authored `triggers`/`agents`/`tools` line always wins.
    """
    meta = dict(meta or {})
    notes = []

    name = (meta.get('name') or '').strip() or skill_id.replace('-', ' ')
    meta['name'] = name

    desc = re.sub(r'\s+', ' ', (meta.get('description') or '')).strip()
    if not desc:
        for line in (body or '').splitlines():
            line = line.strip().lstrip('#').strip()
            if line:
                desc = line[:200]
                break
        notes.append('description taken from the body')
    meta['description'] = desc

    if not meta.get('triggers'):
        derived = derive_triggers(skill_id, name, desc, body)
        if derived:
            meta['triggers'] = derived
            notes.append('triggers derived from the description')
        else:
            meta['triggers'] = [skill_id.replace('-', ' ')]
            notes.append('no triggers found — only /{} activates it'
                         .format(skill_id))

    # `allowed-tools` names Claude Code tools, not T3Lab Revit tools: kept for
    # reference, never promoted to `tools:`.
    if meta.get('allowed-tools') and not meta.get('tools'):
        notes.append('allowed-tools kept as reference (Claude Code tools are '
                     'not T3Lab tools)')

    if source_url:
        meta['source'] = source_url
    return meta, notes


def render_skill_md(meta, body):
    """Serialise normalised meta + body back into a SKILL.md."""
    lines = ['---']
    written = set()
    for key in _META_ORDER:
        if key not in meta:
            continue
        written.add(key)
        lines.append('{}: {}'.format(key, _render_value(meta[key])))
    for key in sorted(meta):
        if key in written or key.startswith('_'):
            continue
        lines.append('{}: {}'.format(key, _render_value(meta[key])))
    lines.append('---')
    lines.append('')
    lines.append((body or '').strip())
    lines.append('')
    return '\n'.join(lines)


def _render_value(value):
    if isinstance(value, (list, tuple)):
        return ', '.join(_one_line(v) for v in value)
    return _one_line(value)


def _one_line(value):
    return re.sub(r'\s+', ' ', '{}'.format(value)).strip()


# ═════════════════════════════════════════════════════════════════════════════
# 5. Installing
# ═════════════════════════════════════════════════════════════════════════════

def _blank_report():
    return {'installed': [], 'updated': [], 'skipped': [], 'notes': [],
            'errors': [], 'repo': None, 'dest': None}


def install_from_zip_bytes(data, source, dest_dir=None, only=None):
    """Install every skill found in a zipball. Returns a report dict.

    `source` is the dict from parse_repo_url(); `only` limits the install to
    a set of skill ids (used by update_all to refresh what is already there).
    """
    report = _blank_report()
    report['repo'] = source.get('url') if source else None
    dest_dir = dest_dir or _user_skills_dir()
    report['dest'] = dest_dir

    if data and len(data) > MAX_ARCHIVE_BYTES:
        report['errors'].append(
            'Archive is {:.1f} MB — over the {} MB limit.'.format(
                len(data) / 1048576.0, MAX_ARCHIVE_BYTES // 1048576))
        return report

    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except Exception as ex:
        report['errors'].append('Not a readable zip archive: {}'.format(ex))
        return report

    try:
        names = zf.namelist()
        skills = discover_skills(names,
                                 subdir=(source or {}).get('subdir') or '',
                                 default_id=_slug((source or {}).get('repo')
                                                  or 'skill'))
        if not skills:
            report['errors'].append(
                'No skills found — expected a SKILL.md folder or a skills/ '
                'folder of .md files.')
            return report
        found_ids = set(s['id'] for s in skills)
        for missing in sorted((only or set()) - found_ids):
            report['notes'].append(
                '{}: no longer in the repo — kept as it is.'.format(missing))
        for skill in skills:
            if only and skill['id'] not in only:
                continue
            others = [s['root'] for s in skills if s['root'] != skill['root']]
            try:
                _install_one(zf, skill, dest_dir, source, report, others)
            except Exception as ex:
                report['skipped'].append((skill['id'], '{}'.format(ex)))
    finally:
        try:
            zf.close()
        except Exception:
            pass
    return report


def _install_one(zf, skill, dest_dir, source, report, other_roots=()):
    """Write one discovered skill into `dest_dir/<id>/`."""
    sid = skill['id']
    raw = zf.read(skill['manifest'])
    text = raw.decode('utf-8', 'replace') if isinstance(raw, bytes) else raw
    meta, body = parse_frontmatter(text)
    if not (body or '').strip():
        report['skipped'].append((sid, 'empty skill body'))
        return

    repo_path = skill['manifest']
    top = strip_top_dir(zf.namelist())
    if top and repo_path.startswith(top):
        repo_path = repo_path[len(top):]

    source_url = None
    if source:
        source_url = '{}/blob/{}/{}'.format(
            source['url'], source.get('ref') or 'HEAD', repo_path)
    meta, notes = normalize_meta(meta, body, sid, source_url)
    for n in notes:
        report['notes'].append('{}: {}'.format(sid, n))

    target = os.path.join(dest_dir, sid)
    existed = os.path.isdir(target)
    # remove_installed() refuses to touch a folder with no source stamp,
    # treating that as "the user's own skill, not ours to delete". Installing
    # must honour the same line: a hand-authored skill that happens to share
    # an id/slug with one from this repo must not be silently overwritten.
    if existed and not os.path.isfile(os.path.join(target, SOURCE_FILE)):
        report['skipped'].append(
            (sid, 'a skill folder with this id already exists and was not '
                  'installed from a repo — rename it or remove it manually '
                  'before installing "{}".'.format(sid)))
        return
    if not os.path.isdir(target):
        os.makedirs(target)

    with io.open(os.path.join(target, 'SKILL.md'), 'w',
                 encoding='utf-8') as f:
        f.write(render_skill_md(meta, body))

    n_files, n_bytes = 1, 0
    if not skill['single_file']:
        n_files, n_bytes = _copy_resources(zf, skill, target, sid, report,
                                           other_roots)

    _write_source_stamp(target, {
        'repo': (source or {}).get('url'),
        'owner': (source or {}).get('owner'),
        'name': (source or {}).get('repo'),
        'ref': (source or {}).get('ref'),
        'subdir': (source or {}).get('subdir') or '',
        'path': repo_path,
        'skill_id': sid,
        'skill_name': meta.get('name'),
        'installed_at': time.strftime('%Y-%m-%d %H:%M:%S'),
        'files': n_files,
    })

    entry = {'id': sid, 'name': meta.get('name'), 'path': target,
             'files': n_files, 'bytes': n_bytes}
    (report['updated'] if existed else report['installed']).append(entry)


def _copy_resources(zf, skill, target, sid, report, other_roots=()):
    """Copy a skill folder's bundled files (scripts/, references/, assets/).

    Blocked extensions, oversized files and anything that would escape the
    target folder are skipped and reported — nothing here is ever executed.
    `other_roots` keeps a repo-root skill from swallowing the skills nested
    below it.
    """
    root = skill['root']
    nested = [r for r in (other_roots or ()) if r.startswith(root) and r != root]
    n_files, n_bytes = 1, 0
    for info in zf.infolist():
        name = info.filename
        if not name.startswith(root) or name.endswith('/'):
            continue
        if any(name.startswith(r) for r in nested):
            continue
        rel = _norm_rel(name[len(root):])
        if not rel or rel.lower() == 'skill.md' or _is_junk(rel):
            continue
        ext = os.path.splitext(rel)[1].lower()
        if ext in BLOCKED_EXT:
            report['skipped'].append(
                ('{}/{}'.format(sid, rel), 'blocked file type'))
            continue
        if info.file_size > MAX_FILE_BYTES:
            report['skipped'].append(
                ('{}/{}'.format(sid, rel), 'file over 1 MB'))
            continue
        if n_files >= MAX_SKILL_FILES or n_bytes + info.file_size > MAX_SKILL_BYTES:
            report['skipped'].append(
                ('{}/{}'.format(sid, rel), 'skill folder size limit reached'))
            continue
        out = os.path.join(target, *rel.split('/'))
        if not _within(target, out):
            report['skipped'].append(
                ('{}/{}'.format(sid, rel), 'unsafe path'))
            continue
        folder = os.path.dirname(out)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        with open(out, 'wb') as f:
            f.write(zf.read(name))
        n_files += 1
        n_bytes += info.file_size
    return n_files, n_bytes


def _within(base, path):
    base_abs = os.path.abspath(base)
    target = os.path.abspath(path)
    return target == base_abs or target.startswith(base_abs + os.sep)


def _write_source_stamp(target, payload):
    try:
        text = json.dumps(payload, indent=2, ensure_ascii=False)
        if isinstance(text, bytes):          # Python 2 / IronPython
            text = text.decode('utf-8')
        with io.open(os.path.join(target, SOURCE_FILE), 'w',
                     encoding='utf-8') as f:
            f.write(text)
    except Exception:
        pass


def install_from_github(text_or_url, dest_dir=None, fetcher=None):
    """Full install: parse the link, download the zipball, install skills."""
    source = (text_or_url if isinstance(text_or_url, dict)
              else parse_repo_url(text_or_url))
    if not source:
        report = _blank_report()
        report['errors'].append(
            'No GitHub repo link found — expected github.com/<owner>/<repo>.')
        return report
    fetch = fetcher or fetch_bytes
    try:
        data = fetch(zipball_url(source['owner'], source['repo'],
                                 source.get('ref')))
    except Exception as ex:
        report = _blank_report()
        report['repo'] = source['url']
        report['errors'].append('{}'.format(ex))
        return report
    return install_from_zip_bytes(data, source, dest_dir=dest_dir)


# ═════════════════════════════════════════════════════════════════════════════
# 6. Inventory & auto-update
# ═════════════════════════════════════════════════════════════════════════════

def list_installed(dest_dir=None):
    """Every skill in the user folder that came from a repo (its stamp)."""
    dest_dir = dest_dir or _user_skills_dir()
    out = []
    try:
        entries = sorted(os.listdir(dest_dir))
    except Exception:
        return out
    for entry in entries:
        stamp = os.path.join(dest_dir, entry, SOURCE_FILE)
        if not os.path.isfile(stamp):
            continue
        try:
            with io.open(stamp, 'r', encoding='utf-8', errors='replace') as f:
                data = json.loads(f.read())
        except Exception:
            continue
        data['skill_id'] = data.get('skill_id') or entry
        data['dir'] = os.path.join(dest_dir, entry)
        out.append(data)
    return out


def update_all(dest_dir=None, fetcher=None):
    """Re-pull every repo-installed skill. One download per repo.

    Returns a merged report — this is what "/skills update" and the settings
    button run, and what keeps imported Claude skills current.
    """
    dest_dir = dest_dir or _user_skills_dir()
    installed = list_installed(dest_dir)
    report = _blank_report()
    report['dest'] = dest_dir
    if not installed:
        report['notes'].append('No repo-installed skills yet.')
        return report

    groups = {}
    for item in installed:
        repo = item.get('repo')
        if not repo:
            continue
        key = (repo, item.get('ref') or '', item.get('subdir') or '')
        groups.setdefault(key, []).append(item)

    fetch = fetcher or fetch_bytes
    for (repo, ref, subdir), items in sorted(groups.items()):
        source = parse_repo_url(repo)
        if not source:
            report['errors'].append('Unreadable source for {}'.format(repo))
            continue
        source['ref'] = ref or None
        source['subdir'] = subdir
        try:
            data = fetch(zipball_url(source['owner'], source['repo'],
                                     source.get('ref')))
        except Exception as ex:
            report['errors'].append('{}: {}'.format(repo, ex))
            continue
        only = set(i['skill_id'] for i in items)
        sub = install_from_zip_bytes(data, source, dest_dir=dest_dir,
                                     only=only)
        for key in ('installed', 'updated', 'skipped', 'notes', 'errors'):
            report[key].extend(sub.get(key) or [])
    # One repo → name it in the report; several → the summary stays generic.
    if len(groups) == 1:
        report['repo'] = list(groups)[0][0]
    return report


def remove_installed(skill_id, dest_dir=None):
    """Delete a repo-installed skill folder. Returns (ok, message)."""
    dest_dir = dest_dir or _user_skills_dir()
    sid = _slug(skill_id or '')
    target = os.path.join(dest_dir, sid)
    if not sid or not os.path.isdir(target):
        return False, 'No installed skill named "{}".'.format(skill_id)
    if not os.path.isfile(os.path.join(target, SOURCE_FILE)):
        return False, ('"{}" was not installed from a repo — delete it from '
                       'the skills folder instead.'.format(sid))
    import shutil
    try:
        shutil.rmtree(target)
    except Exception as ex:
        return False, 'Could not remove "{}": {}'.format(sid, ex)
    return True, 'Removed skill "{}".'.format(sid)


# ═════════════════════════════════════════════════════════════════════════════
# 7. Reporting
# ═════════════════════════════════════════════════════════════════════════════

def format_report(report, viet=False):
    """Human-readable markdown summary of an install/update run."""
    report = report or {}
    lines = []
    repo = report.get('repo')
    installed = report.get('installed') or []
    updated = report.get('updated') or []
    skipped = report.get('skipped') or []
    errors = report.get('errors') or []
    notes = report.get('notes') or []

    if errors and not (installed or updated):
        head = ('Không cài được skill.' if viet
                else 'Could not install any skill.')
        return '\n'.join([head] + ['• {}'.format(e) for e in errors])

    if viet:
        lines.append('**Đã cài skill từ {}**'.format(repo or 'GitHub'))
    else:
        lines.append('**Skills installed from {}**'.format(repo or 'GitHub'))

    for entry in installed:
        lines.append('• `/{}` — {} ({})'.format(
            entry['id'], entry.get('name') or entry['id'],
            ('mới' if viet else 'new')))
    for entry in updated:
        lines.append('• `/{}` — {} ({})'.format(
            entry['id'], entry.get('name') or entry['id'],
            ('cập nhật' if viet else 'updated')))

    if not (installed or updated):
        lines.append('— ' + ('không có skill nào thay đổi.' if viet
                             else 'nothing changed.'))

    if notes:
        lines.append('')
        lines.append('_' + ('Ghi chú' if viet else 'Notes') + ':_')
        for n in notes[:8]:
            lines.append('• {}'.format(n))
    if skipped:
        lines.append('')
        lines.append('_' + ('Đã bỏ qua' if viet else 'Skipped') + ':_')
        for sid, why in skipped[:8]:
            lines.append('• {} — {}'.format(sid, why))
    if errors:
        lines.append('')
        for e in errors[:5]:
            lines.append('⚠ {}'.format(e))

    lines.append('')
    lines.append('Gõ `/skills` để xem danh sách, `/skills update` để cập nhật.'
                 if viet else
                 'Type `/skills` to list them, `/skills update` to refresh.')
    return '\n'.join(lines)


def format_inventory(items, viet=False):
    """Markdown list of repo-installed skills."""
    if not items:
        return ('Chưa có skill nào cài từ GitHub. Dán link repo và nói '
                '"cài skill từ <link>".' if viet else
                'No skills installed from GitHub yet. Paste a repo link and '
                'say "install skills from <link>".')
    lines = ['**{}**'.format('Skill đã cài từ GitHub' if viet
                             else 'Skills installed from GitHub')]
    for item in items:
        lines.append('• `/{}` — {}  ·  {}'.format(
            item.get('skill_id'),
            item.get('skill_name') or item.get('skill_id'),
            item.get('repo') or '?'))
    lines.append('')
    lines.append('`/skills update` — {}'.format(
        'cập nhật tất cả' if viet else 'refresh them all'))
    return '\n'.join(lines)
