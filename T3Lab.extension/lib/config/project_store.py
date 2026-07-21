# -*- coding: utf-8 -*-
"""
ProjectStore — Claude-style project workspaces for T3Lab Assistant.

A project bundles:
    - its own knowledge scope (projects/<pid>/files/ + extra dirs) with a
      project-scoped RAG index (projects/<pid>/rag_index/)
    - custom instructions injected into every specialist prompt
    - its own chat history (projects/<pid>/chats/<doc_key>.json)
    - optional default provider/model applied on activation

Layout: %APPDATA%/T3LabAI/projects/<pid>/project.json
The active project id lives in settings.json ("active_project", managed
by T3LabAISettings). No project selected = fully legacy behavior.

All JSON writes use the ensure_ascii + io.open(utf-8) pattern
(IronPython 2.7 safe with Vietnamese text/paths).

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import hashlib
import io
import json
import os
import shutil
import threading
import time


def _write_json(path, data):
    payload = json.dumps(data, ensure_ascii=True, indent=2)
    if isinstance(payload, bytes):
        payload = payload.decode('ascii')
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(payload)


def _read_json(path, default=None):
    try:
        with io.open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


class ProjectStore(object):
    """Singleton project registry."""

    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(ProjectStore, cls).__new__(cls)
                cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._store_cache = {}     # pid -> KnowledgeStore
        self._initialized = True

    # ── paths ─────────────────────────────────────────────────────────────

    def _root(self):
        base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
        d = os.path.join(base, 'T3LabAI', 'projects')
        if not os.path.isdir(d):
            try:
                os.makedirs(d)
            except Exception:
                pass
        return d

    def project_dir(self, pid):
        return os.path.join(self._root(), pid)

    def _project_json(self, pid):
        return os.path.join(self.project_dir(pid), 'project.json')

    # ── CRUD ──────────────────────────────────────────────────────────────

    def list_projects(self):
        """[{'id','name','created'}] sorted by name."""
        out = []
        try:
            for entry in os.listdir(self._root()):
                meta = _read_json(self._project_json(entry))
                if meta and meta.get('id'):
                    out.append({'id': meta['id'],
                                'name': meta.get('name', entry),
                                'created': meta.get('created', '')})
        except Exception:
            pass
        out.sort(key=lambda p: (p['name'] or '').lower())
        return out

    def create_project(self, name):
        """Create a project; returns its meta dict."""
        seed = (name or 'project') + str(time.time())
        pid = 'p_' + hashlib.md5(seed.encode('utf-8')).hexdigest()[:8]
        meta = {
            'id': pid,
            'name': name or 'Project',
            'created': time.strftime('%Y-%m-%dT%H:%M:%S'),
            'instructions': '',
            'knowledge_dirs': [],
            'provider': None,
            'model': None,
            'skills_disabled': [],
        }
        pdir = self.project_dir(pid)
        for sub in ('', 'files', 'chats', 'skills'):
            d = os.path.join(pdir, sub) if sub else pdir
            if not os.path.isdir(d):
                try:
                    os.makedirs(d)
                except Exception:
                    pass
        _write_json(self._project_json(pid), meta)
        return meta

    def get_project(self, pid):
        if not pid:
            return None
        return _read_json(self._project_json(pid))

    def update_project(self, pid, patch):
        """Merge `patch` into project.json. Returns the new meta or None."""
        meta = self.get_project(pid)
        if not meta:
            return None
        meta.update(patch or {})
        try:
            _write_json(self._project_json(pid), meta)
        except Exception:
            return None
        return meta

    def delete_project(self, pid):
        """Remove the project folder; clears active_project if it was it."""
        if not pid:
            return False
        try:
            if self.get_active_project_id() == pid:
                self.set_active_project(None)
            self._store_cache.pop(pid, None)
            shutil.rmtree(self.project_dir(pid), ignore_errors=True)
            return True
        except Exception:
            return False

    # ── active project ────────────────────────────────────────────────────

    def get_active_project_id(self):
        try:
            from config.settings import get_settings
            pid = get_settings().get_active_project()
            # stale id (folder deleted outside the app) counts as none
            if pid and not os.path.isfile(self._project_json(pid)):
                return None
            return pid
        except Exception:
            return None

    def set_active_project(self, pid):
        try:
            from config.settings import get_settings
            return get_settings().set_active_project(pid)
        except Exception:
            return False

    # ── attachments archive + daily activity log ──────────────────────────

    def _scope_dir(self, pid):
        """projects/<pid> when a project is given, else the T3LabAI root
        (attachments/logs still work with no project selected)."""
        if pid:
            return self.project_dir(pid)
        base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
        return os.path.join(base, 'T3LabAI')

    def attachments_dir(self, pid=None, day=None):
        """<scope>/attachments/<YYYY-MM-DD>/ — created on demand."""
        day = day or time.strftime('%Y-%m-%d')
        d = os.path.join(self._scope_dir(pid), 'attachments', day)
        if not os.path.isdir(d):
            try:
                os.makedirs(d)
            except Exception:
                pass
        return d

    def archive_attachments(self, paths, pid=None):
        """Copy attached files into today's dated folder of the scope.

        Returns the archived paths (a file that fails to copy keeps its
        original path). Existing names get a _N suffix, never overwritten.
        """
        out = []
        dest_dir = self.attachments_dir(pid)
        for src in (paths or []):
            try:
                name = os.path.basename(src)
                dst = os.path.join(dest_dir, name)
                if os.path.abspath(src) == os.path.abspath(dst):
                    out.append(src)
                    continue
                stem, ext = os.path.splitext(name)
                n = 1
                while os.path.exists(dst):
                    dst = os.path.join(dest_dir,
                                       u'{}_{}{}'.format(stem, n, ext))
                    n += 1
                shutil.copy2(src, dst)
                out.append(dst)
            except Exception:
                out.append(src)
        return out

    def activity_log_path(self, pid=None, day=None):
        """<scope>/logs/<YYYY-MM-DD>.md — folder created on demand."""
        day = day or time.strftime('%Y-%m-%d')
        d = os.path.join(self._scope_dir(pid), 'logs')
        if not os.path.isdir(d):
            try:
                os.makedirs(d)
            except Exception:
                pass
        return os.path.join(d, day + '.md')

    def append_activity(self, text, pid=None):
        """Append one timestamped markdown bullet to today's log. Never raises."""
        try:
            path = self.activity_log_path(pid)
            is_new = not os.path.exists(path)
            with io.open(path, 'a', encoding='utf-8') as f:
                if is_new:
                    f.write(u'# T3Lab Assistant Activity Log — {}\n\n'.format(
                        time.strftime('%Y-%m-%d')))
                f.write(u'- **{}** {}\n'.format(
                    time.strftime('%H:%M'), text or u''))
            return True
        except Exception:
            return False

    # ── scoped resources ──────────────────────────────────────────────────

    def history_path(self, pid, doc_key):
        """Project-scoped chat history file for a Revit document."""
        d = os.path.join(self.project_dir(pid), 'chats')
        if not os.path.isdir(d):
            try:
                os.makedirs(d)
            except Exception:
                pass
        return os.path.join(d, '{}.json'.format(doc_key))

    def knowledge_store_for(self, pid):
        """Project-scoped KnowledgeStore (cached). None on failure."""
        meta = self.get_project(pid)
        if not meta:
            return None
        store = self._store_cache.get(pid)
        # lazy import — knowledge_store imports this module for
        # get_active_store(), keep the cycle import-time safe
        from Intelligence.knowledge.knowledge_store import KnowledgeStore
        dirs = [os.path.join(self.project_dir(pid), 'files')]
        for d in meta.get('knowledge_dirs', []):
            if d and d not in dirs:
                dirs.append(d)
        if store is None:
            store = KnowledgeStore(
                os.path.join(self.project_dir(pid), 'rag_index'),
                dirs, scope_label=pid)
            self._store_cache[pid] = store
        else:
            store.source_dirs = dirs
        return store

    def get_prompt_addendum(self, pid):
        meta = self.get_project(pid)
        return (meta or {}).get('instructions', '') or ''

    def get_active_prompt_addendum(self):
        """Instructions of the active project, or ''. Never raises."""
        try:
            pid = self.get_active_project_id()
            if not pid:
                return ''
            return self.get_prompt_addendum(pid)
        except Exception:
            return ''
