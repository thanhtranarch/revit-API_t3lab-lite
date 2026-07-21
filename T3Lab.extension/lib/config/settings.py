# -*- coding: utf-8 -*-
"""
Settings

Configuration settings manager for T3Lab AI integration.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Settings"

import io
import os
import json


class T3LabAISettings(object):
    """Settings manager for T3LabAI"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(T3LabAISettings, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return

        self._settings_file = self._get_settings_path()
        self._settings = self._load_settings()
        self._initialized = True

    def _get_settings_path(self):
        """Get the path to settings file"""
        app_data = os.environ.get('APPDATA', '')
        settings_dir = os.path.join(app_data, 'T3LabAI')
        if not os.path.exists(settings_dir):
            os.makedirs(settings_dir)
        return os.path.join(settings_dir, 'settings.json')

    def _load_settings(self):
        """Load settings from file"""
        if os.path.exists(self._settings_file):
            try:
                with io.open(self._settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass

        return self._get_default_settings()

    def _get_default_settings(self):
        """Get default settings"""
        return {
            'server': {
                'port': 8080,
                'host': 'localhost'
            },
            'providers': [],
            'api_keys': {},
            'active_provider': 'claude',
            'model_preferences': {},
            'username': 'Thạnh',
            'window_state': {
                'left':         None,
                'top':          None,
                'width':        720,
                'height':       580,
                'sidebar_open': False,
            },
            'active_project': None,
            'knowledge': {
                'dirs':               [],
                'embeddings_enabled': True,
                'embed_model':        'nomic-embed-text',
            },
            'agents': {
                'multi_agent':  True,
                'llm_classify': True,
            },
            'skills': {
                'disabled': [],
            },
        }

    def get_window_state(self):
        """Return the last-saved window state dict."""
        defaults = {'left': None, 'top': None,
                    'width': 720, 'height': 580, 'sidebar_open': False}
        saved = self._settings.get('window_state', {})
        defaults.update(saved)
        return defaults

    def save_window_state(self, left, top, width, height, sidebar_open=False):
        """Persist window geometry and sidebar visibility."""
        self._settings['window_state'] = {
            'left':         left,
            'top':          top,
            'width':        width,
            'height':       height,
            'sidebar_open': sidebar_open,
        }
        self.save_settings()

    def save_settings(self):
        """Save settings to file.

        Uses ensure_ascii=True + io.open(utf-8) so non-ASCII values
        (Vietnamese usernames, unicode paths) never break the dump
        under IronPython 2.7.
        """
        try:
            payload = json.dumps(self._settings, indent=2, ensure_ascii=True)
            if isinstance(payload, bytes):
                payload = payload.decode('ascii')
            with io.open(self._settings_file, 'w', encoding='utf-8') as f:
                f.write(payload)
            return True
        except Exception:
            return False

    def get_server_config(self):
        """Get server configuration"""
        return self._settings.get('server', {})

    def get_enabled_providers(self):
        """Get list of enabled providers"""
        return self._settings.get('providers', [])

    def get_api_key(self, provider_name):
        """Get API key for a provider — always reads fresh from the in-memory dict.

        The dict is kept in sync with the file by reload() / set_api_key().
        """
        return self._settings.get('api_keys', {}).get(provider_name)

    def set_api_key(self, provider_name, api_key):
        """Set API key for a provider.

        Reloads the file from disk first so that keys saved by other sessions
        (or other providers) are not accidentally overwritten by stale
        in-memory data.
        """
        # Merge: reload disk → patch → save
        self._settings = self._load_settings()
        if 'api_keys' not in self._settings:
            self._settings['api_keys'] = {}
        self._settings['api_keys'][provider_name] = api_key
        return self.save_settings()

    def get_active_provider(self):
        """Return the name of the last-selected LLM provider ('claude', 'openai', 'ollama')."""
        return self._settings.get('active_provider', 'claude')

    def set_active_provider(self, name):
        """Persist the active provider name."""
        self._settings['active_provider'] = name
        self.save_settings()

    def get_provider_model(self, provider_name):
        """Return the saved model name for a provider, or None."""
        return self._settings.get('model_preferences', {}).get(provider_name)

    def set_provider_model(self, provider_name, model_name):
        """Persist the preferred model name for a provider."""
        if 'model_preferences' not in self._settings:
            self._settings['model_preferences'] = {}
        self._settings['model_preferences'][provider_name] = model_name
        self.save_settings()

    def get_username(self):
        """Return the saved user name, or default 'Thạnh'."""
        return self._settings.get('username', 'Thạnh')

    def set_username(self, username):
        """Persist the user name."""
        self._settings['username'] = username
        self.save_settings()

    # ------------------------------------------------------------------
    # Knowledge directories (RAG sources)
    # ------------------------------------------------------------------

    def get_knowledge_dirs(self):
        """Return the list of user-added knowledge directories."""
        return list(self._settings.get('knowledge', {}).get('dirs', []))

    def add_knowledge_dir(self, path):
        """Add a knowledge directory (reload-merge-save, like set_api_key)."""
        self._settings = self._load_settings()
        know = self._settings.setdefault('knowledge', {})
        dirs = know.setdefault('dirs', [])
        if path not in dirs:
            dirs.append(path)
        return self.save_settings()

    def remove_knowledge_dir(self, path):
        """Remove a knowledge directory."""
        self._settings = self._load_settings()
        know = self._settings.setdefault('knowledge', {})
        dirs = know.setdefault('dirs', [])
        if path in dirs:
            dirs.remove(path)
        return self.save_settings()

    def get_knowledge_option(self, key, default=None):
        """Read a scalar option from the knowledge block."""
        return self._settings.get('knowledge', {}).get(key, default)

    def set_knowledge_option(self, key, value):
        """Persist a scalar option in the knowledge block."""
        self._settings = self._load_settings()
        self._settings.setdefault('knowledge', {})[key] = value
        return self.save_settings()

    # ------------------------------------------------------------------
    # Multi-agent switches
    # ------------------------------------------------------------------

    def is_multi_agent_enabled(self):
        """Kill switch for the specialist dispatcher (default on)."""
        return bool(self._settings.get('agents', {}).get('multi_agent', True))

    def is_llm_classify_enabled(self):
        """Whether the dispatcher may use one small LLM call to classify."""
        return bool(self._settings.get('agents', {}).get('llm_classify', True))

    def get_action_mode(self):
        """Harness action mode for model-editing tools.

        'auto'    = agent executes edits immediately (legacy behavior).
        'confirm' = agent must reply with a plan and wait for the user's OK
                    before any model-modifying tool call.
        """
        mode = self._settings.get('agents', {}).get('action_mode', 'auto')
        return mode if mode in ('auto', 'confirm') else 'auto'

    def get_agent_option(self, key, default=None):
        """Read a scalar switch from the agents block."""
        return self._settings.get('agents', {}).get(key, default)

    def set_agent_option(self, key, value):
        """Persist a switch in the agents block."""
        self._settings = self._load_settings()
        self._settings.setdefault('agents', {})[key] = value
        return self.save_settings()

    # ------------------------------------------------------------------
    # Skills
    # ------------------------------------------------------------------

    def get_disabled_skills(self):
        """Return the list of skill ids the user switched off."""
        return list(self._settings.get('skills', {}).get('disabled', []))

    def set_skill_disabled(self, skill_id, disabled):
        """Toggle a skill on/off (persisted globally)."""
        self._settings = self._load_settings()
        block = self._settings.setdefault('skills', {})
        items = block.setdefault('disabled', [])
        if disabled and skill_id not in items:
            items.append(skill_id)
        elif not disabled and skill_id in items:
            items.remove(skill_id)
        return self.save_settings()

    # ------------------------------------------------------------------
    # Active project
    # ------------------------------------------------------------------

    def get_active_project(self):
        """Return the active project id, or None."""
        return self._settings.get('active_project')

    def set_active_project(self, project_id):
        """Persist the active project id (None = no project)."""
        self._settings = self._load_settings()
        self._settings['active_project'] = project_id
        return self.save_settings()

    def log_model_usage(self, action, provider, model):
        """Log model usage/setup to a log file for audit and fast setup verification."""
        try:
            import datetime
            app_data = os.environ.get('APPDATA', '')
            settings_dir = os.path.join(app_data, 'T3LabAI')
            if not os.path.exists(settings_dir):
                os.makedirs(settings_dir)
            log_file = os.path.join(settings_dir, 'model_setup.log')
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_line = "[{}] Action: {} | Provider: {} | Model: {}\n".format(
                timestamp, action, provider, model
            )
            with open(log_file, 'a') as f:
                f.write(log_line)
        except Exception:
            pass



def get_settings():
    """Get the singleton settings instance"""
    return T3LabAISettings()
