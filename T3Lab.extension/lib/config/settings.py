# -*- coding: utf-8 -*-
"""
Settings

Configuration settings manager for T3Lab AI integration.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Settings"

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
                with open(self._settings_file, 'r') as f:
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
        """Save settings to file"""
        try:
            with open(self._settings_file, 'w') as f:
                json.dump(self._settings, f, indent=2)
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
