"""
T3LabAI Settings Management
Handles configuration for MCP server and providers
"""

import os
import json


class T3LabAISettings:
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
            'api_keys': {}
        }

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
        """Get API key for a provider"""
        return self._settings.get('api_keys', {}).get(provider_name)

    def set_api_key(self, provider_name, api_key):
        """Set API key for a provider"""
        if 'api_keys' not in self._settings:
            self._settings['api_keys'] = {}
        self._settings['api_keys'][provider_name] = api_key
        self.save_settings()


def get_settings():
    """Get the singleton settings instance"""
    return T3LabAISettings()
