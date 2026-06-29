# -*- coding: utf-8 -*-
"""
MCPService — Shared MCP control backend

Single source of truth for all MCP-related operations:
  • MCP HTTP server (start / stop / status)
  • File-based task watcher (start / stop / status)
  • Claude Desktop config snippet generation
  • Data directory management

Import this from any dialog or tool that needs MCP control:
    from Services.mcp_service import MCPService
"""

from __future__ import unicode_literals

import os
import sys

# ─── Path helper ───────────────────────────────────────────────────────────────
def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))   # lib/Services
    lib_dir = os.path.dirname(here)                        # lib
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)


# ─── Lazy imports (avoid crashing if running outside Revit) ────────────────────
def _get_server():
    _ensure_lib_in_path()
    from core.server import get_t3labai_server
    return get_t3labai_server()


def _get_watcher():
    _ensure_lib_in_path()
    from core.file_watcher import get_task_watcher
    return get_task_watcher()


def _get_data_dir():
    _ensure_lib_in_path()
    try:
        from core.file_watcher import T3LAB_DATA_DIR
        return T3LAB_DATA_DIR
    except Exception:
        return os.path.join(os.path.expanduser('~'), 'T3Lab_AI_Data')


def _get_bridge_path():
    """Absolute path to core/bridge.py, used in the Claude Desktop snippet."""
    here        = os.path.dirname(os.path.abspath(__file__))   # lib/Services
    lib_dir     = os.path.dirname(here)                        # lib
    bridge_path = os.path.join(lib_dir, 'core', 'bridge.py')
    return bridge_path.replace('\\', '/')


# ─── MCPService ────────────────────────────────────────────────────────────────

class MCPService(object):
    """
    Stateless helper that wraps MCP server + file watcher operations.

    All methods return simple dicts so callers don't need to handle exceptions
    — failures come back as {'error': '<message>'}.

    Typical usage in a dialog:
        status = MCPService.server_status()
        if not status.get('error'):
            print(status['running'], status['port'])

        ok, err = MCPService.start_server()
        ok, err = MCPService.stop_server()

        ok, err = MCPService.start_watcher()
        ok, err = MCPService.stop_watcher()

        snippet = MCPService.config_snippet()
    """

    # ── MCP HTTP server ────────────────────────────────────────────────────────

    @staticmethod
    def server_status():
        """
        Return status of the MCP HTTP server.

        Returns:
            dict with keys: running (bool), port (int), tools_count (int),
                            commands_processed (int), error (str|None)
        """
        try:
            server = _get_server()
            stats  = server.get_server_stats()
            return {
                'running':             stats.get('running', False),
                'port':                stats.get('port', 48884),
                'tools_count':         stats.get('tools_count', 0),
                'commands_processed':  stats.get('commands_processed', 0),
                'error':               None,
            }
        except Exception as ex:
            return {'running': False, 'port': 48884, 'tools_count': 0,
                    'commands_processed': 0, 'error': str(ex)}

    @staticmethod
    def start_server(port=None):
        """
        Start the MCP HTTP server.

        Args:
            port (int|None): Override port. Uses server default if None.

        Returns:
            (success: bool, error_message: str)
        """
        try:
            server = _get_server()
            if port:
                server.port = int(port)
            ok = server.start_server()
            if ok:
                return True, None
            return False, 'start_server() returned False'
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def stop_server():
        """
        Stop the MCP HTTP server.

        Returns:
            (success: bool, error_message: str)
        """
        try:
            server = _get_server()
            ok = server.stop_server()
            return (True, None) if ok else (False, 'stop_server() returned False')
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def toggle_server(current_port=None):
        """
        Start server if stopped, stop it if running.

        Returns:
            (new_state: 'running'|'stopped', error_message: str|None)
        """
        status = MCPService.server_status()
        if status.get('error'):
            return 'unknown', status['error']
        if status['running']:
            ok, err = MCPService.stop_server()
            return ('stopped' if ok else 'running'), err
        else:
            ok, err = MCPService.start_server(port=current_port)
            return ('running' if ok else 'stopped'), err

    # ── File watcher ───────────────────────────────────────────────────────────

    @staticmethod
    def watcher_status():
        """
        Return status of the file-based task watcher.

        Returns:
            dict with keys: running (bool), data_dir (str),
                            has_ext_event (bool), error (str|None)
        """
        try:
            watcher = _get_watcher()
            info    = watcher.get_status()
            info['error'] = None
            return info
        except Exception as ex:
            return {'running': False, 'data_dir': _get_data_dir(),
                    'has_ext_event': False, 'error': str(ex)}

    @staticmethod
    def start_watcher():
        """
        Start the file task watcher.

        Returns:
            (success: bool, error_message: str)
        """
        try:
            watcher = _get_watcher()
            ok = watcher.start()
            return (True, None) if ok else (False, 'start() returned False')
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def stop_watcher():
        """
        Stop the file task watcher.

        Returns:
            (success: bool, error_message: str)
        """
        try:
            watcher = _get_watcher()
            watcher.stop()
            return True, None
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def toggle_watcher():
        """
        Start watcher if stopped, stop it if running.

        Returns:
            (new_state: 'running'|'stopped', error_message: str|None)
        """
        status = MCPService.watcher_status()
        if status.get('error'):
            return 'unknown', status['error']
        if status['running']:
            ok, err = MCPService.stop_watcher()
            return ('stopped' if ok else 'running'), err
        else:
            ok, err = MCPService.start_watcher()
            return ('running' if ok else 'stopped'), err

    # ── Config & paths ─────────────────────────────────────────────────────────

    @staticmethod
    def config_snippet(port=None):
        """
        Return the Claude Desktop / Cursor mcp_servers JSON block.

        Args:
            port (int|None): Port number to embed. Reads from running server if None.

        Returns:
            str: Formatted JSON snippet ready to paste into claude_desktop_config.json.
        """
        if port is None:
            try:
                server = _get_server()
                port   = server.port
            except Exception:
                port = 48884

        bridge = _get_bridge_path()
        return (
            '{\n'
            '  "mcpServers": {\n'
            '    "t3lab-revit": {\n'
            '      "command": "python",\n'
            '      "args": [\n'
            '        "' + bridge + '",\n'
            '        "' + str(port) + '"\n'
            '      ]\n'
            '    }\n'
            '  }\n'
            '}'
        )

    @staticmethod
    def data_dir():
        """Return the T3Lab_AI_Data directory path (created if absent)."""
        d = _get_data_dir()
        try:
            if not os.path.isdir(d):
                os.makedirs(d)
        except OSError:
            pass
        return d

    @staticmethod
    def open_data_dir():
        """Open the T3Lab_AI_Data directory in the system file explorer."""
        d = MCPService.data_dir()
        try:
            import subprocess
            import platform
            if platform.system() == 'Windows':
                subprocess.Popen(['explorer', d])
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', d])
            else:
                subprocess.Popen(['xdg-open', d])
            return True, None
        except Exception as ex:
            return False, str(ex)

    # ── Claude Desktop auto-configure ─────────────────────────────────────────

    @staticmethod
    def find_claude_desktop_config():
        """Return the expected Claude Desktop config file path for this OS."""
        import platform
        home = os.path.expanduser('~')
        system = platform.system()
        if system == 'Windows':
            appdata = os.environ.get('APPDATA', os.path.join(home, 'AppData', 'Roaming'))
            return os.path.join(appdata, 'Claude', 'claude_desktop_config.json')
        elif system == 'Darwin':
            return os.path.join(home, 'Library', 'Application Support', 'Claude',
                                'claude_desktop_config.json')
        else:
            return os.path.join(home, '.config', 'Claude', 'claude_desktop_config.json')

    @staticmethod
    def claude_desktop_status():
        """
        Check the current Claude Desktop configuration status.

        Returns:
            dict with keys: path (str), file_exists (bool), configured (bool), error (str|None)
        """
        try:
            import json as _json
            path = MCPService.find_claude_desktop_config()
            if not os.path.isfile(path):
                return {'path': path, 'file_exists': False, 'configured': False, 'error': None}
            try:
                import codecs
                with codecs.open(path, 'r', encoding='utf-8') as f:
                    config = _json.loads(f.read())
                servers = config.get('mcpServers', {})
                configured = 't3lab-revit' in servers
                return {'path': path, 'file_exists': True, 'configured': configured, 'error': None}
            except Exception as ex:
                return {'path': path, 'file_exists': True, 'configured': False,
                        'error': 'Parse error: {}'.format(ex)}
        except Exception as ex:
            return {'path': '', 'file_exists': False, 'configured': False, 'error': str(ex)}

    @staticmethod
    def configure_claude_desktop(port=None):
        """
        Auto-write the t3lab-revit entry into Claude Desktop's config JSON.
        Creates the file and directory if absent; merges with existing entries.

        Args:
            port (int|None): Port to embed. Reads from running server if None.

        Returns:
            (success: bool, message: str) — message is config path on success, error on failure
        """
        try:
            import json as _json
            import codecs
            if port is None:
                try:
                    server = _get_server()
                    port = server.port
                except Exception:
                    port = 48884

            bridge = _get_bridge_path()
            path = MCPService.find_claude_desktop_config()

            config = {}
            if os.path.isfile(path):
                try:
                    with codecs.open(path, 'r', encoding='utf-8') as f:
                        raw = f.read().strip()
                    if raw:
                        config = _json.loads(raw)
                except Exception:
                    config = {}

            if 'mcpServers' not in config:
                config['mcpServers'] = {}
            config['mcpServers']['t3lab-revit'] = {
                'command': 'python',
                'args': [bridge, str(port)],
            }

            cfg_dir = os.path.dirname(path)
            if not os.path.isdir(cfg_dir):
                os.makedirs(cfg_dir)

            with codecs.open(path, 'w', encoding='utf-8') as f:
                f.write(_json.dumps(config, indent=2, ensure_ascii=False))

            return True, path
        except Exception as ex:
            return False, str(ex)

    # ── Combined snapshot (for dashboard widgets) ──────────────────────────────

    @staticmethod
    def full_status():
        """
        Return a combined status dict for both server and watcher.

        Useful for status-bar indicators or dashboards that need a single call.

        Returns:
            {
              'server':  {running, port, tools_count, commands_processed, error},
              'watcher': {running, data_dir, has_ext_event, error},
              'config':  '<snippet string>',
            }
        """
        srv = MCPService.server_status()
        wat = MCPService.watcher_status()
        return {
            'server':  srv,
            'watcher': wat,
            'config':  MCPService.config_snippet(port=srv.get('port')),
        }
