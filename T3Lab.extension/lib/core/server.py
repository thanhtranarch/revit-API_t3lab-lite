# -*- coding: utf-8 -*-
"""
MCP Server

Local MCP server implementation for AI-to-Revit communication.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "MCP Server"

import os
import threading
import json
import uuid

from Snippets._compat import eid_value
try:
    from http.server import HTTPServer, BaseHTTPRequestHandler
    from urllib.parse import urlparse, parse_qs
except ImportError:
    from BaseHTTPServer import HTTPServer, BaseHTTPRequestHandler
    from urlparse import urlparse, parse_qs

# External Event Handler for thread-safe Revit API calls
HAS_REVIT_UI = False
try:
    import clr
    clr.AddReference('RevitAPIUI')
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
    
    class MCPExternalEventHandler(IExternalEventHandler):
        def __init__(self, server):
            self.server = server
            self.tool_name = None
            self.arguments = None
            self.result = None
            self.exception = None
            self._lock = threading.Event()

        def Execute(self, app):
            try:
                self.result = self.server._execute_tool_in_context(self.tool_name, self.arguments)
                self.exception = None
            except Exception as e:
                self.exception = e
                self.result = None
            finally:
                self._lock.set()

        def GetName(self):
            return "T3Lab MCP External Event Handler"
            
    HAS_REVIT_UI = True
except Exception as e:
    pass



class MCPRequestHandler(BaseHTTPRequestHandler):
    """HTTP Request Handler for MCP Protocol"""

    protocol_version = 'HTTP/1.1'

    def log_message(self, format, *args):
        """Override to suppress default logging"""
        pass

    def _send_response(self, status_code, content_type, body):
        """Send HTTP response"""
        self.send_response(status_code)
        self.send_header('Content-Type', content_type)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_header('Content-Length', len(body))
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, data, status_code=200):
        """Send JSON response"""
        body = json.dumps(data)
        self._send_response(status_code, 'application/json', body)

    def _send_sse_event(self, event_type, data):
        """Send SSE event"""
        message = "event: {}\ndata: {}\n\n".format(event_type, json.dumps(data))
        self.wfile.write(message.encode('utf-8'))
        self.wfile.flush()

    def do_OPTIONS(self):
        """Handle CORS preflight"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        """Handle GET requests"""
        parsed = urlparse(self.path)
        path = parsed.path

        if path == '/':
            # Server info
            self._send_json({
                'name': 'T3LabAI MCP Server',
                'version': '1.0.0',
                'protocol': 'mcp',
                'status': 'running'
            })

        elif path == '/sse':
            # SSE endpoint for MCP communication
            self._handle_sse()

        elif path == '/health':
            # Health check
            self._send_json({'status': 'ok'})

        elif path in ('/v1/models', '/models'):
            # Tolerate OpenAI-compatible clients that probe for a model list,
            # so they don't repeatedly hit an "unexpected endpoint" 404.
            self._send_json({'object': 'list', 'data': []})

        else:
            self._send_json({'error': 'Not found'}, 404)

    def do_POST(self):
        """Handle POST requests (MCP messages)"""
        parsed = urlparse(self.path)
        path = parsed.path

        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length).decode('utf-8') if content_length > 0 else '{}'

        try:
            request = json.loads(body) if body else {}
        except json.JSONDecodeError:
            self._send_json({'error': 'Invalid JSON'}, 400)
            return

        if path == '/message' or path == '/mcp':
            self._handle_mcp_message(request)
        else:
            self._send_json({'error': 'Not found'}, 404)

    def _handle_sse(self):
        """Handle SSE connection for MCP"""
        self.send_response(200)
        self.send_header('Content-Type', 'text/event-stream')
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('Connection', 'keep-alive')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()

        # Register client
        server = self.server.mcp_server
        client_id = str(uuid.uuid4())
        server._register_client(client_id)

        # Send endpoint event for MCP protocol
        endpoint_url = "http://localhost:{}/message".format(server.port)
        self._send_sse_event('endpoint', endpoint_url)

        try:
            # Keep connection alive
            while server.is_running:
                # Send keep-alive ping every 30 seconds
                import time
                time.sleep(30)
                self._send_sse_event('ping', {'timestamp': time.time()})
        except Exception:
            pass
        finally:
            server._unregister_client(client_id)

    def _handle_mcp_message(self, request):
        """Handle MCP JSON-RPC message"""
        server = self.server.mcp_server

        method = request.get('method', '')
        params = request.get('params', {})
        request_id = request.get('id')

        # ── Auth ────────────────────────────────────────────────────────
        # Reject anything that doesn't carry the shared-secret token (see
        # T3LabAIServer._get_or_create_token). Without this, any local
        # process could reach tools/call — including send_code_to_revit,
        # which executes arbitrary IronPython with full Revit API access.
        expected = 'Bearer ' + server._token
        if self.headers.get('Authorization', '') != expected:
            if 'id' not in request:
                self._send_json({'status': 'unauthorized'}, 401)
            else:
                self._send_json({
                    'jsonrpc': '2.0',
                    'id': request_id,
                    'error': {'code': -32001, 'message': 'Unauthorized: missing or invalid token'}
                }, 401)
            return

        # If it's a notification (no 'id' in request), do not send JSON-RPC response
        if 'id' not in request:
            try:
                if method == 'notifications/initialized':
                    # Handle initialized notification if needed
                    pass
            except Exception:
                pass
            self._send_json({'status': 'ok'})
            return

        response = {
            'jsonrpc': '2.0',
            'id': request_id
        }

        try:
            if method == 'initialize':
                response['result'] = server._handle_initialize(params)
            elif method == 'tools/list':
                response['result'] = server._handle_tools_list()
            elif method == 'tools/call':
                response['result'] = server._handle_tool_call(params)
            else:
                response['error'] = {
                    'code': -32601,
                    'message': 'Method not found: {}'.format(method)
                }
        except Exception as e:
            response['error'] = {
                'code': -32603,
                'message': str(e)
            }

        server._commands_processed += 1
        self._send_json(response)


class T3LabAIServer(object):
    """MCP Server for AI communication with Revit"""

    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super(T3LabAIServer, cls).__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return

        self._port = 48884
        self._is_running = False
        self._client_count = 0
        self._total_clients = 0
        self._commands_processed = 0
        self._server_thread = None
        self._http_server = None
        self._clients = {}
        self._tools = {}
        self._external_event = None
        self._event_handler = None
        self._token = self._get_or_create_token()
        self._pinned_doc_title = None
        # Open TransactionGroup for the current assistant request (B4) — owned
        # and closed on the Revit main thread via the __begin/__end_action_group
        # pseudo-tools below. One request = one group = one Undo entry.
        self._action_group = None
        self._initialized = True

        # Register default Revit tools
        self._register_default_tools()

    def _get_or_create_token(self):
        """Return the shared-secret token every /message and /mcp request
        must carry, creating and persisting one on first run.

        Without this, any local process could hit the HTTP server and call
        tools/call — including send_code_to_revit, which runs arbitrary
        IronPython with full Revit API access. Persisting the token to the
        same %APPDATA%\\T3LabAI directory used by settings.py lets a
        locally-spawned bridge.py (launched by Claude Desktop/Cursor per
        the mcpServers config) read it without any manual setup.
        """
        try:
            app_data = os.environ.get('APPDATA', '')
            token_dir = os.path.join(app_data, 'T3LabAI')
            if not os.path.exists(token_dir):
                os.makedirs(token_dir)
            token_path = os.path.join(token_dir, 'mcp_token.txt')
            if os.path.exists(token_path):
                with open(token_path, 'r') as f:
                    existing = f.read().strip()
                if existing:
                    return existing
            token = uuid.uuid4().hex
            with open(token_path, 'w') as f:
                f.write(token)
            return token
        except Exception:
            # Can't persist (e.g. no APPDATA) — still refuse unauthenticated
            # requests rather than silently accepting everything, just with
            # a token that only this process knows (external bridges won't
            # be able to authenticate until this succeeds).
            return uuid.uuid4().hex

    @property
    def token(self):
        return self._token

    def _register_default_tools(self):
        """Register default Revit tools for MCP"""
        self._tools = {
            'revit_get_active_view': {
                'name': 'revit_get_active_view',
                'description': 'Get information about the currently active view in Revit',
                'inputSchema': {
                    'type': 'object',
                    'properties': {},
                    'required': []
                }
            },
            'revit_get_selected_elements': {
                'name': 'revit_get_selected_elements',
                'description': 'Get information about currently selected elements in Revit',
                'inputSchema': {
                    'type': 'object',
                    'properties': {},
                    'required': []
                }
            },
            'revit_get_project_info': {
                'name': 'revit_get_project_info',
                'description': 'Get project information from the current Revit document',
                'inputSchema': {
                    'type': 'object',
                    'properties': {},
                    'required': []
                }
            },
            'revit_list_views': {
                'name': 'revit_list_views',
                'description': 'List all views in the current Revit document',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'view_type': {
                            'type': 'string',
                            'description': 'Filter by view type (optional)'
                        }
                    },
                    'required': []
                }
            },
            'revit_list_sheets': {
                'name': 'revit_list_sheets',
                'description': 'List all sheets in the current Revit document',
                'inputSchema': {
                    'type': 'object',
                    'properties': {},
                    'required': []
                }
            },
            'revit_get_element_info': {
                'name': 'revit_get_element_info',
                'description': 'Get detailed information about a specific Revit element by ID',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {
                            'type': 'integer',
                            'description': 'The Revit element ID'
                        }
                    },
                    'required': ['element_id']
                }
            },
            'revit_override_color': {
                'name': 'revit_override_color',
                'description': 'Override elements color in the active view',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'color': {
                            'type': 'string',
                            'description': 'Hex color code (e.g. #FF0000) or CSS color name (e.g. red, green, blue)'
                        },
                        'element_ids': {
                            'type': 'array',
                            'items': {
                                'type': 'integer'
                            },
                            'description': 'Optional list of Revit element IDs. If omitted, applies to the currently selected elements.'
                        }
                    },
                    'required': ['color']
                }
            },
            'create_level': {
                'name': 'create_level',
                'description': 'Create a new Level in the Revit model at the specified elevation',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'elevation': {
                            'type': 'number',
                            'description': 'Elevation in meters above project base point'
                        },
                        'name': {
                            'type': 'string',
                            'description': 'Level name (e.g. "Level 3", "Ground Floor")'
                        }
                    },
                    'required': ['elevation']
                }
            },
            'place_wall': {
                'name': 'place_wall',
                'description': 'Place a wall between two points on a specified level',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'start_x': {
                            'type': 'number',
                            'description': 'Start point X coordinate in meters'
                        },
                        'start_y': {
                            'type': 'number',
                            'description': 'Start point Y coordinate in meters'
                        },
                        'end_x': {
                            'type': 'number',
                            'description': 'End point X coordinate in meters'
                        },
                        'end_y': {
                            'type': 'number',
                            'description': 'End point Y coordinate in meters'
                        },
                        'level_name': {
                            'type': 'string',
                            'description': 'Level name to place wall on (default: first level)'
                        },
                        'height': {
                            'type': 'number',
                            'description': 'Wall height in meters (default: 3.0)'
                        },
                        'wall_type_name': {
                            'type': 'string',
                            'description': 'Wall type name (default: first available)'
                        }
                    },
                    'required': ['start_x', 'start_y', 'end_x', 'end_y']
                }
            },
            'get_parameter': {
                'name': 'get_parameter',
                'description': 'Get the value of a named parameter from a Revit element by its ID',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {
                            'type': 'integer',
                            'description': 'Revit element ID'
                        },
                        'parameter_name': {
                            'type': 'string',
                            'description': 'Parameter name to retrieve'
                        }
                    },
                    'required': ['element_id', 'parameter_name']
                }
            },
            'get_elements_by_level': {
                'name': 'get_elements_by_level',
                'description': 'Get all elements of a given category on a specified level (e.g. get all beams on Level 3)',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'level_name': {
                            'type': 'string',
                            'description': 'Level name (e.g. "Level 3", "Ground Floor")'
                        },
                        'category': {
                            'type': 'string',
                            'description': 'Revit category name filter: Walls, Beams, Floors, Columns, Doors, Windows. If omitted, returns all elements on the level.'
                        }
                    },
                    'required': ['level_name']
                }
            },
            # ── Query / read tools ────────────────────────────────────────────
            'get_current_view_info': {
                'name': 'get_current_view_info',
                'description': 'Get detailed info about the current active view including scale, view type, crop region, and discipline',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            'get_current_view_elements': {
                'name': 'get_current_view_elements',
                'description': 'Get all visible elements in the current active view, optionally filtered by category',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {
                            'type': 'string',
                            'description': 'Optional category filter: Walls, Floors, Doors, Windows, Rooms, Columns, Beams, etc.'
                        },
                        'limit': {
                            'type': 'integer',
                            'description': 'Max number of elements to return (default 100)'
                        }
                    },
                    'required': []
                }
            },
            'get_available_family_types': {
                'name': 'get_available_family_types',
                'description': 'Get available family types in the current project, optionally filtered by category',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {
                            'type': 'string',
                            'description': 'Category to filter by (e.g. Doors, Windows, Furniture, Structural Framing)'
                        }
                    },
                    'required': []
                }
            },
            'get_material_quantities': {
                'name': 'get_material_quantities',
                'description': 'Calculate material quantities and takeoffs for elements in the model',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {
                            'type': 'string',
                            'description': 'Category to calculate quantities for: Walls, Floors, Roofs, Ceilings. Default: Walls'
                        },
                        'level_name': {
                            'type': 'string',
                            'description': 'Optional: limit to a specific level'
                        }
                    },
                    'required': []
                }
            },
            'ai_element_filter': {
                'name': 'ai_element_filter',
                'description': 'Intelligent element querying tool for AI assistants — filter by category, parameter name, and value',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {
                            'type': 'string',
                            'description': 'Element category to search (e.g. Walls, Doors, Rooms)'
                        },
                        'parameter_name': {
                            'type': 'string',
                            'description': 'Optional parameter name to filter on'
                        },
                        'parameter_value': {
                            'type': 'string',
                            'description': 'Optional value to match (partial match supported)'
                        },
                        'limit': {
                            'type': 'integer',
                            'description': 'Max results to return (default 50)'
                        }
                    },
                    'required': ['category']
                }
            },
            'analyze_model_statistics': {
                'name': 'analyze_model_statistics',
                'description': 'Analyze model complexity with element counts by category and warnings summary',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            # ── Create tools ──────────────────────────────────────────────────
            'create_point_based_element': {
                'name': 'create_point_based_element',
                'description': 'Create a point-based element (door, window, furniture, fixture) at a specified XYZ location',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'family_type': {
                            'type': 'string',
                            'description': 'Family type name (e.g. "Single-Flush", "36x84")'
                        },
                        'x': {'type': 'number', 'description': 'X coordinate in meters'},
                        'y': {'type': 'number', 'description': 'Y coordinate in meters'},
                        'z': {'type': 'number', 'description': 'Z coordinate in meters (default 0)'},
                        'level_name': {'type': 'string', 'description': 'Level to place element on'},
                        'host_wall_id': {'type': 'integer', 'description': 'Optional host wall element ID for hosted families (doors/windows)'}
                    },
                    'required': ['family_type', 'x', 'y']
                }
            },
            'create_line_based_element': {
                'name': 'create_line_based_element',
                'description': 'Create a line-based element (beam, pipe, duct, structural framing) between two points',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'family_type': {'type': 'string', 'description': 'Family type name (e.g. "W-Wide Flange-Column:W10x33")'},
                        'start_x': {'type': 'number', 'description': 'Start X in meters'},
                        'start_y': {'type': 'number', 'description': 'Start Y in meters'},
                        'start_z': {'type': 'number', 'description': 'Start Z in meters (default 0)'},
                        'end_x': {'type': 'number', 'description': 'End X in meters'},
                        'end_y': {'type': 'number', 'description': 'End Y in meters'},
                        'end_z': {'type': 'number', 'description': 'End Z in meters (default 0)'},
                        'level_name': {'type': 'string', 'description': 'Reference level name'}
                    },
                    'required': ['family_type', 'start_x', 'start_y', 'end_x', 'end_y']
                }
            },
            'create_surface_based_element': {
                'name': 'create_surface_based_element',
                'description': 'Create a surface-based element (floor, ceiling, roof) from a list of boundary points',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_type': {
                            'type': 'string',
                            'description': 'Type: "floor", "ceiling", or "roof"'
                        },
                        'boundary_points': {
                            'type': 'array',
                            'description': 'List of [x, y] pairs in meters defining the boundary polygon',
                            'items': {
                                'type': 'array',
                                'items': {'type': 'number'}
                            }
                        },
                        'level_name': {'type': 'string', 'description': 'Level name'},
                        'type_name': {'type': 'string', 'description': 'Floor/Ceiling/Roof type name (optional)'}
                    },
                    'required': ['element_type', 'boundary_points']
                }
            },
            'create_grid': {
                'name': 'create_grid',
                'description': 'Create a grid system with smart spacing — generates named horizontal and vertical grid lines',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'x_spacings': {
                            'type': 'array',
                            'items': {'type': 'number'},
                            'description': 'List of X-direction spacings in meters (e.g. [6, 6, 6] creates 4 vertical lines)'
                        },
                        'y_spacings': {
                            'type': 'array',
                            'items': {'type': 'number'},
                            'description': 'List of Y-direction spacings in meters'
                        },
                        'origin_x': {'type': 'number', 'description': 'Origin X in meters (default 0)'},
                        'origin_y': {'type': 'number', 'description': 'Origin Y in meters (default 0)'},
                        'x_labels': {
                            'type': 'array',
                            'items': {'type': 'string'},
                            'description': 'Labels for vertical grid lines (e.g. ["A","B","C"]). Auto-generated if omitted.'
                        },
                        'y_labels': {
                            'type': 'array',
                            'items': {'type': 'string'},
                            'description': 'Labels for horizontal grid lines (e.g. ["1","2","3"]). Auto-generated if omitted.'
                        }
                    },
                    'required': ['x_spacings', 'y_spacings']
                }
            },
            'create_room': {
                'name': 'create_room',
                'description': 'Create and place a room at a specified XY location on a level',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'x': {'type': 'number', 'description': 'X coordinate of room placement point in meters'},
                        'y': {'type': 'number', 'description': 'Y coordinate of room placement point in meters'},
                        'level_name': {'type': 'string', 'description': 'Level to place the room on'},
                        'name': {'type': 'string', 'description': 'Room name (e.g. "Office", "Kitchen")'},
                        'number': {'type': 'string', 'description': 'Room number (e.g. "101", "A-01")'}
                    },
                    'required': ['x', 'y']
                }
            },
            'create_structural_framing_system': {
                'name': 'create_structural_framing_system',
                'description': 'Create a structural beam framing system on a grid — places beams at specified bay spacings',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'x_bays': {
                            'type': 'array',
                            'items': {'type': 'number'},
                            'description': 'Bay spacings in X direction in meters'
                        },
                        'y_bays': {
                            'type': 'array',
                            'items': {'type': 'number'},
                            'description': 'Bay spacings in Y direction in meters'
                        },
                        'level_name': {'type': 'string', 'description': 'Level to place beams on'},
                        'beam_type': {'type': 'string', 'description': 'Structural framing family type name (optional)'},
                        'origin_x': {'type': 'number', 'description': 'Origin X in meters (default 0)'},
                        'origin_y': {'type': 'number', 'description': 'Origin Y in meters (default 0)'}
                    },
                    'required': ['x_bays', 'y_bays']
                }
            },
            # ── Modify / delete tools ─────────────────────────────────────────
            'delete_element': {
                'name': 'delete_element',
                'description': ('Delete one or more Revit elements by their IDs. Deletion cascades '
                                '(deleting a wall removes its hosted doors/windows), so prefer a '
                                'dry_run=true preview first to see everything that would be removed.'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_ids': {
                            'type': 'array',
                            'items': {'type': 'integer'},
                            'description': 'List of element IDs to delete'
                        },
                        'dry_run': {
                            'type': 'boolean',
                            'description': 'If true, do NOT delete — return the full list of elements that WOULD be removed (including cascade deletions). Default false.'
                        }
                    },
                    'required': ['element_ids']
                }
            },
            'operate_element': {
                'name': 'operate_element',
                'description': 'Operate on elements: select, hide, isolate, unhide, or setColor in the active view',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'operation': {
                            'type': 'string',
                            'description': 'Operation: "select", "hide", "isolate", "unhide", "reset_color"'
                        },
                        'element_ids': {
                            'type': 'array',
                            'items': {'type': 'integer'},
                            'description': 'Element IDs to operate on'
                        }
                    },
                    'required': ['operation', 'element_ids']
                }
            },
            'color_elements': {
                'name': 'color_elements',
                'description': 'Color elements in the active view based on a parameter value — each unique value gets a distinct color',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {
                            'type': 'string',
                            'description': 'Element category to color (e.g. Walls, Rooms, Floors)'
                        },
                        'parameter_name': {
                            'type': 'string',
                            'description': 'Parameter name to group colors by (e.g. "Type Name", "Level")'
                        }
                    },
                    'required': ['category', 'parameter_name']
                }
            },
            'tag_all_walls': {
                'name': 'tag_all_walls',
                'description': 'Automatically tag all walls in the current active view',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'tag_type': {
                            'type': 'string',
                            'description': 'Wall tag family type name (optional, uses first available)'
                        },
                        'leader': {
                            'type': 'boolean',
                            'description': 'Add tag leader line (default false)'
                        }
                    },
                    'required': []
                }
            },
            'tag_all_rooms': {
                'name': 'tag_all_rooms',
                'description': 'Automatically tag all rooms in the current active view',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'tag_type': {
                            'type': 'string',
                            'description': 'Room tag family type name (optional)'
                        }
                    },
                    'required': []
                }
            },
            # ── Data / storage tools ──────────────────────────────────────────
            'export_room_data': {
                'name': 'export_room_data',
                'description': 'Export all room data from the project (number, name, area, level, department) as structured JSON',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'level_name': {
                            'type': 'string',
                            'description': 'Optional: filter rooms by level name'
                        }
                    },
                    'required': []
                }
            },
            'store_project_data': {
                'name': 'store_project_data',
                'description': 'Store current project metadata (name, number, client, address, status) to a local JSON file for later querying',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            'store_room_data': {
                'name': 'store_room_data',
                'description': 'Store all room metadata to a local JSON file in the project folder',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            'query_stored_data': {
                'name': 'query_stored_data',
                'description': 'Query previously stored project and room data from local JSON files',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'data_type': {
                            'type': 'string',
                            'description': '"project" or "rooms"'
                        }
                    },
                    'required': ['data_type']
                }
            },
            # ── Utility tools ─────────────────────────────────────────────────
            'send_code_to_revit': {
                'name': 'send_code_to_revit',
                'description': ('Execute IronPython code directly in the Revit context. Use with care — full Revit API access. '
                                'Return data by assigning to `result` (or output.append(...)); print() is captured into the '
                                'tool result, and code must NEVER open dialogs (TaskDialog/MessageBox/forms).'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'code': {
                            'type': 'string',
                            'description': ('IronPython 2.7 code to execute. Has access to doc, uidoc, app, and the variables '
                                            '`result` / `output` (a list). Assign your answer to `result` or output.append(...) — '
                                            'do NOT print for display and do NOT show any dialog/window.')
                        }
                    },
                    'required': ['code']
                }
            },
            'say_hello': {
                'name': 'say_hello',
                'description': 'Display a greeting TaskDialog in Revit — use to verify MCP connection is working',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'message': {
                            'type': 'string',
                            'description': 'Message to display (default: "Hello from T3Lab AI!")'
                        }
                    },
                    'required': []
                }
            },
            # ── Parameter writing ─────────────────────────────────────────────
            'set_parameter': {
                'name': 'set_parameter',
                'description': 'Set the value of a named parameter on a Revit element',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {'type': 'integer', 'description': 'Revit element ID'},
                        'parameter_name': {'type': 'string', 'description': 'Parameter name to set'},
                        'value': {'type': 'string', 'description': 'New value as string (will be converted to appropriate type)'}
                    },
                    'required': ['element_id', 'parameter_name', 'value']
                }
            },
            'get_all_parameters': {
                'name': 'get_all_parameters',
                'description': 'Get all parameters (name, value, type, read-only flag) for a given element',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {'type': 'integer', 'description': 'Revit element ID'}
                    },
                    'required': ['element_id']
                }
            },
            # ── Spatial transforms ────────────────────────────────────────────
            'move_elements': {
                'name': 'move_elements',
                'description': 'Move one or more elements by a delta XYZ vector',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_ids': {
                            'type': 'array', 'items': {'type': 'integer'},
                            'description': 'Element IDs to move'
                        },
                        'dx': {'type': 'number', 'description': 'Delta X in meters'},
                        'dy': {'type': 'number', 'description': 'Delta Y in meters'},
                        'dz': {'type': 'number', 'description': 'Delta Z in meters (default 0)'}
                    },
                    'required': ['element_ids', 'dx', 'dy']
                }
            },
            'copy_elements': {
                'name': 'copy_elements',
                'description': 'Copy elements and translate the copies by a delta XYZ vector',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_ids': {
                            'type': 'array', 'items': {'type': 'integer'},
                            'description': 'Element IDs to copy'
                        },
                        'dx': {'type': 'number', 'description': 'Delta X in meters'},
                        'dy': {'type': 'number', 'description': 'Delta Y in meters'},
                        'dz': {'type': 'number', 'description': 'Delta Z in meters (default 0)'}
                    },
                    'required': ['element_ids', 'dx', 'dy']
                }
            },
            'rotate_element': {
                'name': 'rotate_element',
                'description': 'Rotate elements around a Z-axis through a given origin point',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_ids': {
                            'type': 'array', 'items': {'type': 'integer'},
                            'description': 'Element IDs to rotate'
                        },
                        'angle_degrees': {'type': 'number', 'description': 'Rotation angle in degrees (counter-clockwise)'},
                        'origin_x': {'type': 'number', 'description': 'Rotation axis origin X in meters (default 0)'},
                        'origin_y': {'type': 'number', 'description': 'Rotation axis origin Y in meters (default 0)'}
                    },
                    'required': ['element_ids', 'angle_degrees']
                }
            },
            'get_element_bounding_box': {
                'name': 'get_element_bounding_box',
                'description': 'Get the bounding box (min/max XYZ in meters) of a Revit element',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {'type': 'integer', 'description': 'Revit element ID'}
                    },
                    'required': ['element_id']
                }
            },
            # ── View management ───────────────────────────────────────────────
            'create_view': {
                'name': 'create_view',
                'description': 'Create a new view: floor plan, ceiling plan, or 3D isometric view',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'view_type': {
                            'type': 'string',
                            'description': '"floor_plan", "ceiling_plan", or "3d"'
                        },
                        'level_name': {
                            'type': 'string',
                            'description': 'Level name (required for floor/ceiling plan)'
                        },
                        'name': {'type': 'string', 'description': 'Name for the new view'}
                    },
                    'required': ['view_type']
                }
            },
            'set_active_view': {
                'name': 'set_active_view',
                'description': 'Switch the active view in Revit by view name or element ID',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'view_name': {'type': 'string', 'description': 'View name to activate'},
                        'view_id': {'type': 'integer', 'description': 'View element ID (alternative to view_name)'}
                    },
                    'required': []
                }
            },
            'rename_element': {
                'name': 'rename_element',
                'description': 'Rename a view, sheet, level, or any named Revit element',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {'type': 'integer', 'description': 'Element ID to rename'},
                        'new_name': {'type': 'string', 'description': 'New name for the element'}
                    },
                    'required': ['element_id', 'new_name']
                }
            },
            # ── Sheet management ──────────────────────────────────────────────
            'create_sheet': {
                'name': 'create_sheet',
                'description': 'Create a new drawing sheet with a title block',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'sheet_number': {'type': 'string', 'description': 'Sheet number (e.g. "A-101")'},
                        'sheet_name': {'type': 'string', 'description': 'Sheet title (e.g. "Ground Floor Plan")'},
                        'title_block': {'type': 'string', 'description': 'Title block family type name (optional, uses first available)'}
                    },
                    'required': ['sheet_number', 'sheet_name']
                }
            },
            'add_view_to_sheet': {
                'name': 'add_view_to_sheet',
                'description': 'Place a view as a viewport on an existing sheet',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'sheet_id': {'type': 'integer', 'description': 'Sheet element ID'},
                        'view_id': {'type': 'integer', 'description': 'View element ID to place'},
                        'x': {'type': 'number', 'description': 'Viewport center X on sheet in mm (default 297 = center A3)'},
                        'y': {'type': 'number', 'description': 'Viewport center Y on sheet in mm (default 210)'}
                    },
                    'required': ['sheet_id', 'view_id']
                }
            },
            # ── Annotation ────────────────────────────────────────────────────
            'create_text_note': {
                'name': 'create_text_note',
                'description': 'Add a text annotation to the active view at a specified location',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'text': {'type': 'string', 'description': 'Text content'},
                        'x': {'type': 'number', 'description': 'X position in meters (model coordinates)'},
                        'y': {'type': 'number', 'description': 'Y position in meters'},
                        'font_size': {'type': 'number', 'description': 'Text height in mm (default 3.5)'},
                        'text_type': {'type': 'string', 'description': 'Text note type name (optional)'}
                    },
                    'required': ['text', 'x', 'y']
                }
            },
            # ── Model quality ─────────────────────────────────────────────────
            'get_model_warnings': {
                'name': 'get_model_warnings',
                'description': 'Get all Revit model warnings with descriptions and affected element IDs',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'limit': {'type': 'integer', 'description': 'Max warnings to return (default 50)'}
                    },
                    'required': []
                }
            },
            'get_model_health': {
                'name': 'get_model_health',
                'description': 'Get model health summary: warning count, element count, linked files, unused families',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            # ── Collaboration / worksets ──────────────────────────────────────
            'list_worksets': {
                'name': 'list_worksets',
                'description': 'List all user worksets in a workshared model with their open/close status',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            'set_element_workset': {
                'name': 'set_element_workset',
                'description': 'Move one or more elements to a specified workset',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_ids': {
                            'type': 'array', 'items': {'type': 'integer'},
                            'description': 'Element IDs to move to the workset'
                        },
                        'workset_name': {'type': 'string', 'description': 'Target workset name'}
                    },
                    'required': ['element_ids', 'workset_name']
                }
            },
            # ── Datum / navigation ────────────────────────────────────────────
            'list_levels': {
                'name': 'list_levels',
                'description': 'List all levels in the project with their elevations in meters',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            # ── Family management ─────────────────────────────────────────────
            'load_family': {
                'name': 'load_family',
                'description': 'Load a .rfa family file into the current project from a local file path',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'file_path': {
                            'type': 'string',
                            'description': 'Absolute path to the .rfa family file'
                        }
                    },
                    'required': ['file_path']
                }
            },
            # ── Geometry editing ──────────────────────────────────────────────
            'split_curve': {
                'name': 'split_curve',
                'description': ('Split a model or detail curve (line, arc, ellipse or spline) into '
                                'segments while preserving the EXACT original geometry — arcs stay '
                                'arcs, splines stay splines (never flattened to straight lines). '
                                'Provide either "segments" for equal-length division or '
                                '"split_at_ratios" for explicit cut positions.'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {
                            'type': 'integer',
                            'description': 'Element ID of the model curve or detail curve to split'
                        },
                        'segments': {
                            'type': 'integer',
                            'description': 'Number of equal-length pieces to split into (default 2, min 2, max 200)'
                        },
                        'split_at_ratios': {
                            'type': 'array',
                            'items': {'type': 'number'},
                            'description': 'Optional explicit cut positions as normalized fractions 0..1 along the curve (e.g. [0.25, 0.5]). Overrides "segments".'
                        }
                    },
                    'required': ['element_id']
                }
            },
            'split_element': {
                'name': 'split_element',
                'description': ('Split a location-curve element (wall, beam, pipe, duct, model line) '
                                'into two at a point or ratio, preserving the exact curve geometry.'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id': {'type': 'integer', 'description': 'Element ID with a location curve (wall/beam/pipe/etc.)'},
                        'at_ratio': {'type': 'number', 'description': 'Split position as a fraction 0..1 along the curve (default 0.5)'},
                        'x': {'type': 'number', 'description': 'Optional split point X in meters (overrides at_ratio)'},
                        'y': {'type': 'number', 'description': 'Optional split point Y in meters (used with x)'}
                    },
                    'required': ['element_id']
                }
            },
            'join_geometry': {
                'name': 'join_geometry',
                'description': 'Join (or unjoin) the geometry of two elements, e.g. wall-to-floor or wall-to-column.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_id_a': {'type': 'integer', 'description': 'First element ID'},
                        'element_id_b': {'type': 'integer', 'description': 'Second element ID'},
                        'unjoin': {'type': 'boolean', 'description': 'If true, unjoin instead of join (default false)'}
                    },
                    'required': ['element_id_a', 'element_id_b']
                }
            },
            # ── Bulk parameter / selection / tagging ──────────────────────────
            'bulk_set_parameter': {
                'name': 'bulk_set_parameter',
                'description': ('Set one parameter to the same value across MANY elements of a category, '
                                'optionally narrowed by a filter parameter/value substring match.'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {'type': 'string', 'description': 'Category to target (e.g. Walls, Doors, Rooms)'},
                        'parameter_name': {'type': 'string', 'description': 'Parameter to set on each element'},
                        'value': {'type': 'string', 'description': 'New value (string; coerced to the parameter storage type)'},
                        'filter_parameter': {'type': 'string', 'description': 'Optional parameter to filter elements by before setting'},
                        'filter_value': {'type': 'string', 'description': 'Optional value substring the filter_parameter must contain'},
                        'element_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'Optional explicit element IDs (overrides category collection)'},
                        'limit': {'type': 'integer', 'description': 'Max elements to modify (default 500)'}
                    },
                    'required': ['parameter_name', 'value']
                }
            },
            'select_elements': {
                'name': 'select_elements',
                'description': 'Set the Revit selection to elements matched by category + optional parameter filter, or explicit IDs.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {'type': 'string', 'description': 'Category to select (e.g. Walls, Doors)'},
                        'parameter_name': {'type': 'string', 'description': 'Optional parameter to filter on'},
                        'parameter_value': {'type': 'string', 'description': 'Optional value substring to match'},
                        'element_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'Optional explicit IDs to select (overrides category)'},
                        'add_to_selection': {'type': 'boolean', 'description': 'Add to the current selection instead of replacing (default false)'},
                        'limit': {'type': 'integer', 'description': 'Max elements to select (default 500)'},
                        'show': {'type': 'boolean', 'description': 'Also zoom the view onto the selected elements (default false)'}
                    },
                    'required': []
                }
            },
            'tag_elements': {
                'name': 'tag_elements',
                'description': 'Tag all elements of a category in the active view (generic version of tag_all_walls/rooms).',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {'type': 'string', 'description': 'Category to tag (e.g. Doors, Windows, Walls, Rooms)'},
                        'tag_type': {'type': 'string', 'description': 'Tag family type name (optional, uses first available)'},
                        'leader': {'type': 'boolean', 'description': 'Add tag leader line (default false)'}
                    },
                    'required': ['category']
                }
            },
            'create_dimension': {
                'name': 'create_dimension',
                'description': ('Create an aligned dimension across a set of grids (or elements with a '
                                'straight location curve) in the active view.'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'IDs of grids (or line-based elements) to dimension between — at least 2'},
                        'offset': {'type': 'number', 'description': 'Perpendicular offset of the dimension line in meters (default 1.0)'}
                    },
                    'required': ['element_ids']
                }
            },
            # ── Schedules ─────────────────────────────────────────────────────
            'get_schedule_data': {
                'name': 'get_schedule_data',
                'description': 'Read a schedule (ViewSchedule) as a JSON table with a header row and data rows.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'schedule_name': {'type': 'string', 'description': 'Schedule name (partial match allowed)'},
                        'schedule_id': {'type': 'integer', 'description': 'Schedule element ID (alternative to name)'},
                        'limit': {'type': 'integer', 'description': 'Max data rows to return (default 200)'}
                    },
                    'required': []
                }
            },
            'create_schedule': {
                'name': 'create_schedule',
                'description': 'Create a new schedule for a category with an optional explicit field list.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'category': {'type': 'string', 'description': 'Category to schedule (e.g. Walls, Doors, Rooms)'},
                        'fields': {'type': 'array', 'items': {'type': 'string'}, 'description': 'Parameter names to add as columns (optional; adds a sensible default set if omitted)'},
                        'name': {'type': 'string', 'description': 'Schedule name (optional)'}
                    },
                    'required': ['category']
                }
            },
            # ── View / sheet automation ───────────────────────────────────────
            'duplicate_view': {
                'name': 'duplicate_view',
                'description': 'Duplicate a view (plain, with detailing, or as a dependent view).',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'view_id': {'type': 'integer', 'description': 'View element ID to duplicate'},
                        'mode': {'type': 'string', 'description': '"plain" (default), "with_detailing", or "dependent"'},
                        'name': {'type': 'string', 'description': 'Name for the new view (optional)'}
                    },
                    'required': ['view_id']
                }
            },
            'apply_view_template': {
                'name': 'apply_view_template',
                'description': 'Apply a view template to one or more views.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'view_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'View IDs to apply the template to'},
                        'template_name': {'type': 'string', 'description': 'View template name (partial match allowed)'}
                    },
                    'required': ['view_ids', 'template_name']
                }
            },
            'create_view_filter': {
                'name': 'create_view_filter',
                'description': ('Create a rule-based view filter and apply it to a view, optionally hiding '
                                'matching elements or overriding their color.'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'name': {'type': 'string', 'description': 'Filter name'},
                        'categories': {'type': 'array', 'items': {'type': 'string'}, 'description': 'Categories the filter applies to (e.g. ["Walls"])'},
                        'parameter_name': {'type': 'string', 'description': 'Optional parameter to build a "contains" rule on'},
                        'parameter_value': {'type': 'string', 'description': 'Optional value the parameter must contain'},
                        'view_id': {'type': 'integer', 'description': 'View to apply the filter to (default active view)'},
                        'hide': {'type': 'boolean', 'description': 'Hide matching elements (default false)'},
                        'color': {'type': 'string', 'description': 'Optional hex/CSS color to override matching elements'}
                    },
                    'required': ['name', 'categories']
                }
            },
            'place_views_on_sheets': {
                'name': 'place_views_on_sheets',
                'description': 'Place one or more views onto sheets as viewports (one view per sheet by default).',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'view_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'View IDs to place'},
                        'title_block': {'type': 'string', 'description': 'Title block family type name (optional)'},
                        'sheet_id': {'type': 'integer', 'description': 'Existing sheet to place all views on (optional; otherwise a sheet is created per view)'}
                    },
                    'required': ['view_ids']
                }
            },
            'export_dwg': {
                'name': 'export_dwg',
                'description': 'Export sheets or views to DWG files in a folder.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'sheet_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'Sheet IDs to export'},
                        'view_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'View IDs to export (used if sheet_ids omitted)'},
                        'output_folder': {'type': 'string', 'description': 'Destination folder (default: same folder as the .rvt)'}
                    },
                    'required': []
                }
            },
            'export_image': {
                'name': 'export_image',
                'description': 'Export a view (or the active view) to a PNG image.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'view_id': {'type': 'integer', 'description': 'View to export (default active view)'},
                        'output_folder': {'type': 'string', 'description': 'Destination folder (default: same folder as the .rvt)'},
                        'width': {'type': 'integer', 'description': 'Pixel width of the image (default 1600)'}
                    },
                    'required': []
                }
            },
            # ── Standards / model management ──────────────────────────────────
            'create_project_parameter': {
                'name': 'create_project_parameter',
                'description': ('Create a project parameter bound to one or more categories using a shared '
                                'parameter file (created automatically if absent).'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'name': {'type': 'string', 'description': 'Parameter name'},
                        'categories': {'type': 'array', 'items': {'type': 'string'}, 'description': 'Categories to bind to (e.g. ["Walls","Doors"])'},
                        'type': {'type': 'string', 'description': 'Parameter data type: Text (default), Integer, Number, Length, Area, YesNo'},
                        'group': {'type': 'string', 'description': 'Parameter group name (optional, default "Data")'},
                        'instance': {'type': 'boolean', 'description': 'Instance binding (default true) vs type binding'}
                    },
                    'required': ['name', 'categories']
                }
            },
            'room_to_floor': {
                'name': 'room_to_floor',
                'description': 'Create a floor matching the boundary of one or more rooms.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'room_ids': {'type': 'array', 'items': {'type': 'integer'}, 'description': 'Room element IDs to build floors from'},
                        'room_id': {'type': 'integer', 'description': 'Single room element ID (alternative to room_ids)'},
                        'floor_type': {'type': 'string', 'description': 'Floor type name (optional)'}
                    },
                    'required': []
                }
            },
            'purge_unused': {
                'name': 'purge_unused',
                'description': ('Report (and optionally delete) unused family symbols and view templates. '
                                'Defaults to a safe dry run that only reports counts.'),
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'dry_run': {'type': 'boolean', 'description': 'If true (default) only report; if false, delete unused items'}
                    },
                    'required': []
                }
            },
            'audit_model': {
                'name': 'audit_model',
                'description': ('Detailed model audit: warnings by type, imported CAD, groups, in-place '
                                'families, unused types/templates, and basic naming issues.'),
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            'create_workset': {
                'name': 'create_workset',
                'description': 'Create a new user workset (workshared models only).',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'name': {'type': 'string', 'description': 'Workset name'}
                    },
                    'required': ['name']
                }
            },
            # ── Utility / assistant control ───────────────────────────────────
            'file_watcher_status': {
                'name': 'file_watcher_status',
                'description': 'Check whether the T3Lab file-based task watcher is running and get its data directory path',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            'get_revit_context': {
                'name': 'get_revit_context',
                'description': 'Get current Revit context: active view, selected elements, open document info',
                'inputSchema': {'type': 'object', 'properties': {}, 'required': []}
            },
            'show_assistant_pane': {
                'name': 'show_assistant_pane',
                'description': 'Show or hide the T3Lab Assistant dockable pane',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'action': {
                            'type': 'string',
                            'description': '"show" (default) or "hide"'
                        },
                        'message': {
                            'type': 'string',
                            'description': 'Optional message to inject into the pane chat'
                        }
                    },
                    'required': []
                }
            },
            # ── Export ────────────────────────────────────────────────────────
            'export_sheets_pdf': {
                'name': 'export_sheets_pdf',
                'description': 'Export one or more sheets to PDF files in a specified folder',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'sheet_ids': {
                            'type': 'array', 'items': {'type': 'integer'},
                            'description': 'Sheet element IDs to export (omit to export all sheets)'
                        },
                        'output_folder': {
                            'type': 'string',
                            'description': 'Folder path to save PDF files. Default: same folder as .rvt file'
                        },
                        'combined': {
                            'type': 'boolean',
                            'description': 'Combine all sheets into one PDF (default false)'
                        }
                    },
                    'required': []
                }
            }
        }

    @property
    def port(self):
        return self._port

    @port.setter
    def port(self, value):
        self._port = value

    @property
    def is_running(self):
        return self._is_running

    @property
    def client_count(self):
        return len(self._clients)

    def _register_client(self, client_id):
        """Register a connected client"""
        self._clients[client_id] = {'connected': True}
        self._total_clients += 1

    def _unregister_client(self, client_id):
        """Unregister a disconnected client"""
        if client_id in self._clients:
            del self._clients[client_id]

    # ── Active document pinning ────────────────────────────────────────────
    # Every tool call resolves its target document from pyrevit.revit.doc,
    # which follows whichever document/window Revit itself last activated.
    # When several documents are open in this same Revit instance, that
    # "last activated" doc can silently change from under the AI client
    # (e.g. the user clicks into another tab). Pinning lets the user lock
    # tool execution onto one specific open document regardless of which
    # window currently has focus.

    def get_open_documents(self):
        """List non-linked documents open in this Revit instance.

        Returns a list of {'title', 'is_active', 'is_pinned'} dicts. Used by
        the MCP Control dialog to populate the document picker.
        """
        try:
            from pyrevit import HOST_APP, revit
            uiapp = HOST_APP.uiapp
            active_title = None
            try:
                active_title = revit.doc.Title
            except Exception:
                pass

            docs = []
            for d in uiapp.Application.Documents:
                if d.IsLinked:
                    continue
                docs.append({
                    'title': d.Title,
                    'is_active': d.Title == active_title,
                    'is_pinned': d.Title == self._pinned_doc_title,
                })
            return docs
        except Exception:
            return []

    def pin_document(self, title):
        """Pin a document by title so tool calls always target it."""
        self._pinned_doc_title = title or None

    def unpin_document(self):
        """Clear the pin — tool calls fall back to Revit's active document."""
        self._pinned_doc_title = None

    def get_pinned_document(self):
        """Return the pinned document title, or None if unpinned."""
        return self._pinned_doc_title

    def _resolve_target_document(self, doc, uidoc):
        """Return (doc, uidoc) to use for this tool call.

        If a document is pinned and it isn't the currently active one,
        redirect tool execution onto the pinned document: build a UIDocument
        that points at it (so uidoc-dependent tools — active view, selection,
        etc. — operate on the pinned document) and best-effort bring its
        window to the front so the visible Revit view matches. Falls back to
        the live active document if the pinned one is no longer open (e.g. it
        was closed).
        """
        if not self._pinned_doc_title:
            return doc, uidoc
        if doc is not None and doc.Title == self._pinned_doc_title:
            return doc, uidoc
        try:
            from pyrevit import HOST_APP
            uiapp = HOST_APP.uiapp

            target_doc = None
            for d in uiapp.Application.Documents:
                if not d.IsLinked and d.Title == self._pinned_doc_title:
                    target_doc = d
                    break

            if target_doc is None:
                # Pinned document is no longer open — clear the stale pin.
                self._pinned_doc_title = None
                return doc, uidoc

            # Best-effort: bring the pinned document's window to the front so
            # the visible Revit view matches where tools operate. Only saved
            # documents can be re-activated (OpenAndActivateDocument takes a
            # file path — there is NO Document overload; passing the Document
            # object, as the previous implementation did, silently threw and
            # left the pin completely ineffective). Already-open paths are
            # activated in place rather than reopened.
            try:
                path = target_doc.PathName
                if path:
                    uiapp.OpenAndActivateDocument(path)
            except Exception:
                pass

            # Resolve a UIDocument that actually points at the pinned document.
            # This is what genuinely redirects the tool call, independent of
            # whether the window switch above succeeded.
            target_uidoc = uidoc
            try:
                from Autodesk.Revit.UI import UIDocument
                active = uiapp.ActiveUIDocument
                if active is not None and active.Document.Title == self._pinned_doc_title:
                    target_uidoc = active
                else:
                    target_uidoc = UIDocument(target_doc)
            except Exception:
                try:
                    from Autodesk.Revit.UI import UIDocument
                    target_uidoc = UIDocument(target_doc)
                except Exception:
                    target_uidoc = uidoc

            return target_doc, target_uidoc
        except Exception:
            pass
        return doc, uidoc

    def _handle_initialize(self, params):
        """Handle MCP initialize request"""
        return {
            'protocolVersion': '2024-11-05',
            'capabilities': {
                'tools': {}
            },
            'serverInfo': {
                'name': 'T3LabAI Revit MCP Server',
                'version': '1.0.0'
            }
        }

    def _handle_tools_list(self):
        """Handle tools/list request"""
        return {
            'tools': list(self._tools.values())
        }

    def _handle_tool_call(self, params):
        """Handle tools/call request"""
        tool_name = params.get('name', '')
        arguments = params.get('arguments', {})

        if tool_name not in self._tools:
            return {
                'content': [{
                    'type': 'text',
                    'text': 'Error: Unknown tool: {}'.format(tool_name)
                }],
                'isError': True
            }

        # Execute tool and return result
        try:
            result = self._execute_tool(tool_name, arguments)
            return {
                'content': [{
                    'type': 'text',
                    'text': json.dumps(result, indent=2)
                }]
            }
        except Exception as e:
            return {
                'content': [{
                    'type': 'text',
                    'text': 'Error executing tool: {}'.format(str(e))
                }],
                'isError': True
            }

    # Tools that open a Transaction / mutate the model. These MUST run on
    # Revit's main thread via the ExternalEvent — starting a transaction from
    # the HTTP worker thread throws "outside of API context is not allowed".
    _WRITE_TOOLS = frozenset([
        'revit_override_color', 'create_level', 'place_wall', 'set_parameter',
        'create_point_based_element', 'create_line_based_element',
        'create_surface_based_element', 'create_grid', 'create_room',
        'create_structural_framing_system', 'delete_element', 'operate_element',
        'color_elements', 'tag_all_walls', 'tag_all_rooms', 'move_elements',
        'copy_elements', 'rotate_element', 'create_view', 'set_active_view',
        'rename_element', 'create_sheet', 'add_view_to_sheet', 'create_text_note',
        'set_element_workset', 'load_family', 'send_code_to_revit',
        'store_project_data', 'store_room_data', 'export_sheets_pdf', 'say_hello',
        'show_assistant_pane', 'split_curve', 'split_element', 'join_geometry',
        'bulk_set_parameter', 'select_elements', 'tag_elements', 'create_dimension',
        'create_schedule', 'duplicate_view', 'apply_view_template',
        'create_view_filter', 'place_views_on_sheets', 'export_dwg', 'export_image',
        'create_project_parameter', 'room_to_floor', 'purge_unused', 'create_workset',
        # Internal pseudo-tools (agent request TransactionGroup) — NOT in the
        # public registry, but they open/close a TransactionGroup so they must
        # run on the Revit main thread like any other write tool.
        '__begin_action_group', '__end_action_group',
    ])

    def ensure_external_event(self):
        """Create the ExternalEvent used to marshal tool execution onto Revit's
        main thread.

        This MUST be invoked from a valid Revit API context — i.e. the pyRevit
        pushbutton's main (UI) thread. ExternalEvent.Create() throws
        InvalidOperationException when called from a background worker thread,
        which is exactly why start_server() (run from the assistant's
        background startup probe) cannot create it: every write tool would then
        fall back to running its Transaction outside API context and fail.

        Returns (ok: bool, error: str|None).
        """
        if not HAS_REVIT_UI:
            return False, 'Revit UI API not available (RevitAPIUI not loaded)'
        if self._external_event is not None:
            return True, None
        try:
            self._event_handler = MCPExternalEventHandler(self)
            self._external_event = ExternalEvent.Create(self._event_handler)
            return True, None
        except Exception as e:
            self._event_handler = None
            self._external_event = None
            return False, str(e)

    def _execute_tool(self, tool_name, arguments):
        """Execute a Revit tool in a thread-safe manner using External Events."""
        if self._external_event:
            self._event_handler.tool_name = tool_name
            self._event_handler.arguments = arguments
            self._event_handler._lock.clear()
            self._external_event.Raise()

            # Wait for main UI thread execution. Large operations (framing
            # systems, PDF export, bulk deletes) legitimately take far longer
            # than 10s, so allow up to 120s before giving up.
            success = self._event_handler._lock.wait(timeout=120)
            if not success:
                return {'error': 'Execution timed out waiting for Revit thread context', 'tool': tool_name}
            if self._event_handler.exception:
                return {'error': str(self._event_handler.exception), 'tool': tool_name}
            return self._event_handler.result
        else:
            # No ExternalEvent — we're stuck on the HTTP worker thread. Read
            # tools tolerate this, but any write tool would throw a cryptic
            # "transaction outside API context" error. Surface a clear,
            # actionable message instead so the failure is diagnosable.
            if tool_name in self._WRITE_TOOLS:
                return {
                    'error': ('Revit ExternalEvent is not initialised, so model-editing tools '
                              'cannot run on the Revit main thread. Open (or reopen) the T3Lab '
                              'Assistant window once so it can register the event on the UI '
                              'thread, then retry.'),
                    'tool': tool_name,
                    'external_event_ready': False,
                }
            return self._execute_tool_in_context(tool_name, arguments)

    # ── Shared tool helpers ────────────────────────────────────────────────
    def _bic_map(self):
        """Human-facing category name → BuiltInCategory. Superset of the inline
        maps scattered through the older tools, used by the newer bulk/select/
        tag/filter tools so they all accept the same category vocabulary."""
        from Autodesk.Revit.DB import BuiltInCategory as B
        return {
            'Walls': B.OST_Walls, 'Floors': B.OST_Floors, 'Doors': B.OST_Doors,
            'Windows': B.OST_Windows, 'Rooms': B.OST_Rooms, 'Columns': B.OST_Columns,
            'StructuralColumns': B.OST_StructuralColumns, 'Beams': B.OST_StructuralFraming,
            'StructuralFraming': B.OST_StructuralFraming, 'Ceilings': B.OST_Ceilings,
            'Roofs': B.OST_Roofs, 'Furniture': B.OST_Furniture, 'Casework': B.OST_Casework,
            'Grids': B.OST_Grids, 'Levels': B.OST_Levels, 'Sheets': B.OST_Sheets,
            'Stairs': B.OST_Stairs, 'Railings': B.OST_StairsRailing,
            'Pipes': B.OST_PipeCurves, 'Ducts': B.OST_DuctCurves,
            'GenericModel': B.OST_GenericModel,
            'PlumbingFixtures': B.OST_PlumbingFixtures,
            'LightingFixtures': B.OST_LightingFixtures,
            'ElectricalFixtures': B.OST_ElectricalFixtures,
            'MechanicalEquipment': B.OST_MechanicalEquipment,
            'Parking': B.OST_Parking, 'PlantingArea': B.OST_Planting,
            'CurtainPanels': B.OST_CurtainWallPanels,
            'GenericAnnotations': B.OST_GenericAnnotation,
        }

    def _parse_color(self, color_str):
        """Parse a hex (#RRGGBB / #RGB) or CSS-name color into an (r, g, b)
        tuple, or None if unparseable. Mirrors the inline parser in
        revit_override_color so filter overrides accept the same vocabulary."""
        if not color_str:
            return None
        s = color_str.lower().strip()
        css = {
            'red': (255, 0, 0), 'green': (0, 255, 0), 'blue': (0, 0, 255),
            'orange': (255, 165, 0), 'cyan': (0, 255, 255), 'yellow': (255, 255, 0),
            'magenta': (255, 0, 255), 'black': (0, 0, 0), 'white': (255, 255, 255),
            'gray': (128, 128, 128), 'grey': (128, 128, 128), 'pink': (255, 192, 203),
            'purple': (128, 0, 128), 'violet': (238, 130, 238),
        }
        if s in css:
            return css[s]
        if s.startswith('#'):
            h = s[1:]
            try:
                if len(h) == 6:
                    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
                if len(h) == 3:
                    return (int(h[0] * 2, 16), int(h[1] * 2, 16), int(h[2] * 2, 16))
            except ValueError:
                return None
        return None

    def _apply_param_value(self, param, value):
        """Coerce a string value onto a Revit parameter honouring its storage
        type. Returns (ok, error). Mirrors the single-element set_parameter tool
        so bulk_set_parameter behaves identically per element."""
        from Autodesk.Revit.DB import StorageType, ElementId
        if param is None:
            return False, 'parameter not found'
        if param.IsReadOnly:
            return False, 'read-only'
        try:
            st = param.StorageType
            if st == StorageType.String:
                param.Set(value)
            elif st == StorageType.Double:
                parsed = False
                try:
                    parsed = param.SetValueString(value)
                except Exception:
                    parsed = False
                if not parsed:
                    param.Set(float(value))
            elif st == StorageType.Integer:
                param.Set(int(float(value)))
            elif st == StorageType.ElementId:
                param.Set(ElementId(int(value)))
            else:
                return False, 'unsupported storage type'
            return True, None
        except Exception as e:
            return False, str(e)

    def _execute_tool_in_context(self, tool_name, arguments):
        """Execute a Revit tool directly (must be inside Revit context thread)"""
        try:
            from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet,
                                           BuiltInCategory, Level, ElementId,
                                           Transaction, ElementLevelFilter)
            from pyrevit import revit
            doc = revit.doc
            uidoc = revit.uidoc
            doc, uidoc = self._resolve_target_document(doc, uidoc)
        except ImportError:
            return {'error': 'Revit API not available', 'tool': tool_name}

        # ── Agent request TransactionGroup (B4) ──────────────────────────────
        # Pseudo-tools, only reachable through _execute_tool (never listed in
        # the public registry). We are on the Revit main thread here (routed
        # via _WRITE_TOOLS + ExternalEvent), which TransactionGroup requires.
        if tool_name == '__begin_action_group':
            from Autodesk.Revit.DB import TransactionGroup
            # Defensive: never stack groups — close a leftover one first.
            if self._action_group is not None:
                try:
                    self._action_group.Assimilate()
                except Exception:
                    try:
                        self._action_group.RollBack()
                    except Exception:
                        pass
                self._action_group = None
            try:
                title = (arguments or {}).get('title') or 'T3Lab AI actions'
                tg = TransactionGroup(doc, title)
                tg.Start()
                self._action_group = tg
                return {'success': True, 'group': title}
            except Exception as e:
                self._action_group = None
                return {'error': str(e), 'tool': tool_name}

        elif tool_name == '__end_action_group':
            tg = self._action_group
            self._action_group = None
            if tg is None:
                return {'success': True, 'note': 'no open group'}
            try:
                # Assimilate merges every transaction inside into ONE undo item.
                tg.Assimilate()
                return {'success': True}
            except Exception as e:
                try:
                    tg.RollBack()
                except Exception:
                    pass
                return {'error': str(e), 'tool': tool_name}

        if tool_name == 'revit_get_active_view':
            view = doc.ActiveView
            if view is None:
                return {'error': 'No active view in Revit — open or activate a view first.'}
            return {
                'name': view.Name,
                'id': eid_value(view.Id),
                'type': str(view.ViewType)
            }

        elif tool_name == 'revit_get_selected_elements':
            selection = uidoc.Selection.GetElementIds()
            elements = []
            for eid in selection:
                elem = doc.GetElement(eid)
                elements.append({
                    'id': eid_value(eid),
                    'name': elem.Name if hasattr(elem, 'Name') else str(elem),
                    'category': elem.Category.Name if elem.Category else 'Unknown'
                })
            return {'selected_count': len(elements), 'elements': elements}

        elif tool_name == 'revit_get_project_info':
            info = doc.ProjectInformation
            return {
                'name': info.Name,
                'number': info.Number,
                'client': info.ClientName,
                'address': info.Address,
                'status': info.Status
            }

        elif tool_name == 'revit_list_views':
            from Autodesk.Revit.DB import View
            collector = FilteredElementCollector(doc).OfClass(View)
            views = []
            view_type_filter = arguments.get('view_type')
            for v in collector:
                if not v.IsTemplate:
                    vtype = str(v.ViewType)
                    if view_type_filter is None or vtype == view_type_filter:
                        views.append({
                            'name': v.Name,
                            'id': eid_value(v.Id),
                            'type': vtype
                        })
            return {'count': len(views), 'views': views}

        elif tool_name == 'revit_list_sheets':
            collector = FilteredElementCollector(doc).OfClass(ViewSheet)
            sheets = []
            for s in collector:
                sheets.append({
                    'name': s.Name,
                    'number': s.SheetNumber,
                    'id': eid_value(s.Id)
                })
            return {'count': len(sheets), 'sheets': sheets}

        elif tool_name == 'revit_get_element_info':
            from Autodesk.Revit.DB import ElementId
            eid = ElementId(arguments.get('element_id', 0))
            elem = doc.GetElement(eid)
            if elem:
                params = {}
                for p in elem.Parameters:
                    try:
                        params[p.Definition.Name] = p.AsValueString() or p.AsString() or str(p.AsDouble())
                    except Exception:
                        pass
                return {
                    'id': eid_value(elem.Id),
                    'name': elem.Name if hasattr(elem, 'Name') else str(elem),
                    'category': elem.Category.Name if elem.Category else 'Unknown',
                    'parameters': params
                }
            return {'error': 'Element not found'}

        elif tool_name == 'revit_override_color':
            color_str = arguments.get('color')
            element_ids = arguments.get('element_ids')
            
            # If element_ids is omitted or empty, use the active selection
            if not element_ids:
                selection = uidoc.Selection.GetElementIds()
                element_ids = [eid_value(eid) for eid in selection]
                
            if not element_ids:
                return {'error': 'No elements specified and no elements are selected in Revit.'}
                
            # Parse color
            r, g, b = 255, 0, 0 # default red
            if color_str:
                color_str = color_str.lower().strip()
                css_colors = {
                    'red': (255, 0, 0),
                    'green': (0, 255, 0),
                    'blue': (0, 0, 255),
                    'orange': (255, 165, 0),
                    'cyan': (0, 255, 255),
                    'yellow': (255, 255, 0),
                    'magenta': (255, 0, 255),
                    'black': (0, 0, 0),
                    'white': (255, 255, 255),
                    'gray': (128, 128, 128),
                    'grey': (128, 128, 128),
                    'pink': (255, 192, 203),
                    'purple': (128, 0, 128),
                    'violet': (238, 130, 238),
                }
                if color_str in css_colors:
                    r, g, b = css_colors[color_str]
                elif color_str.startswith('#'):
                    hex_val = color_str[1:]
                    if len(hex_val) == 6:
                        try:
                            r = int(hex_val[0:2], 16)
                            g = int(hex_val[2:4], 16)
                            b = int(hex_val[4:6], 16)
                        except ValueError:
                            pass
                    elif len(hex_val) == 3:
                        try:
                            r = int(hex_val[0]*2, 16)
                            g = int(hex_val[1]*2, 16)
                            b = int(hex_val[2]*2, 16)
                        except ValueError:
                            pass
            
            from Autodesk.Revit.DB import Color, OverrideGraphicSettings, ElementId, Transaction
            revit_color = Color(r, g, b)
            
            # Find solid fill pattern for surface fill
            solid_pattern_id = None
            try:
                from Autodesk.Revit.DB import FilteredElementCollector, FillPatternElement
                fill_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement)
                for fp in fill_patterns:
                    pattern = fp.GetFillPattern()
                    if pattern and pattern.IsSolidFill:
                        solid_pattern_id = fp.Id
                        break
            except Exception:
                pass
                
            if solid_pattern_id is None:
                solid_pattern_id = ElementId(-1)
                
            override_settings = OverrideGraphicSettings()
            override_settings.SetProjectionLineColor(revit_color)
            if solid_pattern_id != ElementId(-1):
                try:
                    override_settings.SetSurfaceForegroundPatternId(solid_pattern_id)
                    override_settings.SetSurfaceForegroundPatternColor(revit_color)
                    override_settings.SetCutForegroundPatternId(solid_pattern_id)
                    override_settings.SetCutForegroundPatternColor(revit_color)
                except Exception:
                    pass
            
            view = doc.ActiveView
            t = Transaction(doc, "T3Lab AI Override Color")
            t.Start()
            overridden_count = 0
            for eid_val in element_ids:
                try:
                    eid = ElementId(int(eid_val))
                    view.SetElementOverrides(eid, override_settings)
                    overridden_count += 1
                except Exception:
                    pass
            t.Commit()
            
            return {
                'success': True,
                'overridden_count': overridden_count,
                'color': color_str or 'red',
                'rgb': [r, g, b]
            }

        elif tool_name == 'create_level':
            from Autodesk.Revit.DB import Level, Transaction, ElementId
            elevation_m = float(arguments.get('elevation', 0.0))
            level_name = arguments.get('name')
            # Convert meters to internal Revit units (feet): 1 m = 3.28084 ft
            METERS_TO_FEET = 3.28084
            elevation_ft = elevation_m * METERS_TO_FEET
            t = Transaction(doc, "T3Lab AI: Create Level")
            t.Start()
            try:
                new_level = Level.Create(doc, elevation_ft)
                if level_name:
                    try:
                        new_level.Name = level_name
                    except Exception:
                        # Name already taken — uniquify instead of failing
                        for suffix in range(2, 100):
                            try:
                                new_level.Name = '{} ({})'.format(level_name, suffix)
                                break
                            except Exception:
                                continue
                t.Commit()
                return {
                    'success': True,
                    'level_id': eid_value(new_level.Id),
                    'name': new_level.Name,
                    'elevation_m': elevation_m
                }
            except Exception as ex:
                t.RollBack()
                return {'error': 'Failed to create level: {}'.format(str(ex))}

        elif tool_name == 'place_wall':
            from Autodesk.Revit.DB import (
                Wall, Line, XYZ, Transaction, ElementId,
                FilteredElementCollector, Level, WallType
            )
            METERS_TO_FEET = 3.28084
            start_x = float(arguments.get('start_x', 0.0)) * METERS_TO_FEET
            start_y = float(arguments.get('start_y', 0.0)) * METERS_TO_FEET
            end_x   = float(arguments.get('end_x', 1.0))  * METERS_TO_FEET
            end_y   = float(arguments.get('end_y', 0.0))  * METERS_TO_FEET
            height_m = float(arguments.get('height', 3.0))
            height_ft = height_m * METERS_TO_FEET
            level_name_arg = arguments.get('level_name')
            wall_type_name_arg = arguments.get('wall_type_name')

            # Find level
            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = None
            if level_name_arg:
                for lv in levels:
                    if lv.Name == level_name_arg:
                        target_level = lv
                        break
            if target_level is None and levels:
                target_level = levels[0]
            if target_level is None:
                return {'error': 'No levels found in the document'}

            # Find wall type — exact name, then substring, then first Basic
            # wall (curtain/stacked types can't be created via Wall.Create
            # with an arbitrary line + height).
            from Autodesk.Revit.DB import WallKind
            wall_types = list(FilteredElementCollector(doc).OfClass(WallType).ToElements())
            target_wall_type = None
            if wall_type_name_arg:
                for wt in wall_types:
                    if wt.Name == wall_type_name_arg:
                        target_wall_type = wt
                        break
                if target_wall_type is None:
                    for wt in wall_types:
                        if wall_type_name_arg.lower() in wt.Name.lower():
                            target_wall_type = wt
                            break
                if target_wall_type is None:
                    return {'error': 'Wall type "{}" not found'.format(wall_type_name_arg)}
            if target_wall_type is None:
                for wt in wall_types:
                    try:
                        if wt.Kind == WallKind.Basic:
                            target_wall_type = wt
                            break
                    except Exception:
                        pass
            if target_wall_type is None and wall_types:
                target_wall_type = wall_types[0]
            if target_wall_type is None:
                return {'error': 'No wall types found in the document'}

            # Build geometry
            p1 = XYZ(start_x, start_y, 0.0)
            p2 = XYZ(end_x, end_y, 0.0)
            try:
                wall_line = Line.CreateBound(p1, p2)
            except Exception as ex:
                return {'error': 'Invalid wall line: {}'.format(str(ex))}

            # Compute length in meters for the response
            import math
            dx = float(arguments.get('end_x', 0.0)) - float(arguments.get('start_x', 0.0))
            dy = float(arguments.get('end_y', 0.0)) - float(arguments.get('start_y', 0.0))
            length_m = math.sqrt(dx * dx + dy * dy)

            t = Transaction(doc, "T3Lab AI: Place Wall")
            t.Start()
            try:
                new_wall = Wall.Create(doc, wall_line, target_wall_type.Id, target_level.Id, height_ft, 0.0, False, False)
                t.Commit()
                return {
                    'success': True,
                    'wall_id': eid_value(new_wall.Id),
                    'length_m': round(length_m, 4)
                }
            except Exception as ex:
                t.RollBack()
                return {'error': 'Failed to place wall: {}'.format(str(ex))}

        elif tool_name == 'get_parameter':
            from Autodesk.Revit.DB import ElementId
            element_id_int = int(arguments.get('element_id', 0))
            parameter_name = arguments.get('parameter_name', '')
            eid = ElementId(element_id_int)
            elem = doc.GetElement(eid)
            if elem is None:
                return {'error': 'Element not found: {}'.format(element_id_int)}
            param = elem.LookupParameter(parameter_name)
            if param is None:
                return {'error': 'Parameter "{}" not found on element {}'.format(parameter_name, element_id_int)}
            value = None
            units = ''
            try:
                value = param.AsValueString()
                if value is None:
                    value = param.AsString()
                if value is None:
                    raw = param.AsDouble()
                    value = str(raw)
                    try:
                        units = str(param.DisplayUnitType)
                    except Exception:
                        pass
            except Exception as ex:
                return {'error': 'Failed to read parameter: {}'.format(str(ex))}
            return {
                'element_id': element_id_int,
                'parameter': parameter_name,
                'value': value,
                'units': units
            }

        elif tool_name == 'get_elements_by_level':
            from Autodesk.Revit.DB import (
                FilteredElementCollector, Level,
                ElementLevelFilter, BuiltInCategory
            )
            level_name_arg = arguments.get('level_name', '')
            category_arg = arguments.get('category')

            # Find level by name
            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = None
            for lv in levels:
                if lv.Name == level_name_arg:
                    target_level = lv
                    break
            if target_level is None:
                return {'error': 'Level "{}" not found'.format(level_name_arg)}

            # Category name -> BuiltInCategory mapping
            category_map = {
                'Walls':    BuiltInCategory.OST_Walls,
                'Floors':   BuiltInCategory.OST_Floors,
                'Columns':  BuiltInCategory.OST_Columns,
                'Doors':    BuiltInCategory.OST_Doors,
                'Windows':  BuiltInCategory.OST_Windows,
                'Beams':    BuiltInCategory.OST_StructuralFraming,
            }

            level_filter = ElementLevelFilter(target_level.Id)
            collector = FilteredElementCollector(doc).WherePasses(level_filter)

            if category_arg and category_arg in category_map:
                from Autodesk.Revit.DB import ElementCategoryFilter
                cat_filter = ElementCategoryFilter(category_map[category_arg])
                from Autodesk.Revit.DB import LogicalAndFilter
                combined = LogicalAndFilter(level_filter, cat_filter)
                collector = FilteredElementCollector(doc).WherePasses(combined)

            elements_out = []
            for elem in collector:
                try:
                    cat_name = elem.Category.Name if elem.Category else 'Unknown'
                    elem_name = elem.Name if hasattr(elem, 'Name') else ''
                    type_name = ''
                    try:
                        type_elem = doc.GetElement(elem.GetTypeId())
                        if type_elem is not None:
                            type_name = type_elem.Name if hasattr(type_elem, 'Name') else ''
                    except Exception:
                        pass
                    elements_out.append({
                        'id': eid_value(elem.Id),
                        'name': elem_name,
                        'category': cat_name,
                        'type_name': type_name
                    })
                except Exception:
                    pass

            return {
                'level': level_name_arg,
                'count': len(elements_out),
                'elements': elements_out
            }

        # ── get_current_view_info ───────────────────────────────────────────────
        elif tool_name == 'get_current_view_info':
            view = doc.ActiveView
            if view is None:
                return {'error': 'No active view in Revit — open or activate a view first.'}
            result = {
                'name': view.Name,
                'id': eid_value(view.Id),
                'type': str(view.ViewType),
                'scale': view.Scale if hasattr(view, 'Scale') else None,
                'discipline': str(view.Discipline) if hasattr(view, 'Discipline') else None,
                'detail_level': str(view.DetailLevel) if hasattr(view, 'DetailLevel') else None,
                'is_template': view.IsTemplate,
            }
            try:
                cr = view.CropBox
                result['crop_box'] = {
                    'min_x': round(cr.Min.X * 0.3048, 3),
                    'min_y': round(cr.Min.Y * 0.3048, 3),
                    'max_x': round(cr.Max.X * 0.3048, 3),
                    'max_y': round(cr.Max.Y * 0.3048, 3),
                }
            except Exception:
                pass
            return result

        # ── get_current_view_elements ────────────────────────────────────────
        elif tool_name == 'get_current_view_elements':
            cat_arg   = arguments.get('category')
            limit_arg = int(arguments.get('limit', 100))
            view      = doc.ActiveView
            if view is None:
                return {'error': 'No active view in Revit — open or activate a view first.'}

            CATEGORY_MAP = {
                'Walls': BuiltInCategory.OST_Walls,
                'Floors': BuiltInCategory.OST_Floors,
                'Doors': BuiltInCategory.OST_Doors,
                'Windows': BuiltInCategory.OST_Windows,
                'Rooms': BuiltInCategory.OST_Rooms,
                'Columns': BuiltInCategory.OST_Columns,
                'Beams': BuiltInCategory.OST_StructuralFraming,
                'Ceilings': BuiltInCategory.OST_Ceilings,
                'Roofs': BuiltInCategory.OST_Roofs,
                'Furniture': BuiltInCategory.OST_Furniture,
                'Grids': BuiltInCategory.OST_Grids,
                'Levels': BuiltInCategory.OST_Levels,
            }
            if cat_arg and cat_arg in CATEGORY_MAP:
                collector = FilteredElementCollector(doc, view.Id).OfCategory(CATEGORY_MAP[cat_arg]).WhereElementIsNotElementType()
            else:
                collector = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()

            elements_out = []
            for elem in collector:
                if len(elements_out) >= limit_arg:
                    break
                try:
                    cat_name = elem.Category.Name if elem.Category else 'Unknown'
                    elements_out.append({
                        'id': eid_value(elem.Id),
                        'name': elem.Name if hasattr(elem, 'Name') else '',
                        'category': cat_name,
                    })
                except Exception:
                    pass
            return {'view': view.Name, 'count': len(elements_out), 'elements': elements_out}

        # ── get_available_family_types ───────────────────────────────────────
        elif tool_name == 'get_available_family_types':
            from Autodesk.Revit.DB import FamilySymbol
            cat_arg = arguments.get('category', '')
            collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
            types_out = []
            for sym in collector:
                try:
                    cat_name = sym.Category.Name if sym.Category else ''
                    if cat_arg and cat_name != cat_arg:
                        continue
                    family_name = sym.Family.Name if sym.Family else ''
                    types_out.append({
                        'id': eid_value(sym.Id),
                        'family': family_name,
                        'type': sym.Name,
                        'category': cat_name,
                        'active': sym.IsActive,
                    })
                except Exception:
                    pass
            return {'count': len(types_out), 'types': types_out}

        # ── get_material_quantities ──────────────────────────────────────────
        elif tool_name == 'get_material_quantities':
            cat_arg   = arguments.get('category', 'Walls')
            lvl_arg   = arguments.get('level_name')
            QTY_CATEGORY_MAP = {
                'Walls':    BuiltInCategory.OST_Walls,
                'Floors':   BuiltInCategory.OST_Floors,
                'Roofs':    BuiltInCategory.OST_Roofs,
                'Ceilings': BuiltInCategory.OST_Ceilings,
            }
            bic = QTY_CATEGORY_MAP.get(cat_arg)
            if bic is None:
                return {'error': 'Unsupported category "{}". Use one of: {}'.format(
                    cat_arg, ', '.join(sorted(QTY_CATEGORY_MAP.keys())))}
            collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
            total_area_m2   = 0.0
            total_volume_m3 = 0.0
            items = []
            for elem in collector:
                try:
                    lvl_name = ''
                    try:
                        lv = doc.GetElement(elem.LevelId)
                        lvl_name = lv.Name if lv else ''
                    except Exception:
                        pass
                    if not lvl_name:
                        # Roofs/ceilings expose their level as a parameter
                        try:
                            lp = elem.LookupParameter('Level')
                            if lp:
                                lvl_name = lp.AsValueString() or ''
                        except Exception:
                            pass
                    if lvl_arg and lvl_name != lvl_arg:
                        continue
                    area_ft2 = 0.0
                    vol_ft3  = 0.0
                    try:
                        ap = elem.LookupParameter('Area')
                        if ap:
                            area_ft2 = ap.AsDouble()
                    except Exception:
                        pass
                    try:
                        vp = elem.LookupParameter('Volume')
                        if vp:
                            vol_ft3 = vp.AsDouble()
                    except Exception:
                        pass
                    area_m2 = round(area_ft2 * 0.0929, 3)
                    vol_m3  = round(vol_ft3 * 0.0283, 3)
                    total_area_m2   += area_m2
                    total_volume_m3 += vol_m3
                    items.append({
                        'id': eid_value(elem.Id),
                        'level': lvl_name,
                        'area_m2': area_m2,
                        'volume_m3': vol_m3,
                    })
                except Exception:
                    pass
            return {
                'category': cat_arg,
                'count': len(items),
                'total_area_m2': round(total_area_m2, 3),
                'total_volume_m3': round(total_volume_m3, 3),
                'elements': items,
            }

        # ── ai_element_filter ────────────────────────────────────────────────
        elif tool_name == 'ai_element_filter':
            cat_arg   = arguments.get('category', 'Walls')
            param_arg = arguments.get('parameter_name')
            val_arg   = (arguments.get('parameter_value') or '').lower()
            limit_arg = int(arguments.get('limit', 50))
            CATEGORY_MAP = {
                'Walls': BuiltInCategory.OST_Walls,
                'Floors': BuiltInCategory.OST_Floors,
                'Doors': BuiltInCategory.OST_Doors,
                'Windows': BuiltInCategory.OST_Windows,
                'Rooms': BuiltInCategory.OST_Rooms,
                'Columns': BuiltInCategory.OST_Columns,
                'Beams': BuiltInCategory.OST_StructuralFraming,
                'Ceilings': BuiltInCategory.OST_Ceilings,
                'Roofs': BuiltInCategory.OST_Roofs,
                'Grids': BuiltInCategory.OST_Grids,
            }
            bic = CATEGORY_MAP.get(cat_arg)
            if bic:
                collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
            else:
                collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
            results = []
            for elem in collector:
                if len(results) >= limit_arg:
                    break
                try:
                    match = True
                    param_val = ''
                    if param_arg:
                        p = elem.LookupParameter(param_arg)
                        if p:
                            param_val = p.AsValueString() or p.AsString() or ''
                            if val_arg and val_arg not in param_val.lower():
                                match = False
                        else:
                            match = False
                    if match:
                        results.append({
                            'id': eid_value(elem.Id),
                            'name': elem.Name if hasattr(elem, 'Name') else '',
                            'category': elem.Category.Name if elem.Category else '',
                            'param_value': param_val,
                        })
                except Exception:
                    pass
            return {'category': cat_arg, 'filter_param': param_arg, 'count': len(results), 'elements': results}

        # ── analyze_model_statistics ─────────────────────────────────────────
        elif tool_name == 'analyze_model_statistics':
            from Autodesk.Revit.DB import View, BuiltInCategory as BIC
            stat_cats = [
                ('Walls',    BIC.OST_Walls),
                ('Floors',   BIC.OST_Floors),
                ('Doors',    BIC.OST_Doors),
                ('Windows',  BIC.OST_Windows),
                ('Rooms',    BIC.OST_Rooms),
                ('Columns',  BIC.OST_Columns),
                ('Beams',    BIC.OST_StructuralFraming),
                ('Ceilings', BIC.OST_Ceilings),
                ('Roofs',    BIC.OST_Roofs),
                ('Grids',    BIC.OST_Grids),
                ('Levels',   BIC.OST_Levels),
                ('Sheets',   BIC.OST_Sheets),
            ]
            stats = {}
            for cat_name, bic in stat_cats:
                try:
                    count = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().GetElementCount()
                    stats[cat_name] = count
                except Exception:
                    stats[cat_name] = 0
            total_views = FilteredElementCollector(doc).OfClass(View).GetElementCount()
            return {
                'element_counts': stats,
                'total_views': total_views,
                'project': doc.ProjectInformation.Name,
            }

        # ── create_point_based_element ───────────────────────────────────────
        elif tool_name == 'create_point_based_element':
            from Autodesk.Revit.DB import FamilySymbol, XYZ, Transaction, Wall
            from Autodesk.Revit.DB.Structure import StructuralType
            M2FT = 3.28084
            ftype_name = arguments.get('family_type', '')
            x_ft = float(arguments.get('x', 0)) * M2FT
            y_ft = float(arguments.get('y', 0)) * M2FT
            z_arg = arguments.get('z')
            lvl_name = arguments.get('level_name')
            host_id_arg = arguments.get('host_wall_id')

            # Resolve family symbol: exact type name, "Family:Type", then
            # case-insensitive substring match as a fallback.
            sym = None
            partial = None
            for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
                try:
                    full = '{}:{}'.format(s.Family.Name, s.Name) if s.Family else s.Name
                    if s.Name == ftype_name or full == ftype_name:
                        sym = s
                        break
                    if partial is None and ftype_name.lower() in full.lower():
                        partial = s
                except Exception:
                    pass
            if sym is None:
                sym = partial
            if sym is None:
                return {'error': 'Family type "{}" not found. Use get_available_family_types to list loaded types.'.format(ftype_name)}

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break
            if target_level is None:
                return {'error': 'No levels found in the document'}

            # Default Z to the target level's elevation so the element lands
            # on the level rather than at absolute 0.
            z_ft = float(z_arg) * M2FT if z_arg is not None else target_level.Elevation
            point = XYZ(x_ft, y_ft, z_ft)

            cat_name = sym.Category.Name if sym.Category else ''
            needs_host = cat_name in ('Doors', 'Windows')
            try:
                placement = str(sym.Family.FamilyPlacementType)
                if 'Hosted' in placement:
                    needs_host = True
            except Exception:
                pass

            host_elem = None
            if host_id_arg:
                host_elem = doc.GetElement(ElementId(int(host_id_arg)))
                if host_elem is None:
                    return {'error': 'Host element not found: {}'.format(host_id_arg)}
            elif needs_host:
                # Auto-pick the nearest wall to the placement point.
                best_d = 1e30
                for w in FilteredElementCollector(doc).OfClass(Wall):
                    try:
                        crv = w.Location.Curve
                        d = crv.Distance(XYZ(x_ft, y_ft, crv.GetEndPoint(0).Z))
                        if d < best_d:
                            best_d = d
                            host_elem = w
                    except Exception:
                        pass
                # Reject hosts farther than ~2 m — the point is nowhere near a wall.
                if host_elem is None or best_d > 2.0 * M2FT:
                    return {'error': '"{}" is a hosted family (doors/windows need a wall). '
                                     'Pass host_wall_id or place the point on/near a wall.'.format(ftype_name)}

            struct_type = StructuralType.NonStructural
            if cat_name == 'Structural Columns':
                struct_type = StructuralType.Column

            t = Transaction(doc, 'T3Lab AI Create Element')
            t.Start()
            try:
                if not sym.IsActive:
                    sym.Activate()
                    doc.Regenerate()
                if host_elem is not None:
                    inst = doc.Create.NewFamilyInstance(point, sym, host_elem, target_level, struct_type)
                else:
                    inst = doc.Create.NewFamilyInstance(point, sym, target_level, struct_type)
                t.Commit()
                result = {'success': True, 'element_id': eid_value(inst.Id),
                          'type': '{}:{}'.format(sym.Family.Name, sym.Name) if sym.Family else sym.Name,
                          'level': target_level.Name}
                if host_elem is not None:
                    result['host_id'] = eid_value(host_elem.Id)
                return result
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── create_line_based_element ────────────────────────────────────────
        elif tool_name == 'create_line_based_element':
            from Autodesk.Revit.DB import FamilySymbol, XYZ, Line, Transaction
            from Autodesk.Revit.DB.Structure import StructuralType
            M2FT = 3.28084
            ftype_name = arguments.get('family_type', '')
            sx = float(arguments.get('start_x', 0)) * M2FT
            sy = float(arguments.get('start_y', 0)) * M2FT
            ex = float(arguments.get('end_x', 0)) * M2FT
            ey = float(arguments.get('end_y', 0)) * M2FT
            sz_arg = arguments.get('start_z')
            ez_arg = arguments.get('end_z')
            lvl_name = arguments.get('level_name')

            sym = None
            partial = None
            for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
                try:
                    full = '{}:{}'.format(s.Family.Name, s.Name) if s.Family else s.Name
                    if s.Name == ftype_name or full == ftype_name:
                        sym = s
                        break
                    if partial is None and ftype_name.lower() in full.lower():
                        partial = s
                except Exception:
                    pass
            if sym is None:
                sym = partial
            if sym is None:
                return {'error': 'Family type "{}" not found. Use get_available_family_types to list loaded types.'.format(ftype_name)}

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break
            if target_level is None:
                return {'error': 'No levels found in the document'}

            # Default Z to the reference level's elevation so beams land on
            # the level instead of absolute 0.
            sz = float(sz_arg) * M2FT if sz_arg is not None else target_level.Elevation
            ez = float(ez_arg) * M2FT if ez_arg is not None else target_level.Elevation

            p1 = XYZ(sx, sy, sz)
            p2 = XYZ(ex, ey, ez)
            if p1.DistanceTo(p2) < 0.01:
                return {'error': 'Start and end points are (nearly) identical — cannot create a line element.'}
            curve = Line.CreateBound(p1, p2)

            cat_name = sym.Category.Name if sym.Category else ''
            if cat_name == 'Structural Framing':
                struct_type = StructuralType.Beam
            elif cat_name == 'Structural Columns':
                struct_type = StructuralType.Column
            else:
                struct_type = StructuralType.NonStructural

            t = Transaction(doc, 'T3Lab AI Create Line Element')
            t.Start()
            try:
                if not sym.IsActive:
                    sym.Activate()
                    doc.Regenerate()
                inst = doc.Create.NewFamilyInstance(curve, sym, target_level, struct_type)
                t.Commit()
                return {'success': True, 'element_id': eid_value(inst.Id),
                        'type': '{}:{}'.format(sym.Family.Name, sym.Name) if sym.Family else sym.Name,
                        'level': target_level.Name}
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── create_surface_based_element ─────────────────────────────────────
        elif tool_name == 'create_surface_based_element':
            from Autodesk.Revit.DB import (XYZ, Line, CurveArray, CurveLoop,
                                           Transaction, Floor, FloorType)
            from System.Collections.Generic import List as NetList
            elem_type = (arguments.get('element_type') or 'floor').lower()
            boundary  = arguments.get('boundary_points', [])
            lvl_name  = arguments.get('level_name')
            type_name = arguments.get('type_name')
            M2FT = 3.28084

            if len(boundary) < 3:
                return {'error': 'boundary_points must have at least 3 points'}

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break
            if target_level is None:
                return {'error': 'No levels found in the document'}

            try:
                rev_ver = int(doc.Application.VersionNumber)
            except Exception:
                rev_ver = 0

            # Build the boundary at the level's elevation so the sketch sits
            # on the level (offset 0) instead of at absolute Z=0.
            z_ft = target_level.Elevation
            pts = [XYZ(float(p[0]) * M2FT, float(p[1]) * M2FT, z_ft) for p in boundary]
            segments = []
            for i in range(len(pts)):
                p1 = pts[i]
                p2 = pts[(i + 1) % len(pts)]
                if p1.DistanceTo(p2) > 0.01:
                    segments.append(Line.CreateBound(p1, p2))
            if len(segments) < 3:
                return {'error': 'boundary_points do not form a valid polygon (duplicate/too-close points)'}

            def _pick_type(types):
                picked = types[0] if types else None
                if type_name:
                    for candidate in types:
                        if candidate.Name == type_name:
                            return candidate
                    for candidate in types:
                        if type_name.lower() in candidate.Name.lower():
                            return candidate
                return picked

            t = Transaction(doc, 'T3Lab AI Create Surface Element')
            t.Start()
            try:
                if elem_type == 'ceiling':
                    if rev_ver < 2022:
                        t.RollBack()
                        return {'error': 'Ceiling creation via API requires Revit 2022 or newer (running {}).'.format(rev_ver or 'unknown')}
                    from Autodesk.Revit.DB import Ceiling, CeilingType
                    ceil_types = list(FilteredElementCollector(doc).OfClass(CeilingType).ToElements())
                    ct = _pick_type(ceil_types)
                    if ct is None:
                        t.RollBack()
                        return {'error': 'No ceiling types found in the document'}
                    loop = CurveLoop()
                    for seg in segments:
                        loop.Append(seg)
                    profile = NetList[CurveLoop]()
                    profile.Add(loop)
                    new_elem = Ceiling.Create(doc, profile, ct.Id, target_level.Id)
                    type_used = ct.Name

                elif elem_type == 'roof':
                    from Autodesk.Revit.DB import RoofType, ModelCurveArray, BuiltInCategory as _BIC
                    roof_types = list(FilteredElementCollector(doc)
                                      .OfCategory(_BIC.OST_Roofs)
                                      .OfClass(RoofType).ToElements())
                    rt = _pick_type(roof_types)
                    if rt is None:
                        t.RollBack()
                        return {'error': 'No roof types found in the document'}
                    ca = CurveArray()
                    for seg in segments:
                        ca.Append(seg)
                    ma = ModelCurveArray()
                    new_elem = doc.Create.NewFootPrintRoof(ca, target_level, rt, ma)
                    type_used = rt.Name

                else:  # floor (default)
                    floor_types = list(FilteredElementCollector(doc).OfClass(FloorType).ToElements())
                    ft = _pick_type(floor_types)
                    if ft is None:
                        t.RollBack()
                        return {'error': 'No floor types found in the document'}
                    if rev_ver >= 2022:
                        loop = CurveLoop()
                        for seg in segments:
                            loop.Append(seg)
                        profile = NetList[CurveLoop]()
                        profile.Add(loop)
                        new_elem = Floor.Create(doc, profile, ft.Id, target_level.Id)
                    else:
                        ca = CurveArray()
                        for seg in segments:
                            ca.Append(seg)
                        new_elem = doc.Create.NewFloor(ca, ft, target_level, False)
                    type_used = ft.Name

                t.Commit()
                return {'success': True, 'element_id': eid_value(new_elem.Id),
                        'element_type': elem_type, 'type_used': type_used,
                        'level': target_level.Name}
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── create_grid ──────────────────────────────────────────────────────
        elif tool_name == 'create_grid':
            from Autodesk.Revit.DB import Grid, XYZ, Line, Transaction
            x_spacings = arguments.get('x_spacings', [])
            y_spacings = arguments.get('y_spacings', [])
            ox = float(arguments.get('origin_x', 0)) * 3.28084
            oy = float(arguments.get('origin_y', 0)) * 3.28084
            x_labels = arguments.get('x_labels', [])
            y_labels = arguments.get('y_labels', [])

            def auto_alpha(n):
                labels = []
                for i in range(n):
                    labels.append(chr(ord('A') + i) if i < 26 else 'A{}'.format(i - 25))
                return labels

            if not x_labels:
                x_labels = auto_alpha(len(x_spacings) + 1)
            if not y_labels:
                y_labels = [str(i + 1) for i in range(len(y_spacings) + 1)]

            M2FT = 3.28084
            total_x_ft = sum(x_spacings) * M2FT
            total_y_ft = sum(y_spacings) * M2FT
            # Extend grid lines past the last gridline on each side (min ~3 m)
            margin_ft = max(10.0, 0.15 * max(total_x_ft, total_y_ft))

            def _safe_name(grid, wanted, warnings):
                """Rename a grid, falling back to auto name on duplicates
                instead of failing the whole transaction."""
                try:
                    grid.Name = wanted
                except Exception:
                    warnings.append('Grid name "{}" already in use — kept auto name "{}"'.format(wanted, grid.Name))

            grid_ids = []
            name_warnings = []
            t = Transaction(doc, 'T3Lab AI Create Grid')
            t.Start()
            try:
                # Vertical lines (along Y) at X positions
                x_pos = ox
                for i, spacing in enumerate([0.0] + [s * M2FT for s in x_spacings]):
                    x_pos = ox if i == 0 else x_pos + spacing
                    start = XYZ(x_pos, oy - margin_ft, 0)
                    end   = XYZ(x_pos, oy + total_y_ft + margin_ft, 0)
                    g = Grid.Create(doc, Line.CreateBound(start, end))
                    if i < len(x_labels):
                        _safe_name(g, x_labels[i], name_warnings)
                    grid_ids.append(eid_value(g.Id))

                # Horizontal lines (along X) at Y positions
                y_pos = oy
                for i, spacing in enumerate([0.0] + [s * M2FT for s in y_spacings]):
                    y_pos = oy if i == 0 else y_pos + spacing
                    start = XYZ(ox - margin_ft, y_pos, 0)
                    end   = XYZ(ox + total_x_ft + margin_ft, y_pos, 0)
                    g = Grid.Create(doc, Line.CreateBound(start, end))
                    if i < len(y_labels):
                        _safe_name(g, y_labels[i], name_warnings)
                    grid_ids.append(eid_value(g.Id))

                t.Commit()
                result = {'success': True, 'grid_count': len(grid_ids), 'grid_ids': grid_ids}
                if name_warnings:
                    result['warnings'] = name_warnings
                return result
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── create_room ──────────────────────────────────────────────────────
        elif tool_name == 'create_room':
            from Autodesk.Revit.DB import UV, Transaction
            from Autodesk.Revit.DB.Architecture import Room
            x_m = float(arguments.get('x', 0)) * 3.28084
            y_m = float(arguments.get('y', 0)) * 3.28084
            lvl_name  = arguments.get('level_name')
            room_name = arguments.get('name', 'Room')
            room_num  = arguments.get('number', '')

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break
            if target_level is None:
                return {'error': 'No level found'}

            t = Transaction(doc, 'T3Lab AI Create Room')
            t.Start()
            try:
                pt = UV(x_m, y_m)
                room = doc.Create.NewRoom(target_level, pt)
                if room_name:
                    room.Name = room_name
                if room_num:
                    room.Number = room_num
                doc.Regenerate()
                area_ft2 = 0.0
                try:
                    area_ft2 = room.Area
                except Exception:
                    pass
                t.Commit()
                result = {'success': True, 'room_id': eid_value(room.Id),
                          'name': room.Name, 'number': room.Number,
                          'area_m2': round(area_ft2 * 0.0929, 2),
                          'enclosed': area_ft2 > 0}
                if area_ft2 <= 0:
                    result['warning'] = ('Room was placed but is not enclosed — no bounding walls '
                                         'around ({}, {}) on level "{}".'.format(
                                             arguments.get('x'), arguments.get('y'), target_level.Name))
                return result
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── create_structural_framing_system ─────────────────────────────────
        elif tool_name == 'create_structural_framing_system':
            from Autodesk.Revit.DB import FamilySymbol, XYZ, Line, Transaction
            from Autodesk.Revit.DB import Structure
            x_bays    = arguments.get('x_bays', [])
            y_bays    = arguments.get('y_bays', [])
            lvl_name  = arguments.get('level_name')
            beam_type = arguments.get('beam_type')
            ox = float(arguments.get('origin_x', 0)) * 3.28084
            oy = float(arguments.get('origin_y', 0)) * 3.28084

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break

            framing_syms = []
            for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
                cat = s.Category.Name if s.Category else ''
                if cat == 'Structural Framing':
                    framing_syms.append(s)
            sym = None
            if beam_type:
                for s in framing_syms:
                    if s.Name == beam_type or (s.Family and s.Family.Name == beam_type):
                        sym = s
                        break
                if sym is None:
                    for s in framing_syms:
                        full = '{}:{}'.format(s.Family.Name, s.Name) if s.Family else s.Name
                        if beam_type.lower() in full.lower():
                            sym = s
                            break
                if sym is None:
                    return {'error': 'Structural framing type "{}" not found. Available: {}'.format(
                        beam_type, ', '.join(sorted(set(s.Name for s in framing_syms))[:20]) or '(none)')}
            elif framing_syms:
                sym = framing_syms[0]
            if sym is None:
                return {'error': 'No structural framing family type found. Load a beam family first.'}

            if target_level is None:
                return {'error': 'No levels found in the document'}
            # Place beams at the level's elevation, not absolute Z=0
            z_ft = target_level.Elevation

            x_positions = [ox]
            cur = ox
            for sp in x_bays:
                cur += sp * 3.28084
                x_positions.append(cur)
            y_positions = [oy]
            cur = oy
            for sp in y_bays:
                cur += sp * 3.28084
                y_positions.append(cur)

            beam_ids = []
            t = Transaction(doc, 'T3Lab AI Structural Framing System')
            t.Start()
            try:
                if not sym.IsActive:
                    sym.Activate()
                    doc.Regenerate()
                # Beams in X direction
                for y_pos in y_positions:
                    for i in range(len(x_positions) - 1):
                        p1 = XYZ(x_positions[i], y_pos, z_ft)
                        p2 = XYZ(x_positions[i + 1], y_pos, z_ft)
                        crv = Line.CreateBound(p1, p2)
                        inst = doc.Create.NewFamilyInstance(crv, sym, target_level, Structure.StructuralType.Beam)
                        beam_ids.append(eid_value(inst.Id))
                # Beams in Y direction
                for x_pos in x_positions:
                    for j in range(len(y_positions) - 1):
                        p1 = XYZ(x_pos, y_positions[j], z_ft)
                        p2 = XYZ(x_pos, y_positions[j + 1], z_ft)
                        crv = Line.CreateBound(p1, p2)
                        inst = doc.Create.NewFamilyInstance(crv, sym, target_level, Structure.StructuralType.Beam)
                        beam_ids.append(eid_value(inst.Id))
                t.Commit()
                return {'success': True, 'beams_created': len(beam_ids), 'beam_ids': beam_ids}
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── delete_element ───────────────────────────────────────────────────
        elif tool_name == 'delete_element':
            from Autodesk.Revit.DB import ElementId, Transaction
            ids = arguments.get('element_ids', [])
            if not ids:
                return {'error': 'No element_ids provided'}
            dry_run = bool(arguments.get('dry_run', False))

            def _describe(eid):
                el = doc.GetElement(eid)
                if el is None:
                    return {'id': eid_value(eid), 'name': '(unknown)', 'category': ''}
                return {
                    'id': eid_value(eid),
                    'name': el.Name if hasattr(el, 'Name') and el.Name else str(el.GetType().Name),
                    'category': el.Category.Name if el.Category else '',
                }

            # ── Preview mode: run the delete inside a transaction to collect the
            # full cascade (doc.Delete returns EVERY affected element id), then
            # roll back. After rollback the elements are restored, so we can
            # resolve their names/categories for a readable preview.
            if dry_run:
                requested = [_describe(ElementId(int(v))) for v in ids]
                affected_raw = set()
                t = Transaction(doc, 'T3Lab AI Delete Preview')
                t.Start()
                try:
                    for eid_val in ids:
                        try:
                            removed = doc.Delete(ElementId(int(eid_val)))
                            if removed:
                                for rid in removed:
                                    affected_raw.add(eid_value(rid))
                        except Exception:
                            pass
                finally:
                    t.RollBack()

                requested_ids = set(int(v) for v in ids)
                cascade_ids = sorted(i for i in affected_raw if i not in requested_ids)
                cascade = [_describe(ElementId(i)) for i in cascade_ids]
                return {
                    'dry_run': True,
                    'requested': requested,
                    'requested_count': len(requested),
                    'cascade': cascade,
                    'cascade_count': len(cascade),
                    'total_affected': len(affected_raw),
                    'note': 'Nothing was deleted. Call again with dry_run=false to apply.',
                }

            t = Transaction(doc, 'T3Lab AI Delete Elements')
            t.Start()
            deleted = []
            failed  = []
            try:
                for eid_val in ids:
                    try:
                        doc.Delete(ElementId(int(eid_val)))
                        deleted.append(int(eid_val))
                    except Exception as ex:
                        failed.append({'id': int(eid_val), 'error': str(ex)})
                t.Commit()
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}
            return {'deleted': deleted, 'failed': failed, 'deleted_count': len(deleted)}

        # ── operate_element ──────────────────────────────────────────────────
        elif tool_name == 'operate_element':
            from Autodesk.Revit.DB import ElementId, Transaction
            op      = (arguments.get('operation') or '').lower()
            ids     = arguments.get('element_ids', [])
            view    = doc.ActiveView
            elem_ids = [ElementId(int(i)) for i in ids]

            if op == 'select':
                from System.Collections.Generic import List
                id_list = List[ElementId](elem_ids)
                uidoc.Selection.SetElementIds(id_list)
                return {'success': True, 'operation': 'select', 'count': len(elem_ids)}

            elif op in ('hide', 'isolate', 'unhide'):
                from System.Collections.Generic import List
                id_col = List[ElementId](elem_ids)
                t = Transaction(doc, 'T3Lab AI {} Elements'.format(op.title()))
                t.Start()
                try:
                    if op == 'hide':
                        view.HideElements(id_col)
                    elif op == 'unhide':
                        view.UnhideElements(id_col)
                    elif op == 'isolate':
                        view.IsolateElementsTemporary(id_col)
                    t.Commit()
                    return {'success': True, 'operation': op, 'count': len(elem_ids)}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}

            elif op == 'reset_color':
                from Autodesk.Revit.DB import OverrideGraphicSettings, Transaction
                t = Transaction(doc, 'T3Lab AI Reset Color')
                t.Start()
                try:
                    plain = OverrideGraphicSettings()
                    for eid_obj in elem_ids:
                        view.SetElementOverrides(eid_obj, plain)
                    t.Commit()
                    return {'success': True, 'operation': 'reset_color', 'count': len(elem_ids)}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}

            return {'error': 'Unknown operation: {}'.format(op)}

        # ── color_elements ───────────────────────────────────────────────────
        elif tool_name == 'color_elements':
            from Autodesk.Revit.DB import (Color, OverrideGraphicSettings, Transaction,
                                           FillPatternElement, ElementId)
            cat_arg   = arguments.get('category', 'Rooms')
            param_arg = arguments.get('parameter_name', 'Name')
            view      = doc.ActiveView

            CATEGORY_MAP = {
                'Walls': BuiltInCategory.OST_Walls,
                'Floors': BuiltInCategory.OST_Floors,
                'Rooms': BuiltInCategory.OST_Rooms,
                'Columns': BuiltInCategory.OST_Columns,
                'Beams': BuiltInCategory.OST_StructuralFraming,
            }
            bic = CATEGORY_MAP.get(cat_arg, BuiltInCategory.OST_Rooms)
            collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()

            # Find solid fill pattern
            solid_id = ElementId(-1)
            for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
                pat = fp.GetFillPattern()
                if pat and pat.IsSolidFill:
                    solid_id = fp.Id
                    break

            # Group elements by param value
            groups = {}
            for elem in collector:
                try:
                    p = elem.LookupParameter(param_arg)
                    val = (p.AsValueString() or p.AsString() or 'Unknown') if p else 'Unknown'
                    groups.setdefault(val, []).append(elem.Id)
                except Exception:
                    pass

            # Assign distinct colors
            COLORS = [
                (52, 152, 219), (46, 204, 113), (231, 76, 60),
                (155, 89, 182), (241, 196, 15), (26, 188, 156),
                (230, 126, 34), (149, 165, 166), (52, 73, 94), (127, 140, 141),
            ]
            t = Transaction(doc, 'T3Lab AI Color by Parameter')
            t.Start()
            try:
                for idx, (val, eids) in enumerate(groups.items()):
                    r, g, b = COLORS[idx % len(COLORS)]
                    rev_color = Color(r, g, b)
                    ogs = OverrideGraphicSettings()
                    ogs.SetProjectionLineColor(rev_color)
                    if solid_id != ElementId(-1):
                        try:
                            ogs.SetSurfaceForegroundPatternId(solid_id)
                            ogs.SetSurfaceForegroundPatternColor(rev_color)
                        except Exception:
                            pass
                    for eid_obj in eids:
                        try:
                            view.SetElementOverrides(eid_obj, ogs)
                        except Exception:
                            pass
                t.Commit()
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}
            return {
                'success': True,
                'category': cat_arg,
                'parameter': param_arg,
                'group_count': len(groups),
                'groups': {v: len(ids) for v, ids in groups.items()},
            }

        # ── tag_all_walls ────────────────────────────────────────────────────
        elif tool_name == 'tag_all_walls':
            from Autodesk.Revit.DB import IndependentTag, TagMode, TagOrientation
            from Autodesk.Revit.DB import Wall, Reference
            leader   = bool(arguments.get('leader', False))
            view     = doc.ActiveView
            walls    = FilteredElementCollector(doc, view.Id).OfClass(Wall).ToElements()

            from Autodesk.Revit.DB import UV, XYZ
            t = Transaction(doc, 'T3Lab AI Tag All Walls')
            t.Start()
            tagged = 0
            try:
                for wall in walls:
                    try:
                        loc   = wall.Location
                        mid_pt = loc.Curve.Evaluate(0.5, True)
                        ref   = Reference(wall)
                        uv    = UV(mid_pt.X, mid_pt.Y)
                        tag   = IndependentTag.Create(doc, view.Id, ref, leader, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, mid_pt)
                        tagged += 1
                    except Exception:
                        pass
                t.Commit()
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}
            return {'success': True, 'tagged_count': tagged}

        # ── tag_all_rooms ────────────────────────────────────────────────────
        elif tool_name == 'tag_all_rooms':
            from Autodesk.Revit.DB import IndependentTag, TagMode, TagOrientation, Reference
            from Autodesk.Revit.DB.Architecture import Room
            view  = doc.ActiveView
            rooms = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

            t = Transaction(doc, 'T3Lab AI Tag All Rooms')
            t.Start()
            tagged = 0
            try:
                for room in rooms:
                    try:
                        loc = room.Location
                        if loc is None:
                            continue
                        pt  = loc.Point
                        ref = Reference(room)
                        IndependentTag.Create(doc, view.Id, ref, False, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, pt)
                        tagged += 1
                    except Exception:
                        pass
                t.Commit()
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}
            return {'success': True, 'tagged_count': tagged}

        # ── export_room_data ─────────────────────────────────────────────────
        elif tool_name == 'export_room_data':
            lvl_filter = arguments.get('level_name')
            rooms_out  = []
            for room in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType():
                try:
                    lvl_name = room.Level.Name if room.Level else ''
                    if lvl_filter and lvl_name != lvl_filter:
                        continue
                    area_ft2 = 0.0
                    try:
                        ap = room.LookupParameter('Area')
                        if ap:
                            area_ft2 = ap.AsDouble()
                    except Exception:
                        pass
                    dept = ''
                    try:
                        dp = room.LookupParameter('Department')
                        if dp:
                            dept = dp.AsString() or ''
                    except Exception:
                        pass
                    rooms_out.append({
                        'id': eid_value(room.Id),
                        'number': room.Number,
                        'name': room.Name,
                        'level': lvl_name,
                        'area_m2': round(area_ft2 * 0.0929, 2),
                        'department': dept,
                    })
                except Exception:
                    pass
            return {'count': len(rooms_out), 'rooms': rooms_out}

        # ── store_project_data ───────────────────────────────────────────────
        elif tool_name == 'store_project_data':
            import os, json as _json
            info = doc.ProjectInformation
            data = {
                'name': info.Name,
                'number': info.Number,
                'client': info.ClientName,
                'address': info.Address,
                'status': info.Status,
                'doc_path': doc.PathName,
            }
            out_dir  = os.path.join(os.path.dirname(doc.PathName) if doc.PathName else os.path.expanduser('~'), 'T3Lab_AI_Data')
            try:
                os.makedirs(out_dir)
            except OSError:
                pass
            out_path = os.path.join(out_dir, 'project_data.json')
            with open(out_path, 'w') as f:
                _json.dump(data, f, indent=2)
            return {'success': True, 'file': out_path, 'data': data}

        # ── store_room_data ──────────────────────────────────────────────────
        elif tool_name == 'store_room_data':
            import os, json as _json
            rooms_out = []
            for room in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType():
                try:
                    area_ft2 = 0.0
                    try:
                        ap = room.LookupParameter('Area')
                        if ap:
                            area_ft2 = ap.AsDouble()
                    except Exception:
                        pass
                    rooms_out.append({
                        'id': eid_value(room.Id),
                        'number': room.Number,
                        'name': room.Name,
                        'level': room.Level.Name if room.Level else '',
                        'area_m2': round(area_ft2 * 0.0929, 2),
                    })
                except Exception:
                    pass
            out_dir  = os.path.join(os.path.dirname(doc.PathName) if doc.PathName else os.path.expanduser('~'), 'T3Lab_AI_Data')
            try:
                os.makedirs(out_dir)
            except OSError:
                pass
            out_path = os.path.join(out_dir, 'room_data.json')
            with open(out_path, 'w') as f:
                _json.dump({'rooms': rooms_out}, f, indent=2)
            return {'success': True, 'file': out_path, 'room_count': len(rooms_out)}

        # ── query_stored_data ────────────────────────────────────────────────
        elif tool_name == 'query_stored_data':
            import os, json as _json
            data_type = arguments.get('data_type', 'project')
            out_dir   = os.path.join(os.path.dirname(doc.PathName) if doc.PathName else os.path.expanduser('~'), 'T3Lab_AI_Data')
            fname     = 'project_data.json' if data_type == 'project' else 'room_data.json'
            fpath     = os.path.join(out_dir, fname)
            if not os.path.isfile(fpath):
                return {'error': 'No stored data found. Run store_project_data or store_room_data first.', 'path': fpath}
            with open(fpath, 'r') as f:
                return _json.load(f)

        # ── send_code_to_revit ───────────────────────────────────────────────
        elif tool_name == 'send_code_to_revit':
            code = arguments.get('code', '')
            if not code:
                return {'error': 'No code provided'}
            # Execute directly in this ExternalEvent context (we are already on
            # the Revit main thread here). Also write result to result.json so
            # file-based clients can read it.
            local_ctx = {
                'doc':    doc,
                'uidoc':  uidoc,
                'app':    doc.Application,
                'result': None,
                'output': [],
            }
            # Redirect stdout for the duration of the exec: under pyRevit the
            # FIRST print() pops the output window as a dialog over Revit —
            # LLM-generated code prints habitually, so without this every
            # code-running chat turn threw a raw-text window in the user's
            # face. Captured text is returned as part of the tool result.
            import sys as _sys
            try:
                from StringIO import StringIO as _StringIO      # IronPython 2.7
            except ImportError:
                from io import StringIO as _StringIO            # CPython (tests)
            _old_stdout = _sys.stdout
            _buf = _StringIO()
            try:
                _sys.stdout = _buf
                try:
                    exec(code, local_ctx)   # noqa: S102
                finally:
                    _sys.stdout = _old_stdout
                printed      = _buf.getvalue().strip()
                result_val   = local_ctx.get('result')
                output_lines = local_ctx.get('output', [])
                parts = []
                if printed:
                    parts.append(printed)
                if result_val is not None:
                    parts.append(str(result_val))
                elif output_lines:
                    parts.append('\n'.join(str(x) for x in output_lines))
                out_str = '\n'.join(parts) if parts else 'OK'
                if len(out_str) > 8000:
                    out_str = out_str[:8000] + '... [truncated]'
                # Mirror to file-watcher result files for cross-channel clients
                try:
                    from core.file_watcher import RESULT_FILE, RESULT_TXT
                    import json as _json, time as _time
                    _res = {'task_id': 'mcp', 'status': 'success', 'output': out_str,
                            'error': '', 'timestamp': _time.time()}
                    with open(RESULT_FILE, 'w') as _f:
                        _json.dump(_res, _f, indent=2)
                    with open(RESULT_TXT, 'w') as _f:
                        _f.write(out_str)
                except Exception:
                    pass
                return {'success': True, 'result': out_str}
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                err_str = '{}\n{}'.format(e, tb)
                try:
                    from core.file_watcher import RESULT_FILE, RESULT_TXT
                    import json as _json, time as _time
                    _res = {'task_id': 'mcp', 'status': 'error', 'output': '',
                            'error': err_str, 'timestamp': _time.time()}
                    with open(RESULT_FILE, 'w') as _f:
                        _json.dump(_res, _f, indent=2)
                    with open(RESULT_TXT, 'w') as _f:
                        _f.write('ERROR: {}'.format(err_str))
                except Exception:
                    pass
                return {'success': False, 'error': err_str}

        # ── say_hello ────────────────────────────────────────────────────────
        elif tool_name == 'say_hello':
            msg = arguments.get('message', 'Hello from T3Lab AI!')
            try:
                from Autodesk.Revit.UI import TaskDialog
                TaskDialog.Show('T3Lab AI', msg)
                return {'success': True, 'message': msg}
            except Exception as e:
                return {'error': str(e)}

        # ── set_parameter ────────────────────────────────────────────────────
        elif tool_name == 'set_parameter':
            try:
                from Autodesk.Revit.DB import Transaction, StorageType
                eid   = int(arguments.get('element_id', 0))
                pname = arguments.get('parameter_name', '')
                value = arguments.get('value', '')
                elem  = doc.GetElement(ElementId(eid))
                if not elem:
                    return {'error': 'Element not found: {}'.format(eid)}
                param = elem.LookupParameter(pname)
                if not param:
                    return {'error': 'Parameter not found: {}'.format(pname)}
                if param.IsReadOnly:
                    return {'error': 'Parameter is read-only: {}'.format(pname)}
                t = Transaction(doc, 'T3Lab AI Set Parameter')
                t.Start()
                try:
                    st = param.StorageType
                    if st == StorageType.String:
                        param.Set(value)
                    elif st == StorageType.Double:
                        # Try display-unit parsing first ("3000" mm, "3.5 m",
                        # etc. — matches what the user sees in Revit), then
                        # fall back to a raw internal-unit float.
                        parsed = False
                        try:
                            parsed = param.SetValueString(value)
                        except Exception:
                            parsed = False
                        if not parsed:
                            param.Set(float(value))
                    elif st == StorageType.Integer:
                        param.Set(int(float(value)))
                    elif st == StorageType.ElementId:
                        param.Set(ElementId(int(value)))
                    t.Commit()
                    return {'success': True, 'element_id': eid, 'parameter': pname, 'value': value}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── get_all_parameters ───────────────────────────────────────────────
        elif tool_name == 'get_all_parameters':
            try:
                from Autodesk.Revit.DB import StorageType
                eid  = int(arguments.get('element_id', 0))
                elem = doc.GetElement(ElementId(eid))
                if not elem:
                    return {'error': 'Element not found: {}'.format(eid)}
                params = []
                for p in elem.Parameters:
                    try:
                        st = p.StorageType
                        if st == StorageType.String:
                            val = p.AsString() or ''
                        elif st == StorageType.Double:
                            val = p.AsDouble()
                        elif st == StorageType.Integer:
                            val = p.AsInteger()
                        elif st == StorageType.ElementId:
                            val = eid_value(p.AsElementId())
                        else:
                            val = None
                        params.append({
                            'name': p.Definition.Name,
                            'value': val,
                            'storage_type': str(st),
                            'read_only': p.IsReadOnly
                        })
                    except Exception:
                        pass
                params.sort(key=lambda x: x['name'])
                return {'element_id': eid, 'parameters': params, 'count': len(params)}
            except Exception as e:
                return {'error': str(e)}

        # ── move_elements ────────────────────────────────────────────────────
        elif tool_name == 'move_elements':
            try:
                from Autodesk.Revit.DB import (Transaction, ElementTransformUtils,
                                               XYZ)
                import System.Collections.Generic as SCG
                ft = 3.28084
                dx = float(arguments.get('dx', 0)) * ft
                dy = float(arguments.get('dy', 0)) * ft
                dz = float(arguments.get('dz', 0)) * ft
                ids_raw = arguments.get('element_ids', [])
                id_list = SCG.List[ElementId]([ElementId(int(i)) for i in ids_raw])
                t = Transaction(doc, 'T3Lab AI Move Elements')
                t.Start()
                try:
                    ElementTransformUtils.MoveElements(doc, id_list, XYZ(dx, dy, dz))
                    t.Commit()
                    return {'success': True, 'moved': len(ids_raw), 'delta_m': {'dx': arguments.get('dx'), 'dy': arguments.get('dy'), 'dz': arguments.get('dz', 0)}}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── copy_elements ────────────────────────────────────────────────────
        elif tool_name == 'copy_elements':
            try:
                from Autodesk.Revit.DB import (Transaction, ElementTransformUtils, XYZ)
                import System.Collections.Generic as SCG
                ft = 3.28084
                dx = float(arguments.get('dx', 0)) * ft
                dy = float(arguments.get('dy', 0)) * ft
                dz = float(arguments.get('dz', 0)) * ft
                ids_raw = arguments.get('element_ids', [])
                id_list = SCG.List[ElementId]([ElementId(int(i)) for i in ids_raw])
                t = Transaction(doc, 'T3Lab AI Copy Elements')
                t.Start()
                try:
                    new_ids = ElementTransformUtils.CopyElements(doc, id_list, XYZ(dx, dy, dz))
                    t.Commit()
                    return {'success': True, 'copied': len(ids_raw),
                            'new_element_ids': [eid_value(i) for i in new_ids]}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── rotate_element ───────────────────────────────────────────────────
        elif tool_name == 'rotate_element':
            try:
                from Autodesk.Revit.DB import (Transaction, ElementTransformUtils,
                                               XYZ, Line)
                import System.Collections.Generic as SCG
                import math
                ft = 3.28084
                angle_rad = float(arguments.get('angle_degrees', 0)) * math.pi / 180.0
                ox = float(arguments.get('origin_x', 0)) * ft
                oy = float(arguments.get('origin_y', 0)) * ft
                axis = Line.CreateBound(XYZ(ox, oy, 0), XYZ(ox, oy, 1))
                ids_raw = arguments.get('element_ids', [])
                id_list = SCG.List[ElementId]([ElementId(int(i)) for i in ids_raw])
                t = Transaction(doc, 'T3Lab AI Rotate Elements')
                t.Start()
                try:
                    ElementTransformUtils.RotateElements(doc, id_list, axis, angle_rad)
                    t.Commit()
                    return {'success': True, 'rotated': len(ids_raw), 'angle_degrees': arguments.get('angle_degrees')}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── get_element_bounding_box ─────────────────────────────────────────
        elif tool_name == 'get_element_bounding_box':
            try:
                eid  = int(arguments.get('element_id', 0))
                elem = doc.GetElement(ElementId(eid))
                if not elem:
                    return {'error': 'Element not found: {}'.format(eid)}
                bb = elem.get_BoundingBox(None)
                if not bb:
                    return {'error': 'Element has no bounding box'}
                ft_to_m = 0.3048
                return {
                    'element_id': eid,
                    'min': {'x': round(bb.Min.X * ft_to_m, 4), 'y': round(bb.Min.Y * ft_to_m, 4), 'z': round(bb.Min.Z * ft_to_m, 4)},
                    'max': {'x': round(bb.Max.X * ft_to_m, 4), 'y': round(bb.Max.Y * ft_to_m, 4), 'z': round(bb.Max.Z * ft_to_m, 4)}
                }
            except Exception as e:
                return {'error': str(e)}

        # ── create_view ──────────────────────────────────────────────────────
        elif tool_name == 'create_view':
            try:
                from Autodesk.Revit.DB import (Transaction, FilteredElementCollector,
                                               ViewFamilyType, ViewFamily,
                                               ViewPlan, View3D, Level)
                vtype = (arguments.get('view_type') or 'floor_plan').lower()
                name  = arguments.get('name')
                level_name = arguments.get('level_name')

                vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()

                def find_vft(family):
                    for v in vfts:
                        if v.ViewFamily == family:
                            return v
                    return None

                t = Transaction(doc, 'T3Lab AI Create View')
                t.Start()
                try:
                    if vtype in ('floor_plan', 'ceiling_plan'):
                        family = ViewFamily.FloorPlan if vtype == 'floor_plan' else ViewFamily.CeilingPlan
                        vft = find_vft(family)
                        if not vft:
                            t.RollBack()
                            return {'error': 'No ViewFamilyType for {}'.format(vtype)}
                        # Find level
                        lvl = None
                        if level_name:
                            levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                            for l in levels:
                                if l.Name == level_name:
                                    lvl = l
                                    break
                        if not lvl:
                            lvl = FilteredElementCollector(doc).OfClass(Level).FirstElement()
                        if not lvl:
                            t.RollBack()
                            return {'error': 'No level found'}
                        view = ViewPlan.Create(doc, vft.Id, lvl.Id)
                    else:
                        vft = find_vft(ViewFamily.ThreeDimensional)
                        if not vft:
                            t.RollBack()
                            return {'error': 'No 3D ViewFamilyType found'}
                        view = View3D.CreateIsometric(doc, vft.Id)

                    if name:
                        view.Name = name
                    t.Commit()
                    return {'success': True, 'view_id': eid_value(view.Id), 'view_name': view.Name, 'view_type': vtype}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── set_active_view ──────────────────────────────────────────────────
        elif tool_name == 'set_active_view':
            try:
                from Autodesk.Revit.DB import View
                view_name = arguments.get('view_name')
                view_id   = arguments.get('view_id')
                target_view = None
                if view_id:
                    target_view = doc.GetElement(ElementId(int(view_id)))
                elif view_name:
                    views = FilteredElementCollector(doc).OfClass(View).ToElements()
                    for v in views:
                        if v.Name == view_name:
                            target_view = v
                            break
                if not target_view:
                    return {'error': 'View not found'}
                try:
                    uidoc.ActiveView = target_view
                    return {'success': True, 'view_id': eid_value(target_view.Id), 'view_name': target_view.Name}
                except Exception as e:
                    return {'error': 'Cannot switch view: {}'.format(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── rename_element ───────────────────────────────────────────────────
        elif tool_name == 'rename_element':
            try:
                from Autodesk.Revit.DB import Transaction
                eid      = int(arguments.get('element_id', 0))
                new_name = arguments.get('new_name', '')
                elem     = doc.GetElement(ElementId(eid))
                if not elem:
                    return {'error': 'Element not found: {}'.format(eid)}
                t = Transaction(doc, 'T3Lab AI Rename Element')
                t.Start()
                try:
                    elem.Name = new_name
                    t.Commit()
                    return {'success': True, 'element_id': eid, 'new_name': new_name}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── create_sheet ─────────────────────────────────────────────────────
        elif tool_name == 'create_sheet':
            try:
                from Autodesk.Revit.DB import (Transaction, ViewSheet,
                                               FilteredElementCollector,
                                               FamilySymbol, BuiltInCategory)
                sheet_number = arguments.get('sheet_number', 'A-101')
                sheet_name   = arguments.get('sheet_name', 'New Sheet')
                tb_name      = arguments.get('title_block')

                # Find title block type
                tb_id = ElementId.InvalidElementId
                tb_types = (FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_TitleBlocks)
                            .WhereElementIsElementType()
                            .ToElements())
                if tb_types:
                    if tb_name:
                        for tb in tb_types:
                            if tb.Name == tb_name or (hasattr(tb, 'FamilyName') and tb.FamilyName == tb_name):
                                tb_id = tb.Id
                                break
                    if tb_id == ElementId.InvalidElementId:
                        tb_id = tb_types[0].Id

                t = Transaction(doc, 'T3Lab AI Create Sheet')
                t.Start()
                try:
                    sheet = ViewSheet.Create(doc, tb_id)
                    final_number = sheet_number
                    try:
                        sheet.SheetNumber = sheet_number
                    except Exception:
                        # Duplicate sheet number — uniquify instead of failing
                        final_number = None
                        for suffix in range(2, 100):
                            candidate = '{}-{}'.format(sheet_number, suffix)
                            try:
                                sheet.SheetNumber = candidate
                                final_number = candidate
                                break
                            except Exception:
                                continue
                        if final_number is None:
                            t.RollBack()
                            return {'error': 'Sheet number "{}" already exists and no free variant found'.format(sheet_number)}
                    sheet.Name = sheet_name
                    t.Commit()
                    result = {'success': True, 'sheet_id': eid_value(sheet.Id),
                              'sheet_number': final_number, 'sheet_name': sheet_name}
                    if final_number != sheet_number:
                        result['warning'] = 'Sheet number "{}" was taken — used "{}" instead'.format(sheet_number, final_number)
                    return result
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── add_view_to_sheet ────────────────────────────────────────────────
        elif tool_name == 'add_view_to_sheet':
            try:
                from Autodesk.Revit.DB import (Transaction, Viewport, XYZ)
                sheet_id = int(arguments.get('sheet_id', 0))
                view_id  = int(arguments.get('view_id', 0))
                mm_to_ft = 0.00328084
                x = float(arguments.get('x', 297)) * mm_to_ft
                y = float(arguments.get('y', 210)) * mm_to_ft
                sheet = doc.GetElement(ElementId(sheet_id))
                view  = doc.GetElement(ElementId(view_id))
                if not sheet:
                    return {'error': 'Sheet not found: {}'.format(sheet_id)}
                if not view:
                    return {'error': 'View not found: {}'.format(view_id)}
                try:
                    if not Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
                        return {'error': 'View "{}" cannot be placed on sheet "{}" — it is probably '
                                         'already placed on a sheet, or is a view type (schedule/legend) '
                                         'that needs a different placement method.'.format(view.Name, sheet.SheetNumber)}
                except Exception:
                    pass
                t = Transaction(doc, 'T3Lab AI Add View to Sheet')
                t.Start()
                try:
                    vp = Viewport.Create(doc, sheet.Id, view.Id, XYZ(x, y, 0))
                    t.Commit()
                    return {'success': True, 'viewport_id': eid_value(vp.Id),
                            'sheet_id': sheet_id, 'view_id': view_id}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── create_text_note ─────────────────────────────────────────────────
        elif tool_name == 'create_text_note':
            try:
                from Autodesk.Revit.DB import (Transaction, TextNote,
                                               TextNoteOptions, XYZ,
                                               FilteredElementCollector, TextNoteType)
                text      = arguments.get('text', '')
                if not text:
                    return {'error': 'No text provided'}
                ft = 3.28084
                x  = float(arguments.get('x', 0)) * ft
                y  = float(arguments.get('y', 0)) * ft
                type_name = arguments.get('text_type')
                font_size = arguments.get('font_size')

                active_view = doc.ActiveView
                if not active_view:
                    return {'error': 'No active view'}

                # Resolve text note type: named type > closest font size >
                # document default > first available.
                tn_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
                if not tn_types:
                    return {'error': 'No text note types found in the document'}
                tn_type = None
                if type_name:
                    for tt in tn_types:
                        if tt.Name == type_name:
                            tn_type = tt
                            break
                    if tn_type is None:
                        for tt in tn_types:
                            if type_name.lower() in tt.Name.lower():
                                tn_type = tt
                                break
                if tn_type is None and font_size:
                    # Pick the loaded type whose text height (mm) is closest
                    target_ft = float(font_size) / 304.8
                    best_diff = 1e30
                    for tt in tn_types:
                        try:
                            sp = tt.get_Parameter(
                                __import__('Autodesk.Revit.DB', fromlist=['BuiltInParameter']).BuiltInParameter.TEXT_SIZE)
                            if sp:
                                diff = abs(sp.AsDouble() - target_ft)
                                if diff < best_diff:
                                    best_diff = diff
                                    tn_type = tt
                        except Exception:
                            pass
                if tn_type is None:
                    try:
                        from Autodesk.Revit.DB import ElementTypeGroup
                        default_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
                        if default_id and default_id != ElementId.InvalidElementId:
                            tn_type = doc.GetElement(default_id)
                    except Exception:
                        pass
                if tn_type is None:
                    tn_type = tn_types[0]

                t = Transaction(doc, 'T3Lab AI Create Text Note')
                t.Start()
                try:
                    opts = TextNoteOptions(tn_type.Id)
                    note = TextNote.Create(doc, active_view.Id, XYZ(x, y, 0), text, opts)
                    t.Commit()
                    return {'success': True, 'text_note_id': eid_value(note.Id),
                            'text': text, 'type_used': tn_type.Name}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── get_model_warnings ───────────────────────────────────────────────
        elif tool_name == 'get_model_warnings':
            try:
                limit = int(arguments.get('limit', 50))
                warnings = doc.GetWarnings()
                result = []
                for w in list(warnings)[:limit]:
                    try:
                        failing_ids = [eid_value(i) for i in w.GetFailingElements()]
                        result.append({
                            'description': w.GetDescriptionText(),
                            'failing_element_ids': failing_ids
                        })
                    except Exception:
                        pass
                return {'warnings': result, 'total': len(list(warnings)), 'returned': len(result)}
            except Exception as e:
                return {'error': str(e)}

        # ── get_model_health ─────────────────────────────────────────────────
        elif tool_name == 'get_model_health':
            try:
                from Autodesk.Revit.DB import (FilteredElementCollector,
                                               RevitLinkInstance, Family)
                warning_count = len(list(doc.GetWarnings()))
                elem_count = FilteredElementCollector(doc).WhereElementIsNotElementType().GetElementCount()
                link_count = FilteredElementCollector(doc).OfClass(RevitLinkInstance).GetElementCount()
                family_count = FilteredElementCollector(doc).OfClass(Family).GetElementCount()
                return {
                    'warning_count': warning_count,
                    'element_count': elem_count,
                    'linked_files': link_count,
                    'loaded_families': family_count
                }
            except Exception as e:
                return {'error': str(e)}

        # ── list_worksets ────────────────────────────────────────────────────
        elif tool_name == 'list_worksets':
            try:
                from Autodesk.Revit.DB import FilteredWorksetCollector, WorksetKind
                if not doc.IsWorkshared:
                    return {'workshared': False, 'worksets': []}
                worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
                result = []
                for ws in worksets:
                    result.append({
                        'id': ws.Id.IntegerValue,
                        'name': ws.Name,
                        'is_open': ws.IsOpen,
                        'owner': ws.Owner or ''
                    })
                return {'workshared': True, 'worksets': result, 'count': len(result)}
            except Exception as e:
                return {'error': str(e)}

        # ── set_element_workset ──────────────────────────────────────────────
        elif tool_name == 'set_element_workset':
            try:
                from Autodesk.Revit.DB import (Transaction, FilteredWorksetCollector,
                                               WorksetKind, WorksetId)
                if not doc.IsWorkshared:
                    return {'error': 'Document is not workshared'}
                ws_name  = arguments.get('workset_name', '')
                ids_raw  = arguments.get('element_ids', [])
                # Find target workset
                worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
                target_ws = None
                for ws in worksets:
                    if ws.Name == ws_name:
                        target_ws = ws
                        break
                if not target_ws:
                    return {'error': 'Workset not found: {}'.format(ws_name)}
                t = Transaction(doc, 'T3Lab AI Set Workset')
                t.Start()
                try:
                    count = 0
                    for raw_id in ids_raw:
                        elem = doc.GetElement(ElementId(int(raw_id)))
                        if elem:
                            ws_param = elem.get_Parameter(
                                __import__('Autodesk.Revit.DB', fromlist=['BuiltInParameter']).BuiltInParameter.ELEM_PARTITION_PARAM
                            )
                            if ws_param and not ws_param.IsReadOnly:
                                ws_param.Set(target_ws.Id.IntegerValue)
                                count += 1
                    t.Commit()
                    return {'success': True, 'moved_count': count, 'workset': ws_name}
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── list_levels ──────────────────────────────────────────────────────
        elif tool_name == 'list_levels':
            try:
                from Autodesk.Revit.DB import Level
                levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                ft_to_m = 0.3048
                result = sorted(
                    [{'id': eid_value(l.Id), 'name': l.Name, 'elevation_m': round(l.Elevation * ft_to_m, 4)} for l in levels],
                    key=lambda x: x['elevation_m']
                )
                return {'levels': result, 'count': len(result)}
            except Exception as e:
                return {'error': str(e)}

        # ── load_family ──────────────────────────────────────────────────────
        elif tool_name == 'load_family':
            try:
                from Autodesk.Revit.DB import Transaction
                import os as _os
                file_path = arguments.get('file_path', '')
                if not _os.path.isfile(file_path):
                    return {'error': 'File not found: {}'.format(file_path)}
                t = Transaction(doc, 'T3Lab AI Load Family')
                t.Start()
                try:
                    success = doc.LoadFamily(file_path)
                    t.Commit()
                    result = {'success': bool(success), 'file_path': file_path}
                    if not success:
                        result['note'] = ('LoadFamily returned False — the family is probably '
                                          'already loaded in this project (Revit does not '
                                          'overwrite without IFamilyLoadOptions).')
                    return result
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e)}
            except Exception as e:
                return {'error': str(e)}

        # ── file_watcher_status ──────────────────────────────────────────────
        elif tool_name == 'file_watcher_status':
            try:
                from core.file_watcher import get_task_watcher
                watcher = get_task_watcher()
                return watcher.get_status()
            except Exception as e:
                return {'error': str(e), 'running': False}

        # ── get_revit_context ────────────────────────────────────────────────
        elif tool_name == 'get_revit_context':
            try:
                ft_to_m = 0.3048
                ctx = {
                    'document': doc.Title,
                    'is_workshared': doc.IsWorkshared,
                    'file_path': doc.PathName or '(unsaved)',
                }
                view = doc.ActiveView
                if view:
                    ctx['active_view'] = {
                        'name':      view.Name,
                        'type':      str(view.ViewType),
                        'id':        eid_value(view.Id),
                    }
                    try:
                        bb = view.get_BoundingBox(view)
                        if bb:
                            ctx['active_view']['view_range'] = {
                                'min': {'x': round(bb.Min.X * ft_to_m, 2),
                                        'y': round(bb.Min.Y * ft_to_m, 2)},
                                'max': {'x': round(bb.Max.X * ft_to_m, 2),
                                        'y': round(bb.Max.Y * ft_to_m, 2)},
                            }
                    except Exception:
                        pass
                try:
                    sel = uidoc.Selection.GetElementIds()
                    if sel and sel.Count > 0:
                        sel_info = []
                        for sid in list(sel)[:10]:
                            el = doc.GetElement(sid)
                            if el:
                                sel_info.append({
                                    'id':       eid_value(el.Id),
                                    'category': el.Category.Name if el.Category else '?',
                                    'name':     el.Name or '',
                                })
                        ctx['selection'] = {'count': sel.Count, 'elements': sel_info}
                    else:
                        ctx['selection'] = {'count': 0, 'elements': []}
                except Exception:
                    ctx['selection'] = {'count': 0, 'elements': []}
                return ctx
            except Exception as e:
                return {'error': str(e)}

        # ── show_assistant_pane ──────────────────────────────────────────────
        elif tool_name == 'show_assistant_pane':
            try:
                from System import Guid as SysGuid
                from Autodesk.Revit.UI import DockablePaneId
                from pyrevit import HOST_APP
                action  = arguments.get('action', 'show').lower()
                message = arguments.get('message', '')
                pane_guid = SysGuid('7F3A9B2E-C4D1-4E8F-A6B5-1234567890AB')
                pane_id   = DockablePaneId(pane_guid)
                uiapp     = HOST_APP.uiapp
                pane      = uiapp.GetDockablePane(pane_id)
                if pane is None:
                    return {'error': 'DockablePane not registered. Restart Revit after installing T3Lab.'}
                if action == 'hide':
                    pane.Hide()
                    return {'success': True, 'action': 'hide'}
                else:
                    pane.Show()
                    result = {'success': True, 'action': 'show'}
                # Inject message into pane if provided
                if message:
                    try:
                        from GUI.AssistantPaneControl import get_pane_controller
                        ctrl = get_pane_controller()
                        if ctrl:
                            ctrl.add_message(message, is_user=False)
                            result['message_injected'] = True
                    except Exception:
                        result['message_injected'] = False
                return result
            except Exception as e:
                return {'error': str(e)}

        # ── export_sheets_pdf ────────────────────────────────────────────────
        elif tool_name == 'export_sheets_pdf':
            try:
                from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet,
                                               PDFExportOptions, ViewSet,
                                               ExportPaperFormat, RasterQualityType,
                                               ExportColorType)
                import os as _os
                sheet_ids_raw  = arguments.get('sheet_ids', [])
                output_folder  = arguments.get('output_folder', '')
                combined       = bool(arguments.get('combined', False))

                if not output_folder:
                    doc_path = doc.PathName
                    output_folder = _os.path.dirname(doc_path) if doc_path else _os.path.expanduser('~')

                if not _os.path.isdir(output_folder):
                    try:
                        _os.makedirs(output_folder)
                    except Exception:
                        return {'error': 'Cannot create output folder: {}'.format(output_folder)}

                # Collect sheets
                all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
                if sheet_ids_raw:
                    export_ids = [int(i) for i in sheet_ids_raw]
                    sheets = [s for s in all_sheets if eid_value(s.Id) in export_ids]
                else:
                    sheets = list(all_sheets)

                if not sheets:
                    return {'error': 'No sheets to export'}

                opts = PDFExportOptions()
                opts.Combine = combined

                view_set = ViewSet()
                for s in sheets:
                    view_set.Insert(s)

                doc.Export(output_folder, 'T3Lab_Export', view_set, opts)
                return {'success': True, 'sheet_count': len(sheets), 'output_folder': output_folder, 'combined': combined}
            except Exception as e:
                return {'error': str(e)}

        # ── split_curve ──────────────────────────────────────────────────────
        elif tool_name == 'split_curve':
            from Autodesk.Revit.DB import (ElementId, Transaction, CurveElement,
                                           ModelCurve, DetailCurve)
            try:
                raw_id = arguments.get('element_id')
                if raw_id is None:
                    return {'error': 'element_id is required'}

                el = doc.GetElement(ElementId(int(raw_id)))
                if el is None:
                    return {'error': 'No element with id {}'.format(raw_id)}
                if not isinstance(el, CurveElement):
                    return {'error': ('Element {} is a {}, not a curve element. Only model curves '
                                      'and detail curves (line / arc / spline) can be split.'
                                      ).format(raw_id, el.GetType().Name)}

                is_model  = isinstance(el, ModelCurve)
                is_detail = isinstance(el, DetailCurve)
                if not (is_model or is_detail):
                    return {'error': ('Unsupported curve element type {} — only model curves and '
                                      'detail curves can be split.').format(el.GetType().Name)}

                curve = el.GeometryCurve
                if curve is None:
                    return {'error': 'Element has no geometry curve.'}
                if not curve.IsBound:
                    return {'error': ('Curve is periodic / unbound (e.g. a full circle or ellipse) '
                                      '— it has no endpoints to split at.')}

                # Build the list of normalized cut fractions (0..1 inclusive).
                ratios = arguments.get('split_at_ratios')
                if ratios:
                    interior = sorted(set(round(float(r), 9) for r in ratios if 0.0 < float(r) < 1.0))
                    if not interior:
                        return {'error': 'split_at_ratios must contain at least one value strictly between 0 and 1.'}
                    fracs = [0.0] + interior + [1.0]
                else:
                    n = int(arguments.get('segments', 2))
                    if n < 2:
                        return {'error': 'segments must be >= 2 (or provide split_at_ratios).'}
                    if n > 200:
                        return {'error': 'segments too large (max 200).'}
                    fracs = [float(i) / n for i in range(n + 1)]

                # Map the normalized fractions to RAW curve parameters (Revit maps
                # [0,1] linearly onto [GetEndParameter(0), GetEndParameter(1)]), then
                # slice the ORIGINAL curve with Clone()+MakeBound. This reuses the
                # source geometry sub-range verbatim, so an arc stays an arc and a
                # spline stays a spline — unlike rebuilding endpoints with
                # Line.CreateBound, which flattens every curve into a straight line.
                rp0 = curve.GetEndParameter(0)
                rp1 = curve.GetEndParameter(1)
                sub_curves = []
                for i in range(len(fracs) - 1):
                    a = rp0 + (rp1 - rp0) * fracs[i]
                    b = rp0 + (rp1 - rp0) * fracs[i + 1]
                    lo, hi = (a, b) if a < b else (b, a)
                    seg = curve.Clone()
                    seg.MakeBound(lo, hi)
                    sub_curves.append(seg)

                curve_kind = type(curve).__name__

                # Host context needed to recreate sibling curve elements.
                sketch_plane = el.SketchPlane if is_model else None
                owner_view = None
                if is_detail:
                    try:
                        owner_view = doc.GetElement(el.OwnerViewId)
                    except Exception:
                        owner_view = None
                try:
                    orig_style = el.LineStyle
                except Exception:
                    orig_style = None

                new_ids = []
                t = Transaction(doc, 'T3Lab AI Split Curve')
                t.Start()
                try:
                    # Reshape the original element onto the first segment, then add
                    # one new sibling per remaining segment (same style / host).
                    el.GeometryCurve = sub_curves[0]
                    for seg in sub_curves[1:]:
                        if is_model:
                            new_el = doc.Create.NewModelCurve(seg, sketch_plane)
                        else:
                            if owner_view is None:
                                t.RollBack()
                                return {'error': 'Cannot resolve owner view for the detail curve.'}
                            new_el = doc.Create.NewDetailCurve(owner_view, seg)
                        if orig_style is not None:
                            try:
                                new_el.LineStyle = orig_style
                            except Exception:
                                pass
                        new_ids.append(eid_value(new_el.Id))
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}

                return {
                    'success': True,
                    'curve_type': curve_kind,
                    'geometry_preserved': True,
                    'segment_count': len(sub_curves),
                    'original_element_id': eid_value(el.Id),
                    'new_element_ids': new_ids,
                }
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── split_element ────────────────────────────────────────────────────
        elif tool_name == 'split_element':
            from Autodesk.Revit.DB import (ElementId, Transaction, XYZ,
                                           LocationCurve, ElementTransformUtils)
            try:
                raw_id = arguments.get('element_id')
                if raw_id is None:
                    return {'error': 'element_id is required'}
                el = doc.GetElement(ElementId(int(raw_id)))
                if el is None:
                    return {'error': 'No element with id {}'.format(raw_id)}
                loc = el.Location
                if not isinstance(loc, LocationCurve):
                    return {'error': ('Element {} has no location curve (type {}). Only '
                                      'location-curve elements (wall, beam, pipe, duct, line) '
                                      'can be split; use split_curve for model/detail curves.'
                                      ).format(raw_id, el.GetType().Name)}
                curve = loc.Curve
                if not curve.IsBound:
                    return {'error': 'Location curve is unbound — cannot split.'}
                rp0 = curve.GetEndParameter(0)
                rp1 = curve.GetEndParameter(1)

                # Determine the raw split parameter from an XY point or a ratio.
                if arguments.get('x') is not None and arguments.get('y') is not None:
                    M2FT = 3.28084
                    pt = XYZ(float(arguments['x']) * M2FT, float(arguments['y']) * M2FT, 0)
                    try:
                        proj = curve.Project(pt)
                        split_raw = proj.Parameter
                    except Exception:
                        return {'error': 'Could not project the given point onto the curve.'}
                else:
                    ratio = float(arguments.get('at_ratio', 0.5))
                    if ratio <= 0.0 or ratio >= 1.0:
                        return {'error': 'at_ratio must be strictly between 0 and 1.'}
                    split_raw = rp0 + (rp1 - rp0) * ratio

                lo, hi = (rp0, rp1) if rp0 < rp1 else (rp1, rp0)
                if not (lo < split_raw < hi):
                    return {'error': 'Split position lies outside the curve span.'}

                first = curve.Clone();  first.MakeBound(min(rp0, split_raw), max(rp0, split_raw))
                second = curve.Clone(); second.MakeBound(min(split_raw, rp1), max(split_raw, rp1))

                t = Transaction(doc, 'T3Lab AI Split Element')
                t.Start()
                try:
                    copied = ElementTransformUtils.CopyElement(doc, el.Id, XYZ(0, 0, 0))
                    new_id = list(copied)[0] if copied else None
                    loc.Curve = first
                    if new_id is not None:
                        new_el = doc.GetElement(new_id)
                        new_el.Location.Curve = second
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {
                    'success': True,
                    'original_element_id': eid_value(el.Id),
                    'new_element_id': eid_value(new_id) if new_id is not None else None,
                    'category': el.Category.Name if el.Category else '',
                }
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── join_geometry ────────────────────────────────────────────────────
        elif tool_name == 'join_geometry':
            from Autodesk.Revit.DB import ElementId, Transaction, JoinGeometryUtils
            try:
                a = doc.GetElement(ElementId(int(arguments.get('element_id_a', 0))))
                b = doc.GetElement(ElementId(int(arguments.get('element_id_b', 0))))
                if a is None or b is None:
                    return {'error': 'Both element_id_a and element_id_b must resolve to elements.'}
                unjoin = bool(arguments.get('unjoin', False))
                t = Transaction(doc, 'T3Lab AI Join Geometry')
                t.Start()
                try:
                    joined = JoinGeometryUtils.AreElementsJoined(doc, a, b)
                    if unjoin:
                        if joined:
                            JoinGeometryUtils.UnjoinGeometry(doc, a, b)
                        action = 'unjoined'
                    else:
                        if not joined:
                            JoinGeometryUtils.JoinGeometry(doc, a, b)
                        action = 'joined'
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'action': action,
                        'element_id_a': eid_value(a.Id), 'element_id_b': eid_value(b.Id)}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── bulk_set_parameter ───────────────────────────────────────────────
        elif tool_name == 'bulk_set_parameter':
            from Autodesk.Revit.DB import Transaction, ElementId
            try:
                pname  = arguments.get('parameter_name')
                value  = arguments.get('value')
                if not pname or value is None:
                    return {'error': 'parameter_name and value are required.'}
                value  = str(value)
                fparam = arguments.get('filter_parameter')
                fval   = (arguments.get('filter_value') or '').lower()
                limit  = int(arguments.get('limit', 500))
                ids    = arguments.get('element_ids')

                if ids:
                    elements = [doc.GetElement(ElementId(int(i))) for i in ids]
                    elements = [e for e in elements if e is not None]
                else:
                    cat = arguments.get('category')
                    bic = self._bic_map().get(cat) if cat else None
                    if cat and bic is None:
                        return {'error': 'Unknown category "{}". Known: {}'.format(
                            cat, ', '.join(sorted(self._bic_map().keys())))}
                    coll = FilteredElementCollector(doc).WhereElementIsNotElementType()
                    if bic is not None:
                        coll = coll.OfCategory(bic)
                    elements = list(coll)

                modified, skipped, errors = 0, 0, 0
                t = Transaction(doc, 'T3Lab AI Bulk Set Parameter')
                t.Start()
                try:
                    for elem in elements:
                        if modified >= limit:
                            break
                        try:
                            if fparam:
                                fp = elem.LookupParameter(fparam)
                                pv = ''
                                if fp:
                                    pv = (fp.AsValueString() or fp.AsString() or '')
                                if fval and fval not in pv.lower():
                                    continue
                            p = elem.LookupParameter(pname)
                            ok, _err = self._apply_param_value(p, value)
                            if ok:
                                modified += 1
                            elif p is None:
                                skipped += 1
                            else:
                                errors += 1
                        except Exception:
                            errors += 1
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'parameter': pname, 'value': value,
                        'modified': modified, 'skipped_no_param': skipped, 'errors': errors}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── select_elements ──────────────────────────────────────────────────
        elif tool_name == 'select_elements':
            from Autodesk.Revit.DB import ElementId
            from System.Collections.Generic import List as NetList
            try:
                limit = int(arguments.get('limit', 500))
                ids   = arguments.get('element_ids')
                if ids:
                    target = [ElementId(int(i)) for i in ids][:limit]
                else:
                    cat = arguments.get('category')
                    bic = self._bic_map().get(cat) if cat else None
                    if cat and bic is None:
                        return {'error': 'Unknown category "{}".'.format(cat)}
                    coll = FilteredElementCollector(doc).WhereElementIsNotElementType()
                    if bic is not None:
                        coll = coll.OfCategory(bic)
                    pname = arguments.get('parameter_name')
                    pval  = (arguments.get('parameter_value') or '').lower()
                    target = []
                    for elem in coll:
                        if len(target) >= limit:
                            break
                        if pname:
                            p = elem.LookupParameter(pname)
                            if not p:
                                continue
                            v = (p.AsValueString() or p.AsString() or '')
                            if pval and pval not in v.lower():
                                continue
                        target.append(elem.Id)

                if bool(arguments.get('add_to_selection', False)):
                    current = list(uidoc.Selection.GetElementIds())
                    seen = set(eid_value(i) for i in current)
                    for i in target:
                        if eid_value(i) not in seen:
                            current.append(i)
                    target = current
                net = NetList[ElementId]()
                for i in target:
                    net.Add(i)
                uidoc.Selection.SetElementIds(net)
                if bool(arguments.get('show', False)) and net.Count:
                    # Zoom the view onto the selection (element-link clicks).
                    try:
                        uidoc.ShowElements(net)
                    except Exception:
                        pass
                return {'success': True, 'selected_count': net.Count}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── tag_elements ─────────────────────────────────────────────────────
        elif tool_name == 'tag_elements':
            from Autodesk.Revit.DB import (Transaction, IndependentTag, TagMode,
                                           TagOrientation, Reference, LocationCurve,
                                           LocationPoint)
            try:
                cat = arguments.get('category')
                bic = self._bic_map().get(cat)
                if bic is None:
                    return {'error': 'Unknown category "{}".'.format(cat)}
                leader = bool(arguments.get('leader', False))
                view = doc.ActiveView
                elems = FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType().ToElements()
                t = Transaction(doc, 'T3Lab AI Tag Elements')
                t.Start()
                tagged, failed = 0, 0
                try:
                    for elem in elems:
                        try:
                            loc = elem.Location
                            if isinstance(loc, LocationPoint):
                                pt = loc.Point
                            elif isinstance(loc, LocationCurve):
                                pt = loc.Curve.Evaluate(0.5, True)
                            else:
                                failed += 1
                                continue
                            IndependentTag.Create(doc, view.Id, Reference(elem), leader,
                                                  TagMode.TM_ADDBY_CATEGORY,
                                                  TagOrientation.Horizontal, pt)
                            tagged += 1
                        except Exception:
                            failed += 1
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'category': cat, 'tagged_count': tagged, 'failed': failed}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── create_dimension ─────────────────────────────────────────────────
        elif tool_name == 'create_dimension':
            from Autodesk.Revit.DB import (Transaction, ElementId, Grid, Reference,
                                           ReferenceArray, Line, XYZ)
            try:
                ids = arguments.get('element_ids', [])
                if len(ids) < 2:
                    return {'error': 'element_ids must contain at least 2 grids/line elements.'}
                M2FT = 3.28084
                offset = float(arguments.get('offset', 1.0)) * M2FT
                view = doc.ActiveView

                grids = []
                for i in ids:
                    e = doc.GetElement(ElementId(int(i)))
                    if isinstance(e, Grid):
                        grids.append(e)
                if len(grids) < 2:
                    return {'error': 'Need at least 2 grids. Only grids are supported by create_dimension.'}

                refs = ReferenceArray()
                pts = []
                for g in grids:
                    refs.Append(Reference(g))
                    pts.append(g.Curve.GetEndPoint(0))

                c0 = grids[0].Curve
                gdir = (c0.GetEndPoint(1) - c0.GetEndPoint(0)).Normalize()
                perp = XYZ(-gdir.Y, gdir.X, 0)
                origin = pts[0] + gdir.Multiply(offset)
                vals = [(p - origin).DotProduct(perp) for p in pts]
                lo, hi = min(vals), max(vals)
                if abs(hi - lo) < 1e-6:
                    return {'error': 'Selected grids are coincident — nothing to dimension.'}
                start = origin + perp.Multiply(lo)
                end   = origin + perp.Multiply(hi)
                dim_line = Line.CreateBound(start, end)

                t = Transaction(doc, 'T3Lab AI Create Dimension')
                t.Start()
                try:
                    dim = doc.Create.NewDimension(view, dim_line, refs)
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'dimension_id': eid_value(dim.Id), 'grid_count': len(grids)}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── get_schedule_data ────────────────────────────────────────────────
        elif tool_name == 'get_schedule_data':
            from Autodesk.Revit.DB import ViewSchedule, ElementId, SectionType
            try:
                sched = None
                sid = arguments.get('schedule_id')
                sname = arguments.get('schedule_name')
                if sid is not None:
                    cand = doc.GetElement(ElementId(int(sid)))
                    if isinstance(cand, ViewSchedule):
                        sched = cand
                if sched is None:
                    all_sched = [s for s in FilteredElementCollector(doc).OfClass(ViewSchedule)
                                 if not s.IsTemplate]
                    if sname:
                        for s in all_sched:
                            if s.Name == sname:
                                sched = s; break
                        if sched is None:
                            for s in all_sched:
                                if sname.lower() in s.Name.lower():
                                    sched = s; break
                    elif all_sched:
                        sched = all_sched[0]
                if sched is None:
                    return {'error': 'No matching schedule found.'}

                defn = sched.Definition
                headers = []
                for i in range(defn.GetFieldCount()):
                    try:
                        headers.append(defn.GetField(i).GetName())
                    except Exception:
                        headers.append('Field{}'.format(i))

                limit = int(arguments.get('limit', 200))
                body = sched.GetTableData().GetSectionData(SectionType.Body)
                n_rows = body.NumberOfRows
                n_cols = body.NumberOfColumns
                rows = []
                for r in range(n_rows):
                    if len(rows) >= limit:
                        break
                    row = []
                    for c in range(n_cols):
                        try:
                            row.append(sched.GetCellText(SectionType.Body, r, c))
                        except Exception:
                            row.append('')
                    rows.append(row)
                return {'schedule': sched.Name, 'id': eid_value(sched.Id),
                        'headers': headers, 'row_count': len(rows), 'rows': rows}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── create_schedule ──────────────────────────────────────────────────
        elif tool_name == 'create_schedule':
            from Autodesk.Revit.DB import (Transaction, ViewSchedule, ElementId,
                                           SchedulableField)
            try:
                cat = arguments.get('category')
                bic = self._bic_map().get(cat)
                if bic is None:
                    return {'error': 'Unknown category "{}".'.format(cat)}
                fields = arguments.get('fields') or []
                name = arguments.get('name')

                t = Transaction(doc, 'T3Lab AI Create Schedule')
                t.Start()
                try:
                    sched = ViewSchedule.CreateSchedule(doc, ElementId(bic))
                    defn = sched.Definition
                    available = {}
                    for sf in defn.GetSchedulableFields():
                        try:
                            available[sf.GetName(doc)] = sf
                        except Exception:
                            pass
                    added = []
                    if fields:
                        for fname in fields:
                            sf = available.get(fname)
                            if sf is None:
                                for k, v in available.items():
                                    if fname.lower() in k.lower():
                                        sf = v; break
                            if sf is not None:
                                try:
                                    defn.AddField(sf)
                                    added.append(fname)
                                except Exception:
                                    pass
                    else:
                        # No fields specified — add a handful of common ones.
                        for k in ['Family and Type', 'Type', 'Level', 'Count', 'Comments', 'Mark']:
                            sf = available.get(k)
                            if sf is not None:
                                try:
                                    defn.AddField(sf); added.append(k)
                                except Exception:
                                    pass
                    if name:
                        try:
                            sched.Name = name
                        except Exception:
                            pass
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'schedule_id': eid_value(sched.Id),
                        'name': sched.Name, 'fields_added': added}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── duplicate_view ───────────────────────────────────────────────────
        elif tool_name == 'duplicate_view':
            from Autodesk.Revit.DB import Transaction, ElementId, View, ViewDuplicateOption
            try:
                view = doc.GetElement(ElementId(int(arguments.get('view_id', 0))))
                if not isinstance(view, View):
                    return {'error': 'view_id does not refer to a view.'}
                mode = (arguments.get('mode') or 'plain').lower()
                opt = ViewDuplicateOption.Duplicate
                if mode == 'with_detailing':
                    opt = ViewDuplicateOption.WithDetailing
                elif mode == 'dependent':
                    opt = ViewDuplicateOption.AsDependent
                if not view.CanViewBeDuplicated(opt):
                    return {'error': 'This view cannot be duplicated with mode "{}".'.format(mode)}
                name = arguments.get('name')
                t = Transaction(doc, 'T3Lab AI Duplicate View')
                t.Start()
                try:
                    new_id = view.Duplicate(opt)
                    if name:
                        try:
                            doc.GetElement(new_id).Name = name
                        except Exception:
                            pass
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                new_v = doc.GetElement(new_id)
                return {'success': True, 'new_view_id': eid_value(new_id), 'name': new_v.Name}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── apply_view_template ──────────────────────────────────────────────
        elif tool_name == 'apply_view_template':
            from Autodesk.Revit.DB import Transaction, ElementId, View
            try:
                tpl_name = arguments.get('template_name', '')
                view_ids = arguments.get('view_ids', [])
                if not view_ids:
                    return {'error': 'view_ids is required.'}
                templates = [v for v in FilteredElementCollector(doc).OfClass(View) if v.IsTemplate]
                tpl = None
                for v in templates:
                    if v.Name == tpl_name:
                        tpl = v; break
                if tpl is None:
                    for v in templates:
                        if tpl_name.lower() in v.Name.lower():
                            tpl = v; break
                if tpl is None:
                    return {'error': 'View template "{}" not found.'.format(tpl_name)}
                t = Transaction(doc, 'T3Lab AI Apply View Template')
                t.Start()
                applied, failed = 0, 0
                try:
                    for vid in view_ids:
                        try:
                            v = doc.GetElement(ElementId(int(vid)))
                            v.ViewTemplateId = tpl.Id
                            applied += 1
                        except Exception:
                            failed += 1
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'template': tpl.Name, 'applied': applied, 'failed': failed}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── create_view_filter ───────────────────────────────────────────────
        elif tool_name == 'create_view_filter':
            from Autodesk.Revit.DB import (Transaction, ElementId, View,
                                           ParameterFilterElement, ElementParameterFilter,
                                           ParameterFilterRuleFactory, OverrideGraphicSettings,
                                           Color)
            from System.Collections.Generic import List as NetList
            try:
                name = arguments.get('name')
                cats = arguments.get('categories', [])
                cat_ids = NetList[ElementId]()
                for c in cats:
                    bic = self._bic_map().get(c)
                    if bic is not None:
                        cat_ids.Add(ElementId(bic))
                if cat_ids.Count == 0:
                    return {'error': 'No valid categories resolved from {}.'.format(cats)}

                view = doc.GetElement(ElementId(int(arguments['view_id']))) if arguments.get('view_id') else doc.ActiveView

                # Optional single "contains" rule on a parameter.
                elem_filter = None
                pname = arguments.get('parameter_name')
                pval  = arguments.get('parameter_value')
                if pname and pval is not None:
                    # Resolve the parameter id from a sample element in one of the categories.
                    pid = None
                    for c in cats:
                        bic = self._bic_map().get(c)
                        if bic is None:
                            continue
                        for e in FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType():
                            p = e.LookupParameter(pname)
                            if p:
                                pid = p.Id; break
                        if pid:
                            break
                    if pid is not None:
                        rule = None
                        try:
                            rule = ParameterFilterRuleFactory.CreateContainsRule(pid, str(pval))
                        except Exception:
                            try:
                                rule = ParameterFilterRuleFactory.CreateContainsRule(pid, str(pval), False)
                            except Exception:
                                rule = None
                        if rule is not None:
                            elem_filter = ElementParameterFilter(rule)

                t = Transaction(doc, 'T3Lab AI Create View Filter')
                t.Start()
                try:
                    if elem_filter is not None:
                        pfe = ParameterFilterElement.Create(doc, name, cat_ids, elem_filter)
                    else:
                        pfe = ParameterFilterElement.Create(doc, name, cat_ids)
                    view.AddFilter(pfe.Id)
                    if bool(arguments.get('hide', False)):
                        view.SetFilterVisibility(pfe.Id, False)
                    color = arguments.get('color')
                    if color:
                        rgb = self._parse_color(color)
                        if rgb:
                            ogs = OverrideGraphicSettings()
                            col = Color(rgb[0], rgb[1], rgb[2])
                            ogs.SetProjectionLineColor(col)
                            ogs.SetSurfaceForegroundPatternColor(col)
                            view.SetFilterOverrides(pfe.Id, ogs)
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'filter_id': eid_value(pfe.Id), 'name': name,
                        'view': view.Name}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── place_views_on_sheets ────────────────────────────────────────────
        elif tool_name == 'place_views_on_sheets':
            from Autodesk.Revit.DB import (Transaction, ElementId, ViewSheet, Viewport,
                                           XYZ, FamilySymbol, BuiltInCategory)
            try:
                view_ids = arguments.get('view_ids', [])
                if not view_ids:
                    return {'error': 'view_ids is required.'}
                tb_name = arguments.get('title_block')
                existing_sheet_id = arguments.get('sheet_id')

                tb = None
                tbs = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).OfClass(FamilySymbol))
                if tb_name:
                    for s in tbs:
                        full = '{}:{}'.format(s.Family.Name, s.Name) if s.Family else s.Name
                        if s.Name == tb_name or full == tb_name or tb_name.lower() in full.lower():
                            tb = s; break
                if tb is None and tbs:
                    tb = tbs[0]

                results = []
                t = Transaction(doc, 'T3Lab AI Place Views On Sheets')
                t.Start()
                try:
                    if existing_sheet_id:
                        sheet = doc.GetElement(ElementId(int(existing_sheet_id)))
                        col, row = 0, 0
                        for vid in view_ids:
                            v_eid = ElementId(int(vid))
                            if Viewport.CanAddViewToSheet(doc, sheet.Id, v_eid):
                                center = XYZ(0.5 + col * 0.9, 0.9 - row * 0.7, 0)
                                vp = Viewport.Create(doc, sheet.Id, v_eid, center)
                                results.append({'sheet_id': eid_value(sheet.Id), 'viewport_id': eid_value(vp.Id)})
                                col += 1
                                if col > 1:
                                    col = 0; row += 1
                    else:
                        if tb is None:
                            t.RollBack()
                            return {'error': 'No title block available to create sheets.'}
                        if not tb.IsActive:
                            tb.Activate(); doc.Regenerate()
                        for vid in view_ids:
                            v_eid = ElementId(int(vid))
                            sheet = ViewSheet.Create(doc, tb.Id)
                            if Viewport.CanAddViewToSheet(doc, sheet.Id, v_eid):
                                vp = Viewport.Create(doc, sheet.Id, v_eid, XYZ(1.0, 0.7, 0))
                                results.append({'sheet_id': eid_value(sheet.Id),
                                                'sheet_number': sheet.SheetNumber,
                                                'viewport_id': eid_value(vp.Id)})
                            else:
                                results.append({'sheet_id': eid_value(sheet.Id), 'error': 'view could not be placed'})
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'placed': len(results), 'results': results}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── export_dwg ───────────────────────────────────────────────────────
        elif tool_name == 'export_dwg':
            from Autodesk.Revit.DB import DWGExportOptions, ElementId
            from System.Collections.Generic import List as NetList
            import os as _os
            try:
                ids = arguments.get('sheet_ids') or arguments.get('view_ids') or []
                if not ids:
                    return {'error': 'Provide sheet_ids or view_ids to export.'}
                folder = arguments.get('output_folder', '')
                if not folder:
                    dp = doc.PathName
                    folder = _os.path.dirname(dp) if dp else _os.path.expanduser('~')
                if not _os.path.isdir(folder):
                    _os.makedirs(folder)
                id_list = NetList[ElementId]()
                for i in ids:
                    id_list.Add(ElementId(int(i)))
                opts = DWGExportOptions()
                ok = doc.Export(folder, 'T3Lab_Export', id_list, opts)
                return {'success': bool(ok), 'count': id_list.Count, 'output_folder': folder}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── export_image ─────────────────────────────────────────────────────
        elif tool_name == 'export_image':
            from Autodesk.Revit.DB import (ImageExportOptions, ExportRange, ImageFileType,
                                           ImageResolution, ElementId)
            from System.Collections.Generic import List as NetList
            import os as _os
            try:
                view = doc.GetElement(ElementId(int(arguments['view_id']))) if arguments.get('view_id') else doc.ActiveView
                folder = arguments.get('output_folder', '')
                if not folder:
                    dp = doc.PathName
                    folder = _os.path.dirname(dp) if dp else _os.path.expanduser('~')
                if not _os.path.isdir(folder):
                    _os.makedirs(folder)
                safe = ''.join(ch for ch in view.Name if ch.isalnum() or ch in ' _-').strip() or 'view'
                base = _os.path.join(folder, 'T3Lab_' + safe)
                opts = ImageExportOptions()
                opts.FilePath = base
                opts.ExportRange = ExportRange.SetOfViews
                view_ids = NetList[ElementId]()
                view_ids.Add(view.Id)
                opts.SetViewsAndSheets(view_ids)
                opts.HLRandWFViewsFileType = ImageFileType.PNG
                opts.ShadowViewsFileType = ImageFileType.PNG
                opts.ImageResolution = ImageResolution.DPI_150
                try:
                    opts.PixelSize = int(arguments.get('width', 1600))
                except Exception:
                    pass
                # Revit appends " - <ViewType> - <ViewName>" to the base name,
                # so the exact output path isn't known up-front. Snapshot the
                # matching files before, export, and report what changed —
                # callers (assistant vision capture) need the real file path.
                prefix = 'T3Lab_' + safe
                before = {}
                try:
                    for fn in _os.listdir(folder):
                        if fn.startswith(prefix):
                            p = _os.path.join(folder, fn)
                            before[fn] = _os.path.getmtime(p)
                except Exception:
                    pass
                doc.ExportImage(opts)
                files = []
                try:
                    for fn in _os.listdir(folder):
                        if not fn.startswith(prefix):
                            continue
                        if not fn.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff')):
                            continue
                        p = _os.path.join(folder, fn)
                        if fn not in before or _os.path.getmtime(p) != before[fn]:
                            files.append(p)
                except Exception:
                    pass
                return {'success': True, 'view': view.Name,
                        'output_folder': folder, 'files': files}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── create_project_parameter ─────────────────────────────────────────
        elif tool_name == 'create_project_parameter':
            from Autodesk.Revit.DB import (Transaction, BuiltInParameterGroup,
                                           ExternalDefinitionCreationOptions)
            import os as _os
            try:
                name = arguments.get('name')
                cats = arguments.get('categories', [])
                if not name or not cats:
                    return {'error': 'name and categories are required.'}
                group_name = arguments.get('group', 'Data')
                instance = bool(arguments.get('instance', True))
                type_name = (arguments.get('type') or 'Text')

                app = doc.Application
                # Ensure a shared-parameter file exists.
                sp_path = app.SharedParametersFilename
                if not sp_path or not _os.path.isfile(sp_path):
                    data_dir = _os.path.join(_os.path.expanduser('~'), 'T3Lab_AI_Data')
                    if not _os.path.isdir(data_dir):
                        _os.makedirs(data_dir)
                    sp_path = _os.path.join(data_dir, 'T3Lab_SharedParameters.txt')
                    if not _os.path.isfile(sp_path):
                        open(sp_path, 'w').close()
                    app.SharedParametersFilename = sp_path
                def_file = app.OpenSharedParameterFile()
                if def_file is None:
                    return {'error': 'Could not open shared parameter file.'}

                grp = def_file.Groups.get_Item('T3Lab') or def_file.Groups.Create('T3Lab')

                # Resolve the data type across Revit versions (SpecTypeId vs ParameterType).
                spec = None
                try:
                    from Autodesk.Revit.DB import SpecTypeId
                    spec_map = {
                        'text': SpecTypeId.String.Text, 'integer': SpecTypeId.Int.Integer,
                        'number': SpecTypeId.Number, 'length': SpecTypeId.Length,
                        'area': SpecTypeId.Area, 'yesno': SpecTypeId.Boolean.YesNo,
                    }
                    spec = spec_map.get(type_name.lower(), SpecTypeId.String.Text)
                    ext_opts = ExternalDefinitionCreationOptions(name, spec)
                except Exception:
                    from Autodesk.Revit.DB import ParameterType
                    pt_map = {
                        'text': ParameterType.Text, 'integer': ParameterType.Integer,
                        'number': ParameterType.Number, 'length': ParameterType.Length,
                        'area': ParameterType.Area, 'yesno': ParameterType.YesNo,
                    }
                    ext_opts = ExternalDefinitionCreationOptions(name, pt_map.get(type_name.lower(), ParameterType.Text))

                ext_def = None
                for d in grp.Definitions:
                    if d.Name == name:
                        ext_def = d; break
                if ext_def is None:
                    ext_def = grp.Definitions.Create(ext_opts)

                cat_set = app.Create.NewCategorySet()
                for c in cats:
                    bic = self._bic_map().get(c)
                    if bic is None:
                        continue
                    try:
                        cat_set.Insert(doc.Settings.Categories.get_Item(bic))
                    except Exception:
                        pass
                if cat_set.IsEmpty:
                    return {'error': 'No valid categories resolved.'}

                binding = (app.Create.NewInstanceBinding(cat_set) if instance
                           else app.Create.NewTypeBinding(cat_set))

                t = Transaction(doc, 'T3Lab AI Create Project Parameter')
                t.Start()
                try:
                    ok = doc.ParameterBindings.Insert(ext_def, binding, BuiltInParameterGroup.PG_DATA)
                    if not ok:
                        ok = doc.ParameterBindings.ReInsert(ext_def, binding, BuiltInParameterGroup.PG_DATA)
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'parameter': name, 'binding': 'instance' if instance else 'type',
                        'categories': cats}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── room_to_floor ────────────────────────────────────────────────────
        elif tool_name == 'room_to_floor':
            from Autodesk.Revit.DB import (Transaction, ElementId, Floor, FloorType,
                                           SpatialElementBoundaryOptions, CurveLoop)
            from System.Collections.Generic import List as NetList
            try:
                ids = arguments.get('room_ids') or ([arguments['room_id']] if arguments.get('room_id') else [])
                if not ids:
                    return {'error': 'Provide room_id or room_ids.'}
                ftype_name = arguments.get('floor_type')
                ftypes = list(FilteredElementCollector(doc).OfClass(FloorType))
                ftype = None
                if ftype_name:
                    for f in ftypes:
                        if f.Name == ftype_name or ftype_name.lower() in f.Name.lower():
                            ftype = f; break
                if ftype is None and ftypes:
                    ftype = ftypes[0]
                if ftype is None:
                    return {'error': 'No floor types available.'}

                bopts = SpatialElementBoundaryOptions()
                created = []
                t = Transaction(doc, 'T3Lab AI Room To Floor')
                t.Start()
                try:
                    for rid in ids:
                        room = doc.GetElement(ElementId(int(rid)))
                        if room is None:
                            continue
                        loops = room.GetBoundarySegments(bopts)
                        if not loops or loops.Count == 0:
                            continue
                        loop = CurveLoop()
                        for seg in loops[0]:
                            loop.Append(seg.GetCurve())
                        level_id = room.LevelId
                        try:
                            profile = NetList[CurveLoop]()
                            profile.Add(loop)
                            fl = Floor.Create(doc, profile, ftype.Id, level_id)
                            created.append(eid_value(fl.Id))
                        except Exception:
                            # Older Revit fallback (NewFloor) — best effort.
                            try:
                                from Autodesk.Revit.DB import CurveArray
                                ca = CurveArray()
                                for seg in loops[0]:
                                    ca.Append(seg.GetCurve())
                                fl = doc.Create.NewFloor(ca, False)
                                created.append(eid_value(fl.Id))
                            except Exception:
                                pass
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'floors_created': len(created), 'floor_ids': created}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── purge_unused ─────────────────────────────────────────────────────
        elif tool_name == 'purge_unused':
            from Autodesk.Revit.DB import (Transaction, FamilySymbol, View,
                                           ElementId)
            try:
                dry_run = bool(arguments.get('dry_run', True))
                used_type_ids = set()
                from Autodesk.Revit.DB import FamilyInstance
                for fi in FilteredElementCollector(doc).OfClass(FamilyInstance):
                    try:
                        used_type_ids.add(eid_value(fi.GetTypeId()))
                    except Exception:
                        pass
                unused_syms = [s for s in FilteredElementCollector(doc).OfClass(FamilySymbol)
                               if eid_value(s.Id) not in used_type_ids]

                views = list(FilteredElementCollector(doc).OfClass(View))
                used_tpl = set()
                for v in views:
                    try:
                        if not v.IsTemplate and eid_value(v.ViewTemplateId) != -1:
                            used_tpl.add(eid_value(v.ViewTemplateId))
                    except Exception:
                        pass
                unused_tpl = [v for v in views if v.IsTemplate and eid_value(v.Id) not in used_tpl]

                report = {
                    'unused_family_types': len(unused_syms),
                    'unused_family_type_names': [s.Name for s in unused_syms[:30]],
                    'unused_view_templates': len(unused_tpl),
                    'unused_view_template_names': [v.Name for v in unused_tpl[:30]],
                }
                if dry_run:
                    report['dry_run'] = True
                    report['note'] = 'Set dry_run=false to delete these items.'
                    return report

                deleted = 0
                t = Transaction(doc, 'T3Lab AI Purge Unused')
                t.Start()
                try:
                    for s in unused_syms:
                        try:
                            doc.Delete(s.Id); deleted += 1
                        except Exception:
                            pass
                    for v in unused_tpl:
                        try:
                            doc.Delete(v.Id); deleted += 1
                        except Exception:
                            pass
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                report['dry_run'] = False
                report['deleted'] = deleted
                return report
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── audit_model ──────────────────────────────────────────────────────
        elif tool_name == 'audit_model':
            from Autodesk.Revit.DB import (View, FamilyInstance, FamilySymbol,
                                           ImportInstance, Group, BuiltInCategory as _BIC)
            try:
                # Warnings grouped by description.
                warn_groups = {}
                try:
                    for w in doc.GetWarnings():
                        d = w.GetDescriptionText()
                        warn_groups[d] = warn_groups.get(d, 0) + 1
                except Exception:
                    pass
                top_warnings = sorted(warn_groups.items(), key=lambda kv: kv[1], reverse=True)[:15]

                imported_cad = FilteredElementCollector(doc).OfClass(ImportInstance).GetElementCount()
                model_groups = FilteredElementCollector(doc).OfCategory(_BIC.OST_IOSModelGroups).WhereElementIsNotElementType().GetElementCount()
                in_place = 0
                for fi in FilteredElementCollector(doc).OfClass(FamilyInstance):
                    try:
                        if fi.Symbol and fi.Symbol.Family and fi.Symbol.Family.IsInPlace:
                            in_place += 1
                    except Exception:
                        pass

                used_type_ids = set()
                for fi in FilteredElementCollector(doc).OfClass(FamilyInstance):
                    try:
                        used_type_ids.add(eid_value(fi.GetTypeId()))
                    except Exception:
                        pass
                unused_syms = sum(1 for s in FilteredElementCollector(doc).OfClass(FamilySymbol)
                                  if eid_value(s.Id) not in used_type_ids)

                views = list(FilteredElementCollector(doc).OfClass(View))
                default_named = sum(1 for v in views if not v.IsTemplate and (
                    v.Name.startswith('Copy of') or 'Copy 1' in v.Name))

                return {
                    'warnings_total': sum(warn_groups.values()),
                    'warnings_by_type': [{'description': d, 'count': c} for d, c in top_warnings],
                    'imported_cad_instances': imported_cad,
                    'model_groups': model_groups,
                    'in_place_families': in_place,
                    'unused_family_types': unused_syms,
                    'views_with_default_names': default_named,
                    'total_views': len(views),
                }
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        # ── create_workset ───────────────────────────────────────────────────
        elif tool_name == 'create_workset':
            from Autodesk.Revit.DB import Transaction, Workset, WorksetTable, FilteredWorksetCollector, WorksetKind
            try:
                if not doc.IsWorkshared:
                    return {'error': 'Document is not workshared — cannot create worksets.'}
                name = arguments.get('name')
                if not name:
                    return {'error': 'name is required.'}
                existing = [w.Name for w in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)]
                if name in existing:
                    return {'error': 'Workset "{}" already exists.'.format(name)}
                t = Transaction(doc, 'T3Lab AI Create Workset')
                t.Start()
                try:
                    ws = Workset.Create(doc, name)
                    t.Commit()
                except Exception as e:
                    t.RollBack()
                    return {'error': str(e), 'tool': tool_name}
                return {'success': True, 'workset': name, 'workset_id': eid_value(ws.Id)}
            except Exception as e:
                return {'error': str(e), 'tool': tool_name}

        return {'error': 'Tool not implemented'}

    def start_server(self):
        """Start the MCP server"""
        if self._is_running:
            return False

        # Check port and increment dynamically
        def is_port_in_use(port):
            import socket
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            try:
                s.bind(('127.0.0.1', port))
                s.close()
                return False
            except socket.error:
                return True

        port = 48884
        while is_port_in_use(port):
            port += 1
            if port > 48894:
                raise Exception("No free port available in range 48884-48894")
        self._port = port

        # Update pyRevit user config
        try:
            from pyrevit import userconfigs
            userconfigs.set_config_value("routes", "port", str(self._port))
        except Exception:
            pass

        # Initialize External Event for thread safety. NOTE: this only
        # succeeds when start_server() is itself called on Revit's main
        # thread (e.g. from an MCP Control dialog button). When the assistant
        # auto-starts the server from a background probe thread,
        # ExternalEvent.Create fails here and must instead be created up-front
        # via ensure_external_event() on the pushbutton UI thread.
        if HAS_REVIT_UI and not self._external_event:
            self.ensure_external_event()

        # Activate pyRevit routes server
        try:
            from pyrevit.routes.server import activate_server
            activate_server()
        except Exception:
            pass

        def run_server():
            try:
                self._http_server = HTTPServer(('localhost', self._port), MCPRequestHandler)
                self._http_server.mcp_server = self
                self._is_running = True
                self._http_server.serve_forever()
            except Exception as e:
                self._is_running = False
                raise e

        self._server_thread = threading.Thread(target=run_server)
        self._server_thread.daemon = True
        self._server_thread.start()

        # Wait a moment for server to start
        import time
        time.sleep(0.5)

        return self._is_running

    def stop_server(self):
        """Stop the MCP server"""
        if not self._is_running:
            return False

        self._is_running = False

        # Deactivate pyRevit routes server
        try:
            from pyrevit.routes.server import deactivate_server
            deactivate_server()
        except Exception:
            pass

        if self._http_server:
            self._http_server.shutdown()
            self._http_server.server_close()
            self._http_server = None

        if self._server_thread:
            self._server_thread.join(timeout=5)
            self._server_thread = None

        return True

    def get_server_stats(self):
        """Get server statistics"""
        return {
            'running': self._is_running,
            'port': self._port,
            'total_clients': self._total_clients,
            'commands_processed': self._commands_processed,
            'current_clients': len(self._clients),
            'tools_count': len(self._tools),
            'external_event_ready': self._external_event is not None,
        }

    def register_tool(self, name, description, input_schema, handler):
        """Register a custom tool"""
        self._tools[name] = {
            'name': name,
            'description': description,
            'inputSchema': input_schema
        }


def get_t3labai_server():
    """Get the singleton T3LabAI server instance"""
    return T3LabAIServer()
