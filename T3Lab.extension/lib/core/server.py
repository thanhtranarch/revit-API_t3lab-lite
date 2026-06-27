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
            elif method == 'notifications/initialized':
                response['result'] = {}
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


class T3LabAIServer:
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
        self._initialized = True

        # Register default Revit tools
        self._register_default_tools()

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
                'description': 'Delete one or more Revit elements by their IDs',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'element_ids': {
                            'type': 'array',
                            'items': {'type': 'integer'},
                            'description': 'List of element IDs to delete'
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
                'description': 'Execute IronPython code directly in the Revit context. Use with care — full Revit API access.',
                'inputSchema': {
                    'type': 'object',
                    'properties': {
                        'code': {
                            'type': 'string',
                            'description': 'IronPython 2.7 code to execute. Has access to doc, uidoc, app.'
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

    def _execute_tool(self, tool_name, arguments):
        """Execute a Revit tool in a thread-safe manner using External Events."""
        if self._external_event:
            self._event_handler.tool_name = tool_name
            self._event_handler.arguments = arguments
            self._event_handler._lock.clear()
            self._external_event.Raise()
            
            # Wait for main UI thread execution (10s timeout)
            success = self._event_handler._lock.wait(timeout=10)
            if not success:
                return {'error': 'Execution timed out waiting for Revit thread context', 'tool': tool_name}
            if self._event_handler.exception:
                return {'error': str(self._event_handler.exception), 'tool': tool_name}
            return self._event_handler.result
        else:
            # Fallback if external event is not available (e.g. running headlessly or test context)
            return self._execute_tool_in_context(tool_name, arguments)

    def _execute_tool_in_context(self, tool_name, arguments):
        """Execute a Revit tool directly (must be inside Revit context thread)"""
        try:
            from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet,
                                           BuiltInCategory, Level, ElementId,
                                           Transaction, ElementLevelFilter)
            from pyrevit import revit
            doc = revit.doc
            uidoc = revit.uidoc
        except ImportError:
            return {'error': 'Revit API not available', 'tool': tool_name}

        if tool_name == 'revit_get_active_view':
            view = doc.ActiveView
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
                    new_level.Name = level_name
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

            # Find wall type
            wall_types = list(FilteredElementCollector(doc).OfClass(WallType).ToElements())
            target_wall_type = None
            if wall_type_name_arg:
                for wt in wall_types:
                    if wt.Name == wall_type_name_arg:
                        target_wall_type = wt
                        break
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
            return {'view': doc.ActiveView.Name, 'count': len(elements_out), 'elements': elements_out}

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
            from Autodesk.Revit.DB import Wall, Floor, Roof
            cat_arg   = arguments.get('category', 'Walls')
            lvl_arg   = arguments.get('level_name')
            CAT_TO_CLASS = {
                'Walls': (Wall, 'Area', 'Volume'),
                'Floors': (Floor, 'Area', 'Volume'),
            }
            clz, area_p, vol_p = CAT_TO_CLASS.get(cat_arg, (Wall, 'Area', 'Volume'))
            collector = FilteredElementCollector(doc).OfClass(clz).WhereElementIsNotElementType()
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
                    if lvl_arg and lvl_name != lvl_arg:
                        continue
                    area_ft2 = 0.0
                    vol_ft3  = 0.0
                    try:
                        ap = elem.LookupParameter(area_p)
                        if ap:
                            area_ft2 = ap.AsDouble()
                    except Exception:
                        pass
                    try:
                        vp = elem.LookupParameter(vol_p)
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
            from Autodesk.Revit.DB import FamilySymbol, FamilyInstance, XYZ, Transaction
            from Autodesk.Revit.DB import Structure
            ftype_name = arguments.get('family_type', '')
            x_m = float(arguments.get('x', 0)) * 3.28084
            y_m = float(arguments.get('y', 0)) * 3.28084
            z_m = float(arguments.get('z', 0)) * 3.28084
            lvl_name = arguments.get('level_name')
            point = XYZ(x_m, y_m, z_m)

            sym = None
            for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
                if s.Name == ftype_name or (s.Family and '{}:{}'.format(s.Family.Name, s.Name) == ftype_name):
                    sym = s
                    break
            if sym is None:
                return {'error': 'Family type "{}" not found'.format(ftype_name)}

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break

            t = Transaction(doc, 'T3Lab AI Create Element')
            t.Start()
            try:
                if not sym.IsActive:
                    sym.Activate()
                    doc.Regenerate()
                inst = doc.Create.NewFamilyInstance(
                    point, sym, target_level,
                    Structure.StructuralType.NonStructural
                )
                t.Commit()
                return {'success': True, 'element_id': eid_value(inst.Id), 'type': ftype_name}
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── create_line_based_element ────────────────────────────────────────
        elif tool_name == 'create_line_based_element':
            from Autodesk.Revit.DB import FamilySymbol, XYZ, Line, Transaction
            from Autodesk.Revit.DB import Structure
            ftype_name = arguments.get('family_type', '')
            sx = float(arguments.get('start_x', 0)) * 3.28084
            sy = float(arguments.get('start_y', 0)) * 3.28084
            sz = float(arguments.get('start_z', 0)) * 3.28084
            ex = float(arguments.get('end_x', 0)) * 3.28084
            ey = float(arguments.get('end_y', 0)) * 3.28084
            ez = float(arguments.get('end_z', 0)) * 3.28084
            lvl_name = arguments.get('level_name')

            sym = None
            for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
                if s.Name == ftype_name or (s.Family and '{}:{}'.format(s.Family.Name, s.Name) == ftype_name):
                    sym = s
                    break
            if sym is None:
                return {'error': 'Family type "{}" not found'.format(ftype_name)}

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break

            curve = Line.CreateBound(XYZ(sx, sy, sz), XYZ(ex, ey, ez))
            t = Transaction(doc, 'T3Lab AI Create Line Element')
            t.Start()
            try:
                if not sym.IsActive:
                    sym.Activate()
                    doc.Regenerate()
                inst = doc.Create.NewFamilyInstance(
                    curve, sym, target_level,
                    Structure.StructuralType.Beam
                )
                t.Commit()
                return {'success': True, 'element_id': eid_value(inst.Id)}
            except Exception as e:
                t.RollBack()
                return {'error': str(e)}

        # ── create_surface_based_element ─────────────────────────────────────
        elif tool_name == 'create_surface_based_element':
            from Autodesk.Revit.DB import XYZ, Line, CurveArray, Transaction, Floor, FloorType
            elem_type = (arguments.get('element_type') or 'floor').lower()
            boundary  = arguments.get('boundary_points', [])
            lvl_name  = arguments.get('level_name')
            type_name = arguments.get('type_name')

            if len(boundary) < 3:
                return {'error': 'boundary_points must have at least 3 points'}

            levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
            target_level = levels[0] if levels else None
            if lvl_name:
                for lv in levels:
                    if lv.Name == lvl_name:
                        target_level = lv
                        break

            def m_to_ft(v): return v * 3.28084

            curves = CurveArray()
            pts = [[m_to_ft(p[0]), m_to_ft(p[1])] for p in boundary]
            for i in range(len(pts)):
                p1 = XYZ(pts[i][0], pts[i][1], 0)
                p2 = XYZ(pts[(i + 1) % len(pts)][0], pts[(i + 1) % len(pts)][1], 0)
                curves.Append(Line.CreateBound(p1, p2))

            t = Transaction(doc, 'T3Lab AI Create Surface Element')
            t.Start()
            try:
                floor_types = list(FilteredElementCollector(doc).OfClass(FloorType).ToElements())
                ft = floor_types[0] if floor_types else None
                if type_name:
                    for flt in floor_types:
                        if flt.Name == type_name:
                            ft = flt
                            break
                new_elem = doc.Create.NewFloor(curves, ft, target_level, False)
                t.Commit()
                return {'success': True, 'element_id': eid_value(new_elem.Id), 'type': elem_type}
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

            line_len_ft = 100.0
            grid_ids = []
            t = Transaction(doc, 'T3Lab AI Create Grid')
            t.Start()
            try:
                # Vertical lines (along Y) at X positions
                x_pos = ox
                for i, (spacing) in enumerate([0.0] + [s * 3.28084 for s in x_spacings]):
                    if i > 0:
                        x_pos += spacing
                    else:
                        x_pos = ox
                    start = XYZ(x_pos, oy - line_len_ft / 2, 0)
                    end   = XYZ(x_pos, oy + line_len_ft / 2, 0)
                    g = Grid.Create(doc, Line.CreateBound(start, end))
                    if i < len(x_labels):
                        g.Name = x_labels[i]
                    grid_ids.append(eid_value(g.Id))

                # Horizontal lines (along X) at Y positions
                y_pos = oy
                for i, (spacing) in enumerate([0.0] + [s * 3.28084 for s in y_spacings]):
                    if i > 0:
                        y_pos += spacing
                    else:
                        y_pos = oy
                    start = XYZ(ox - line_len_ft / 2, y_pos, 0)
                    end   = XYZ(ox + line_len_ft / 2, y_pos, 0)
                    g = Grid.Create(doc, Line.CreateBound(start, end))
                    if i < len(y_labels):
                        g.Name = y_labels[i]
                    grid_ids.append(eid_value(g.Id))

                t.Commit()
                return {'success': True, 'grid_count': len(grid_ids), 'grid_ids': grid_ids}
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
                t.Commit()
                return {'success': True, 'room_id': eid_value(room.Id), 'name': room.Name, 'number': room.Number}
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

            sym = None
            for s in FilteredElementCollector(doc).OfClass(FamilySymbol):
                cat = s.Category.Name if s.Category else ''
                if cat == 'Structural Framing':
                    if beam_type is None or s.Name == beam_type or (s.Family and s.Family.Name == beam_type):
                        sym = s
                        break
            if sym is None:
                return {'error': 'No structural framing family type found. Load a beam family first.'}

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
                        p1 = XYZ(x_positions[i], y_pos, 0)
                        p2 = XYZ(x_positions[i + 1], y_pos, 0)
                        crv = Line.CreateBound(p1, p2)
                        inst = doc.Create.NewFamilyInstance(crv, sym, target_level, Structure.StructuralType.Beam)
                        beam_ids.append(eid_value(inst.Id))
                # Beams in Y direction
                for x_pos in x_positions:
                    for j in range(len(y_positions) - 1):
                        p1 = XYZ(x_pos, y_positions[j], 0)
                        p2 = XYZ(x_pos, y_positions[j + 1], 0)
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
            local_ctx = {'doc': doc, 'uidoc': uidoc, 'app': doc.Application, 'result': None}
            try:
                exec(code, local_ctx)
                result = local_ctx.get('result')
                return {'success': True, 'result': str(result) if result is not None else 'OK'}
            except Exception as e:
                return {'success': False, 'error': str(e)}

        # ── say_hello ────────────────────────────────────────────────────────
        elif tool_name == 'say_hello':
            msg = arguments.get('message', 'Hello from T3Lab AI!')
            try:
                from Autodesk.Revit.UI import TaskDialog
                TaskDialog.Show('T3Lab AI', msg)
                return {'success': True, 'message': msg}
            except Exception as e:
                return {'error': str(e)}

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

        # Initialize External Event for thread safety
        if HAS_REVIT_UI and not self._external_event:
            try:
                self._event_handler = MCPExternalEventHandler(self)
                self._external_event = ExternalEvent.Create(self._event_handler)
            except Exception as e:
                pass

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
            'tools_count': len(self._tools)
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
