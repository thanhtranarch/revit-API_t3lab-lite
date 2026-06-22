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
            from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet
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
