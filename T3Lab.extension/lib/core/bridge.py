"""T3Lab MCP stdio bridge.

Forwards JSON-RPC requests from an MCP client (Claude Desktop / Claude Code)
to the T3Lab HTTP server running inside Revit.

Deployment: this file is copied by the extension to a machine-stable
location — %APPDATA%/T3LabAI/bridge.py — and the Claude config points at
that copy (see Services/mcp_service.py, MCPService.deploy_bridge). Users can
download, move or update the extension freely; the MCP config never breaks
because it never references a path inside the extension folder.

Attach-proof protocol handling: the MCP handshake no longer requires a live
Revit instance —

  - initialize / ping      -> answered locally by the bridge, so the client
                              always attaches even when Revit is closed.
  - tools/list             -> served live from Revit when reachable (and the
                              manifest is cached to
                              %APPDATA%/T3LabAI/tools_cache.json); when Revit
                              is down the cached manifest is served instead.
  - tools/call, Revit down -> structured error payload telling the model to
                              start Revit / the T3Lab server, instead of a
                              broken session.

Multi-window (multi-instance) support: each running Revit instance hosts its
own server on the first free port in 48884-48894. The bridge tracks a
"current" port and re-routes automatically:

  - list_open_documents   -> fans out to every alive instance and merges the
                             results, so documents in other Revit windows are
                             visible to the AI client.
  - switch_active_document-> if the requested document is open in another
                             Revit window, the bridge switches its current
                             port to that instance; every subsequent tool
                             call then targets that window.
  - any other call        -> forwarded to the current port; if that instance
                             is gone (connection refused, or it stopped
                             answering /health — hung zombie socket, foreign
                             server such as pyRevit Routes on the port), the
                             bridge fails over to another alive instance.

Bridge-level tools (served entirely by the bridge, injected into tools/list
on top of the in-Revit server's manifest):

  - list_revit_instances  -> probe the port range and report every running
                             Revit instance, its open documents, and which
                             instance the bridge currently targets.
  - switch_revit_instance -> explicitly retarget ALL subsequent tool calls at
                             another instance, by port or by a document open
                             in it (the by-document form delegates to the
                             cross-instance switch_active_document handling,
                             so the target window is activated too).
"""

import sys
import os
import json
import urllib.request
import urllib.error
from concurrent.futures import ThreadPoolExecutor

BRIDGE_VERSION = '2.2.0'

PORT_MIN, PORT_MAX = 48884, 48894
PROBE_TIMEOUT = 1.0    # /health probe per port (run in parallel)
TARGET_PROBE_TIMEOUT = 4.0  # explicit single-port checks get a fairer wait —
                            # a busy instance (single-threaded server mid-tool)
                            # can miss the fast range scan yet still be alive
CALL_TIMEOUT = 130     # matches the server's 120s ExternalEvent wait + margin
LIST_TIMEOUT = 15      # tools/list must answer fast or fall back to cache

# Protocol versions this bridge can echo back to the client. The in-Revit
# server is version-agnostic (tools/list + tools/call only), so accepting
# the client's requested revision here is safe.
SUPPORTED_PROTOCOLS = ('2025-06-18', '2025-03-26', '2024-11-05')
DEFAULT_PROTOCOL = '2024-11-05'

# Tools implemented by the bridge itself (multi-instance control lives here —
# a server inside one Revit process cannot see the other processes). Injected
# into every tools/list response on top of the server manifest; never cached.
BRIDGE_TOOLS = [
    {
        'name': 'list_revit_instances',
        'description': ('List running Revit instances. Each open Revit window is a separate '
                        'process hosting its own T3Lab server on a port in {}-{}; this reports '
                        'every alive instance, the documents open in it, and which instance '
                        'tool calls currently target. Use switch_revit_instance (by port) or '
                        'switch_active_document (by document) to retarget.'
                        ).format(PORT_MIN, PORT_MAX),
        'inputSchema': {'type': 'object', 'properties': {}, 'required': []},
    },
    {
        'name': 'switch_revit_instance',
        'description': ('Point ALL subsequent tool calls at another running Revit instance '
                        '(separate Revit window/process). Pass port from list_revit_instances, '
                        'or document (a title/file name open in the target instance) — the '
                        'by-document form also activates that document\'s window, like '
                        'switch_active_document.'),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'port': {
                    'type': 'integer',
                    'description': 'Instance port from list_revit_instances ({}-{})'.format(
                        PORT_MIN, PORT_MAX),
                },
                'document': {
                    'type': 'string',
                    'description': ('Document title, file name or .rvt path open in the '
                                    'target instance (alternative to port)'),
                },
            },
            'required': [],
        },
    },
]


def _data_dir():
    """%APPDATA%/T3LabAI — same folder as mcp_token.txt / mcp_paths.json."""
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def _read_token():
    """Read the shared-secret token T3LabAIServer writes on first run to
    %APPDATA%\\T3LabAI\\mcp_token.txt. All Revit instances on the machine
    share this file, so one token authenticates against every instance."""
    try:
        with open(os.path.join(_data_dir(), 'mcp_token.txt'), 'r') as f:
            return f.read().strip()
    except Exception:
        return ''


# ─── Tools-manifest cache (lets tools/list answer with Revit closed) ─────────

def _cache_file():
    return os.path.join(_data_dir(), 'tools_cache.json')


def _load_cached_tools():
    try:
        with open(_cache_file(), 'r', encoding='utf-8') as f:
            tools = json.load(f).get('tools')
        return tools if isinstance(tools, list) else []
    except Exception:
        return []


def _save_cached_tools(tools):
    try:
        os.makedirs(_data_dir(), exist_ok=True)
        tmp = _cache_file() + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump({'bridge_version': BRIDGE_VERSION, 'tools': tools}, f)
        os.replace(tmp, _cache_file())
    except Exception:
        pass


def _post_mcp(port, request, token, timeout=CALL_TIMEOUT):
    """Forward one JSON-RPC request dict to an instance; returns parsed reply."""
    # 127.0.0.1, NOT localhost: on Windows "localhost" resolves to ::1 first
    # and the IPv6 connect attempt takes ~2s to fail before falling back to
    # IPv4 — adding a flat 2s to EVERY tool call (the server binds IPv4 only).
    data = json.dumps(request).encode('utf-8')
    headers = {'Content-Type': 'application/json'}
    if token:
        headers['Authorization'] = 'Bearer ' + token
    req = urllib.request.Request(
        f"http://127.0.0.1:{port}/mcp", data=data, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode('utf-8'))


def _is_conn_refused(exc):
    if isinstance(exc, urllib.error.URLError):
        return isinstance(exc.reason, ConnectionRefusedError)
    return isinstance(exc, ConnectionRefusedError)


def _port_alive(port, timeout=PROBE_TIMEOUT):
    """True if a T3Lab server answers /health on this port."""
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/health")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status == 200
    except Exception:
        return False


def _alive_ports():
    """Probe the whole port range in parallel; ~PROBE_TIMEOUT total."""
    ports = range(PORT_MIN, PORT_MAX + 1)
    with ThreadPoolExecutor(max_workers=len(ports)) as pool:
        return [p for p, ok in zip(ports, pool.map(_port_alive, ports)) if ok]


def _tool_payload(response):
    """Parse the JSON text payload of a tools/call response, or None."""
    try:
        return json.loads(response['result']['content'][0]['text'])
    except Exception:
        return None


def _tool_response(request_id, payload):
    """Build a tools/call JSON-RPC response around a payload dict."""
    return {
        'jsonrpc': '2.0',
        'id': request_id,
        'result': {'content': [{'type': 'text',
                                'text': json.dumps(payload, indent=2)}]},
    }


def _rpc_error(request_id, message, code=-32603):
    return {'jsonrpc': '2.0', 'id': request_id,
            'error': {'code': code, 'message': message}}


def _documents_on(port, token):
    """list_open_documents on one instance; [] on any failure (including
    instances still running an older build without the tool)."""
    try:
        resp = _post_mcp(port, {
            'jsonrpc': '2.0', 'id': 'bridge-list',
            'method': 'tools/call',
            'params': {'name': 'list_open_documents', 'arguments': {}},
        }, token, timeout=10)
        return (_tool_payload(resp) or {}).get('documents', []) or []
    except Exception:
        return []


def _documents_map(ports, token):
    """{port: [docs]} across instances, probed in parallel so one busy
    Revit (single-threaded server mid-tool) can't stall the whole fan-out."""
    with ThreadPoolExecutor(max_workers=max(len(ports), 1)) as pool:
        return dict(zip(ports, pool.map(
            lambda p: _documents_on(p, token), ports)))


def _match_tier(doc, query):
    """Mirror the server's match precedence. Lower tier = stronger match;
    None = no match. Tiers: 0 exact title, 1 exact path, 2 file name
    (with/without extension), 3 title substring."""
    q = (query or '').strip().lower()
    if not q:
        return None
    title = (doc.get('title') or '').lower()
    path = doc.get('path') or ''
    if path == '(unsaved)':
        path = ''
    if title == q:
        return 0
    try:
        if path and os.path.normcase(os.path.normpath(path)) == \
                os.path.normcase(os.path.normpath(query.strip())):
            return 1
    except Exception:
        pass
    base = os.path.basename(path).lower() if path else ''
    if base and (base == q or os.path.splitext(base)[0] == q):
        return 2
    if q in title:
        return 3
    return None


def _ensure_alive_target(current):
    """One-time (per bridge process) validation of the startup target port.

    Port 48884 (PORT_MIN) is not guaranteed to host a T3Lab server: a killed
    Revit can leave a hung zombie socket on it, and the pyRevit Routes server
    binds 0.0.0.0:48884 when enabled. Both accept connections without ever
    answering MCP — a request then dies with a timeout or an HTTP error, NOT
    connection-refused, so the conn-refused failover never fired and
    tools/list silently fell back to the (possibly empty) cache. Probe
    /health once up front and retarget to the first alive instance."""
    if current.get('validated'):
        return
    current['validated'] = True
    if _port_alive(current['port'], timeout=PROBE_TIMEOUT):
        return
    ports = _alive_ports()
    if ports:
        current['port'] = ports[0]


def _forward(request, current, token, timeout=CALL_TIMEOUT):
    """Forward to the current instance; fail over to another alive instance
    when the current one is gone — connection refused (that Revit was
    closed) or no longer answering /health (hung socket, foreign server on
    the port). Errors from a still-healthy instance (HTTP 500, a long tool
    call timing out) are NOT retried elsewhere — re-running the call on a
    different instance would target the wrong document."""
    _ensure_alive_target(current)
    try:
        return _post_mcp(current['port'], request, token, timeout)
    except Exception as e:
        if not _is_conn_refused(e) and _port_alive(current['port']):
            raise
        for port in _alive_ports():
            if port != current['port']:
                resp = _post_mcp(port, request, token, timeout)
                current['port'] = port
                return resp
        raise


def _handle_list(request, current, token):
    """Fan list_open_documents out to every alive instance. With a single
    instance the plain server response passes through untouched."""
    ports = _alive_ports()
    if len(ports) <= 1:
        return _forward(request, current, token)

    docs_map = _documents_map(ports, token)
    instances, total = [], 0
    for port in ports:
        docs = docs_map.get(port, [])
        for d in docs:
            d['instance_port'] = port
        total += len(docs)
        instances.append({
            'instance_port': port,
            'is_current_connection': port == current['port'],
            'documents': docs,
        })
    if total == 0:
        # Every instance reported nothing — most likely they run an older
        # build without list_open_documents. Pass the plain server response
        # through instead of a misleading empty merge.
        return _forward(request, current, token)
    return _tool_response(request.get('id'), {
        'instances': instances,
        'document_count': total,
        'note': ('Several Revit windows (instances) are open. '
                 'switch_active_document works on documents in ANY of them — '
                 'the bridge re-routes tool calls to the right window '
                 'automatically.'),
    })


def _handle_switch(request, current, token):
    """Route switch_active_document to whichever instance has the document.

    Order matters: other instances are checked BEFORE letting the current
    instance open the path from disk, otherwise a full path to a file that
    is already open in another Revit window would be re-opened (read-only)
    in the current one instead of switching windows.
    """
    rid = request.get('id')
    args = (request.get('params') or {}).get('arguments') or {}
    query = args.get('path_or_title') or ''

    ports = _alive_ports()
    multi = len(ports) > 1

    # Single window — plain forward (server opens from disk if needed).
    if not multi:
        return _forward(request, current, token)

    docs_map = _documents_map(ports, token)

    # Current instance wins when it has any match (the server applies the
    # exact precedence and ambiguity rules itself).
    if any(_match_tier(d, query) is not None
           for d in docs_map.get(current['port'], [])):
        return _forward(request, current, token)

    # Look for the document in the other Revit windows, best tier first.
    matches = []
    for port in ports:
        if port == current['port']:
            continue
        for d in docs_map.get(port, []):
            tier = _match_tier(d, query)
            if tier is not None:
                matches.append((tier, port, d))
    if matches:
        best = min(m[0] for m in matches)
        matches = [m for m in matches if m[0] == best]

    if len(matches) == 1:
        _, port, _doc = matches[0]
        resp = _post_mcp(port, request, token)
        payload = _tool_payload(resp)
        if payload and payload.get('success'):
            current['port'] = port
            payload['instance_port'] = port
            payload['switched_window'] = True
            payload['note'] = ('Document is open in another Revit window — the '
                               'bridge now routes ALL tool calls to that window '
                               '(port {}).').format(port)
            return _tool_response(rid, payload)
        return resp

    if len(matches) > 1:
        return _tool_response(rid, {
            'error': ('Ambiguous: "{}" matches documents in several Revit '
                      'windows.').format(query),
            'candidates': [{'instance_port': p,
                            'title': d.get('title'),
                            'path': d.get('path')} for _, p, d in matches],
        })

    # No match in any other window — let the current instance handle it
    # (open from disk, or report a proper error). Enrich that error with
    # what the other windows DO have open, so the AI can self-correct.
    resp = _forward(request, current, token)
    payload = _tool_payload(resp)
    if payload and payload.get('error'):
        other = {str(port): [d.get('title') for d in docs_map.get(port, [])]
                 for port in ports if port != current['port']}
        if any(other.values()):
            payload['documents_in_other_windows'] = other
            return _tool_response(rid, payload)
    return resp


def _handle_initialize_local(request):
    """Answer initialize without touching Revit — attach always succeeds."""
    params = request.get('params') or {}
    protocol = params.get('protocolVersion')
    if protocol not in SUPPORTED_PROTOCOLS:
        protocol = DEFAULT_PROTOCOL
    return {
        'jsonrpc': '2.0',
        'id': request.get('id'),
        'result': {
            'protocolVersion': protocol,
            'capabilities': {'tools': {}},
            'serverInfo': {'name': 'T3Lab Revit MCP Server',
                           'version': BRIDGE_VERSION},
        },
    }


def _handle_tools_list(request, current, token):
    """Live manifest from Revit when reachable (refreshing the cache);
    cached manifest when not — so a client that attached while Revit was
    closed still gets the full tool list. The bridge's own multi-instance
    tools are appended on top either way (only the server manifest is
    cached, so they never duplicate)."""
    try:
        resp = _forward(request, current, token, timeout=LIST_TIMEOUT)
        tools = (resp.get('result') or {}).get('tools')
        if isinstance(tools, list) and tools:
            _save_cached_tools(tools)
        else:
            tools = []
    except Exception:
        tools = _load_cached_tools()
    return {'jsonrpc': '2.0', 'id': request.get('id'),
            'result': {'tools': tools + BRIDGE_TOOLS}}


def _handle_instances(request, current, token):
    """list_revit_instances — probe the port range and report every running
    Revit instance with its open documents. Served entirely by the bridge."""
    rid = request.get('id')
    ports = _alive_ports()
    if not ports:
        return _revit_down_response(rid)

    docs_map = _documents_map(ports, token)
    instances = []
    for port in ports:
        docs = docs_map.get(port, [])
        instances.append({
            'port': port,
            'is_current': port == current['port'],
            'active_document': next(
                (d.get('title') for d in docs if d.get('is_active')), None),
            'documents': [{'title': d.get('title'), 'path': d.get('path')}
                          for d in docs],
        })

    payload = {
        'instances': instances,
        'count': len(instances),
        'current_port': current['port'],
        'note': ('Tool calls go to the instance marked is_current. Retarget with '
                 'switch_revit_instance (by port) or switch_active_document (by '
                 'document — re-routes across instances automatically).'),
    }
    if current['port'] not in ports:
        payload['warning'] = ('The instance the bridge last targeted (port {}) is gone — '
                              'the next tool call fails over to another alive instance '
                              'automatically.').format(current['port'])
    return _tool_response(rid, payload)


def _handle_switch_instance(request, current, token):
    """switch_revit_instance — explicit retarget by port, or by document
    (delegated to the cross-instance switch_active_document handling so the
    target document's window is activated as well)."""
    rid = request.get('id')
    args = (request.get('params') or {}).get('arguments') or {}
    port = args.get('port')
    document = (args.get('document') or '').strip()

    if port is None and not document:
        return _tool_response(rid, {
            'error': 'Pass port (from list_revit_instances) or document.',
        })

    if port is not None:
        try:
            port = int(port)
        except (TypeError, ValueError):
            return _tool_response(rid, {'error': 'port must be an integer.'})
        # Probe the requested port directly with a generous timeout — a busy
        # instance can miss the fast parallel range scan yet still be alive,
        # and refusing a switch to a live instance is worse than waiting.
        if not _port_alive(port, timeout=TARGET_PROBE_TIMEOUT):
            ports = _alive_ports()
            if not ports:
                return _revit_down_response(rid)
            return _tool_response(rid, {
                'error': 'No running Revit instance on port {}.'.format(port),
                'alive_ports': ports,
                'hint': 'Call list_revit_instances to see what is running.',
            })
        already = port == current['port']
        current['port'] = port
        docs = _documents_on(port, token)
        return _tool_response(rid, {
            'success': True,
            'port': port,
            'already_current': already,
            'active_document': next(
                (d.get('title') for d in docs if d.get('is_active')), None),
            'documents': [d.get('title') for d in docs],
            'note': ('All tool calls now go to the Revit instance on port {}. '
                     'They target its ACTIVE document — use switch_active_document '
                     'to activate a different one.').format(port),
        })

    # By document — same semantics as switch_active_document, which already
    # finds the document across instances, switches the bridge's port and
    # activates the window.
    synthetic = {
        'jsonrpc': '2.0',
        'id': rid,
        'method': 'tools/call',
        'params': {'name': 'switch_active_document',
                   'arguments': {'path_or_title': document}},
    }
    return _handle_switch(synthetic, current, token)


def _revit_down_response(request_id):
    return _tool_response(request_id, {
        'error': 'Revit is not reachable — no T3Lab MCP server found on '
                 'ports {}-{}.'.format(PORT_MIN, PORT_MAX),
        'hint': ('Start Revit and make sure the T3Lab server is running '
                 '(it auto-starts with the T3Lab extension; otherwise use '
                 'the MCP Control button on the T3Lab ribbon). Then retry '
                 'this tool — no reconnect/restart of the client is needed.'),
    })


def _handle(request, current, token):
    method = request.get('method')
    if method == 'initialize':
        return _handle_initialize_local(request)
    if method == 'ping':
        return {'jsonrpc': '2.0', 'id': request.get('id'), 'result': {}}
    if method == 'tools/list':
        return _handle_tools_list(request, current, token)
    if method == 'tools/call':
        name = (request.get('params') or {}).get('name')
        try:
            if name == 'list_revit_instances':
                return _handle_instances(request, current, token)
            if name == 'switch_revit_instance':
                return _handle_switch_instance(request, current, token)
            if name == 'list_open_documents':
                return _handle_list(request, current, token)
            if name == 'switch_active_document':
                return _handle_switch(request, current, token)
            return _forward(request, current, token)
        except Exception as e:
            if _is_conn_refused(e):
                return _revit_down_response(request.get('id'))
            raise
    return _forward(request, current, token)


def main():
    port = PORT_MIN
    # Read port from arguments if provided
    for arg in sys.argv[1:]:
        try:
            port = int(arg)
            break
        except ValueError:
            pass

    current = {'port': port}
    token = _read_token()

    while True:
        line = sys.stdin.readline()
        if not line:
            break
        try:
            request = json.loads(line)
        except Exception:
            continue

        # Notifications (no 'id') are handled locally — the in-Revit server
        # keeps no MCP session state, so forwarding them only adds a failed
        # connect per notification whenever Revit is closed.
        if 'id' not in request:
            continue

        try:
            response = _handle(request, current, token)
        except Exception as e:
            response = _rpc_error(
                request.get('id'),
                f"T3Lab Revit server connection failed: {str(e)}")

        sys.stdout.write(json.dumps(response) + '\n')
        sys.stdout.flush()


if __name__ == '__main__':
    main()
