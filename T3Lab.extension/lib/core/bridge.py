"""T3Lab MCP stdio bridge.

Forwards JSON-RPC requests from an MCP client (Claude Desktop / Claude Code)
to the T3Lab HTTP server running inside Revit.

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
                             was closed (connection refused), the bridge
                             fails over to another alive instance.
"""

import sys
import os
import json
import urllib.request
import urllib.error
from concurrent.futures import ThreadPoolExecutor

PORT_MIN, PORT_MAX = 48884, 48894
PROBE_TIMEOUT = 0.5    # /health probe per port (run in parallel)
CALL_TIMEOUT = 130     # matches the server's 120s ExternalEvent wait + margin


def _read_token():
    """Read the shared-secret token T3LabAIServer writes on first run to
    %APPDATA%\\T3LabAI\\mcp_token.txt. All Revit instances on the machine
    share this file, so one token authenticates against every instance."""
    try:
        app_data = os.environ.get('APPDATA', '')
        token_path = os.path.join(app_data, 'T3LabAI', 'mcp_token.txt')
        with open(token_path, 'r') as f:
            return f.read().strip()
    except Exception:
        return ''


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


def _alive_ports():
    """Probe the whole port range in parallel; ~PROBE_TIMEOUT total."""
    def probe(port):
        try:
            req = urllib.request.Request(f"http://127.0.0.1:{port}/health")
            with urllib.request.urlopen(req, timeout=PROBE_TIMEOUT) as r:
                return port if r.status == 200 else None
        except Exception:
            return None

    ports = range(PORT_MIN, PORT_MAX + 1)
    with ThreadPoolExecutor(max_workers=len(ports)) as pool:
        return [p for p in pool.map(probe, ports) if p is not None]


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


def _forward(request, current, token, timeout=CALL_TIMEOUT):
    """Forward to the current instance; on connection refused (that Revit
    was closed) fail over to another alive instance and retry once."""
    try:
        return _post_mcp(current['port'], request, token, timeout)
    except Exception as e:
        if not _is_conn_refused(e):
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


def _handle(request, current, token):
    if request.get('method') == 'tools/call':
        name = (request.get('params') or {}).get('name')
        if name == 'list_open_documents':
            return _handle_list(request, current, token)
        if name == 'switch_active_document':
            return _handle_switch(request, current, token)
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

        # Notifications (no 'id') are forwarded fire-and-forget
        is_notification = 'id' not in request

        try:
            response = _handle(request, current, token)
        except Exception as e:
            response = _rpc_error(
                request.get('id'),
                f"T3Lab Revit server connection failed: {str(e)}")

        if not is_notification:
            sys.stdout.write(json.dumps(response) + '\n')
            sys.stdout.flush()


if __name__ == '__main__':
    main()
