import sys
import os
import json
import urllib.request
import urllib.error

def _read_token():
    """Read the shared-secret token T3LabAIServer writes on first run to
    %APPDATA%\\T3LabAI\\mcp_token.txt. The server rejects any /mcp request
    without it, so this bridge needs to forward it on every call."""
    try:
        app_data = os.environ.get('APPDATA', '')
        token_path = os.path.join(app_data, 'T3LabAI', 'mcp_token.txt')
        with open(token_path, 'r') as f:
            return f.read().strip()
    except Exception:
        return ''

def main():
    port = 48884
    # Read port from arguments if provided
    for arg in sys.argv[1:]:
        try:
            port = int(arg)
            break
        except ValueError:
            pass

    url = f"http://localhost:{port}/mcp"
    token = _read_token()

    while True:
        line = sys.stdin.readline()
        if not line:
            break
        try:
            request = json.loads(line)
        except Exception:
            continue

        # Check if it is a notification (no 'id' parameter)
        is_notification = 'id' not in request

        # Forward the JSON-RPC request to Revit's in-process HTTP server
        data = json.dumps(request).encode('utf-8')
        headers = {'Content-Type': 'application/json'}
        if token:
            headers['Authorization'] = 'Bearer ' + token
        req = urllib.request.Request(
            url,
            data=data,
            headers=headers
        )
        try:
            with urllib.request.urlopen(req) as response:
                res_body = response.read().decode('utf-8')
                if not is_notification:
                    sys.stdout.write(res_body + '\n')
                    sys.stdout.flush()
        except Exception as e:
            if not is_notification:
                # Return JSON-RPC error if connection is refused (e.g. Revit closed/server stopped)
                error_res = {
                    "jsonrpc": "2.0",
                    "id": request.get("id"),
                    "error": {
                        "code": -32603,
                        "message": f"T3Lab Revit server connection failed: {str(e)}"
                    }
                }
                sys.stdout.write(json.dumps(error_res) + '\n')
                sys.stdout.flush()

if __name__ == '__main__':
    main()
