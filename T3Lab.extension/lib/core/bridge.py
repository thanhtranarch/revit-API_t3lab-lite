import sys
import json
import urllib.request
import urllib.error

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
        req = urllib.request.Request(
            url,
            data=data,
            headers={'Content-Type': 'application/json'}
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
