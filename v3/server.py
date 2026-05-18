import json
import os
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler
from utils import get_config

class DashboardAPIHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'X-Requested-With, Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed_path = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed_path.query)
        
        # エンドポイントのルーティング
        if parsed_path.path.startswith('/api/'):
            self.handle_api_request(parsed_path.path, query)
        else:
            # フロントエンドの静的ファイルを提供
            super().do_GET()

    def handle_api_request(self, path, query):
        brand = query.get('brand', ['FRV'])[0]
        cfg = get_config()
        output_dir = cfg["paths"]["output_dir"]
        json_path = os.path.join(output_dir, f"dashboard_data_{brand}.json")
        
        if not os.path.exists(json_path):
            self.send_error_response(404, f"Data not found for brand: {brand}")
            return
            
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            self.send_error_response(500, f"Error reading data: {e}")
            return

        if path == '/api/dashboard':
            store = query.get('store', [None])[0]
            if store:
                # 特定店舗のデータのみを返す
                store_data = data.get("stores", {}).get(store)
                if store_data:
                    self.send_json_response(200, store_data)
                else:
                    self.send_error_response(404, f"Store not found: {store}")
            else:
                # ブランド全体のデータを返す (storesは除外またはサマリーのみにしても良いが、今回はそのまま返す)
                self.send_json_response(200, data)
                
        elif path == '/api/stores':
            # 店舗リストを返す
            stores = list(data.get("stores", {}).keys())
            self.send_json_response(200, {"stores": stores})
            
        else:
            self.send_error_response(404, "Endpoint not found")

    def send_json_response(self, status_code, payload):
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        response_json = json.dumps(payload, ensure_ascii=False)
        self.wfile.write(response_json.encode('utf-8'))

    def send_error_response(self, status_code, message):
        self.send_json_response(status_code, {"error": message})


def run_server(port=8000):
    server_address = ('', port)
    httpd = HTTPServer(server_address, DashboardAPIHandler)
    print(f"Starting lightweight API server on port {port}...")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        httpd.server_close()

if __name__ == '__main__':
    run_server()
