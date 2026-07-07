import http.server
import socketserver
import webbrowser
import os
import sys

# Set encoding
sys.stdout.reconfigure(encoding='utf-8')

PORT = 8080
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

import subprocess

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)
        
    def do_GET(self):
        # Whenever a request is made, check if Excel has been modified
        if self.path == "/" or self.path.endswith("index.html") or self.path.endswith("data.js"):
            self.check_and_regenerate()
        return super().do_GET()

    def check_and_regenerate(self):
        f5 = 'D:\\AVEAM\\Antigravity2\\Test-LiveArtifact\\5_PN_Job_Relationship_rev00-test.xlsx'
        f1 = os.path.join(DIRECTORY, '1-List of each sub-assembly.xlsx')
        data_js = os.path.join(DIRECTORY, 'data.js')
        
        should_regen = not os.path.exists(data_js)
        
        if not should_regen:
            js_mtime = os.path.getmtime(data_js)
            for f in [f5, f1]:
                if os.path.exists(f):
                    if os.path.getmtime(f) > js_mtime:
                        should_regen = True
                        break
                        
        if should_regen:
            print("\n[REALTIME UPDATE] Excel file change detected. Regenerating data.js...")
            try:
                script_path = os.path.join(DIRECTORY, 'update_prototype_data.py')
                result = subprocess.run([sys.executable, script_path], capture_output=True, text=True, encoding='utf-8')
                if result.returncode == 0:
                    print("[REALTIME UPDATE] data.js regenerated successfully.")
                else:
                    print("[ERROR] Failed to regenerate data.js:")
                    print(result.stderr)
            except Exception as e:
                print(f"[ERROR] Exception during regeneration: {e}")

    def end_headers(self):
        # Prevent caching for live updates in dev
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

import socket

def get_local_ip():
    try:
        # Create a dummy socket connection to get the active local network interface IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def main():
    os.chdir(DIRECTORY)
    
    # Configure socket server with reuse address allowed
    socketserver.TCPServer.allow_reuse_address = True
    
    try:
        local_ip = get_local_ip()
        with socketserver.TCPServer(("", PORT), Handler) as httpd:
            print(f"==================================================")
            print(f"    AVEAM REAL-TIME DASHBOARD SERVER RUNNING")
            print(f"==================================================")
            print(f" URL (คอมพิวเตอร์เครื่องนี้): http://localhost:{PORT}")
            print(f" URL (สำหรับต่อจอทีวี/มือถือ): http://{local_ip}:{PORT}")
            print(f" โฟลเดอร์ระบบ: {DIRECTORY}")
            print(f"==================================================")
            print(f" กด Ctrl+C เพื่อหยุดการทำงานของเซิร์ฟเวอร์...")
            
            # Open browser automatically for dashboard
            webbrowser.open(f"http://localhost:{PORT}/dashboard.html")
            
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nหยุดการทำงานเซิร์ฟเวอร์เรียบร้อยแล้ว.")
    except Exception as e:
        print(f"Error starting server: {e}")

if __name__ == '__main__':
    main()
