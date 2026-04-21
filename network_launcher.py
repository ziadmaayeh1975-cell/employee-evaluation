import subprocess
import sys
import os
import time
import webbrowser
import socket
import threading

APP_FILE = "app.py"
APP_NAME = "Fannoun System - Network Mode"
PORT = 8501
HOST = "0.0.0.0"

def get_app_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "172.130.10.9"

def is_port_available(port):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(1)
        result = sock.connect_ex(('localhost', port))
        sock.close()
        return result != 0
    except:
        return True

def find_available_port(start_port=8501):
    for port in range(start_port, start_port + 100):
        if is_port_available(port):
            return port
    return start_port

def find_streamlit():
    python_dir = os.path.dirname(sys.executable)
    scripts_dir = os.path.join(python_dir, 'Scripts')
    possible_paths = [
        os.path.join(scripts_dir, 'streamlit.exe'),
        os.path.join(scripts_dir, 'streamlit'),
        'streamlit'
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return 'streamlit'

def open_browser(port):
    time.sleep(4)
    webbrowser.open(f"http://localhost:{port}")

def main():
    local_ip = get_local_ip()

    print("=" * 55)
    print("    Fannoun System - Network Mode")
    print("=" * 55)
    print()

    app_dir = get_app_path()
    os.chdir(app_dir)

    app_path = os.path.join(app_dir, APP_FILE)
    if not os.path.exists(app_path):
        print(f"ERROR: {APP_FILE} not found!")
        input("Press Enter to close...")
        sys.exit(1)

    port = find_available_port(PORT)
    streamlit_path = find_streamlit()

    print(f"  Server IP  : {local_ip}")
    print(f"  Port       : {port}")
    print()
    print("=" * 55)
    print("  LINKS FOR USERS:")
    print(f"  This PC    : http://localhost:{port}")
    print(f"  Network    : http://{local_ip}:{port}")
    print("=" * 55)
    print()
    print("  Share this link with other users:")
    print(f"  >>> http://{local_ip}:{port} <<<")
    print()
    print("  DO NOT close this window!")
    print("  To stop: press Ctrl+C")
    print("-" * 55)

    t = threading.Thread(target=open_browser, args=(port,), daemon=True)
    t.start()

    cmd = []
    cmd.append(streamlit_path)
    cmd.append("run")
    cmd.append(app_path)
    cmd.append(f"--server.port={port}")
    cmd.append("--server.headless=true")
    cmd.append("--browser.gatherUsageStats=false")
    cmd.append(f"--server.address={HOST}")

    try:
        subprocess.run(cmd)
    except KeyboardInterrupt:
        print("\nSystem stopped")
    except FileNotFoundError:
        print("ERROR: streamlit not installed!")
        input("Press Enter to close...")

if __name__ == "__main__":
    main()