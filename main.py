import subprocess
import sys
import importlib.util
import os
import logging

REQUIRED_PACKAGES = ["flask", "pywebview", "xlrd"]

def install_and_launch():
    for package in REQUIRED_PACKAGES:
        search_name = "webview" if package == "pywebview" else package
        if importlib.util.find_spec(search_name) is None:
            print(f"[*] Setting up internal tool: Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

    os.environ['WERKZEUG_RUN_MAIN'] = 'true'
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)

    import webview
    from app import app 
    import threading

    def run_flask():
        app.run(port=5000, debug=False, use_reloader=False)

    print("[+] System Ready. Launching Excel Comparison Tool...")
    
    t = threading.Thread(target=run_flask)
    t.daemon = True
    t.start()

    webview.create_window('Excel Comparison Tool', 'http://127.0.0.1:5000', 
                          width=1300, height=800)
    webview.start()

if __name__ == "__main__":
    try:
        install_and_launch()
    except Exception as e:
        print(f"\n[!] Error: {e}")
        input("Press Enter to close...")