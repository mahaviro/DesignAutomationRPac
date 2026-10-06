import subprocess
import sys
import os
import time
import threading
import webview

def run_streamlit():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    subprocess.run([sys.executable, "-m", "streamlit", "run", "batching.py", "--server.headless=true"])

if __name__ == "__main__":
    t = threading.Thread(target=run_streamlit)
    t.daemon = True
    t.start()

    time.sleep(3)

    webview.create_window("Plate Batching Automation", "http://localhost:8501", width=1200, height=800)
    webview.start()