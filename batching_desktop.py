import logging
import multiprocessing
import os
import sys
import time
import webview

# Matikan log warning Streamlit agar konsol/latar belakang tetap bersih[cite: 10]
logging.getLogger(
    "streamlit.runtime.scriptrunner_utils.script_run_context"
).setLevel(logging.ERROR)

def run_streamlit():
    # Menentukan jalur direktori saat berjalan sebagai .exe maupun skrip biasa[cite: 10]
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    # Menghubungkan langsung ke file utama batching.py[cite: 10]
    app_path = os.path.join(base_path, "batching.py")

    # Konfigurasi argumen server Streamlit[cite: 10]
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--server.headless=true",
        "--server.port=8501",
        "--global.developmentMode=false",
        "--logger.level=error",
    ]

    from streamlit.web import cli as stcli
    stcli.main()

if __name__ == "__main__":
    multiprocessing.freeze_support()

    # Menjalankan Streamlit menggunakan Process terisolasi (seperti referensi OpenApp.py)[cite: 10]
    p = multiprocessing.Process(target=run_streamlit)
    p.start()

    # Beri waktu jeda 3 detik agar server siap menerima koneksi[cite: 10]
    time.sleep(3)

    # Membuka jendela aplikasi desktop pywebview[cite: 10]
    window = webview.create_window(
        title="Plate Batching Automation - PT. r-pac Packaging Indonesia © 2026",
        url="http://localhost:8501",
        width=1280,
        height=800,
        resizable=True,
    )

    webview.start()
    
    # Mematikan proses server saat jendela ditutup[cite: 10]
    p.terminate()