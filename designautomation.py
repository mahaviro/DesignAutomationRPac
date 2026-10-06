import os
import time
import shutil
import subprocess
import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Base Directory 
BASE_DIR = r"D:\Automation File Maker\Testing"

# Sub Directory
WATCH_FOLDER = os.path.join(BASE_DIR, "Data_Import")
FMP_FOLDER = os.path.join(BASE_DIR, "Label_Key")
PRINT_POOL = os.path.join(BASE_DIR, "Print_Pool")
MASTERS_FOLDER = os.path.join(BASE_DIR, "Masters")
LAYOUT_OUTPUT = os.path.join(BASE_DIR, "Layout")
JOBFILE_OUTPUT = os.path.join(BASE_DIR, "JobFile")
ARCHIVE_FOLDER = os.path.join(WATCH_FOLDER, "Archive")

class ExcelHandler(FileSystemEventHandler):
    def __init__(self, app):
        self.app = app

    def on_created(self, event):
        if self.app.is_paused or event.is_directory:
            return
        if event.src_path.endswith('.xlsx'):
            filename = os.path.basename(event.src_path)
            self.app.root.after(0, self.app.log_message, f"[{filename}] Detected new Excel file.")
            time.sleep(1) 
            self.app.root.after(0, self.app.process_file, filename)

class AutoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Design Automation System")
        
        try:
            self.root.state('zoomed')
        except Exception:
            self.root.geometry("1000x700")
        
        self.observer = None
        self.is_paused = False
        
        self.lbl_status = tk.Label(root, text="Status: STOPPED", font=("Helvetica", 16, "bold"), fg="red")
        self.lbl_status.pack(pady=15)
        
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)
        
        btn_font = ("Helvetica", 12, "bold")
        self.btn_run = tk.Button(btn_frame, text="▶ RUN", bg="#4CAF50", fg="black", font=btn_font, width=14, height=2, command=self.start)
        self.btn_run.pack(side=tk.LEFT, padx=10)
        
        self.btn_pause = tk.Button(btn_frame, text="⏸ PAUSE", bg="#FF9800", fg="black", font=btn_font, width=14, height=2, command=self.pause, state=tk.DISABLED)
        self.btn_pause.pack(side=tk.LEFT, padx=10)
        
        self.btn_stop = tk.Button(btn_frame, text="⏹ STOP", bg="#F44336", fg="black", font=btn_font, width=14, height=2, command=self.stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=10)
        
        log_label = tk.Label(root, text="Process Console Log:", font=("Helvetica", 12, "bold"))
        log_label.pack(anchor="w", padx=25, pady=(20, 5))
        
        self.txt_log = scrolledtext.ScrolledText(root, width=110, height=22, state=tk.DISABLED, font=("Consolas", 11), bg="#F4F4F4")
        self.txt_log.pack(padx=25, pady=5)
        
        footer_label = tk.Label(root, text="PT. r-pac Packaging Indonesia © 2026 — Design Automation System", font=("Helvetica", 10, "italic"), fg="#555555")
        footer_label.pack(side=tk.BOTTOM, pady=15)
        
        self.log_message("System initialized. Waiting to start...")
        
    def log_message(self, message):
        self.txt_log.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] {message}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def process_file(self, filename):
        try:
            db_name = filename.split('_Data')[0]
            
            os.makedirs(PRINT_POOL, exist_ok=True)
            existing_folders = [d for d in os.listdir(PRINT_POOL) if os.path.isdir(os.path.join(PRINT_POOL, d)) and d.startswith(f"{db_name}_")]
            running_number = len(existing_folders) + 1
            
            folder_name = f"{db_name}_{running_number:03d}" 
            target_dir = os.path.join(PRINT_POOL, folder_name)
            os.makedirs(target_dir, exist_ok=True)
            self.log_message(f"[{filename}] Created print folder: {folder_name}")

            tab_filename = f"{db_name}_tabfile.tab"
            tab_os_path = os.path.join(target_dir, tab_filename)
            
            relative_path = f"{folder_name}/{tab_filename}"
            encoded_param = relative_path.replace(" ", "%20")

            fmp_filename = f"{db_name}_Label_Key.fmp12"
            fm_path = os.path.join(FMP_FOLDER, fmp_filename)
            
            if os.path.exists(fm_path):
                # Launch FileMaker minimized
                self.log_message(f"[{filename}] Opening database '{fmp_filename}' (Minimized)...")
                subprocess.Popen(f'start /min "" "{fm_path}"', shell=True)
                time.sleep(4) 
                
                self.log_message(f"[{filename}] Triggering Master_Process...")
                os.startfile(f"fmp://$/{fmp_filename}?script=Master_Process&param={encoded_param}")
                
                self.log_message(f"[{filename}] Waiting for .tab file generation...")
                timeout = 45 
                while not os.path.exists(tab_os_path) and timeout > 0:
                    time.sleep(1)
                    timeout -= 1
                
                if os.path.exists(tab_os_path):
                    self.log_message(f"[{filename}] SUCCESS: File '{tab_filename}' created.")
                    
                    os.makedirs(LAYOUT_OUTPUT, exist_ok=True)
                    os.makedirs(JOBFILE_OUTPUT, exist_ok=True)
                    
                    # C. Adobe InDesign Automation for Layout Template (.indt -> .indd)
                    layout_indd_name = f"{db_name}_layout.indt"
                    layout_indd_path = os.path.join(MASTERS_FOLDER, layout_indd_name)
                    indd_layout_name = f"{db_name}_layout_{running_number:03d}.indd"
                    indd_layout_path = os.path.join(LAYOUT_OUTPUT, indd_layout_name)
                    pdf_layout_name = f"{db_name}_layout_{running_number:03d}.pdf"
                    pdf_layout_path = os.path.join(LAYOUT_OUTPUT, pdf_layout_name)
                    
                    if os.path.exists(layout_indd_path):
                        self.log_message(f"[{filename}] Processing InDesign Layout template (.indt)...")
                        
                        vbs_layout_indt = layout_indd_path.replace(os.sep, '/')
                        vbs_tab_path = tab_os_path.replace(os.sep, '/')
                        vbs_indd_layout = indd_layout_path.replace(os.sep, '/')
                        vbs_pdf_layout = pdf_layout_path.replace(os.sep, '/')
                        
                        vbs_layout_code = f"""
                        Option Explicit
                        Dim myInDesign, doc, myRecordCount
                        Const idPDFType = 1952403524
                        
                        On Error Resume Next
                        Set myInDesign = CreateObject("InDesign.Application")
                        If Err.Number <> 0 Then
                            WScript.Quit
                        End If
                        On Error GoTo 0
                        
                        myInDesign.Activate
                        Set doc = myInDesign.Open("{vbs_layout_indt}")
                        
                        On Error Resume Next
                        myRecordCount = myInDesign.importDataFromFile("{vbs_tab_path}", "Prototype", "Label", "")
                        
                        ' Minimize InDesign window automatically
                        myInDesign.ActiveWindow.Minimize
                        On Error GoTo 0
                        
                        ' Wait for InData database transaction to close completely
                        WScript.Sleep 3000
                        
                        ' Export to PDF
                        doc.Export idPDFType, "{vbs_pdf_layout}", False
                        
                        ' Save as normal .indd document
                        doc.Save "{vbs_indd_layout}"
                        
                        ' Close document only (InDesign application remains open)
                        doc.Close
                        """
                        
                        vbs_layout_file = os.path.join(target_dir, "process_layout.vbs")
                        with open(vbs_layout_file, "w", encoding="utf-8") as vf:
                            vf.write(vbs_layout_code)
                        
                        os.startfile(vbs_layout_file)
                        time.sleep(8)
                        self.log_message(f"[{filename}] Layout processed and saved as .indd (InDesign minimized).")
                    else:
                        self.log_message(f"[{filename}] WARNING: Template '{layout_indd_name}' not found in Masters folder.")

                    # D. Adobe InDesign Automation for JobFile Template (.indt -> .indd)
                    jobfile_indd_name = f"{db_name}_JobFile.indt"
                    jobfile_indd_path = os.path.join(MASTERS_FOLDER, jobfile_indd_name)
                    indd_jobfile_name = f"{db_name}_JobFile_{running_number:03d}.indd"
                    indd_jobfile_path = os.path.join(JOBFILE_OUTPUT, indd_jobfile_name)
                    pdf_jobfile_name = f"{db_name}_JobFile_{running_number:03d}.pdf"
                    pdf_jobfile_path = os.path.join(JOBFILE_OUTPUT, pdf_jobfile_name)
                    
                    if os.path.exists(jobfile_indd_path):
                        self.log_message(f"[{filename}] Processing InDesign JobFile template (.indt)...")
                        
                        vbs_jobfile_indt = jobfile_indd_path.replace(os.sep, '/')
                        vbs_tab_path = tab_os_path.replace(os.sep, '/')
                        vbs_indd_jobfile = indd_jobfile_path.replace(os.sep, '/')
                        vbs_pdf_jobfile = pdf_jobfile_path.replace(os.sep, '/')
                        
                        vbs_jobfile_code = f"""
                        Option Explicit
                        Dim myInDesign, doc, myRecordCount
                        Const idPDFType = 1952403524
                        
                        On Error Resume Next
                        Set myInDesign = CreateObject("InDesign.Application")
                        If Err.Number <> 0 Then
                            WScript.Quit
                        End If
                        On Error GoTo 0
                        
                        myInDesign.Activate
                        Set doc = myInDesign.Open("{vbs_jobfile_indt}")
                        
                        On Error Resume Next
                        myRecordCount = myInDesign.importDataFromFile("{vbs_tab_path}", "Prototype", "Label", "")
                        
                        ' Minimize InDesign window automatically
                        myInDesign.ActiveWindow.Minimize
                        On Error GoTo 0
                        
                        ' Wait for InData database transaction to close completely
                        WScript.Sleep 3000
                        
                        ' Export to PDF
                        doc.Export idPDFType, "{vbs_pdf_jobfile}", False
                        
                        ' Save as normal .indd document
                        doc.Save "{vbs_indd_jobfile}"
                        
                        ' Close document only (InDesign application remains open)
                        doc.Close
                        """
                        
                        vbs_jobfile_file = os.path.join(target_dir, "process_jobfile.vbs")
                        with open(vbs_jobfile_file, "w", encoding="utf-8") as vf:
                            vf.write(vbs_jobfile_code)
                        
                        os.startfile(vbs_jobfile_file)
                        time.sleep(8)
                        self.log_message(f"[{filename}] JobFile processed and saved as .indd (InDesign minimized).")
                    else:
                        self.log_message(f"[{filename}] WARNING: Template '{jobfile_indd_name}' not found in Masters folder.")
                    
                    # E. Move & Rename .xlsx to Archive
                    os.makedirs(ARCHIVE_FOLDER, exist_ok=True)
                    source_excel = os.path.join(WATCH_FOLDER, filename)
                    name_part, ext_part = os.path.splitext(filename)
                    new_archive_name = f"{name_part}_{running_number:03d}{ext_part}"
                    dest_excel = os.path.join(ARCHIVE_FOLDER, new_archive_name)
                    
                    if os.path.exists(dest_excel):
                        os.remove(dest_excel)
                        
                    shutil.move(source_excel, dest_excel)
                    self.log_message(f"[{filename}] SUCCESS: Excel moved & renamed to Archive as '{new_archive_name}'.")
                else:
                    self.log_message(f"[{filename}] ERROR: File .tab not found. Check FileMaker process.")
                
            else:
                self.log_message(f"[{filename}] ERROR: '{fmp_filename}' not found in Label_Key folder.")
                
            self.log_message("-" * 60) 
            
        except Exception as e:
            self.log_message(f"[{filename}] SYSTEM ERROR: {str(e)}")

    def start(self):
        os.makedirs(WATCH_FOLDER, exist_ok=True)
        os.makedirs(FMP_FOLDER, exist_ok=True)
        os.makedirs(PRINT_POOL, exist_ok=True)
        os.makedirs(MASTERS_FOLDER, exist_ok=True)
        os.makedirs(LAYOUT_OUTPUT, exist_ok=True)
        os.makedirs(JOBFILE_OUTPUT, exist_ok=True)
        os.makedirs(ARCHIVE_FOLDER, exist_ok=True)
        
        self.lbl_status.config(text="Status: RUNNING", fg="green")
        self.btn_run.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL, text="⏸ PAUSE")
        self.btn_stop.config(state=tk.NORMAL)
        
        existing_files = [f for f in os.listdir(WATCH_FOLDER) if f.endswith('.xlsx')]
        if existing_files:
            self.log_message(f"Found {len(existing_files)} existing Excel file(s). Processing...")
            for filename in existing_files:
                self.process_file(filename)

        self.is_paused = False
        self.event_handler = ExcelHandler(self)
        self.observer = Observer()
        self.observer.schedule(self.event_handler, WATCH_FOLDER, recursive=False)
        self.observer.start()
        self.log_message("Monitoring active. Waiting for incoming files...")
        
    def pause(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.lbl_status.config(text="Status: PAUSED", fg="orange")
            self.btn_pause.config(text="▶ RESUME")
            self.log_message("Monitoring paused. New files will be ignored.")
        else:
            self.lbl_status.config(text="Status: RUNNING", fg="green")
            self.btn_pause.config(text="⏸ PAUSE")
            self.log_message("Monitoring resumed.")
            
    def stop(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
        self.is_paused = False
        self.lbl_status.config(text="Status: STOPPED", fg="red")
        self.btn_run.config(state=tk.NORMAL)
        self.btn_pause.config(state=tk.DISABLED, text="⏸ PAUSE")
        self.btn_stop.config(state=tk.DISABLED)
        self.log_message("Monitoring stopped.")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconbitmap(r"D:\Automation File Maker\app_icon.ico")
    except Exception:
        pass    
    app = AutoApp(root)
    root.mainloop()