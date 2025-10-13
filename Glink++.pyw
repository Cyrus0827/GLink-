import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import subprocess
import ctypes
import sys
import threading
import queue

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

class ProgressWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Operation Progress")
        self.geometry("400x150")
        self.progress_queue = queue.Queue()
        
        # 整体进度条
        ttk.Label(self, text="Total Progress:").pack(pady=(10,0))
        self.total_progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.total_progress.pack()
        
        # 当前文件进度条
        ttk.Label(self, text="Current File Progress:").pack(pady=(10,0))
        self.file_progress = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.file_progress.pack()
        
        # 状态标签
        self.status_label = ttk.Label(self, text="Initializing...")
        self.status_label.pack(pady=5)
        
        self.check_queue()

    def check_queue(self):
        try:
            while True:
                msg_type, content = self.progress_queue.get_nowait()
                if msg_type == "total_progress":
                    self.total_progress['value'] = content
                elif msg_type == "file_progress":
                    self.file_progress['value'] = content
                elif msg_type == "status":
                    self.status_label.config(text=content)
                elif msg_type == "complete":
                    self.destroy()
                    messagebox.showinfo("Complete", content)
                elif msg_type == "error":
                    self.destroy()
                    self.master.event_generate("<<ShowError>>", when="tail")
        except queue.Empty:
            pass
        self.after(100, self.check_queue)

class LinkCreatorApp:
    def __init__(self, master):
        self.master = master
        master.title("Folder Link Creator")
        self.setup_ui()
        self.stop_event = threading.Event()
        self.master.bind("<<ShowError>>", self.show_error)
        
    def setup_ui(self):
        # 源神
        ttk.Label(self.master, text="Source Folder:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.src_path = tk.StringVar()
        ttk.Entry(self.master, textvariable=self.src_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.master, text="Browse", command=self.browse_source).grid(row=0, column=2, padx=5, pady=5)

        # 目标文件夹
        ttk.Label(self.master, text="Target Folder:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.dst_path = tk.StringVar()
        ttk.Entry(self.master, textvariable=self.dst_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.master, text="Browse", command=self.browse_target).grid(row=1, column=2, padx=5, pady=5)

        # Execute 按钮
        ttk.Button(self.master, text="Create Link", command=self.execute_operations).grid(row=2, column=1, pady=10)

    def browse_source(self):
        path = filedialog.askdirectory()
        if path:
            self.src_path.set(os.path.normpath(path))

    def browse_target(self):
        path = filedialog.askdirectory()
        if path:
            self.dst_path.set(os.path.normpath(path))

    def show_error(self, event=None):
        messagebox.showerror("Error", self.error_message)

    def execute_operations(self):
        src = self.src_path.get()
        final_dst = self.dst_path.get()

        # 验证
        if not all([src, final_dst]):
            messagebox.showerror("Error", "Both fields are required!")
            return
            
        if not os.path.isdir(src):
            messagebox.showerror("Error", "Source folder does not exist!")
            return

        if os.path.exists(final_dst):
            messagebox.showerror("Error", "Target folder already exists!")
            return

        self.progress_window = ProgressWindow(self.master)
        worker = threading.Thread(
            target=self._perform_operations,
            args=(src, final_dst),
            daemon=True
        )
        worker.start()

    def _perform_operations(self, src, final_dst):
        try:
            # 阶段1：扫描文件
            self._update_progress("status", "Scanning files...")
            file_list, total_size = self._scan_files(src)
            
            # 阶段2：复制文件
            self._update_progress("status", "Copying files...")
            copied_size = 0
            chunk_size = 1024 * 1024  # 1MB
            
            for file_info in file_list:
                if self.stop_event.is_set():
                    raise RuntimeError("Operation cancelled by user")
                
                src_path = file_info['path']
                rel_path = os.path.relpath(src_path, src)
                dst_path = os.path.join(final_dst, rel_path)
                os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                
                # 更新进度
                self._update_progress("status", f"Copying: {os.path.basename(src_path)}")
                
                # 分块
                with open(src_path, 'rb') as f_src, open(dst_path, 'wb') as f_dst:
                    file_copied = 0
                    while True:
                        chunk = f_src.read(chunk_size)
                        if not chunk:
                            break
                        f_dst.write(chunk)
                        file_copied += len(chunk)
                        copied_size += len(chunk)
                        
                        # 进度条
                        file_percent = (file_copied / file_info['size']) * 100
                        total_percent = (copied_size / total_size) * 100
                        self._update_progress("file_progress", file_percent)
                        self._update_progress("total_progress", total_percent)
            
            # 3：删除源文件
            self._update_progress("status", "Deleting source...")
            shutil.rmtree(src)
            
            # 4：符号链接
            self._update_progress("status", "Creating link...")
            subprocess.run(
                ['cmd', '/c', 'mklink', '/D', src, final_dst],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                shell=True
            )
            
            self._update_progress("complete", "Operation completed successfully!")
            
        except Exception as e:
            self.error_message = str(e)
            self._update_progress("error", None)
            # 清理可能已复制的文件
            if os.path.exists(final_dst):
                shutil.rmtree(final_dst, ignore_errors=True)

    def _scan_files(self, path):
        file_list = []
        total_size = 0
        for root, dirs, files in os.walk(path):
            for file in files:
                file_path = os.path.join(root, file)
                file_size = os.path.getsize(file_path)
                file_list.append({
                    'path': file_path,
                    'size': file_size
                })
                total_size += file_size
        return file_list, total_size

    def _update_progress(self, msg_type, content):
        if hasattr(self, 'progress_window'):
            self.progress_window.progress_queue.put((msg_type, content))

if __name__ == "__main__":
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()
        
    root = tk.Tk()
    app = LinkCreatorApp(root)
    root.mainloop()
