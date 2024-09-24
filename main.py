import os
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
from pptx import Presentation
from openpyxl import load_workbook
import win32file
import string
from threading import Thread
import time
import heapq
import json
import shutil
import win32com.client
import win32api
import tempfile
import pythoncom
import mimetypes

class FileManager:
    def __init__(self, root):
        self.root = root
        self.root.title("文件管理器")
        self.root.geometry("1200x700")

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.paned_window = ttk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        self.left_frame = ttk.Frame(self.paned_window)
        self.middle_frame = ttk.Frame(self.paned_window)
        self.right_frame = ttk.Frame(self.paned_window)

        self.paned_window.add(self.left_frame, weight=1)
        self.paned_window.add(self.middle_frame, weight=2)
        self.paned_window.add(self.right_frame, weight=2)

        # 搜索框
        self.search_frame = ttk.Frame(self.left_frame)
        self.search_frame.pack(fill=tk.X, pady=(0, 5))
        self.search_entry = ttk.Entry(self.search_frame)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.search_button = ttk.Button(self.search_frame, text="搜索", command=self.search_files)
        self.search_button.pack(side=tk.RIGHT)

        # 文件树
        self.tree = ttk.Treeview(self.left_frame, columns=("size", "date_modified", "path"), selectmode="extended")
        self.tree.heading("#0", text="文件名", anchor=tk.W)
        self.tree.heading("size", text="大小", anchor=tk.W)
        self.tree.heading("date_modified", text="修改日期", anchor=tk.W)
        self.tree.column("#0", width=150, minwidth=150)
        self.tree.column("size", width=100, minwidth=100)
        self.tree.column("date_modified", width=150, minwidth=150)
        self.tree.column("path", width=0, stretch=tk.NO)

        self.tree_scrollbar = ttk.Scrollbar(self.left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 项目集列表
        self.project_frame = ttk.Frame(self.middle_frame)
        self.project_frame.pack(fill=tk.BOTH, expand=True)

        self.project_label = ttk.Label(self.project_frame, text="项目集", font=("Arial", 12, "bold"))
        self.project_label.pack(pady=(0, 5))

        self.project_tree = ttk.Treeview(self.project_frame, selectmode="extended")
        self.project_tree.heading("#0", text="项目集", anchor=tk.W)
        self.project_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.project_scrollbar = ttk.Scrollbar(self.project_frame, orient=tk.VERTICAL, command=self.project_tree.yview)
        self.project_tree.configure(yscrollcommand=self.project_scrollbar.set)
        self.project_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.add_project_button = ttk.Button(self.middle_frame, text="新建项目集", command=self.add_project)
        self.add_project_button.pack(pady=10)

        # 预览窗口
        self.preview_frame = ttk.Frame(self.right_frame)
        self.preview_frame.pack(fill=tk.BOTH, expand=True)

        self.preview_label = ttk.Label(self.preview_frame, text="文件预览", font=("Arial", 12, "bold"))
        self.preview_label.pack(pady=(0, 5))

        self.preview_canvas = tk.Canvas(self.preview_frame)
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.preview_scrollbar = ttk.Scrollbar(self.preview_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=self.preview_scrollbar.set)
        self.preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.preview_button = ttk.Button(self.right_frame, text="预览选中文件", command=self.preview_selected_file)
        self.preview_button.pack(pady=10)

        self.progress = ttk.Progressbar(self.main_frame, orient=tk.HORIZONTAL, length=200, mode='indeterminate')
        self.progress.pack(pady=10)

        self.file_info_label = ttk.Label(self.main_frame, text="文件信息", font=("Arial", 10))
        self.file_info_label.pack(pady=5)

        self.tree.bind("<ButtonRelease-1>", self.on_file_select)
        self.tree.bind("<Double-1>", self.open_file)
        self.project_tree.bind("<ButtonRelease-1>", self.on_project_select)
        self.project_tree.bind("<Double-1>", self.open_project_file)

        # 拖放功能
        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_release)
        self.project_tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.project_tree.bind("<B1-Motion>", self.on_drag_motion)
        self.project_tree.bind("<ButtonRelease-1>", self.on_drag_release)

        # 右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="从项目集中移除", command=self.remove_from_project)
        self.context_menu.add_command(label="在桌面创建文件夹", command=self.create_folder_on_desktop)
        self.project_tree.bind("<Button-3>", self.show_context_menu)

        self.projects = {}
        self.load_projects()
        self.update_project_tree()

        self.scan_recent_files()

    def search_files(self):
        search_term = self.search_entry.get().lower()
        self.tree.delete(*self.tree.get_children())
        for file_info in self.all_files:
            if search_term in file_info[1].lower():
                self.tree.insert("", "end", text=file_info[1], values=(f"{file_info[2]} bytes", file_info[3], file_info[4]))

    def scan_recent_files(self):
        self.progress.start()
        Thread(target=self.scan_recent_files_thread).start()

    def scan_recent_files_thread(self):
        drives = win32file.GetLogicalDrives()
        available_drives = [f"{d}:\\" for d in string.ascii_uppercase if drives & (1 << ord(d) - ord('A'))]
        
        self.all_files = []
        for drive in available_drives:
            if win32file.GetDriveType(drive) == win32file.DRIVE_FIXED:
                self.all_files.extend(self.scan_directory(drive))
        
        self.all_files = heapq.nlargest(100, self.all_files, key=lambda x: x[0])
        
        self.root.after(0, self.update_tree, self.all_files)
        self.progress.stop()

    def scan_directory(self, path):
        recent_files = []
        try:
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.lower().endswith(('.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx', '.pdf')):
                        full_path = os.path.join(root, file)
                        try:
                            access_time = os.path.getatime(full_path)
                            size = os.path.getsize(full_path)
                            date_modified = time.ctime(os.path.getmtime(full_path))
                            recent_files.append((access_time, file, size, date_modified, full_path))
                        except (PermissionError, FileNotFoundError):
                            pass
        except (PermissionError, FileNotFoundError):
            pass
        return recent_files

    def update_tree(self, recent_files):
        for _, file, size, date_modified, full_path in recent_files:
            self.tree.insert("", "end", text=file, values=(f"{size} bytes", date_modified, full_path))

    def add_project(self):
        project_name = simpledialog.askstring("新建项目集", "请输入项目集名称：")
        if project_name:
            if project_name not in self.projects:
                self.projects[project_name] = []
                self.update_project_tree()
                self.save_projects()
            else:
                messagebox.showwarning("添加失败", "项目集名称已存在")

    def update_project_tree(self):
        self.project_tree.delete(*self.project_tree.get_children())
        for project_name, files in self.projects.items():
            project_node = self.project_tree.insert("", "end", text=project_name)
            for file_path in files:
                self.project_tree.insert(project_node, "end", text=os.path.basename(file_path), values=(file_path,))

    def load_projects(self):
        try:
            with open("projects.json", "r") as f:
                self.projects = json.load(f)
        except FileNotFoundError:
            self.projects = {}

    def save_projects(self):
        with open("projects.json", "w") as f:
            json.dump(self.projects, f)

    def on_file_select(self, event):
        selected_items = self.tree.selection()
        if selected_items:
            item = selected_items[0]
            file_path = self.tree.item(item, 'values')[2]
            file_info = f"文件: {os.path.basename(file_path)}\n大小: {self.tree.item(item, 'values')[0]}\n修改日期: {self.tree.item(item, 'values')[1]}"
            self.file_info_label.config(text=file_info)

    def on_project_select(self, event):
        selected_items = self.project_tree.selection()
        if selected_items:
            item = selected_items[0]
            if self.project_tree.parent(item):
                file_path = self.project_tree.item(item, 'values')[0]
                file_info = f"文件: {os.path.basename(file_path)}\n路径: {file_path}"
                self.file_info_label.config(text=file_info)
            else:
                project_name = self.project_tree.item(item, 'text')
                file_count = len(self.projects[project_name])
                project_info = f"项目集: {project_name}\n文件数量: {file_count}"
                self.file_info_label.config(text=project_info)

    def open_file(self, event):
        selected_item = self.tree.selection()[0]
        file_path = self.tree.item(selected_item, 'values')[2]
        os.startfile(file_path)

    def open_project_file(self, event):
        selected_item = self.project_tree.selection()[0]
        if self.project_tree.parent(selected_item):
            file_path = self.project_tree.item(selected_item, 'values')[0]
            os.startfile(file_path)

    # 拖放功能
    def on_drag_start(self, event):
        widget = event.widget
        if widget == self.tree:
            selection = widget.selection()
            if selection:
                self._drag_data = {'items': selection, 'source': 'tree'}
                widget.bind("<Motion>", self.on_drag_motion)
                widget.bind("<ButtonRelease-1>", self.on_drag_release)
        elif widget == self.project_tree:
            selection = widget.selection()
            if selection:
                self._drag_data = {'items': selection, 'source': 'project_tree'}
                widget.bind("<Motion>", self.on_drag_motion)
                widget.bind("<ButtonRelease-1>", self.on_drag_release)

    def on_drag_motion(self, event):
        pass  # 可以在这里添加视觉反馈，如果需要的话

    def on_drag_release(self, event):
        widget = event.widget
        if hasattr(self, '_drag_data'):
            target = event.widget.winfo_containing(event.x_root, event.y_root)
            if target == self.project_tree and self._drag_data['source'] == 'tree':
                self.add_files_to_project(self._drag_data['items'])
            widget.unbind("<Motion>")
            widget.unbind("<ButtonRelease-1>")
            del self._drag_data

    def add_files_to_project(self, items):
        project_selection = self.project_tree.selection()
        if project_selection:
            project_item = project_selection[0]
            if not self.project_tree.parent(project_item):  # 确保选中的是项目集而不是文件
                project_name = self.project_tree.item(project_item, 'text')
                for item in items:
                    file_path = self.tree.item(item, 'values')[2]
                    if file_path not in self.projects[project_name]:
                        self.projects[project_name].append(file_path)
                self.update_project_tree()
                self.save_projects()
                messagebox.showinfo("添加成功", f"文件已添加到项目集 '{project_name}'")
            else:
                messagebox.showwarning("添加失败", "请选择一个项目集而不是文件")
        else:
            messagebox.showwarning("添加失败", "请先选择一个项目集")

    # 右键菜单
    def show_context_menu(self, event):
        self.context_menu.post(event.x_root, event.y_root)

    def remove_from_project(self):
        selection = self.project_tree.selection()
        if selection:
            for selected_item in selection:
                parent_item = self.project_tree.parent(selected_item)
                if parent_item:
                    project_name = self.project_tree.item(parent_item, 'text')
                    file_path = self.project_tree.item(selected_item, 'values')[0]
                    self.projects[project_name].remove(file_path)
            self.update_project_tree()
            self.save_projects()
            messagebox.showinfo("移除成功", "选中的文件已从项目集中移除")
        else:
            messagebox.showwarning("移除失败", "请先选择一个或多个文件")

    def create_folder_on_desktop(self):
        selection = self.project_tree.selection()
        if selection:
            selected_item = selection[0]
            if not self.project_tree.parent(selected_item):  # 确保选中的是项目集而不是文件
                project_name = self.project_tree.item(selected_item, 'text')
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                project_folder_path = os.path.join(desktop_path, project_name)
                
                try:
                    os.makedirs(project_folder_path, exist_ok=True)
                    for file_path in self.projects[project_name]:
                        shutil.copy2(file_path, project_folder_path)
                    messagebox.showinfo("成功", f"项目集 '{project_name}' 已在桌面创建文件夹，并复制了所有文件")
                except Exception as e:
                    messagebox.showerror("错误", f"创建文件夹或复制文件时出错：{str(e)}")
            else:
                messagebox.showwarning("操作失败", "请选择一个项目集而不是文件")
        else:
            messagebox.showwarning("操作失败", "请先选择一个项目集")

    def preview_selected_file(self):
        selected_items = self.tree.selection() or self.project_tree.selection()
        if selected_items:
            selected_item = selected_items[0]
            if self.tree.selection():
                file_path = self.tree.item(selected_item, 'values')[2]
            else:
                file_path = self.project_tree.item(selected_item, 'values')[0]
            self.preview_file(file_path)

    def preview_file(self, file_path):
        try:
            file_extension = os.path.splitext(file_path)[1].lower()
            if file_extension == '.pdf':
                self.preview_pdf(file_path)
            else:
                self.show_file_info(file_path)
        except Exception as e:
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(10, 10, anchor=tk.NW, text=f"无法预览文件: {e}")

    def preview_pdf(self, file_path):
        try:
            doc = fitz.open(file_path)
            page = doc.load_page(0)  # 预览第一页
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((800, 800))  # 调整图片大小
            img = ImageTk.PhotoImage(img)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=img)
            self.preview_canvas.image = img  # 保持引用
            self.preview_canvas.config(scrollregion=self.preview_canvas.bbox(tk.ALL))
        except Exception as e:
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(10, 10, anchor=tk.NW, text=f"无法预览PDF文件: {e}")

    def show_file_info(self, file_path):
        try:
            file_stats = os.stat(file_path)
            file_size = file_stats.st_size
            mod_time = time.ctime(file_stats.st_mtime)
            mime_type, _ = mimetypes.guess_type(file_path)

            info_text = f"文件名: {os.path.basename(file_path)}\n"
            info_text += f"文件大小: {self.format_size(file_size)}\n"
            info_text += f"修改时间: {mod_time}\n"
            info_text += f"MIME类型: {mime_type or '未知'}\n"

            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(10, 10, anchor=tk.NW, text=info_text, font=("Arial", 12))
        except Exception as e:
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(10, 10, anchor=tk.NW, text=f"无法获取文件信息: {e}")

    def format_size(self, size):
        # 将文件大小转换为人类可读的格式
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.2f} {unit}"
            size /= 1024.0

if __name__ == "__main__":
    root = tk.Tk()
    FileManager(root)
    root.mainloop()
