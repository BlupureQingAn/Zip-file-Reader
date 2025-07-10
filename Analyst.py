import zipfile
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.font import Font
import threading
import traceback
import ctypes

# 设置Windows控制台编码为UTF-8
if sys.platform == "win32":
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleOutputCP(65001)
    kernel32.SetConsoleCP(65001)


def decode_zip_name(name_bytes, encoding='gbk'):
    """尝试多种编码方式解码ZIP中的文件名"""
    # 如果已经是字符串，直接返回
    if isinstance(name_bytes, str):
        return name_bytes

    encodings = ['utf-8', 'gbk', 'gb18030', 'big5', 'cp950', 'latin1', 'cp437']

    for enc in encodings:
        try:
            return name_bytes.decode(enc)
        except (UnicodeDecodeError, AttributeError):
            continue

    # 所有编码都失败时使用回退方案
    try:
        return name_bytes.decode('utf-8', errors='replace')
    except:
        try:
            return name_bytes.decode('latin1', errors='replace')
        except:
            return str(name_bytes)


class ZipViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ZIP文件结构查看器")
        self.root.geometry("900x600")
        self.root.minsize(700, 500)

        # 设置应用图标
        try:
            if os.path.exists("zip_icon.ico"):
                self.root.iconbitmap(default="zip_icon.ico")
        except:
            pass

        # 创建字体
        try:
            self.title_font = Font(family="Microsoft YaHei", size=16, weight="bold")
            self.normal_font = Font(family="Microsoft YaHei", size=10)
            self.markdown_font = Font(family="Consolas", size=10)  # Markdown显示字体
        except:
            # 回退字体
            self.title_font = Font(size=16, weight="bold")
            self.normal_font = Font(size=10)
            self.markdown_font = Font(family="Courier", size=10)

        # 创建UI
        self.create_widgets()

        # 设置初始状态
        self.current_zip = None
        self.markdown_content = ""

        # 兼容的拖放支持
        self.setup_drag_drop()

    def setup_drag_drop(self):
        """设置兼容的拖放支持"""
        if sys.platform == 'win32':
            # Windows 平台使用兼容方法
            self.root.bind("<Button-1>", self.on_drag_start)
            self.root.bind("<B1-Motion>", self.on_drag_motion)
            self.root.bind("<ButtonRelease-1>", self.on_drop)
        else:
            # 其他平台尝试标准方法
            try:
                self.root.drop_target_register(tk.DND_FILES)
                self.root.dnd_bind('<<Drop>>', self.on_drop)
            except AttributeError:
                self.status.config(text="当前平台不支持拖放功能")

    def create_widgets(self):
        # 创建顶部框架
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        # 标题
        title_label = ttk.Label(top_frame, text="ZIP文件结构查看器", font=self.title_font)
        title_label.pack(side=tk.LEFT, padx=5)

        # 操作按钮
        btn_frame = ttk.Frame(top_frame)
        btn_frame.pack(side=tk.RIGHT, padx=5)

        open_btn = ttk.Button(btn_frame, text="打开ZIP文件", command=self.open_zip)
        open_btn.pack(side=tk.LEFT, padx=5)

        save_btn = ttk.Button(btn_frame, text="保存Markdown", command=self.save_markdown)
        save_btn.pack(side=tk.LEFT, padx=5)

        # 文件信息显示
        info_frame = ttk.LabelFrame(self.root, text="文件信息")
        info_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.path_label = ttk.Label(info_frame, text="未选择文件", font=self.normal_font)
        self.path_label.pack(side=tk.LEFT, padx=10, pady=5, fill=tk.X, expand=True)

        self.size_label = ttk.Label(info_frame, text="", font=self.normal_font)
        self.size_label.pack(side=tk.RIGHT, padx=10, pady=5)

        # 创建Markdown显示区域
        markdown_frame = ttk.LabelFrame(self.root, text="文件结构 (Markdown格式)")
        markdown_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # 添加滚动文本框
        self.markdown_text = scrolledtext.ScrolledText(
            markdown_frame,
            wrap=tk.WORD,
            font=self.markdown_font,
            padx=10,
            pady=10
        )
        self.markdown_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.markdown_text.config(state=tk.DISABLED)  # 初始为只读

        # 状态栏
        self.status = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(side=tk.BOTTOM, fill=tk.X, padx=1, pady=1)

    def open_zip(self, file_path=None):
        """打开ZIP文件"""
        if not file_path:
            file_path = filedialog.askopenfilename(
                filetypes=[("ZIP文件", "*.zip"), ("所有文件", "*.*")]
            )

        if not file_path:
            return

        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在！")
            return

        # 在新线程中加载ZIP文件
        threading.Thread(target=self.load_zip_file, args=(file_path,), daemon=True).start()

    def load_zip_file(self, file_path):
        """在新线程中加载ZIP文件并生成Markdown"""
        try:
            # 更新状态
            self.status.config(text="正在加载ZIP文件并生成Markdown...")
            self.root.update()

            # 清空现有内容
            self.markdown_content = ""

            # 加载新ZIP文件
            self.current_zip = file_path

            # 生成Markdown内容
            self.markdown_content = "# ZIP文件内容结构\n\n"
            self.markdown_content += f"**文件名**: `{os.path.basename(file_path)}`\n\n"

            # 获取文件大小
            file_size = os.path.getsize(file_path)
            size_mb = file_size / (1024 * 1024)
            self.markdown_content += f"**文件大小**: {size_mb:.2f} MB\n\n"

            # 在后台解析ZIP文件
            with zipfile.ZipFile(file_path, 'r') as zf:
                # 构建目录树
                tree = {}

                for info in zf.infolist():
                    try:
                        # 获取原始文件名字节并解码
                        orig_bytes = getattr(info, '_orig_filename', info.filename.encode('cp437'))
                        name = decode_zip_name(orig_bytes)

                        # 处理目录条目
                        is_dir = name.endswith('/')
                        parts = [p for p in name.rstrip('/').split('/') if p]

                        if parts:
                            current = tree
                            for part in parts:
                                if part not in current:
                                    current[part] = {}
                                current = current[part]
                    except Exception as e:
                        # 记录错误但继续处理其他文件
                        error_msg = f"处理文件时出错: {traceback.format_exc()}"
                        self.root.after(0, lambda: self.status.config(text=error_msg[:100] + "..."))

                # 生成Markdown树形结构
                self.markdown_content += "## 文件结构\n```\n"
                self.markdown_content += self.generate_markdown_tree(tree, 0)
                self.markdown_content += "```\n\n"

                # 生成文件列表
                self.markdown_content += "## 文件列表\n"
                self.markdown_content += "| 文件名 | 大小 | 压缩后大小 | 压缩比 |\n"
                self.markdown_content += "|--------|------|------------|--------|\n"

                for info in zf.infolist():
                    try:
                        # 获取原始文件名字节并解码
                        orig_bytes = getattr(info, '_orig_filename', info.filename.encode('cp437'))
                        name = decode_zip_name(orig_bytes)

                        if not name.endswith('/'):  # 只处理文件，跳过目录
                            # 计算压缩比
                            compress_ratio = 100
                            if info.file_size > 0:
                                compress_ratio = (info.compress_size / info.file_size) * 100

                            # 添加文件信息行
                            self.markdown_content += (
                                f"| `{name}` | "
                                f"{self.format_size(info.file_size)} | "
                                f"{self.format_size(info.compress_size)} | "
                                f"{compress_ratio:.1f}% |\n"
                            )
                    except Exception as e:
                        # 记录错误但继续处理其他文件
                        pass

            # 在UI线程中更新Markdown显示
            self.root.after(0, self.update_markdown_display)

            # 更新文件信息
            self.root.after(0, self.update_status,
                            f"文件路径: {file_path}",
                            f"文件大小: {size_mb:.2f} MB",
                            f"加载完成: {os.path.basename(file_path)}")

        except Exception as e:
            error_msg = f"无法读取ZIP文件: {traceback.format_exc()}"
            self.root.after(0, self.show_error, "错误", error_msg)
            self.root.after(0, lambda: self.status.config(text="加载失败"))

    def generate_markdown_tree(self, tree, indent_level):
        """递归生成Markdown树形结构"""
        result = ""
        indent = "    " * indent_level
        items = sorted(tree.items())

        for i, (name, children) in enumerate(items):
            # 判断是否为最后一个节点
            is_last = (i == len(items) - 1)

            # 当前节点的前缀符号
            connector = "└── " if is_last else "├── "
            prefix = indent + connector

            # 添加当前节点
            result += prefix + name + "\n"

            # 递归添加子节点
            if children:
                new_indent = indent + ("    " if is_last else "│   ")
                result += self.generate_markdown_tree(children, indent_level + 1)

        return result

    def update_markdown_display(self):
        """更新Markdown显示区域"""
        self.markdown_text.config(state=tk.NORMAL)  # 启用编辑
        self.markdown_text.delete(1.0, tk.END)  # 清空内容
        self.markdown_text.insert(tk.END, self.markdown_content)
        self.markdown_text.config(state=tk.DISABLED)  # 恢复只读

    def save_markdown(self):
        """保存Markdown到文件"""
        if not self.markdown_content:
            messagebox.showinfo("提示", "没有可保存的内容")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown文件", "*.md"), ("所有文件", "*.*")]
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.markdown_content)
            messagebox.showinfo("保存成功", f"Markdown文件已保存到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("保存错误", f"保存文件时出错:\n{str(e)}")

    def show_error(self, title, message):
        """显示错误信息，处理可能的多行消息"""
        # 截断过长的错误消息
        if len(message) > 1000:
            message = message[:1000] + "\n... [消息过长被截断] ..."
        messagebox.showerror(title, message)

    def update_status(self, path_text, size_text, status_text):
        """更新UI状态"""
        self.path_label.config(text=path_text)
        self.size_label.config(text=size_text)
        self.status.config(text=status_text)

    def format_size(self, size):
        """格式化文件大小"""
        if size is None:
            return "未知"
        if size < 1024:
            return f"{size} 字节"
        elif size < 1024 * 1024:
            return f"{size / 1024:.1f} KB"
        else:
            return f"{size / (1024 * 1024):.1f} MB"

    def on_drag_start(self, event):
        """开始拖拽"""
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def on_drag_motion(self, event):
        """拖拽中"""
        # 简单实现 - 在Windows上不需要特别处理
        pass

    def on_drop(self, event):
        """处理文件拖放事件"""
        # Windows兼容的拖放处理
        try:
            if sys.platform == 'win32':
                # 在Windows上，我们可以从剪贴板获取文件路径
                self.root.clipboard_clear()
                self.root.clipboard_append("")
                self.root.update()

                # 获取剪贴板内容
                clipboard_content = self.root.clipboard_get()

                # 清理可能的额外字符
                file_path = clipboard_content.strip().replace('{', '').replace('}', '')

                # 检查是否是有效路径
                if os.path.exists(file_path) and file_path.lower().endswith('.zip'):
                    self.open_zip(file_path)
        except Exception as e:
            messagebox.showerror("拖放错误", f"无法处理拖放的文件:\n{str(e)}")


if __name__ == "__main__":
    # 设置系统编码为UTF-8
    if sys.version_info >= (3, 7):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')

    root = tk.Tk()
    app = ZipViewerApp(root)
    root.mainloop()