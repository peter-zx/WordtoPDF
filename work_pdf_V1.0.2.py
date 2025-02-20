import os
from win32com.client import Dispatch, gencache
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Frame, Canvas, BooleanVar

class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Word转PDF工具")
        self.root.geometry("800x600")
        
        # 初始化变量
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.file_vars = []
        
        # 主布局
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 输入目录选择
        input_group = ttk.LabelFrame(main_frame, text="输入设置")
        input_group.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_group, text="文档目录:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(input_group, textvariable=self.input_dir, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(input_group, text="选择目录", command=self.select_input_dir).grid(row=0, column=2, padx=2)
        ttk.Button(input_group, text="刷新列表", command=self.show_file_list).grid(row=0, column=3, padx=2)
        
        # 输出目录选择
        output_group = ttk.LabelFrame(main_frame, text="输出设置")
        output_group.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_group, text="PDF目录:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(output_group, textvariable=self.output_dir, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(output_group, text="选择目录", command=self.select_output_dir).grid(row=0, column=2, padx=2)
        
        # 文件列表区域
        list_group = ttk.LabelFrame(main_frame, text="待转换文件列表（自动保持目录结构）")
        list_group.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 创建带滚动条的Canvas
        self.canvas = Canvas(list_group, borderwidth=0)
        scrollbar = ttk.Scrollbar(list_group, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = ttk.Frame(self.canvas)
        
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.canvas.create_window((0,0), window=self.scroll_frame, anchor="nw")
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        # 转换控制区域
        control_group = ttk.Frame(main_frame)
        control_group.pack(fill=tk.X, pady=5)
        
        self.progress = ttk.Progressbar(control_group, orient='horizontal', length=300, mode='determinate')
        self.progress.pack(side=tk.LEFT, expand=True)
        
        ttk.Button(control_group, text="开始转换", command=self.start_conversion).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_group, text="全选", command=self.select_all).pack(side=tk.RIGHT)
        ttk.Button(control_group, text="全不选", command=self.deselect_all).pack(side=tk.RIGHT)
        
        # 状态栏
        self.status = ttk.Label(main_frame, text="就绪", relief=tk.SUNKEN)
        self.status.pack(fill=tk.X)

    # 以下是修正后的方法定义（统一缩进）
    def select_input_dir(self):
        """选择输入目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.input_dir.set(dir_path)
            self.output_dir.set(dir_path + "_PDF")  # 自动设置输出目录
            self.show_file_list()

    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir.set(dir_path)

    def show_file_list(self):
        """显示文件列表"""
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.file_vars.clear()
        
        input_dir = self.input_dir.get()
        if not input_dir or not os.path.exists(input_dir):
            return
        
        file_list = []
        for root_dir, _, files in os.walk(input_dir):
            for file in files:
                if file.lower().endswith(('.doc', '.docx', '.docm','.dotx','.dotm','.odt','.rtf')):
                    rel_path = os.path.relpath(root_dir, input_dir)
                    display_path = os.path.join(rel_path, file)
                    file_list.append((display_path, os.path.join(root_dir, file)))
        
        for display_path, full_path in file_list:
            var = BooleanVar(value=True)
            self.file_vars.append((full_path, var))
            cb = ttk.Checkbutton(self.scroll_frame, 
                                text=f"{display_path}", 
                                variable=var)
            cb.pack(anchor=tk.W, fill=tk.X)

    def select_all(self):
        """全选文件"""
        for _, var in self.file_vars:
            var.set(True)

    def deselect_all(self):
        """全不选文件"""
        for _, var in self.file_vars:
            var.set(False)

    def start_conversion(self):
        """开始转换"""
        input_dir = self.input_dir.get()
        output_dir = self.output_dir.get()
        
        if not all([input_dir, output_dir]):
            messagebox.showwarning("警告", "请先选择输入和输出目录！")
            return
            
        selected_files = [path for path, var in self.file_vars if var.get()]
        if not selected_files:
            messagebox.showwarning("警告", "请至少选择一个要转换的文件！")
            return

        total = len(selected_files)
        success = 0
        errors = []
        
        self.progress["value"] = 0
        self.status["text"] = "开始转换..."
        
        for idx, file_path in enumerate(selected_files, 1):
            try:
                # 生成保持结构的输出路径
                rel_path = os.path.relpath(file_path, input_dir)
                pdf_path = os.path.join(output_dir, os.path.splitext(rel_path)[0] + ".pdf")
                
                # 自动创建目录
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
                
                if convert_file(file_path, pdf_path):
                    success += 1
                else:
                    errors.append(os.path.basename(file_path))
            except Exception as e:
                errors.append(f"{os.path.basename(file_path)} ({str(e)})")
            
            # 更新进度
            self.progress["value"] = (idx / total) * 100
            self.status["text"] = f"正在转换 {os.path.basename(file_path)}... ({idx}/{total})"
            self.root.update_idletasks()
        
        # 显示结果
        self.progress["value"] = 0
        result_msg = f"转换完成！成功 {success} 个，失败 {len(errors)} 个"
        if errors:
            result_msg += "\n失败文件：\n" + "\n".join(errors)
        self.status["text"] = result_msg
        
        if success > 0:
            if messagebox.askyesno("完成", "是否打开输出目录？"):
                os.startfile(output_dir)

def convert_file(doc_path, pdf_path):
    """转换单个文件"""
    word = init_word_app()
    try:
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        return True
    except Exception as e:
        print(f"转换失败：{doc_path} -> {str(e)}")
        return False

def init_word_app():
    """初始化Word应用（单例模式）"""
    if not hasattr(init_word_app, "instance"):
        gencache.EnsureDispatch('Word.Application')
        init_word_app.instance = Dispatch('Word.Application')
        init_word_app.instance.Visible = False
    return init_word_app.instance

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterGUI(root)
    root.mainloop()
