# -*- coding: utf-8 -*-
"""主窗口UI模块"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import traceback
from core.logger import logger

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")


class MainWindow:
    """主窗口"""

    def __init__(self, root):
        self.root = root
        self.root.title("文档整理工具 v1.0")
        
        # 设置窗口大小和最小尺寸，确保三栏布局不被遮盖
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 自适应窗口大小，但确保最小尺寸
        window_width = min(1200, screen_width - 100)
        window_height = min(900, screen_height - 100)
        
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.minsize(1000, 700)  # 设置最小尺寸避免布局被破坏
        self.root.resizable(True, True)
        
        # 配置根窗口的网格权重，确保子组件可以扩展
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # 核心组件引用（由外部注入）
        self.excel_parser = None
        self.folder_manager = None
        self.file_copier = None
        self.word_converter = None

        # UI变量
        self.excel_path_var = tk.StringVar()
        self.source_folder_var = tk.StringVar()
        self.file_types_var = tk.StringVar(value=".docx .doc .pdf")
        self.project_name_var = tk.StringVar(value="文档整理")
        self.target_folder_var = tk.StringVar(value=DESKTOP)
        self.selected_folder_var = tk.StringVar(value="(请先选择目标文件夹)")
        self.folder_count_var = tk.StringVar(value="选中文件夹: 0")
        self.selected_count_var = tk.StringVar(value="已选择: 0 个文件")
        self.selected_info_var = tk.StringVar(value="选中: 0/0")  # 新增选中数量显示

        # 树形组件引用
        self.folder_tree = None
        self.file_tree = None
        self.result_text = None
        
        # 文件列表item映射
        self.file_item_map = {}  # item_id -> file index
        
        # 文件夹树选择状态管理
        self.folder_item_map = {}  # item_id -> folder_path

        # 目标根路径
        self.target_root_path = ""

    def set_components(self, excel_parser, folder_manager, file_copier, word_converter):
        """设置核心组件"""
        self.excel_parser = excel_parser
        self.folder_manager = folder_manager
        self.file_copier = file_copier
        self.word_converter = word_converter

    def setup_ui(self):
        """设置UI"""
        # 标题
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(title_frame, text="文档整理工具", font=("微软雅黑", 20, "bold")).pack()

        # 主框架 - 三栏网格布局
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        # 配置三列等宽布局
        main_frame.columnconfigure(0, weight=1, uniform="main_col")
        main_frame.columnconfigure(1, weight=1, uniform="main_col")
        main_frame.columnconfigure(2, weight=1, uniform="main_col")

        # 左侧
        self._setup_left_panel(main_frame)
        # 中间
        self._setup_mid_panel(main_frame)
        # 右侧
        self._setup_right_panel(main_frame)

        # 底部
        self._setup_bottom_panel()

        # 结果区
        self._setup_result_panel()

    def _setup_left_panel(self, parent):
        """左侧面板"""
        left_frame = ttk.LabelFrame(parent, text="第1步: 选择Excel和源文件夹", padding=15)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left_frame.columnconfigure(0, weight=1)
        
        # 设置最小宽度避免内容被遮盖
        left_frame.configure(width=300)

        # Excel
        tk.Label(left_frame, text="1. 选择Excel表格:", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(0, 5))
        excel_frame = ttk.Frame(left_frame)
        excel_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Entry(excel_frame, textvariable=self.excel_path_var, width=35).pack(side=tk.LEFT, padx=5)
        tk.Button(excel_frame, text="选择Excel", bg="#2196F3", fg="white",
                  command=self._on_select_excel).pack(side=tk.LEFT)

        # 源文件夹
        tk.Label(left_frame, text="2. 选择源文件夹:", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        source_frame = ttk.Frame(left_frame)
        source_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Entry(source_frame, textvariable=self.source_folder_var, width=35).pack(side=tk.LEFT, padx=5)
        tk.Button(source_frame, text="选择文件夹", bg="#2196F3", fg="white",
                  command=self._on_select_source_folder).pack(side=tk.LEFT)

        # 文件类型
        filter_frame = ttk.Frame(left_frame)
        filter_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(filter_frame, text="文件类型:").pack(side=tk.LEFT)
        ttk.Entry(filter_frame, textvariable=self.file_types_var, width=15).pack(side=tk.LEFT, padx=5)

        # 项目名称
        tk.Label(left_frame, text="3. 项目名称:", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        project_frame = ttk.Frame(left_frame)
        project_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Entry(project_frame, textvariable=self.project_name_var, width=25).pack(side=tk.LEFT, padx=5)
        tk.Label(project_frame, text="(最外层文件夹)", foreground="gray").pack(side=tk.LEFT)

        # 目标位置
        tk.Label(left_frame, text="4. 保存位置:", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W, pady=(10, 5))
        target_frame = ttk.Frame(left_frame)
        target_frame.pack(fill=tk.X)
        ttk.Entry(target_frame, textvariable=self.target_folder_var, width=35).pack(side=tk.LEFT, padx=5)
        tk.Button(target_frame, text="选择位置", bg="#2196F3", fg="white",
                  command=self._on_select_target_folder).pack(side=tk.LEFT)

    def _setup_mid_panel(self, parent):
        """中间面板 - 文件夹树"""
        mid_frame = ttk.LabelFrame(parent, text="第2步: 选择目标文件夹", padding=15)
        mid_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10))
        mid_frame.columnconfigure(0, weight=1)
        
        # 设置最小宽度避免内容被遮盖
        mid_frame.configure(width=300)

        # 按钮
        btn_frame = ttk.Frame(mid_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Button(btn_frame, text="选择目标文件夹", bg="#2196F3", fg="white",
                  command=self._on_select_target_from_tree).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="从Excel创建结构", bg="#4CAF50", fg="white",
                  command=self._on_create_from_excel).pack(side=tk.LEFT, padx=5)

        # 树 - 使用简单的树形结构
        self.folder_tree = ttk.Treeview(mid_frame, show="tree", height=15)
        self.folder_tree.pack(fill=tk.BOTH, expand=True)
        
        # 绑定事件
        self.folder_tree.bind("<<TreeviewSelect>>", self._on_folder_select)

        scrollbar = ttk.Scrollbar(mid_frame, orient=tk.VERTICAL, command=self.folder_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.configure(yscrollcommand=scrollbar.set)

        # 显示
        tk.Label(mid_frame, text="当前选中:", foreground="blue").pack(anchor=tk.W, pady=(5, 0))
        tk.Label(mid_frame, textvariable=self.selected_folder_var, foreground="green").pack(anchor=tk.W)
        tk.Label(mid_frame, textvariable=self.folder_count_var, foreground="red", font=("微软雅黑", 10, "bold")).pack(anchor=tk.W, pady=(5, 0))

    def _setup_right_panel(self, parent):
        """右侧面板 - 文件列表"""
        right_frame = ttk.LabelFrame(parent, text="第3步: 勾选需要复制的文件", padding=15)
        right_frame.grid(row=0, column=2, sticky="nsew")
        right_frame.columnconfigure(0, weight=1)
        
        # 设置最小宽度避免内容被遮盖
        right_frame.configure(width=300)

        # 按钮 - 放在标题下面
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Button(btn_frame, text="全选", bg="#4CAF50", fg="white", font=("微软雅黑", 9),
                  command=self._on_select_all).pack(side=tk.LEFT, padx=3, ipadx=8)
        tk.Button(btn_frame, text="取消全选", bg="#f44336", fg="white", font=("微软雅黑", 9),
                  command=self._on_deselect_all).pack(side=tk.LEFT, padx=3, ipadx=8)
        tk.Button(btn_frame, text="反选", bg="#2196F3", fg="white", font=("微软雅黑", 9),
                  command=self._on_invert_selection).pack(side=tk.LEFT, padx=3, ipadx=8)
        
        # 选中数量显示
        self.selected_info_var = tk.StringVar(value="选中: 0/0")
        selected_label = tk.Label(btn_frame, textvariable=self.selected_info_var, 
                                foreground="blue", font=("微软雅黑", 10, "bold"))
        selected_label.pack(side=tk.RIGHT, anchor=tk.E)

        # 树 - 使用自定义列
        columns = ("选择", "文件名", "大小", "类型")
        self.file_tree = ttk.Treeview(right_frame, columns=columns, show="headings", height=12, selectmode="extended")
        self.file_tree.heading("选择", text="□")
        self.file_tree.heading("文件名", text="文件名")
        self.file_tree.heading("大小", text="大小")
        self.file_tree.heading("类型", text="类型")
        self.file_tree.column("选择", width=80, anchor="center")  # 加大选择框宽度
        self.file_tree.column("文件名", width=200)
        self.file_tree.column("大小", width=80, anchor="center")
        self.file_tree.column("类型", width=60, anchor="center")
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 绑定事件
        self.file_tree.bind("<Button-1>", self._on_file_click)
        self.file_tree.bind("<space>", self._on_space_press)
        self.file_tree.bind("<Return>", self._on_space_press)

        scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        # 说明文本
        help_text = "使用说明: 点击文件行任意位置选中文件 • 按空格键切换选中状态 • 支持多选和框选"
        help_label = tk.Label(right_frame, text=help_text, foreground="gray", font=("微软雅黑", 9))
        help_label.pack(anchor=tk.W, pady=(5, 0))

    def _setup_bottom_panel(self):
        """底部按钮 - 三栏卡片式布局"""
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=20, pady=15)

        # 使用grid布局确保三个按钮长度一致
        bottom_frame.columnconfigure(0, weight=1, uniform="button_col")
        bottom_frame.columnconfigure(1, weight=1, uniform="button_col")
        bottom_frame.columnconfigure(2, weight=1, uniform="button_col")

        # 按钮1：从Excel创建文件夹
        btn1 = tk.Button(bottom_frame, text="从Excel创建文件夹", font=("微软雅黑", 12, "bold"),
                        bg="#4CAF50", fg="white", height=2, command=self._on_create_from_excel)
        btn1.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        # 按钮2：复制文件到选中文件夹
        btn2 = tk.Button(bottom_frame, text="复制文件到选中文件夹", font=("微软雅黑", 12, "bold"),
                        bg="#FF9800", fg="white", height=2, command=self._on_copy_files)
        btn2.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        # 按钮3：Word转PDF
        btn3 = tk.Button(bottom_frame, text="Word转PDF", font=("微软雅黑", 12, "bold"),
                        bg="#E91E63", fg="white", height=2, command=self._on_show_pdf_dialog)
        btn3.grid(row=0, column=2, sticky="ew")

    def _setup_result_panel(self):
        """结果面板"""
        result_frame = ttk.LabelFrame(self.root, text="执行结果", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))

        self.result_text = tk.Text(result_frame, height=6, font=("微软雅黑", 10))
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=scrollbar.set)

    # ==================== 事件处理 ====================
    def _on_select_excel(self):
        path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if path:
            self.excel_path_var.set(path)
            try:
                logger.info(f"用户选择Excel文件: {path}")
                self.excel_parser.load_from_file(path)
                messagebox.showinfo("完成", "Excel加载成功！")
                logger.info("Excel加载成功")
            except Exception as e:
                error_msg = f"Excel加载失败: {str(e)}"
                logger.error(error_msg)
                messagebox.showerror("错误", error_msg)

    def _on_select_source_folder(self):
        path = filedialog.askdirectory(title="选择源文件夹")
        if path:
            logger.info(f"用户选择源文件夹: {path}")
            self.source_folder_var.set(path)
            self._scan_files(path)

    def _on_select_target_folder(self):
        path = filedialog.askdirectory(title="选择保存位置")
        if path:
            self.target_folder_var.set(path)

    def _on_select_target_from_tree(self):
        path = filedialog.askdirectory(title="选择目标文件夹")
        if path:
            self._load_folder_structure(path)

    def _on_create_from_excel(self):
        if not self.excel_parser.get_structure():
            if self.excel_path_var.get():
                try:
                    self.excel_parser.load_from_file(self.excel_path_var.get())
                except Exception as e:
                    messagebox.showerror("错误", f"加载Excel失败: {str(e)}")
                    return
            else:
                messagebox.showwarning("警告", "请先选择Excel文件！")
                return

        target = self.target_folder_var.get().strip()
        project = self.project_name_var.get().strip() or "文档整理"

        if project not in target:
            target = os.path.join(target, project)

        self.folder_manager.set_base_path(os.path.dirname(target))
        self.folder_manager.set_structure(self.excel_parser.get_structure())

        try:
            root_path, count = self.folder_manager.create_structure(project)
            self._load_folder_structure(root_path)
            self._log(f"✓ 已创建 {count} 个文件夹")
            messagebox.showinfo("完成", f"已创建 {count} 个文件夹！")
        except Exception as e:
            messagebox.showerror("错误", f"创建失败: {str(e)}")

    def _on_folder_select(self, event):
        """文件夹选中事件"""
        selection = self.folder_tree.selection()
        if selection:
            item = selection[0]
            folder_path = self.folder_item_map.get(item, "")
            depth = folder_path.replace(self.target_root_path, '').count(os.sep) if self.target_root_path else 0
            self.selected_folder_var.set(f"当前选中: {os.path.basename(folder_path)}")
            self.folder_count_var.set(f"层级: {depth + 1} 级 | 路径: {folder_path}")

    def _load_folder_structure(self, folder_path):
        """加载文件夹结构"""
        # 清空树和映射
        self.folder_item_map.clear()
        for item in self.folder_tree.get_children():
            self.folder_tree.delete(item)

        root_name = os.path.basename(folder_path)
        root_item = self.folder_tree.insert("", tk.END, text=f"📁 {root_name}", open=True)
        self.folder_item_map[root_item] = folder_path

        count = self._add_folders_to_tree(root_item, folder_path)

        self.target_root_path = folder_path
        self.selected_folder_var.set(f"已选择: 0 个文件夹")
        self.folder_count_var.set(f"共 {count + 1} 个文件夹")

    def _add_folders_to_tree(self, parent_item, folder_path):
        """递归添加文件夹"""
        count = 0
        try:
            items = os.listdir(folder_path)
            for item in sorted(items):
                item_path = os.path.join(folder_path, item)
                if os.path.isdir(item_path):
                    count += 1
                    child = self.folder_tree.insert(parent_item, tk.END, text=f"📁 {item}", open=True)
                    self.folder_item_map[child] = item_path
                    count += self._add_folders_to_tree(child, item_path)
        except PermissionError:
            pass
        return count

    def _scan_files(self, folder_path):
        """扫描文件"""
        file_types = self.file_types_var.get().split()
        try:
            logger.info(f"开始扫描文件，类型: {file_types}")
            files = self.file_copier.scan_folder(folder_path, file_types)

            # 清空树和映射
            self.file_item_map.clear()
            for item in self.file_tree.get_children():
                self.file_tree.delete(item)

            # 添加文件到树，保存item_id到index的映射
            for i, f in enumerate(files):
                item_id = self.file_tree.insert("", tk.END, values=("□", f["name"], f["size_str"], f["type"]))
                self.file_item_map[item_id] = i

            self._update_file_count()
            success_msg = f"共扫描到 {len(files)} 个文件"
            messagebox.showinfo("完成", success_msg)
            logger.info(success_msg)
        except Exception as e:
            error_msg = f"扫描失败: {str(e)}"
            logger.error(error_msg)
            messagebox.showerror("错误", error_msg)

    def _on_file_click(self, event):
        """处理文件列表点击 - 简化版本"""
        # 获取点击的行
        item_id = self.file_tree.identify_row(event.y)
        if not item_id:
            return
        
        # 获取点击的列
        column = self.file_tree.identify_column(event.x)
        
        # 如果点击的是第一列（选择列），直接切换状态
        if column == "#1":
            self._toggle_item(item_id)
        else:
            # 点击其他列时，如果当前行未被选中，则选中它
            if item_id not in self.file_tree.selection():
                self.file_tree.selection_set(item_id)
        
        # 不阻止默认行为，让Treeview处理选中高亮

    def _on_space_press(self, event):
        """空格键切换选中状态"""
        # 获取当前选中的所有行
        selected_items = self.file_tree.selection()
        if selected_items:
            for item_id in selected_items:
                self._toggle_item(item_id)
        return "break"

    def _toggle_item(self, item_id):
        """切换单个项目的选中状态"""
        if item_id not in self.file_item_map:
            return
            
        file_idx = self.file_item_map[item_id]
        
        # 获取当前状态
        current = self.file_tree.set(item_id, column="#1")
        new_state = (current != "✓")
        
        # 更新显示
        new_display = "✓" if new_state else "□"
        self.file_tree.set(item_id, column="#1", value=new_display)
        
        # 更新文件选择状态
        if 0 <= file_idx < len(self.file_copier.scanned_files):
            self.file_copier.scanned_files[file_idx]["selected"] = new_state
            self.file_copier.check_states[file_idx] = new_state
        
        self._update_file_count()

    def _on_select_all(self):
        self.file_copier.select_all()
        for item_id, idx in self.file_item_map.items():
            self.file_tree.set(item_id, column="#1", value="✓")
        self._update_file_count()

    def _on_deselect_all(self):
        self.file_copier.deselect_all()
        for item_id in self.file_item_map:
            self.file_tree.set(item_id, column="#1", value="□")
        self._update_file_count()

    def _on_invert_selection(self):
        """反选"""
        for item_id, file_idx in self.file_item_map.items():
            current = self.file_tree.set(item_id, column="#1")
            new_state = (current != "✓")
            
            self.file_tree.set(item_id, column="#1", value="✓" if new_state else "□")
            
            if 0 <= file_idx < len(self.file_copier.scanned_files):
                self.file_copier.scanned_files[file_idx]["selected"] = new_state
                self.file_copier.check_states[file_idx] = new_state
        
        self._update_file_count()

    def _update_file_count(self):
        count = self.file_copier.get_selected_count()
        total = len(self.file_copier.scanned_files)
        self.selected_count_var.set(f"已选择: {count} 个文件")
        self.selected_info_var.set(f"选中: {count}/{total}")

    def _on_copy_files(self):
        """复制文件到选中的文件夹"""
        files = self.file_copier.get_selected_files()
        if not files:
            warning_msg = "请先选择需要复制的文件！"
            logger.warning(warning_msg)
            messagebox.showwarning("警告", warning_msg)
            return

        # 获取当前选中的文件夹（单选）
        selection = self.folder_tree.selection()
        if not selection:
            warning_msg = "请先选择目标文件夹！点击文件夹进行选择。"
            logger.warning(warning_msg)
            messagebox.showwarning("警告", warning_msg)
            return

        selected_folder = self.folder_item_map.get(selection[0])
        if not selected_folder or not os.path.exists(selected_folder):
            error_msg = f"目标文件夹不存在: {selected_folder}"
            logger.error(error_msg)
            messagebox.showerror("错误", error_msg)
            return

        logger.info(f"开始复制操作，选中 {len(files)} 个文件到: {selected_folder}")

        # 执行复制
        results, success, fail = self.file_copier.copy_to_selected_leaf_folders(files, selected_folder)
        
        # 记录操作结果
        result_log = f"复制操作完成: 成功 {success}, 失败 {fail}"
        logger.info(result_log)
        
        self._log("\n".join(results))
        messagebox.showinfo("完成", f"复制成功: {success}, 失败: {fail}")

    def _get_all_folders(self, tree_item):
        """获取树中所有文件夹"""
        folders = []

        def get_children(parent):
            children = self.folder_tree.get_children(parent)
            for child in children:
                values = self.folder_tree.item(child, "values")
                if values:
                    folders.append(values[0])
                get_children(child)

        get_children(tree_item)
        return folders

    def _on_show_pdf_dialog(self):
        """显示PDF转换对话框 - 简化可靠版本"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Word转PDF")
        
        # 简化界面尺寸
        dialog.geometry("700x600")
        dialog.minsize(600, 400)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 配置网格权重
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=0)  # 标题
        dialog.rowconfigure(1, weight=0)  # 设置
        dialog.rowconfigure(2, weight=0)  # 按钮
        dialog.rowconfigure(3, weight=1)  # 结果

        # 标题
        title_frame = ttk.Frame(dialog)
        title_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=10)
        tk.Label(title_frame, text="Word转PDF转换", font=("微软雅黑", 16, "bold"), 
                foreground="#E91E63").pack()

        # 设置区域
        settings_frame = ttk.LabelFrame(dialog, text="转换设置", padding=15)
        settings_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        settings_frame.columnconfigure(0, weight=1)

        # 源文件夹
        source_var = tk.StringVar()
        tk.Label(settings_frame, text="源文件夹:", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=5)
        
        source_frame = ttk.Frame(settings_frame)
        source_frame.grid(row=1, column=0, sticky="ew", pady=5)
        source_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(source_frame, textvariable=source_var, width=50).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        tk.Button(source_frame, text="选择", bg="#2196F3", fg="white",
                  command=lambda: self._select_folder(source_var)).grid(row=0, column=1)

        # 输出文件夹
        output_var = tk.StringVar(value=DESKTOP)
        tk.Label(settings_frame, text="输出文件夹:", font=("微软雅黑", 10)).grid(row=2, column=0, sticky="w", pady=5)
        
        output_frame = ttk.Frame(settings_frame)
        output_frame.grid(row=3, column=0, sticky="ew", pady=5)
        output_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_frame, textvariable=output_var, width=50).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        tk.Button(output_frame, text="选择", bg="#2196F3", fg="white",
                  command=lambda: self._select_folder(output_var)).grid(row=0, column=1)

        # 选项
        keep_var = tk.BooleanVar(value=True)
        auto_wrap_var = tk.BooleanVar(value=True)
        
        option_frame = ttk.Frame(settings_frame)
        option_frame.grid(row=4, column=0, sticky="w", pady=10)
        
        ttk.Checkbutton(option_frame, text="保持文件夹结构", variable=keep_var).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(option_frame, text="自动创建顶层文件夹", variable=auto_wrap_var).pack(side=tk.LEFT)

        # 转换按钮 - 绝对可见
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        button_frame.columnconfigure(0, weight=1)
        
        # 醒目的转换按钮
        convert_btn = tk.Button(button_frame, text="🚀 开始转换", bg="#E91E63", fg="white", 
                               font=("微软雅黑", 12, "bold"), height=2,
                               command=lambda: self._convert_pdf_check(dialog, source_var, output_var, 
                                                                       keep_var.get(), auto_wrap_var.get()))
        convert_btn.pack(fill=tk.X, expand=True)

        # 状态显示
        status_var = tk.StringVar(value="准备就绪")
        status_label = tk.Label(button_frame, textvariable=status_var, foreground="blue", 
                               font=("微软雅黑", 9))
        status_label.pack(pady=5)
        
        # 结果区域
        result_frame = ttk.LabelFrame(dialog, text="转换结果", padding=10)
        result_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        result_text = tk.Text(result_frame, font=("微软雅黑", 9), wrap=tk.WORD)
        result_text.grid(row=0, column=0, sticky="nsew")
        
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=result_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        result_text.configure(yscrollcommand=scrollbar.set)
        
        dialog.result_text = result_text
        dialog.status_var = status_var

    def _select_folder(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _convert_pdf_check(self, dialog, source_var, output_var, keep, auto_wrap):
        """转换前检查"""
        source = source_var.get().strip()
        output = output_var.get().strip()
        
        # 检查源文件夹
        if not source:
            dialog.status_var.set("❌ 请选择源文件夹")
            messagebox.showwarning("警告", "请选择源文件夹！")
            return
            
        if not os.path.exists(source):
            dialog.status_var.set("❌ 源文件夹不存在")
            messagebox.showerror("错误", f"源文件夹不存在: {source}")
            return
            
        # 检查Word转换器是否可用
        if not self.word_converter.is_available():
            dialog.status_var.set("❌ Word转换器不可用")
            messagebox.showerror("错误", "Word转换器不可用，请检查pywin32和comtypes库的安装")
            return
            
        dialog.status_var.set("🔄 正在转换中...")
        self._convert_pdf(dialog, source, output, keep, auto_wrap)

    def _convert_pdf(self, dialog, source, output, keep, auto_wrap):
        """执行PDF转换"""
        try:
            results, success, fail = self.word_converter.convert_folder(source, output, keep, auto_wrap)
            dialog.result_text.delete(1.0, tk.END)
            dialog.result_text.insert(tk.END, "\n".join(results))
            
            # 滚动到顶部
            dialog.result_text.see("1.0")
            
            dialog.status_var.set(f"✅ 转换完成: 成功{success}, 失败{fail}")
            messagebox.showinfo("完成", f"转换成功: {success}, 失败: {fail}")
            
        except ImportError:
            dialog.status_var.set("❌ 依赖库缺失")
            messagebox.showerror("错误", "需要安装 pywin32 和 comtypes 库！")
        except FileNotFoundError as e:
            dialog.status_var.set("❌ 文件路径错误")
            messagebox.showerror("错误", f"路径错误: {str(e)}")
        except Exception as e:
            dialog.status_var.set("❌ 转换失败")
            error_msg = f"转换失败: {str(e)}"
            logger.error(error_msg)
            messagebox.showerror("错误", error_msg)

    def _log(self, message):
        """显示日志"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, message)
