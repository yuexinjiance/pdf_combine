import sys
import os
from PyPDF2 import PdfMerger
from docx2pdf import convert
import tkinter as tk
from tkinter import messagebox, Button, Label, filedialog, Frame
from tkinter import ttk, font
import threading
import glob

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title('越鑫证书转换和合并工具')
        self.root.configure(bg='#f0f0f0')  # 设置背景色
        
        # 创建自定义样式
        self.create_custom_style()
        
        # 创建选项卡控件
        self.notebook = ttk.Notebook(root, style='Custom.TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建两个选项卡页面 - 使用与图片中按钮相同的样式
        self.convert_tab = Frame(self.notebook, bg='#00BFA5')  # 使用青绿色背景
        self.merge_tab = Frame(self.notebook, bg='#00BFA5')    # 使用青绿色背景
        
        # 添加选项卡（使用Unicode字符作为图标）
        self.notebook.add(self.convert_tab, text='Word转PDF ')
        self.notebook.add(self.merge_tab, text='PDF合并 ')
        
        # 初始化两个选项卡的内容
        self.init_convert_tab()
        self.init_merge_tab()
        
        # 底部版权信息
        footer = Label(root, text='© 2025 越鑫证书转换和合并工具', font=('Arial', 8),
                      bg='#f0f0f0', fg='#999999')
        footer.pack(side=tk.BOTTOM, pady=5)
        
        # 绑定选项卡切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
    
    def create_custom_style(self):
        """创建自定义样式"""
        style = ttk.Style()
        
        # 配置Notebook样式 - 与图片中按钮样式保持一致
        style.configure('Custom.TNotebook', background='#f0f0f0')
        style.configure('Custom.TNotebook.Tab', padding=[20, 10], 
                       font=('Arial', 10, 'bold'), 
                       borderwidth=0,
                       relief='flat')
        
        # 配置选项卡样式 - 使用图片中的青绿色，但保持字体颜色为黑色
        style.map('Custom.TNotebook.Tab',
                 background=[('selected', '#00BFA5'), ('!selected', '#e0e0e0')],
                 foreground=[('selected', 'black'), ('!selected', 'black')])
        
        # 确保选项卡内容区域的背景色是正确的
        style.configure('TNotebook', background='#f0f0f0')
        style.configure('TFrame', background='#f0f0f0')
        
        # 配置进度条样式 - 也使用相同的青绿色
        style.configure("TProgressbar", thickness=25, background='#00BFA5')
    
    def init_convert_tab(self):
        # 设置主框架 - 确保背景色正确
        main_frame = Frame(self.convert_tab, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = Label(main_frame, text='Word文档转PDF', 
                           font=('Arial', 16, 'bold'), bg='#f0f0f0', fg='#333333')
        title_label.pack(pady=10)
        
        # 说明文字
        instruction = Label(main_frame, text='请选择包含Word文档的文件夹', 
                           font=('Arial', 10), bg='#f0f0f0', fg='#555555')
        instruction.pack(pady=5)
        
        # 文件夹选择框架
        folder_frame = Frame(main_frame, bg='#f0f0f0')
        folder_frame.pack(pady=10, fill=tk.X)
        
        self.convert_folder_path = tk.StringVar()
        self.convert_folder_display = Label(folder_frame, textvariable=self.convert_folder_path, 
                                   width=40, anchor='w', bg='#ffffff', relief='sunken',
                                   padx=5, pady=5)
        self.convert_folder_display.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        browse_button = Button(folder_frame, text='浏览...', command=self.browse_convert_folder,
                              bg='#e0e0e0', fg='#333333', padx=10,
                              activebackground='#d0d0d0', relief='raised')
        browse_button.pack(side=tk.RIGHT)
        
        # 进度条
        progress_frame = Frame(main_frame, bg='#f0f0f0')
        progress_frame.pack(pady=15, fill=tk.X)
        
        self.convert_progress = ttk.Progressbar(progress_frame, orient="horizontal", length=300, mode="determinate")
        self.convert_progress.pack(fill=tk.X)
        
        # 状态标签
        self.convert_status_label = Label(main_frame, text="准备就绪", bg='#f0f0f0', fg='#555555',
                                 font=('Arial', 9))
        self.convert_status_label.pack(pady=5)
        
        # 当前处理文件标签
        self.convert_current_file_label = Label(main_frame, text="", bg='#f0f0f0', fg='#555555',
                                      font=('Arial', 9))
        self.convert_current_file_label.pack(pady=5)
        
        # 按钮框架
        button_frame = Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(pady=10)
        
        # 使用与图片中相同的青绿色
        convert_button = Button(button_frame, text='开始转换', command=self.start_conversion,
                               bg='#00BFA5', fg='white', font=('Arial', 10, 'bold'),
                               padx=15, pady=8, relief='flat',
                               activebackground='#00A090')
        convert_button.pack()
        
        # 存储选择的文件夹路径
        self.convert_selected_folder = ""
    
    def init_merge_tab(self):
        # 设置主框架 - 确保背景色正确
        main_frame = Frame(self.merge_tab, bg='#f0f0f0', padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = Label(main_frame, text='PDF文件合并', 
                           font=('Arial', 16, 'bold'), bg='#f0f0f0', fg='#333333')
        title_label.pack(pady=10)
        
        # 说明文字
        instruction = Label(main_frame, text='请选择包含PDF文件的文件夹', 
                           font=('Arial', 10), bg='#f0f0f0', fg='#555555')
        instruction.pack(pady=5)
        
        # 文件夹选择框架
        folder_frame = Frame(main_frame, bg='#f0f0f0')
        folder_frame.pack(pady=10, fill=tk.X)
        
        self.merge_folder_path = tk.StringVar()
        self.merge_folder_display = Label(folder_frame, textvariable=self.merge_folder_path, 
                                   width=40, anchor='w', bg='#ffffff', relief='sunken',
                                   padx=5, pady=5)
        self.merge_folder_display.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        browse_button = Button(folder_frame, text='浏览...', command=self.browse_merge_folder,
                              bg='#e0e0e0', fg='#333333', padx=10,
                              activebackground='#d0d0d0', relief='raised')
        browse_button.pack(side=tk.RIGHT)
        
        # 进度条
        progress_frame = Frame(main_frame, bg='#f0f0f0')
        progress_frame.pack(pady=15, fill=tk.X)
        
        self.merge_progress = ttk.Progressbar(progress_frame, orient="horizontal", length=300, mode="determinate")
        self.merge_progress.pack(fill=tk.X)
        
        # 状态标签
        self.merge_status_label = Label(main_frame, text="准备就绪", bg='#f0f0f0', fg='#555555',
                                 font=('Arial', 9))
        self.merge_status_label.pack(pady=5)
        
        # 当前处理文件标签
        self.merge_current_file_label = Label(main_frame, text="", bg='#f0f0f0', fg='#555555',
                                      font=('Arial', 9))
        self.merge_current_file_label.pack(pady=5)
        
        # 按钮框架
        button_frame = Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(pady=10)
        
        # 使用与图片中相同的青绿色
        merge_button = Button(button_frame, text='开始合并', command=self.start_merge,
                               bg='#00BFA5', fg='white', font=('Arial', 10, 'bold'),
                               padx=15, pady=8, relief='flat',
                               activebackground='#00A090')
        merge_button.pack()
        
        # 存储选择的文件夹路径
        self.merge_selected_folder = ""
    
    def browse_convert_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含Word文档的文件夹")
        if folder_path:
            self.convert_selected_folder = folder_path
            # 显示文件夹路径，如果太长则截断
            if len(folder_path) > 40:
                display_path = "..." + folder_path[-37:]
            else:
                display_path = folder_path
            self.convert_folder_path.set(display_path)
            self.update_convert_status("已选择文件夹，点击转换按钮开始处理", 0)
            self.convert_current_file_label.config(text="")
    
    def browse_merge_folder(self):
        folder_path = filedialog.askdirectory(title="选择包含PDF文件的文件夹")
        if folder_path:
            self.merge_selected_folder = folder_path
            # 显示文件夹路径，如果太长则截断
            if len(folder_path) > 40:
                display_path = "..." + folder_path[-37:]
            else:
                display_path = folder_path
            self.merge_folder_path.set(display_path)
            self.update_merge_status("已选择文件夹，点击合并按钮开始处理", 0)
            self.merge_current_file_label.config(text="")
    
    def start_conversion(self):
        if not self.convert_selected_folder:
            messagebox.showwarning('错误', '请先选择文件夹!')
            return
        # 使用线程执行转换过程，避免界面卡死
        threading.Thread(target=self.convert_process, daemon=True).start()
    
    def start_merge(self):
        if not self.merge_selected_folder:
            messagebox.showwarning('错误', '请先选择文件夹!')
            return
        # 使用线程执行合并过程，避免界面卡死
        threading.Thread(target=self.merge_process, daemon=True).start()
    
    def update_convert_status(self, message, progress_value=None, current_file=None):
        # 更新状态标签和进度条
        self.convert_status_label.config(text=message)
        if progress_value is not None:
            self.convert_progress["value"] = progress_value
        if current_file is not None:
            self.convert_current_file_label.config(text=f"当前处理: {current_file}")
        self.root.update()
    
    def update_merge_status(self, message, progress_value=None, current_file=None):
        # 更新状态标签和进度条
        self.merge_status_label.config(text=message)
        if progress_value is not None:
            self.merge_progress["value"] = progress_value
        if current_file is not None:
            self.merge_current_file_label.config(text=f"当前处理: {current_file}")
        self.root.update()
    
    def convert_docx_to_pdf_with_progress(self, source_folder, target_folder):
        # 获取所有docx文件
        docx_files = []
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                if file.endswith('.docx') and not file.startswith('~$'):  # 排除临时文件
                    docx_files.append(os.path.join(root, file))
        
        total_files = len(docx_files)
        if total_files == 0:
            self.update_convert_status("未找到Word文档", 0)
            return
        
        # 创建目标文件夹
        os.makedirs(target_folder, exist_ok=True)
        
        # 逐个转换文件并更新进度
        for i, docx_file in enumerate(docx_files):
            file_name = os.path.basename(docx_file)
            relative_path = os.path.relpath(docx_file, source_folder)
            
            # 更新进度条和当前处理文件
            progress = (i / total_files * 100)
            self.update_convert_status(f"正在转换 ({i+1}/{total_files})", progress, file_name)
            
            # 转换单个文件
            try:
                # 创建目标文件的目录结构
                target_dir = os.path.dirname(os.path.join(target_folder, relative_path))
                os.makedirs(target_dir, exist_ok=True)
                
                # 使用docx2pdf转换单个文件
                convert(docx_file, os.path.join(target_folder, os.path.splitext(relative_path)[0] + '.pdf'))
            except Exception as e:
                self.update_convert_status(f"转换文件 {file_name} 时出错: {str(e)}", None)
    
    def convert_process(self):
        try:
            # 重置进度条
            self.convert_progress["value"] = 0
            
            # 检查文件夹是否存在
            if not os.path.exists(self.convert_selected_folder):
                raise FileNotFoundError(f"文件夹不存在")
            
            # 在选择的文件夹内创建PDF文件夹
            new_pdf_dir = os.path.join(self.convert_selected_folder, f'PDF版报告')
            
            # 更新状态 - 开始转换
            self.update_convert_status("正在扫描Word文件...", 0)
            
            # 使用自定义函数转换Word文件并显示进度
            self.convert_docx_to_pdf_with_progress(self.convert_selected_folder, new_pdf_dir)
            
            # 更新状态 - 转换完成
            self.update_convert_status("Word转换完成!", 100)
            self.convert_current_file_label.config(text="")
            messagebox.showinfo('完成', f"Word文档已转换为PDF，保存在:\n{os.path.abspath(new_pdf_dir)}")
            
            # 重置状态
            self.update_convert_status("准备就绪", 0)
            
        except Exception as e:
            self.update_convert_status(f"错误: {str(e)}", 0)
            self.convert_current_file_label.config(text="")
            messagebox.showerror('错误', f"操作失败：\n{str(e)}")
    
    def merge_process(self):
        try:
            # 重置进度条
            self.merge_progress["value"] = 0
            
            # 检查文件夹是否存在
            if not os.path.exists(self.merge_selected_folder):
                raise FileNotFoundError(f"文件夹不存在")
            
            # 获取文件夹名称
            folder_name = os.path.basename(self.merge_selected_folder)
            
            # 更新状态 - 开始合并
            self.update_merge_status("正在扫描PDF文件...", 10)
            
            # 合并PDF
            merger = PdfMerger()
            pdf_files = []
            for root, dirs, files in os.walk(self.merge_selected_folder):
                for file in files:
                    if file.endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))
            
            if not pdf_files:
                raise FileNotFoundError(f"在文件夹中没有找到PDF文件")
            
            # 更新进度条随着PDF合并进度
            total_files = len(pdf_files)
            for i, file in enumerate(pdf_files):
                file_name = os.path.basename(file)
                merger.append(file)
                progress = 10 + (i + 1) / total_files * 80  # 从10%到90%
                self.update_merge_status(f"正在合并PDF: {i+1}/{total_files}", progress, file_name)
                    
            # 最终PDF保存在选择的文件夹中
            output_file = os.path.join(self.merge_selected_folder, f"{folder_name}合并版.pdf")
            
            self.update_merge_status("正在写入最终PDF文件...", 90)
            merger.write(output_file)
            merger.close()
            
            self.update_merge_status("合并完成!", 100)
            self.merge_current_file_label.config(text="")
            messagebox.showinfo('完成', f"PDF文件已合并，保存为:\n{os.path.abspath(output_file)}")
            
            # 重置状态
            self.update_merge_status("准备就绪", 0)
            
        except Exception as e:
            self.update_merge_status(f"错误: {str(e)}", 0)
            self.merge_current_file_label.config(text="")
            messagebox.showerror('错误', f"操作失败：\n{str(e)}")
    
    def on_tab_changed(self, event):
        """当选项卡切换时调用此函数"""
        # 获取当前选中的选项卡
        current_tab = self.notebook.select()
        tab_id = self.notebook.index(current_tab)
        
        # 根据选中的选项卡设置背景色
        if tab_id == 0:  # Word转PDF选项卡
            self.convert_tab.configure(bg='#f0f0f0')
            for widget in self.convert_tab.winfo_children():
                if isinstance(widget, Frame):
                    widget.configure(bg='#f0f0f0')
        else:  # PDF合并选项卡
            self.merge_tab.configure(bg='#f0f0f0')
            for widget in self.merge_tab.winfo_children():
                if isinstance(widget, Frame):
                    widget.configure(bg='#f0f0f0')

if __name__ == '__main__':
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.geometry('550x500')  # 调整窗口大小以适应选项卡
    root.resizable(False, False)  # 禁止调整窗口大小
    root.mainloop()
