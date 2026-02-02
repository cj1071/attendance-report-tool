#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
员工工时报表生成工具 - 苹果风格GUI界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import threading
from datetime import datetime
from excel_report_generator_fixed import ExcelReportGenerator
from run_attendance_stats import generate_attendance_stats

class ModernButton(tk.Button):
    """现代化按钮样式 - 兼容macOS"""
    def __init__(self, parent, **kwargs):
        # 提取自定义参数
        bg_color = kwargs.pop('bg', '#007AFF')
        hover_color = kwargs.pop('hover_color', None)
        
        # 默认样式
        default_style = {
            'font': ('PingFang SC', 13),
            'bg': bg_color,
            'fg': 'white',
            'relief': 'flat',
            'bd': 0,
            'padx': 20,
            'pady': 8,
            'cursor': 'hand2',
            'activebackground': hover_color if hover_color else self._darken_color(bg_color),
            'activeforeground': 'white',
            'highlightthickness': 0
        }
        default_style.update(kwargs)
        super().__init__(parent, **default_style)
        
        # 保存颜色用于悬停效果
        self.original_bg = bg_color
        self.hover_bg = hover_color if hover_color else self._darken_color(bg_color)
        
        # 悬停效果
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
    
    def _darken_color(self, hex_color):
        """将颜色变暗用于悬停效果"""
        # 移除 # 号
        hex_color = hex_color.lstrip('#')
        # 转换为 RGB
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        # 变暗 20%
        r, g, b = int(r * 0.8), int(g * 0.8), int(b * 0.8)
        # 转回十六进制
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def _on_enter(self, event):
        self.config(bg=self.hover_bg)
    
    def _on_leave(self, event):
        self.config(bg=self.original_bg)

class ProgressWindow:
    """进度窗口"""
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("处理中...")
        self.window.geometry("400x150")
        self.window.resizable(False, False)
        self.window.configure(bg='#F2F2F7')

        # 居中显示
        self.window.transient(parent)
        self.window.grab_set()

        # 计算居中位置
        parent.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - 200
        y = parent.winfo_y() + (parent.winfo_height() // 2) - 75
        self.window.geometry(f"400x150+{x}+{y}")
        
        # 进度条
        self.progress = ttk.Progressbar(
            self.window, 
            mode='indeterminate',
            length=300
        )
        self.progress.pack(pady=30)
        
        # 状态标签
        self.status_label = tk.Label(
            self.window,
            text="正在处理Excel文件...",
            font=('SimSun', 12),
            bg='#F2F2F7',
            fg='#1C1C1E'
        )
        self.status_label.pack(pady=10)
        
        self.progress.start(10)
    
    def update_status(self, text):
        self.status_label.config(text=text)
        self.window.update()
    
    def close(self):
        self.progress.stop()
        self.window.destroy()

class ExcelReportApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("员工工时报表生成工具")
        self.root.geometry("800x700")
        self.root.configure(bg='#F2F2F7')
        self.root.resizable(True, True)  # 允许拉伸
        self.root.minsize(900, 680)      # 设置最小尺寸
        
        # 设置图标（如果有的话）
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        self.selected_file = None
        self.output_dir = os.getcwd()
        self.generated_work_hours_files = []  # 保存生成的工时报表文件路径
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 创建主容器框架
        container = tk.Frame(self.root, bg='#F2F2F7')
        container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # 创建Canvas和Scrollbar
        canvas = tk.Canvas(container, bg='#F2F2F7', highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient='vertical', command=canvas.yview)
        
        # 创建可滚动的框架
        scrollable_frame = tk.Frame(canvas, bg='#F2F2F7')
        
        # 绑定滚动事件
        scrollable_frame.bind(
            '<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox('all'))
        )
        
        # 创建窗口
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 确保可滚动框架宽度匹配Canvas宽度
        def _configure_canvas_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        canvas.bind('<Configure>', _configure_canvas_width)
        
        # 布局Canvas和Scrollbar
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)  # Windows
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))  # Linux
        
        # 使用scrollable_frame作为主框架
        main_frame = scrollable_frame

        # 主标题
        title_label = tk.Label(
            main_frame,
            text="📊 员工工时报表生成工具",
            font=('SimSun', 18, 'bold'),
            bg='#F2F2F7',
            fg='#1C1C1E'
        )
        title_label.pack(pady=(15, 8))

        # 副标题
        subtitle_label = tk.Label(
            main_frame,
            text="将劳务签到表转换为按公司分组的月度考勤报表",
            font=('SimSun', 12),
            bg='#F2F2F7',
            fg='#8E8E93'
        )
        subtitle_label.pack(pady=(0, 25))
        
        # 文件选择区域
        file_frame = tk.Frame(main_frame, bg='#F2F2F7')
        file_frame.pack(pady=15, padx=20, fill='x')

        file_label = tk.Label(
            file_frame,
            text="📁 选择Excel文件",
            font=('SimSun', 14, 'bold'),
            bg='#F2F2F7',
            fg='#1C1C1E'
        )
        file_label.pack(anchor='w', pady=(0, 8))
        
        # 文件选择按钮和显示
        file_select_frame = tk.Frame(file_frame, bg='#F2F2F7')
        file_select_frame.pack(fill='x')
        
        self.file_display = tk.Label(
            file_select_frame,
            text="未选择文件",
            font=('SimSun', 11),
            bg='white',
            fg='#8E8E93',
            relief='solid',
            bd=1,
            padx=12,
            pady=8,
            anchor='w'
        )
        self.file_display.pack(side='left', fill='x', expand=True, padx=(0, 8))

        select_btn = ModernButton(
            file_select_frame,
            text="选择文件",
            font=('SimSun', 11),
            width=10,
            command=self.select_file
        )
        select_btn.pack(side='right')
        
        # 输出目录区域
        output_frame = tk.Frame(main_frame, bg='#F2F2F7')
        output_frame.pack(pady=15, padx=20, fill='x')

        output_label = tk.Label(
            output_frame,
            text="📂 输出目录",
            font=('SimSun', 14, 'bold'),
            bg='#F2F2F7',
            fg='#1C1C1E'
        )
        output_label.pack(anchor='w', pady=(0, 8))

        output_select_frame = tk.Frame(output_frame, bg='#F2F2F7')
        output_select_frame.pack(fill='x')

        self.output_display = tk.Label(
            output_select_frame,
            text=self.output_dir,
            font=('SimSun', 11),
            bg='white',
            fg='#1C1C1E',
            relief='solid',
            bd=1,
            padx=12,
            pady=8,
            anchor='w'
        )
        self.output_display.pack(side='left', fill='x', expand=True, padx=(0, 8))

        output_btn = ModernButton(
            output_select_frame,
            text="选择目录",
            font=('SimSun', 11),
            bg='#34C759',
            width=10,
            command=self.select_output_dir
        )
        output_btn.pack(side='right')
        
        # 功能特性展示
        features_frame = tk.Frame(main_frame, bg='#F2F2F7')
        features_frame.pack(pady=15, padx=20, fill='x')

        features_label = tk.Label(
            features_frame,
            text="✨ 功能特性",
            font=('SimSun', 14, 'bold'),
            bg='#F2F2F7',
            fg='#1C1C1E'
        )
        features_label.pack(anchor='w', pady=(0, 8))

        features_text = [
            "• 智能处理跨年数据",
            "• 多次签到动态扩展列",
            "• 斑马纹区分不同员工",
            "• 按原始数据顺序排列",
            "• 自动按公司分组生成报表"
        ]

        for feature in features_text:
            feature_label = tk.Label(
                features_frame,
                text=feature,
                font=('SimSun', 11),
                bg='#F2F2F7',
                fg='#8E8E93'
            )
            feature_label.pack(anchor='w', pady=1)
        
        # 按钮区域
        button_frame = tk.Frame(main_frame, bg='#F2F2F7')
        button_frame.pack(pady=25, padx=20)

        # 统一的按钮样式参数
        button_style = {
            'padx': 30,
            'pady': 10,
            'width': 16
        }

        # 第一行按钮
        button_row1 = tk.Frame(button_frame, bg='#F2F2F7')
        button_row1.pack(pady=(0, 12))

        # 🔵 生成工时报表按钮（蓝色 - 主要功能）
        generate_btn = ModernButton(
            button_row1,
            text="生成工时报表",
            bg='#007AFF',
            command=self.generate_reports,
            **button_style
        )
        generate_btn.pack(side='left', padx=6)

        # 🟠 生成考勤统计按钮（橙色 - 辅助功能）
        self.stats_btn = ModernButton(
            button_row1,
            text="生成考勤统计",
            bg='#FF9500',
            command=self.generate_attendance_stats,
            **button_style
        )
        self.stats_btn.pack(side='left', padx=6)

        # 第二行按钮
        button_row2 = tk.Frame(button_frame, bg='#F2F2F7')
        button_row2.pack()

        # 🟢 一键生成全部按钮（绿色 - 快捷功能）
        self.all_btn = ModernButton(
            button_row2,
            text="一键生成全部",
            bg='#34C759',
            command=self.generate_all_reports,
            **button_style
        )
        self.all_btn.pack(side='left', padx=6)

        # 🟣 打开文件夹按钮（紫色 - 辅助功能）
        self.open_folder_btn = ModernButton(
            button_row2,
            text="打开输出文件夹",
            bg='#5856D6',
            command=self.open_output_folder,
            **button_style
        )
        self.open_folder_btn.pack(side='left', padx=6)

        # 状态栏
        status_frame = tk.Frame(self.root, bg='#F2F2F7')
        status_frame.pack(side='bottom', fill='x', pady=5)

        self.status_label = tk.Label(
            status_frame,
            text="准备就绪",
            font=('SimSun', 10),
            bg='#F2F2F7',
            fg='#8E8E93'
        )
        self.status_label.pack(pady=5)
    
    def select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file = file_path
            filename = os.path.basename(file_path)
            self.file_display.config(text=filename, fg='#1C1C1E')
            self.status_label.config(text=f"已选择文件: {filename}")
    
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.output_dir
        )

        if dir_path:
            self.output_dir = dir_path
            self.output_display.config(text=dir_path)
            self.status_label.config(text=f"输出目录: {dir_path}")

    def open_output_folder(self):
        """打开输出文件夹"""
        try:
            import subprocess
            import platform

            if platform.system() == "Windows":
                os.startfile(self.output_dir)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", self.output_dir])
            else:  # Linux
                subprocess.run(["xdg-open", self.output_dir])

            self.status_label.config(text=f"已打开文件夹: {self.output_dir}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {e}")
    
    def generate_reports(self):
        """生成报表"""
        if not self.selected_file:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        
        if not os.path.exists(self.selected_file):
            messagebox.showerror("错误", "选择的文件不存在")
            return
        
        # 在新线程中执行生成任务
        thread = threading.Thread(target=self._generate_reports_thread)
        thread.daemon = True
        thread.start()
    
    def _generate_reports_thread(self):
        """在后台线程中生成报表"""
        progress_window = None
        
        try:
            # 显示进度窗口
            self.root.after(0, lambda: self._show_progress())
            
            # 创建进度窗口
            progress_window = ProgressWindow(self.root)
            
            # 更新状态
            progress_window.update_status("正在读取Excel文件...")
            
            # 创建报表生成器
            generator = ExcelReportGenerator()
            generator.read_input_excel(self.selected_file)
            
            if not generator.raw_data:
                raise Exception("没有读取到有效数据")
            
            # 生成报表
            generated_files = []
            companies = sorted(generator.companies)
            
            for i, company in enumerate(companies):
                progress_window.update_status(f"正在生成 {company} 的工时报表... ({i+1}/{len(companies)})")
                
                report_info = generator.generate_company_report(company)
                if report_info:
                    filepath = generator.save_company_report(report_info, self.output_dir)
                    generated_files.append(filepath)
            
            # 保存生成的文件路径（用于后续生成考勤统计）
            self.generated_work_hours_files = generated_files
            
            # 关闭进度窗口
            progress_window.close()
            
            # 显示成功消息
            success_msg = f"✅ 工时报表生成完成!\n\n共生成 {len(generated_files)} 个文件:\n"
            for filepath in generated_files:
                success_msg += f"• {os.path.basename(filepath)}\n"
            success_msg += f"\n📁 保存位置: {self.output_dir}"
            
            messagebox.showinfo("成功", success_msg)
            
            # 更新状态
            self.root.after(0, lambda: self.status_label.config(text=f"已生成 {len(generated_files)} 个工时报表文件"))
            
        except Exception as e:
            if progress_window:
                progress_window.close()
            
            error_msg = f"生成报表时发生错误:\n\n{str(e)}"
            messagebox.showerror("错误", error_msg)
            
            self.root.after(0, lambda: self.status_label.config(text="生成失败"))
    
    def _show_progress(self):
        """显示进度状态"""
        self.status_label.config(text="正在生成报表...")
    
    def generate_attendance_stats(self):
        """生成考勤统计报表"""
        if not self.selected_file:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        
        if not os.path.exists(self.selected_file):
            messagebox.showerror("错误", "选择的文件不存在")
            return
        
        # 在新线程中执行生成任务
        thread = threading.Thread(target=self._generate_attendance_stats_thread)
        thread.daemon = True
        thread.start()
    
    def _generate_attendance_stats_thread(self):
        """在后台线程中生成考勤统计"""
        progress_window = None
        
        try:
            # 创建进度窗口
            progress_window = ProgressWindow(self.root)
            progress_window.update_status("正在生成考勤统计报表...")
            
            # 调用考勤统计生成函数
            generate_attendance_stats(self.selected_file, self.output_dir)
            
            # 关闭进度窗口
            progress_window.close()
            
            # 显示成功消息
            success_msg = f"✅ 考勤统计报表生成完成!\n\n📁 保存位置: {self.output_dir}"
            messagebox.showinfo("成功", success_msg)
            
            # 更新状态
            self.root.after(0, lambda: self.status_label.config(text="考勤统计报表已生成"))
            
        except Exception as e:
            if progress_window:
                progress_window.close()
            
            error_msg = f"生成考勤统计时发生错误:\n\n{str(e)}"
            messagebox.showerror("错误", error_msg)
            
            self.root.after(0, lambda: self.status_label.config(text="生成失败"))
    
    def generate_all_reports(self):
        """一键生成所有报表（工时报表 + 考勤统计）"""
        if not self.selected_file:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        
        if not os.path.exists(self.selected_file):
            messagebox.showerror("错误", "选择的文件不存在")
            return
        
        # 在新线程中执行生成任务
        thread = threading.Thread(target=self._generate_all_reports_thread)
        thread.daemon = True
        thread.start()
    
    def _generate_all_reports_thread(self):
        """在后台线程中生成所有报表"""
        progress_window = None
        
        try:
            # 创建进度窗口
            progress_window = ProgressWindow(self.root)
            
            # ===== 第一步：生成工时报表 =====
            progress_window.update_status("正在读取Excel文件...")
            
            generator = ExcelReportGenerator()
            generator.read_input_excel(self.selected_file)
            
            if not generator.raw_data:
                raise Exception("没有读取到有效数据")
            
            # 生成工时报表
            work_hours_files = []
            companies = sorted(generator.companies)
            
            for i, company in enumerate(companies):
                progress_window.update_status(f"[1/2] 正在生成 {company} 的工时报表... ({i+1}/{len(companies)})")
                
                report_info = generator.generate_company_report(company)
                if report_info:
                    filepath = generator.save_company_report(report_info, self.output_dir)
                    work_hours_files.append(filepath)
            
            # ===== 第二步：生成考勤统计 =====
            progress_window.update_status("[2/2] 正在生成考勤统计报表...")
            
            generate_attendance_stats(self.selected_file, self.output_dir)
            
            # 关闭进度窗口
            progress_window.close()
            
            # 显示成功消息
            success_msg = f"✅ 所有报表生成完成!\n\n"
            success_msg += f"📊 工时报表: {len(work_hours_files)} 个文件\n"
            success_msg += f"📈 考勤统计: {len(companies)} 个文件\n"
            success_msg += f"\n📁 保存位置: {self.output_dir}"
            
            messagebox.showinfo("成功", success_msg)
            
            # 更新状态
            total_files = len(work_hours_files) + len(companies)
            self.root.after(0, lambda: self.status_label.config(text=f"已生成 {total_files} 个报表文件"))
            
        except Exception as e:
            if progress_window:
                progress_window.close()
            
            error_msg = f"生成报表时发生错误:\n\n{str(e)}"
            messagebox.showerror("错误", error_msg)
            
            self.root.after(0, lambda: self.status_label.config(text="生成失败"))
    
    def run(self):
        """运行应用"""
        # 居中显示窗口
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
        
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelReportApp()
    app.run()
