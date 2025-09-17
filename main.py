#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT转换为图片幻灯片工具 - 背景版本（最终GUI版）
使用Win32 COM接口设置背景，避免python-pptx的兼容性问题
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import tempfile
import shutil
from PIL import Image
import win32com.client
import threading
import queue

class PPTToImageSlidesGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PPT转图片幻灯片工具 - 背景版")
        self.root.geometry("700x700")
        self.root.resizable(True, True)
        
        # 消息队列用于线程间通信
        self.message_queue = queue.Queue()
        
        # 创建GUI界面
        self.create_widgets()
        
        # 设置关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 启动消息处理
        self.process_queue()
        
    def create_widgets(self):
        """创建GUI组件"""
        # 主标题
        title_frame = tk.Frame(self.root)
        title_frame.pack(pady=10, fill=tk.X)
        
        title_label = tk.Label(title_frame, text="PPT转图片幻灯片工具", 
                              font=("Microsoft YaHei", 18, "bold"), fg="#2E86AB")
        title_label.pack()
        
        subtitle_label = tk.Label(title_frame, text="背景填充版本 - 让图片作为幻灯片背景", 
                                 font=("Microsoft YaHei", 10), fg="#666666")
        subtitle_label.pack()
        
        # 分隔线
        separator1 = ttk.Separator(self.root, orient='horizontal')
        separator1.pack(fill=tk.X, padx=20, pady=10)
        
        # 功能说明
        info_frame = tk.Frame(self.root)
        info_frame.pack(pady=10, padx=20, fill=tk.X)
        
        info_text = """✨ 功能特色：
• 将PPT的每张幻灯片转换为图片，然后作为背景填充到新的幻灯片中
• 支持 .ppt 和 .pptx 格式
• 图片将作为背景而非前景对象，提供更好的视觉效果
• 自动处理图片尺寸和比例，确保完美填充"""
        
        info_label = tk.Label(info_frame, text=info_text, justify=tk.LEFT, 
                             font=("Microsoft YaHei", 9), fg="#444444")
        info_label.pack(anchor=tk.W)
        
        # 分隔线
        separator2 = ttk.Separator(self.root, orient='horizontal')
        separator2.pack(fill=tk.X, padx=20, pady=10)
        
        # 文件选择区域
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, padx=20, fill=tk.X)
        
        tk.Label(file_frame, text="📁 选择PPT文件：", 
                font=("Microsoft YaHei", 11, "bold")).pack(anchor=tk.W)
        
        select_frame = tk.Frame(file_frame)
        select_frame.pack(fill=tk.X, pady=5)
        
        self.file_var = tk.StringVar(value="未选择文件")
        file_display = tk.Label(select_frame, textvariable=self.file_var, 
                               relief=tk.SUNKEN, anchor=tk.W, 
                               font=("Microsoft YaHei", 9), bg="#F8F9FA")
        file_display.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        select_btn = tk.Button(select_frame, text="浏览文件", command=self.select_file,
                              font=("Microsoft YaHei", 9), bg="#28A745", fg="white",
                              width=12, height=1)
        select_btn.pack(side=tk.RIGHT)
        
        # 转换按钮区域
        convert_frame = tk.Frame(self.root)
        convert_frame.pack(pady=20)
        
        self.convert_btn = tk.Button(convert_frame, text="🚀 开始转换", 
                                   command=self.start_conversion,
                                   font=("Microsoft YaHei", 12, "bold"),
                                   bg="#007BFF", fg="white",
                                   width=20, height=2,
                                   relief=tk.RAISED, bd=2)
        self.convert_btn.pack()
        
        # 进度条
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(pady=10, padx=20, fill=tk.X)
        
        tk.Label(progress_frame, text="转换进度：", 
                font=("Microsoft YaHei", 10)).pack(anchor=tk.W)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="准备就绪")
        status_label = tk.Label(progress_frame, textvariable=self.status_var,
                               font=("Microsoft YaHei", 9), fg="#666666")
        status_label.pack(anchor=tk.W)
        
        # 日志区域
        log_frame = tk.LabelFrame(self.root, text="📋 转换日志", 
                                 font=("Microsoft YaHei", 10, "bold"))
        log_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        # 创建文本框和滚动条
        text_frame = tk.Frame(log_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(text_frame, height=16, wrap=tk.WORD,
                               font=("Consolas", 9), bg="#F8F9FA")
        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部信息
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=10)
        
        tk.Label(bottom_frame, text="💡 提示：转换完成后会自动打开文件保存对话框",
                font=("Microsoft YaHei", 8), fg="#6C757D").pack(anchor=tk.W)
        
        self.selected_file = None
        
        # 初始化日志
        self.log("PPT转图片幻灯片工具已启动")
        self.log("请选择要转换的PPT文件")
        
    def log(self, message):
        """添加日志消息到队列"""
        self.message_queue.put(('log', message))
        
    def update_status(self, status):
        """更新状态消息到队列"""
        self.message_queue.put(('status', status))
        
    def update_progress(self, action):
        """更新进度条到队列"""
        self.message_queue.put(('progress', action))
        
    def process_queue(self):
        """处理消息队列"""
        try:
            while True:
                msg_type, msg_data = self.message_queue.get_nowait()
                
                if msg_type == 'log':
                    self.log_text.insert(tk.END, f"{msg_data}\n")
                    self.log_text.see(tk.END)
                    
                elif msg_type == 'status':
                    self.status_var.set(msg_data)
                    
                elif msg_type == 'progress':
                    if msg_data == 'start':
                        self.progress.start(10)
                    elif msg_data == 'stop':
                        self.progress.stop()
                        
                elif msg_type == 'conversion_complete':
                    success, output_file = msg_data
                    self.on_conversion_complete(success, output_file)
                    
        except queue.Empty:
            pass
        
        # 每100ms检查一次队列
        self.root.after(100, self.process_queue)
        
    def select_file(self):
        """选择PPT文件"""
        file_types = [
            ("PowerPoint文件", "*.ppt *.pptx"),
            ("PowerPoint 97-2003", "*.ppt"),
            ("PowerPoint 2007+", "*.pptx"),
            ("所有文件", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择PPT文件", 
            filetypes=file_types,
            initialdir=os.getcwd()
        )
        
        if filename:
            self.selected_file = filename
            display_name = os.path.basename(filename)
            if len(display_name) > 50:
                display_name = display_name[:47] + "..."
            self.file_var.set(display_name)
            self.log(f"已选择文件: {filename}")
            self.convert_btn.config(state=tk.NORMAL)
        
    def start_conversion(self):
        """开始转换（在新线程中）"""
        if not self.selected_file:
            messagebox.showerror("错误", "请先选择PPT文件")
            return
            
        # 自动生成输出文件路径，与原PPT在同一目录
        input_dir = os.path.dirname(self.selected_file)
        input_basename = os.path.splitext(os.path.basename(self.selected_file))[0]
        output_file = os.path.join(input_dir, f"{input_basename}_背景版.pptx")
        
        # 如果文件已存在，生成不重复的文件名
        counter = 1
        while os.path.exists(output_file):
            output_file = os.path.join(input_dir, f"{input_basename}_背景版({counter}).pptx")
            counter += 1
        
        self.log(f"输出文件路径: {output_file}")
            
        # 禁用转换按钮
        self.convert_btn.config(state=tk.DISABLED)
        self.update_status("正在转换...")
        self.update_progress('start')
        
        # 在新线程中执行转换
        conversion_thread = threading.Thread(
            target=self.convert_in_thread,
            args=(self.selected_file, output_file)
        )
        conversion_thread.daemon = True
        conversion_thread.start()
        
    def convert_in_thread(self, input_ppt, output_ppt):
        """在线程中执行转换"""
        try:
            success = self.convert_ppt_to_image_slides(input_ppt, output_ppt)
            self.message_queue.put(('conversion_complete', (success, output_ppt)))
        except Exception as e:
            self.log(f"转换过程发生异常: {e}")
            self.message_queue.put(('conversion_complete', (False, output_ppt)))
        
    def on_conversion_complete(self, success, output_file):
        """转换完成回调"""
        self.update_progress('stop')
        self.convert_btn.config(state=tk.NORMAL)
        
        if success:
            self.update_status("转换完成！")
            self.log("=" * 50)
            self.log("🎉 转换成功完成！")
            self.log(f"输出文件: {output_file}")
            
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file) / 1024 / 1024  # MB
                self.log(f"文件大小: {file_size:.2f} MB")
                
        else:
            self.update_status("转换失败")
            self.log("❌ 转换失败，请检查上面的日志信息")
            messagebox.showerror("转换失败", "转换过程中发生错误，请查看日志获取详细信息")
    
    def convert_ppt_to_image_slides(self, input_ppt, output_ppt):
        """转换PPT为图片幻灯片（背景模式）"""
        temp_dir = None
        powerpoint = None
        
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(prefix="ppt_to_image_")
            self.log(f"创建临时目录: {temp_dir}")
            
            # 1. 启动PowerPoint
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            self.log("PowerPoint COM接口初始化成功")
            
            # 2. 打开原PPT
            presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt))
            slide_count = presentation.Slides.Count
            self.log(f"成功打开PPT，共 {slide_count} 张幻灯片")
            
            # 获取幻灯片尺寸信息
            slide_width = presentation.PageSetup.SlideWidth
            slide_height = presentation.PageSetup.SlideHeight
            self.log(f"幻灯片尺寸: {slide_width:.1f} x {slide_height:.1f} 点")
            
            # 3. 导出为图片
            self.log("开始导出幻灯片为图片...")
            image_files = []
            
            for i in range(1, slide_count + 1):
                image_path = os.path.join(temp_dir, f"slide_{i:03d}.png")
                self.log(f"导出幻灯片 {i}/{slide_count}: slide_{i:03d}.png")
                
                try:
                    presentation.Slides(i).Export(image_path, "PNG")
                    if os.path.exists(image_path):
                        image_files.append(image_path)
                        self.update_status(f"已导出 {i}/{slide_count} 张幻灯片")
                    else:
                        self.log(f"警告：幻灯片 {i} 导出后文件不存在")
                except Exception as e:
                    self.log(f"导出幻灯片 {i} 失败: {e}")
                    continue
            
            if not image_files:
                self.log("错误：没有成功导出任何图片")
                return False
                
            self.log(f"成功导出 {len(image_files)} 张图片")
            presentation.Close()
            
            # 4. 重新打开PPT作为模板，设置背景
            self.log("重新打开PPT设置背景...")
            template_presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt))
            
            # 处理每张幻灯片
            processed_count = 0
            for i, image_file in enumerate(image_files, 1):
                if i <= template_presentation.Slides.Count:
                    slide = template_presentation.Slides(i)
                    
                    try:
                        self.log(f"处理幻灯片 {i}/{len(image_files)}...")
                        self.update_status(f"设置背景 {i}/{len(image_files)}")
                        
                        # 清空幻灯片内容
                        shape_count = slide.Shapes.Count
                        for j in range(shape_count, 0, -1):
                            try:
                                slide.Shapes(j).Delete()
                            except:
                                pass
                            
                        # 设置背景图片
                        slide.Background.Fill.UserPicture(os.path.abspath(image_file))
                        self.log(f"✓ 成功设置幻灯片 {i} 背景")
                        processed_count += 1

                    except Exception as e:
                        self.log(f"✗ 处理幻灯片 {i} 时发生错误: {e}")
                        continue
            
            if processed_count == 0:
                self.log("错误：没有成功处理任何幻灯片")
                return False
            
            # 5. 保存新PPT
            self.log("保存处理后的PPT...")
            self.update_status("正在保存文件...")
            template_presentation.SaveAs(os.path.abspath(output_ppt))
            template_presentation.Close()
            
            self.log(f"成功处理 {processed_count} 张幻灯片")
            self.log("PPT保存完成")
            return True
            
        except Exception as e:
            self.log(f"转换过程发生错误: {e}")
            import traceback
            self.log("详细错误信息:")
            self.log(traceback.format_exc())
            return False
            
        finally:
            # 清理资源
            try:
                if powerpoint:
                    powerpoint.Quit()
                    self.log("PowerPoint COM接口已关闭")
            except:
                pass
                
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                    self.log(f"清理临时目录: {temp_dir}")
                except Exception as e:
                    self.log(f"清理临时目录失败: {e}")
        
    def on_closing(self):
        """程序关闭时的处理"""
        self.root.quit()
        self.root.destroy()
        
    def run(self):
        """启动GUI"""
        self.root.mainloop()

if __name__ == "__main__":
    app = PPTToImageSlidesGUI()
    app.run()
