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
        output_file = os.path.join(input_dir, f"{input_basename}_image.pptx")
        
        # 如果文件已存在，生成不重复的文件名
        counter = 1
        while os.path.exists(output_file):
            output_file = os.path.join(input_dir, f"{input_basename}_image({counter}).pptx")
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
    
    def verify_background_set(self, slide):
        """验证幻灯片背景是否成功设置"""
        try:
            # 检查背景填充类型
            fill_type = slide.Background.Fill.Type
            # 如果是图片填充类型，说明背景设置成功
            if fill_type == 6:  # msoFillPicture = 6
                return True
            
            # 备用验证：检查是否有背景相关的形状
            try:
                if slide.Shapes.Count > 0:
                    # 检查是否有图片形状
                    for i in range(1, slide.Shapes.Count + 1):
                        shape = slide.Shapes(i)
                        if hasattr(shape, 'Type') and shape.Type == 13:  # msoShapeTypePicture = 13
                            return True
                return False
            except:
                return False
                
        except Exception as e:
            # 如果验证过程出错，假设设置失败
            return False
    
    def validate_image_file(self, image_path):
        """验证图片文件的有效性"""
        try:
            # 检查文件是否存在
            if not os.path.exists(image_path):
                return False
                
            # 检查文件大小（空文件或太小的文件可能有问题）
            file_size = os.path.getsize(image_path)
            if file_size < 1000:  # 小于1KB的图片文件可能有问题
                return False
                
            # 尝试用PIL打开图片验证其有效性
            with Image.open(image_path) as img:
                # 验证图片尺寸
                width, height = img.size
                if width < 10 or height < 10:  # 尺寸太小的图片可能有问题
                    return False
                    
                # 验证图片模式
                if img.mode not in ['RGB', 'RGBA', 'L', 'P']:
                    return False
                    
            return True
            
        except Exception as e:
            return False
    
    def convert_ppt_to_image_slides(self, input_ppt, output_ppt):
        """转换PPT为图片幻灯片（背景模式）"""
        temp_dir = None
        powerpoint = None
        
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(prefix="ppt_to_image_")
            self.log(f"创建临时目录: {temp_dir}")
            
            # 1. 启动PowerPoint（修复版本兼容性问题）
            try:
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                self.log("PowerPoint COM接口创建成功")
                
                # 尝试设置PowerPoint属性（某些版本可能不支持隐藏窗口）
                try:
                    powerpoint.DisplayAlerts = False  # 禁用警告对话框
                    self.log("已禁用PowerPoint警告对话框")
                except Exception as alert_error:
                    self.log(f"设置DisplayAlerts失败: {alert_error}")
                
                # 谨慎处理Visible属性（某些版本不允许隐藏）
                try:
                    # 先尝试获取当前状态
                    current_visible = powerpoint.Visible
                    self.log(f"PowerPoint当前可见状态: {current_visible}")
                    
                    # 如果当前不可见，尝试设置为可见（避免兼容性问题）
                    if not current_visible:
                        powerpoint.Visible = True
                        self.log("PowerPoint窗口已设置为可见")
                    
                except Exception as visible_error:
                    self.log(f"设置Visible属性失败，使用默认设置: {visible_error}")
                
                self.log("PowerPoint COM接口初始化完成")
                
            except Exception as pp_error:
                self.log(f"PowerPoint初始化失败: {pp_error}")
                return False
            
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
                    
                    # 验证导出的图片文件
                    if self.validate_image_file(image_path):
                        image_files.append(image_path)
                        self.update_status(f"已导出 {i}/{slide_count} 张幻灯片")
                        self.log(f"✓ 幻灯片 {i} 导出成功并验证通过")
                    else:
                        self.log(f"✗ 幻灯片 {i} 导出失败或文件无效")
                        # 尝试重新导出一次
                        try:
                            import time
                            time.sleep(0.5)  # 等待一下
                            presentation.Slides(i).Export(image_path, "PNG")
                            if self.validate_image_file(image_path):
                                image_files.append(image_path)
                                self.log(f"✓ 幻灯片 {i} 重新导出成功")
                            else:
                                self.log(f"✗ 幻灯片 {i} 重新导出仍然失败")
                        except:
                            self.log(f"✗ 幻灯片 {i} 重新导出时发生异常")
                            
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
            
            # 确保模板幻灯片数量与图片数量匹配
            template_slide_count = template_presentation.Slides.Count
            image_count = len(image_files)
            self.log(f"模板幻灯片数量: {template_slide_count}, 图片数量: {image_count}")
            
            if template_slide_count != image_count:
                self.log(f"警告：幻灯片数量({template_slide_count})与图片数量({image_count})不匹配")
                # 如果模板幻灯片少于图片数量，添加幻灯片
                while template_presentation.Slides.Count < image_count:
                    # 复制最后一张幻灯片
                    last_slide = template_presentation.Slides(template_presentation.Slides.Count)
                    new_slide = last_slide.Duplicate()
                    self.log(f"添加了新幻灯片，当前总数: {template_presentation.Slides.Count}")
            
            # 处理每张幻灯片
            processed_count = 0
            for i, image_file in enumerate(image_files, 1):
                if i <= template_presentation.Slides.Count:
                    slide = template_presentation.Slides(i)
                    
                    try:
                        self.log(f"处理幻灯片 {i}/{len(image_files)}...")
                        self.update_status(f"设置背景 {i}/{len(image_files)}")
                        
                        # 关键修复：设置FollowMasterBackground为False
                        try:
                            slide.FollowMasterBackground = False
                            self.log(f"✓ 幻灯片 {i} 已禁用跟随母版背景")
                        except Exception as e:
                            self.log(f"设置FollowMasterBackground失败: {e}")
                        
                        # 设置为空白版式（避免占位符文本）
                        try:
                            slide.Layout = 12  # ppLayoutBlank = 12
                            self.log(f"✓ 幻灯片 {i} 已设置为空白版式")
                        except Exception as e:
                            self.log(f"设置空白版式失败: {e}")
                        
                        # 彻底清空幻灯片内容（包括占位符）
                        try:
                            shape_count = slide.Shapes.Count
                            deleted_count = 0
                            # 从后往前删除所有形状，包括占位符
                            for j in range(shape_count, 0, -1):
                                try:
                                    shape = slide.Shapes(j)
                                    # 删除所有形状，包括占位符
                                    shape.Delete()
                                    deleted_count += 1
                                except:
                                    pass
                            self.log(f"清空了 {deleted_count} 个元素（包括占位符）")
                        except Exception as e:
                            self.log(f"清空幻灯片内容时出错: {e}")
                        
                        # 设置背景图片
                        background_set = False
                        abs_image_path = os.path.abspath(image_file)
                        
                        # 方法1：使用UserPicture设置背景
                        try:
                            slide.Background.Fill.UserPicture(abs_image_path)
                            # 等待一下让设置生效
                            import time
                            time.sleep(0.1)
                            background_set = self.verify_background_set(slide)
                            if background_set:
                                self.log(f"✓ 方法1成功：幻灯片 {i} 背景设置完成")
                            else:
                                self.log(f"方法1设置后验证失败")
                        except Exception as e:
                            self.log(f"方法1失败：{e}")
                        
                        # 如果UserPicture失败，使用备用方案
                        if not background_set:
                            try:
                                # 获取幻灯片尺寸
                                slide_width = template_presentation.PageSetup.SlideWidth
                                slide_height = template_presentation.PageSetup.SlideHeight
                                
                                # 添加图片铺满整个幻灯片
                                picture = slide.Shapes.AddPicture(abs_image_path, False, True, 0, 0, slide_width, slide_height)
                                # 将图片移到最底层（作为背景）
                                try:
                                    picture.ZOrder(0)  # 发送到底层
                                except:
                                    pass
                                background_set = True
                                self.log(f"✓ 备用方案成功：幻灯片 {i} 图片作为背景添加完成")
                            except Exception as e:
                                self.log(f"备用方案失败：{e}")
                        
                        # 确保没有其他内容（最终检查）
                        if background_set:
                            try:
                                # 检查是否有新的占位符或形状被意外添加
                                current_shape_count = slide.Shapes.Count
                                if current_shape_count > 1:  # 应该只有背景图片
                                    for j in range(current_shape_count, 1, -1):  # 保留第一个形状（背景）
                                        try:
                                            shape = slide.Shapes(j)
                                            # 删除任何额外的形状（包括可能重新出现的占位符）
                                            shape.Delete()
                                            self.log(f"删除了额外的形状/占位符")
                                        except:
                                            pass
                            except Exception as e:
                                self.log(f"最终清理时出错: {e}")
                        
                        if background_set:
                            processed_count += 1
                            self.log(f"✓ 幻灯片 {i} 处理完成，仅保留纯净背景")
                        else:
                            self.log(f"✗ 幻灯片 {i} 所有背景设置方法都失败")

                    except Exception as e:
                        self.log(f"✗ 处理幻灯片 {i} 时发生严重错误: {e}")
                        continue
            
            if processed_count == 0:
                self.log("错误：没有成功处理任何幻灯片")
                return False
            
            # 5. 保存新PPT（改进的错误处理）
            self.log("保存处理后的PPT...")
            self.update_status("正在保存文件...")
            
            save_success = False
            try:
                # 确保保存路径目录存在
                output_dir = os.path.dirname(os.path.abspath(output_ppt))
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                # 尝试保存
                abs_output_path = os.path.abspath(output_ppt)
                self.log(f"保存到: {abs_output_path}")
                
                template_presentation.SaveAs(abs_output_path)
                save_success = True
                self.log("PPT保存成功")
                
            except Exception as save_error:
                self.log(f"保存失败，尝试备用保存方法: {save_error}")
                try:
                    # 备用保存方法：使用ExportAsFixedFormat
                    backup_path = output_ppt.replace('.pptx', '_backup.pptx')
                    template_presentation.SaveAs(os.path.abspath(backup_path))
                    save_success = True
                    self.log(f"备用保存成功: {backup_path}")
                except Exception as backup_error:
                    self.log(f"备用保存也失败: {backup_error}")
            
            # 安全关闭演示文稿
            try:
                if save_success:
                    # 等待保存完成
                    import time
                    time.sleep(0.5)
                
                # 尝试关闭演示文稿
                template_presentation.Close()
                self.log("演示文稿已关闭")
                
            except Exception as close_error:
                self.log(f"关闭演示文稿时出错（可能已经关闭）: {close_error}")
                # 尝试强制关闭
                try:
                    powerpoint.Presentations.Close()
                except:
                    pass
            
            if save_success:
                self.log(f"成功处理 {processed_count} 张幻灯片")
                self.log("PPT转换完成")
                return True
            else:
                self.log("保存失败，转换未完成")
                return False
            
        except Exception as e:
            self.log(f"转换过程发生错误: {e}")
            import traceback
            self.log("详细错误信息:")
            self.log(traceback.format_exc())
            return False
            
        finally:
            # 改进的资源清理
            self.log("开始清理资源...")
            
            # 1. 安全关闭所有演示文稿
            try:
                if 'powerpoint' in locals() and powerpoint:
                    # 关闭所有打开的演示文稿
                    presentations_count = powerpoint.Presentations.Count
                    self.log(f"发现 {presentations_count} 个打开的演示文稿")
                    
                    for i in range(presentations_count, 0, -1):
                        try:
                            presentation = powerpoint.Presentations(i)
                            presentation.Close()
                            self.log(f"已关闭演示文稿 {i}")
                        except Exception as close_err:
                            self.log(f"关闭演示文稿 {i} 失败: {close_err}")
                    
                    # 等待一下再退出PowerPoint
                    import time
                    time.sleep(0.5)
                    
            except Exception as cleanup_error:
                self.log(f"清理演示文稿时出错: {cleanup_error}")
            
            # 2. 安全退出PowerPoint
            try:
                if 'powerpoint' in locals() and powerpoint:
                    powerpoint.Quit()
                    self.log("PowerPoint COM接口已关闭")
                    
                    # 释放COM对象引用
                    del powerpoint
                    
            except Exception as quit_error:
                self.log(f"退出PowerPoint时出错: {quit_error}")
            
            # 3. 清理临时目录
            if 'temp_dir' in locals() and temp_dir and os.path.exists(temp_dir):
                try:
                    # 等待一下确保文件不被占用
                    import time
                    time.sleep(0.5)
                    shutil.rmtree(temp_dir)
                    self.log(f"清理临时目录: {temp_dir}")
                except Exception as e:
                    self.log(f"清理临时目录失败: {e}")
                    # 尝试强制清理
                    try:
                        import subprocess
                        subprocess.run(['rmdir', '/s', '/q', temp_dir], shell=True, check=False)
                        self.log("强制清理临时目录完成")
                    except:
                        self.log("强制清理也失败，临时文件可能需要手动删除")
            
            self.log("资源清理完成")
        
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
