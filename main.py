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
import time
import traceback

class PPTToImageSlidesGUI:
    def __init__(self):
        # 优先尝试用tkinterdnd2增强主窗口
        self.dnd_enabled = False
        try:
            from tkinterdnd2 import TkinterDnD
            self.root = TkinterDnD.Tk()
            self.dnd_enabled = True
        except ImportError:
            self.root = tk.Tk()
        self.root.title("PPT转图片幻灯片工具 - 背景版")
        self.root.geometry("700x700")
        self.root.resizable(True, True)

        # 消息队列用于线程间通信
        self.message_queue = queue.Queue()

        # 转换状态标志：PowerPoint 为单实例 COM，禁止并发转换
        self.converting = False

        # 创建GUI界面
        self.create_widgets()

        # 拖拽支持
        self.add_drag_and_drop_support()

        # 设置关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 启动消息处理
        self.process_queue()
    def add_drag_and_drop_support(self):
        """为主窗口添加文件拖拽支持，仅支持PPT文件"""
        if getattr(self, 'dnd_enabled', False):
            # 用tkinterdnd2增强拖拽体验
            self.root.drop_target_register('DND_Files')
            self.root.dnd_bind('<<Drop>>', self.on_drop_file)
        else:
            # 降级为普通Tk，提示用户
            self.log("未检测到tkinterdnd2库，拖拽功能受限。可通过 pip install tkinterdnd2 获得更好体验。")
            # 绑定简单的拖拽事件（大多数Tk无效，仅保留提示）
            pass

    def _set_convert_enabled(self, enabled):
        """根据"入参意图 + 转换状态 + 文件选择情况"安全设置转换按钮可用性。

        转换进行中始终禁用，避免用户在转换期间再次触发并发转换。
        """
        try:
            should_enable = bool(enabled) and not self.converting and bool(self.selected_file)
            self.convert_btn.config(state=tk.NORMAL if should_enable else tk.DISABLED)
        except tk.TclError:
            pass

    def on_drop_file(self, event):
        """拖拽文件到窗口时的处理，支持带空格路径"""
        import re
        data = event.data.strip()
        # 用正则提取所有大括号包裹的路径，否则按空格分割
        paths = re.findall(r'\{([^}]*)\}', data)
        if not paths:
            # 没有大括号，直接按空格分割
            paths = data.split()
        if not paths:
            self.log("未检测到有效的文件路径")
            return
        file_path = paths[0]
        # 检查扩展名
        if file_path.lower().endswith(('.ppt', '.pptx')):
            self.selected_file = file_path
            display_name = os.path.basename(file_path)
            if len(display_name) > 50:
                display_name = display_name[:47] + "..."
            self.file_var.set(display_name)
            self.log(f"已拖入文件: {file_path}")
            self._set_convert_enabled(True)
        else:
            self.log(f"拖入的文件不是PPT: {file_path}")
            messagebox.showwarning("文件类型不支持", "请拖入PPT或PPTX文件！")
        
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
        
    def _handle_queue_message(self, msg_type, msg_data):
        """处理单条队列消息（与轮询解耦，便于异常隔离）"""
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

    def process_queue(self):
        """处理消息队列。

        单条消息处理异常不应中断整个轮询循环，否则日志/进度/完成回调会永久失效；
        因此异常被隔离在单条消息层面，且重新调度放在 finally 中确保一定执行。
        """
        try:
            while True:
                try:
                    msg_type, msg_data = self.message_queue.get_nowait()
                except queue.Empty:
                    break

                try:
                    self._handle_queue_message(msg_type, msg_data)
                except Exception as e:
                    try:
                        self.log_text.insert(tk.END, f"[内部错误] 处理界面消息失败: {e}\n")
                        self.log_text.see(tk.END)
                    except Exception:
                        pass
        finally:
            # 每100ms检查一次队列；窗口已销毁时忽略重调度异常
            try:
                self.root.after(100, self.process_queue)
            except tk.TclError:
                pass
        
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
            self._set_convert_enabled(True)
        
    def start_conversion(self):
        """开始转换（在新线程中）"""
        if not self.selected_file:
            messagebox.showerror("错误", "请先选择PPT文件")
            return

        # 防止并发转换：PowerPoint 为单实例 COM，多个转换线程会相互干扰
        if self.converting:
            messagebox.showwarning("正在转换", "已有转换任务正在进行，请等待其完成")
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
            
        # 进入转换状态并禁用转换按钮
        self.converting = True
        self._set_convert_enabled(False)
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
        """在线程中执行转换，需初始化COM"""
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                success = self.convert_ppt_to_image_slides(input_ppt, output_ppt)
                self.message_queue.put(('conversion_complete', (success, output_ppt)))
            except Exception as e:
                self.log(f"转换过程发生异常: {e}")
                self.message_queue.put(('conversion_complete', (False, output_ppt)))
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
        
    def on_conversion_complete(self, success, output_file):
        """转换完成回调"""
        self.update_progress('stop')
        self.converting = False
        self._set_convert_enabled(True)
        
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
    
    def _clear_slide_shapes(self, slide):
        """删除幻灯片上的所有形状（含占位符），返回删除数量。

        仅当"存在任意图片形状就判定成功"这类启发式判据可能误报，
        此方法用于确定性清空，供正常流程与失败回退复用。
        """
        deleted_count = 0
        try:
            for j in range(slide.Shapes.Count, 0, -1):
                try:
                    slide.Shapes(j).Delete()
                    deleted_count += 1
                except Exception:
                    pass
        except Exception as e:
            self.log(f"清空幻灯片形状时出错: {e}")
        return deleted_count

    def verify_background_set(self, slide):
        """验证幻灯片背景是否成功设置为图片填充。

        仅以背景填充类型作为判据（msoFillPicture = 6）。
        刻意不使用"存在图片形状"作为兜底判据：该判据会把"背景未设置成功"
        误判为成功，从而跳过更可靠的 AddPicture 回退方案。
        验证失败时返回 False，交由调用方走回退逻辑，行为可预期。
        """
        try:
            return slide.Background.Fill.Type == 6  # msoFillPicture = 6
        except Exception as e:
            self.log(f"验证背景设置时出错（按失败处理）: {e}")
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
            # 每项为 (幻灯片序号, JPG路径)：保留原始页码，避免部分导出失败时图片错位
            image_files = []
            
            for i in range(1, slide_count + 1):
                png_path = os.path.join(temp_dir, f"slide_{i:03d}_tmp.png")  # 临时PNG
                jpg_path = os.path.join(temp_dir, f"slide_{i:03d}.jpg")
                self.log(f"导出幻灯片 {i}/{slide_count}: slide_{i:03d}.jpg")

                try:
                    presentation.Slides(i).Export(png_path, "PNG")

                    # 验证临时PNG图片
                    if self.validate_image_file(png_path):
                        # 转为JPG
                        try:
                            with Image.open(png_path) as img:
                                rgb_img = img.convert("RGB")
                                rgb_img.save(jpg_path, "JPG", quality=95, optimize=True)
                            # 验证JPG
                            if self.validate_image_file(jpg_path):
                                image_files.append((i, jpg_path))
                                self.update_status(f"已导出 {i}/{slide_count} 张幻灯片")
                                self.log(f"✓ 幻灯片 {i} JPG 转换成功")
                            else:
                                self.log(f"✗ 幻灯片 {i} JPG 转换失败")
                        except Exception as imgconv_e:
                            self.log(f"✗ 幻灯片 {i} PNG转JPG失败: {imgconv_e}")
                    else:
                        self.log(f"✗ 幻灯片 {i} PNG临时文件导出失败或文件无效")
                        # 尝试重新导出一次
                        try:
                            time.sleep(0.5)
                            presentation.Slides(i).Export(png_path, "PNG")
                            if self.validate_image_file(png_path):
                                with Image.open(png_path) as img:
                                    rgb_img = img.convert("RGB")
                                    rgb_img.save(jpg_path, "JPG", quality=95, optimize=True)
                                if self.validate_image_file(jpg_path):
                                    image_files.append((i, jpg_path))
                                    self.log(f"✓ 幻灯片 {i} 重新导出并转JPG成功")
                                else:
                                    self.log(f"✗ 幻灯片 {i} 重新导出转JPG仍然失败")
                            else:
                                self.log(f"✗ 幻灯片 {i} 重新导出PNG临时文件仍然失败")
                        except Exception as retry_e:
                            self.log(f"✗ 幻灯片 {i} 重新导出时发生异常: {retry_e}")

                except Exception as e:
                    self.log(f"导出幻灯片 {i} 失败: {e}")
                    continue

            if not image_files:
                self.log("错误：没有成功导出任何图片")
                return False
                
            self.log(f"成功导出 {len(image_files)} 张JPG图片")
            presentation.Close()
            
            # 4. 重新打开PPT作为模板，设置背景
            self.log("重新打开PPT，设置JPG图片为背景...")
            template_presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt))
            
            # 建立 幻灯片序号 -> 图片路径 的映射，确保图片与原始页严格对应
            template_slide_count = template_presentation.Slides.Count
            image_by_slide = {idx: path for idx, path in image_files}
            image_count = len(image_files)
            self.log(f"模板幻灯片数量: {template_slide_count}, 成功导出图片数量: {image_count}")

            # 若成功导出的最大页码超出模板页数，防御性补齐（正常情况下不会发生）
            if image_by_slide:
                while template_presentation.Slides.Count < max(image_by_slide):
                    last_slide = template_presentation.Slides(template_presentation.Slides.Count)
                    last_slide.Duplicate()
                    self.log(f"添加了新幻灯片，当前总数: {template_presentation.Slides.Count}")

            # 清空没有对应图片的幻灯片（导出失败或多余的页），避免输出残留原始内容
            for i in range(1, template_presentation.Slides.Count + 1):
                if i not in image_by_slide:
                    try:
                        self.log(f"警告：幻灯片 {i} 无对应图片，已清空其内容以避免残留原始内容")
                        deleted = self._clear_slide_shapes(template_presentation.Slides(i))
                        self.log(f"已清空幻灯片 {i} 的 {deleted} 个元素")
                    except Exception as e:
                        self.log(f"清空幻灯片 {i} 时出错: {e}")
            
            # 处理每张幻灯片
            processed_count = 0
            for i, image_file in sorted(image_by_slide.items()):
                if i <= template_presentation.Slides.Count:
                    slide = template_presentation.Slides(i)
                    
                    try:
                        self.log(f"处理幻灯片 {i}/{len(image_files)} (JPG)...")
                        self.update_status(f"设置JPG背景 {i}/{len(image_files)}")
                        
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
                        deleted_count = self._clear_slide_shapes(slide)
                        self.log(f"清空了 {deleted_count} 个元素（包括占位符）")
                        
                        # 设置背景图片
                        background_set = False
                        abs_image_path = os.path.abspath(image_file)
                        bg_shape_id = None  # 备用方案中背景图的唯一Id，供最终清理精确保留

                        # 方法1：使用UserPicture设置JPG背景
                        try:
                            slide.Background.Fill.UserPicture(abs_image_path)
                            # 等待一下让设置生效
                            time.sleep(0.1)
                            background_set = self.verify_background_set(slide)
                            if background_set:
                                self.log(f"✓ 方法1成功：幻灯片 {i} JPG背景设置完成")
                            else:
                                self.log(f"方法1设置JPG后验证失败")
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
                                # 记录背景图唯一Id：后续按Id精确保留，不依赖形状顺序（ZOrder 可能失败）
                                try:
                                    bg_shape_id = picture.Id
                                except Exception:
                                    bg_shape_id = None
                                # 将图片移到最底层（作为背景）
                                try:
                                    picture.ZOrder(0)  # 发送到底层
                                except Exception as zorder_error:
                                    self.log(f"ZOrder调整失败（不影响使用）: {zorder_error}")
                                background_set = True
                                self.log(f"✓ 备用方案成功：幻灯片 {i} JPG图片作为背景添加完成")
                            except Exception as e:
                                self.log(f"备用方案失败：{e}")
                        
                        # 最终清理：UserPicture 成功时幻灯片应无形状；AddPicture 成功时应仅保留背景图。
                        # 按 Shape.Id 精确匹配保留背景图，避免依赖形状顺序（ZOrder 可能失败）
                        if background_set:
                            try:
                                for j in range(slide.Shapes.Count, 0, -1):
                                    try:
                                        shape = slide.Shapes(j)
                                        shape_id = None
                                        try:
                                            shape_id = shape.Id
                                        except Exception:
                                            shape_id = None
                                        if bg_shape_id is not None and shape_id == bg_shape_id:
                                            continue
                                        shape.Delete()
                                        self.log("删除了额外的形状/占位符")
                                    except Exception:
                                        pass
                            except Exception as e:
                                self.log(f"最终清理时出错: {e}")
                        
                        if background_set:
                            processed_count += 1
                            self.log(f"✓ 幻灯片 {i} 处理完成，仅保留纯净JPG背景")
                        else:
                            self.log(f"✗ 幻灯片 {i} 所有JPG背景设置方法都失败")

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
                    time.sleep(0.5)
                
                # 尝试关闭演示文稿
                template_presentation.Close()
                self.log("演示文稿已关闭")
                
            except Exception as close_error:
                self.log(f"关闭演示文稿时出错（可能已经关闭）: {close_error}")
                # 尝试强制关闭
                try:
                    powerpoint.Presentations.Close()
                except Exception:
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
                    time.sleep(0.5)
                    # ignore_errors=True：避免个别文件仍被占用导致整体失败
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    if os.path.exists(temp_dir):
                        self.log(f"临时目录未能完全删除，可能需要手动清理: {temp_dir}")
                    else:
                        self.log(f"清理临时目录: {temp_dir}")
                except Exception as e:
                    self.log(f"清理临时目录失败: {e}")
            
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
