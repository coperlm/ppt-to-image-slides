#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT转换为图片幻灯片工具 - 背景版本（简化）
基于原版代码，仅添加背景设置功能
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import tempfile
import shutil
from PIL import Image
import win32com.client

class PPTToImageSlidesSimple:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PPT转图片幻灯片 - 背景版")
        self.root.geometry("600x500")
        
        # 创建GUI界面
        self.create_widgets()
        
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="PPT转图片幻灯片工具 - 背景版", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 说明文字
        info_text = """
将PPT转换为图片幻灯片，图片将作为背景填充整个幻灯片。
支持的格式：.ppt, .pptx
        """
        info_label = tk.Label(self.root, text=info_text, justify=tk.LEFT)
        info_label.pack(pady=10)
        
        # 文件选择按钮
        select_frame = tk.Frame(self.root)
        select_frame.pack(pady=10)
        
        tk.Button(select_frame, text="选择PPT文件", command=self.select_file, width=20, height=2).pack(side=tk.LEFT, padx=5)
        
        # 显示选中的文件
        self.file_label = tk.Label(self.root, text="未选择文件", width=60, relief=tk.SUNKEN)
        self.file_label.pack(pady=10)
        
        # 转换按钮
        tk.Button(self.root, text="开始转换", command=self.convert, width=20, height=2, bg="#4CAF50", fg="white").pack(pady=20)
        
        # 日志文本框
        log_frame = tk.Frame(self.root)
        log_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        tk.Label(log_frame, text="转换日志:", anchor="w").pack(fill=tk.X)
        
        self.log_text = tk.Text(log_frame, height=15, width=70)
        scrollbar = tk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.selected_file = None
        
    def log(self, message):
        """记录日志"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def select_file(self):
        """选择PPT文件"""
        file_types = [("PowerPoint files", "*.ppt *.pptx"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="选择PPT文件", filetypes=file_types)
        
        if filename:
            self.selected_file = filename
            self.file_label.config(text=os.path.basename(filename))
            self.log(f"已选择文件: {filename}")
        
    def convert(self):
        """执行转换"""
        if not self.selected_file:
            messagebox.showerror("错误", "请先选择PPT文件")
            return
            
        try:
            # 选择输出文件
            output_file = filedialog.asksaveasfilename(
                title="保存转换后的PPT",
                defaultextension=".pptx",
                filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
            )
            
            if not output_file:
                return
                
            self.log("开始转换...")
            self.log(f"输入文件: {self.selected_file}")
            self.log(f"输出文件: {output_file}")
            
            success = self.convert_ppt_to_image_slides(self.selected_file, output_file)
            
            if success:
                self.log("✓ 转换完成！")
                messagebox.showinfo("成功", f"转换完成！文件已保存为：\n{output_file}")
            else:
                self.log("✗ 转换失败")
                messagebox.showerror("失败", "转换过程中发生错误，请查看日志")
                
        except Exception as e:
            self.log(f"✗ 发生异常: {e}")
            messagebox.showerror("错误", f"发生异常：{str(e)}")
    
    def convert_ppt_to_image_slides(self, input_ppt, output_ppt):
        """转换PPT为图片幻灯片"""
        temp_dir = None
        powerpoint = None
        
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(prefix="ppt_to_image_")
            self.log(f"临时目录: {temp_dir}")
            
            # 1. 启动PowerPoint
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # 不设置Visible属性，避免COM错误
            self.log("PowerPoint COM接口初始化成功")
            
            # 2. 打开原PPT
            presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt))
            self.log(f"找到 {presentation.Slides.Count} 张幻灯片")
            
            # 获取幻灯片尺寸信息
            slide_width = presentation.PageSetup.SlideWidth
            slide_height = presentation.PageSetup.SlideHeight
            self.log(f"幻灯片尺寸: {slide_width} x {slide_height} 点")
            
            # 3. 导出为图片
            image_files = []
            for i in range(1, presentation.Slides.Count + 1):
                image_path = os.path.join(temp_dir, f"slide_{i:03d}.png")
                self.log(f"导出幻灯片 {i}/{presentation.Slides.Count}: slide_{i:03d}.png")
                
                try:
                    presentation.Slides(i).Export(image_path, "PNG")
                    image_files.append(image_path)
                except Exception as e:
                    self.log(f"导出幻灯片 {i} 失败: {e}")
                    continue
            
            presentation.Close()
            self.log(f"成功导出 {len(image_files)} 张图片")
            
            if not image_files:
                self.log("没有成功导出任何图片")
                return False
            
            # 4. 创建新的PPT，使用背景模式
            self.log("创建新的PPT，使用背景填充模式...")
            
            # 重新打开第一个文件作为模板
            template_presentation = powerpoint.Presentations.Open(os.path.abspath(input_ppt))
            
            # 修改每张幻灯片
            for i, image_file in enumerate(image_files, 1):
                if i <= template_presentation.Slides.Count:
                    slide = template_presentation.Slides(i)
                    
                    try:
                        self.log(f"设置幻灯片 {i} 背景...")
                        
                        # 清空幻灯片内容
                        while slide.Shapes.Count > 0:
                            slide.Shapes(1).Delete()
                        
                        # 尝试设置背景
                        try:
                            # 方法1：使用背景填充
                            slide.Background.Fill.UserPicture(os.path.abspath(image_file))
                            self.log(f"✓ 成功设置幻灯片 {i} 背景")
                        except:
                            # 方法2：作为形状添加并调整大小
                            self.log(f"背景设置失败，使用图片对象方式...")
                            shape = slide.Shapes.AddPicture(
                                os.path.abspath(image_file),
                                LinkToFile=False,
                                SaveWithDocument=True,
                                Left=0,
                                Top=0,
                                Width=slide_width,
                                Height=slide_height
                            )
                            # 发送到最底层
                            shape.ZOrder(1)  # 发送到底层
                            self.log(f"✓ 使用图片对象方式设置幻灯片 {i}")
                            
                    except Exception as e:
                        self.log(f"✗ 设置幻灯片 {i} 失败: {e}")
                        continue
            
            # 保存新PPT
            template_presentation.SaveAs(os.path.abspath(output_ppt))
            template_presentation.Close()
            
            self.log("PPT保存完成")
            return True
            
        except Exception as e:
            self.log(f"转换过程中发生错误: {e}")
            import traceback
            self.log(f"完整错误信息: {traceback.format_exc()}")
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
                except:
                    pass
    
    def run(self):
        """运行GUI"""
        self.root.mainloop()

if __name__ == "__main__":
    app = PPTToImageSlidesSimple()
    app.run()
