#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用示例脚本
"""

import os
import sys

def show_usage_examples():
    """显示使用示例"""
    print("=" * 60)
    print("PPT转图片幻灯片工具 - 使用示例")
    print("=" * 60)
    
    print("\n🖥️  图形化界面版本（推荐）:")
    print("   python ppt_to_image_slides_gui.py")
    print("   → 打开图形化界面，点击操作即可")
    print("   → 支持文件拖拽、实时进度显示")
    print("   → 自动设置输出文件名")
    
    print("\n⌨️  命令行版本:")
    print("\n1. 基本用法（最简单）：")
    print("   python ppt_to_image_slides.py presentation.pptx")
    print("   → 输出：presentation_images.pptx")
    
    print("\n2. 指定输出文件名：")
    print("   python ppt_to_image_slides.py presentation.pptx -o output.pptx")
    
    print("\n3. 指定图片格式：")
    print("   python ppt_to_image_slides.py presentation.pptx -f JPG")
    
    print("\n4. 完整示例（包含中文路径）：")
    print('   python ppt_to_image_slides.py "我的演示文稿.pptx" -o "图片版.pptx" -f PNG')
    
    print("\n5. 查看帮助信息：")
    print("   python ppt_to_image_slides.py -h")
    
    print("\n" + "=" * 60)
    print("💡 使用建议：")
    print("• 新手用户推荐使用图形化界面版本")
    print("• 批量处理或脚本自动化使用命令行版本")
    print("• PNG格式质量更好，JPG格式文件更小")
    print("• 确保已安装Microsoft PowerPoint")
    print("• 支持中文文件名和包含空格的路径")
    print("• 输出目录会自动创建")
    print("• 图片会自动缩放以填满整个幻灯片")
    print("=" * 60)

if __name__ == "__main__":
    show_usage_examples()
