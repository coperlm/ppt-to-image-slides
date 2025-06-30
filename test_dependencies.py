#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试脚本：验证所有依赖包是否正确安装
"""

def test_imports():
    """测试所有必需的模块是否能正确导入"""
    try:
        print("测试导入模块...")
        
        import win32com.client
        print("✓ win32com.client 导入成功")
        
        from pptx import Presentation
        print("✓ python-pptx 导入成功")
        
        from PIL import Image
        print("✓ Pillow 导入成功")
        
        import os
        import sys
        import argparse
        import tempfile
        import shutil
        from pathlib import Path
        print("✓ 标准库模块导入成功")
        
        print("\n所有依赖模块导入成功！")
        return True
        
    except ImportError as e:
        print(f"✗ 导入失败: {e}")
        return False

def test_powerpoint_com():
    """测试PowerPoint COM接口是否可用"""
    try:
        print("\n测试PowerPoint COM接口...")
        import win32com.client
        
        # 尝试创建PowerPoint应用程序对象
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        
        # 获取PowerPoint版本信息
        version = powerpoint.Version
        print(f"✓ PowerPoint COM接口可用 (版本: {version})")
        
        # 关闭PowerPoint
        powerpoint.Quit()
        print("✓ PowerPoint 正常关闭")
        
        return True
        
    except Exception as e:
        print(f"✗ PowerPoint COM接口测试失败: {e}")
        print("请确保已安装Microsoft PowerPoint")
        return False

def main():
    print("=" * 50)
    print("PPT转图片幻灯片工具 - 依赖测试")
    print("=" * 50)
    
    # 测试模块导入
    import_ok = test_imports()
    
    if import_ok:
        # 测试PowerPoint COM接口
        com_ok = test_powerpoint_com()
        
        if com_ok:
            print("\n🎉 所有测试通过！工具已准备就绪。")
            print("\n使用示例：")
            print("python ppt_to_image_slides.py your_presentation.pptx")
        else:
            print("\n⚠️  PowerPoint COM接口不可用，请检查PowerPoint安装")
    else:
        print("\n❌ 依赖模块导入失败，请安装所需包：")
        print("pip install pywin32 python-pptx Pillow")

if __name__ == "__main__":
    main()
