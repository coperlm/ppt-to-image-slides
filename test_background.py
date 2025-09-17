#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试背景功能的简化脚本
"""

import sys
import os

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_background_functionality():
    """测试背景功能"""
    try:
        from ppt_to_image_slides_background_gui import PPTToImageSlidesBackground
        print("✓ 背景版模块导入成功")
        
        # 测试创建转换器实例
        converter = PPTToImageSlidesBackground()
        print("✓ 转换器实例创建成功")
        
        print("✓ 所有基本测试通过！")
        return True
        
    except Exception as e:
        print(f"✗ 测试失败: {e}")
        return False

if __name__ == "__main__":
    print("=" * 50)
    print("PPT背景功能测试")
    print("=" * 50)
    
    if test_background_functionality():
        print("\n🎉 测试成功！背景功能准备就绪。")
    else:
        print("\n❌ 测试失败！请检查代码。")
