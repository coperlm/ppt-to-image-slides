#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试最终版本的背景功能
"""

import os
import sys
import traceback

def test_background_conversion():
    """测试背景转换功能"""
    try:
        print("开始测试背景转换功能...")
        
        # 导入背景版本的转换器
        from ppt_to_image_slides_background_gui import PPTToImageSlidesBackground
        
        # 检查测试文件
        test_ppt = "测试PPT.pptx"
        if not os.path.exists(test_ppt):
            print(f"错误：测试文件 {test_ppt} 不存在")
            return False
        
        print(f"找到测试文件: {test_ppt}")
        
        # 创建转换器实例
        converter = PPTToImageSlidesBackground()
        print("转换器实例创建成功")
        
        # 执行转换 - 使用简单的日志回调
        def simple_log(message):
            print(f"转换日志: {message}")
        
        converter.log = simple_log
        
        # 准备输出文件名
        output_file = "测试输出_背景版本.pptx"
        
        print(f"开始转换，输出文件: {output_file}")
        
        # 执行转换
        success = converter.convert_ppt_to_image_slides(test_ppt, output_file)
        
        if success:
            print(f"✓ 转换成功！输出文件: {output_file}")
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"✓ 输出文件存在，大小: {file_size} 字节")
                return True
            else:
                print("✗ 输出文件不存在")
                return False
        else:
            print("✗ 转换失败")
            return False
            
    except Exception as e:
        print(f"✗ 测试过程中出现异常: {e}")
        print(f"完整错误信息:")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("背景功能最终测试")
    print("=" * 60)
    
    success = test_background_conversion()
    
    print("=" * 60)
    if success:
        print("测试结果: ✓ 成功")
    else:
        print("测试结果: ✗ 失败")
    print("=" * 60)
