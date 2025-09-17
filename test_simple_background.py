#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试简化版背景功能
"""

import os
import sys
import traceback

def test_simple_background():
    """测试简化版背景转换功能"""
    try:
        print("开始测试简化版背景转换功能...")
        
        # 导入简化版转换器
        from ppt_to_image_slides_simple_background import PPTToImageSlidesSimple
        
        # 检查测试文件
        test_ppt = "测试PPT.pptx"
        if not os.path.exists(test_ppt):
            print(f"错误：测试文件 {test_ppt} 不存在")
            return False
        
        print(f"找到测试文件: {test_ppt}")
        
        # 创建转换器实例
        converter = PPTToImageSlidesSimple()
        print("转换器实例创建成功")
        
        # 设置日志回调，将日志输出到控制台
        def console_log(message):
            print(f"[转换日志] {message}")
        
        # 临时替换log方法
        converter.log = console_log
        
        # 准备输出文件名
        output_file = "测试输出_简化背景版本.pptx"
        
        print(f"开始转换，输出文件: {output_file}")
        
        # 执行转换 - 直接调用转换方法，绕过GUI
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
    print("简化版背景功能测试")
    print("=" * 60)
    
    success = test_simple_background()
    
    print("=" * 60)
    if success:
        print("测试结果: ✓ 成功")
    else:
        print("测试结果: ✗ 失败")
    print("=" * 60)
