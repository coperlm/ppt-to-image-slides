@echo off
chcp 65001 >nul
echo ====================================
echo   PowerPoint转图片幻灯片工具
echo       高级选项菜单
echo ====================================
echo.
echo 选择功能：
echo.
echo [1] 启动图形化界面
echo [2] 命令行界面说明
echo [3] 安装/更新依赖包
echo [4] 运行依赖测试
echo [5] 查看使用示例
echo [6] 查看版本信息
echo [7] 退出
echo.
set /p choice="请输入选择 (1-7): "

if "%choice%"=="1" (
    echo.
    echo 启动图形化界面...
    python ppt_to_image_slides_gui.py
) else if "%choice%"=="2" (
    echo.
    echo ================================
    echo     命令行界面使用说明
    echo ================================
    echo.
    echo 基本语法：
    echo python ppt_to_image_slides.py [输入文件] [选项]
    echo.
    echo 常用示例：
    echo python ppt_to_image_slides.py presentation.pptx
    echo python ppt_to_image_slides.py "我的PPT.pptx" -o "输出.pptx" -f PNG
    echo.
    echo 参数说明：
    echo -o, --output    指定输出文件路径
    echo -f, --format    图片格式 (PNG/JPG)
    echo -t, --temp-dir  临时目录
    echo -h, --help      显示完整帮助
    echo.
    echo 查看完整帮助：
    python ppt_to_image_slides.py -h
) else if "%choice%"=="3" (
    echo.
    echo 启动依赖包安装程序...
    call install_dependencies.bat
) else if "%choice%"=="4" (
    echo.
    echo 运行依赖测试...
    python test_dependencies.py
) else if "%choice%"=="5" (
    echo.
    echo 查看使用示例...
    python examples.py
) else if "%choice%"=="6" (
    echo.
    echo ================================
    echo      版本信息
    echo ================================
    echo.
    echo PowerPoint转图片幻灯片工具 v2.0.0
    echo 发布日期: 2025年6月30日
    echo.
    echo 主要功能：
    echo • 图形化界面，简单易用
    echo • 智能DPI检测和页面尺寸保持
    echo • 支持PNG和JPG格式
    echo • 中文路径和空格支持
    echo • 实时进度显示
    echo.
    echo 系统要求：
    echo • Windows操作系统
    echo • Python 3.6+
    echo • Microsoft PowerPoint
    echo.
    echo 更多信息请查看 README.md 和 CHANGELOG.md
) else if "%choice%"=="7" (
    echo.
    echo 再见！
    exit /b 0
) else (
    echo.
    echo 无效选择，请重新运行脚本。
)

echo.
pause
