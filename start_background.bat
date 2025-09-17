@echo off
chcp 65001 >nul
echo ====================================
echo   PowerPoint转图片幻灯片工具 - 背景版
echo ====================================
echo.
echo 新功能：图片背景填充模式
echo 正在启动背景版图形化界面...
echo.

REM 检查Python是否可用
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python环境！
    echo 请确保已安装Python并添加到系统PATH。
    echo.
    pause
    exit /b 1
)

REM 启动背景版图形化界面
python ppt_to_image_slides_background_gui.py

REM 如果程序正常退出，显示完成信息
if %errorlevel% equ 0 (
    echo.
    echo 程序已正常退出。
) else (
    echo.
    echo 程序运行遇到问题，错误代码: %errorlevel%
    echo 如果是依赖包问题，请运行 install_dependencies.bat 安装依赖。
)

echo.
echo 按任意键退出...
pause >nul
