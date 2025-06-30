@echo off
chcp 65001 >nul
echo ====================================
echo   依赖包安装脚本
echo ====================================
echo.
echo 此脚本将安装PowerPoint转图片幻灯片工具所需的依赖包
echo.

REM 检查Python是否可用
echo [1/4] 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误：未找到Python环境！
    echo.
    echo 请先安装Python，并确保：
    echo 1. Python版本为3.6或更高
    echo 2. Python已添加到系统PATH环境变量
    echo 3. 可以在命令行中运行 python 命令
    echo.
    echo 下载地址：https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo ✅ Python环境检查通过: %PYTHON_VERSION%
echo.

REM 检查pip是否可用
echo [2/4] 检查pip包管理器...
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误：pip不可用！
    echo 请重新安装Python，确保包含pip。
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python -m pip --version') do set PIP_VERSION=%%i
echo ✅ pip检查通过: %PIP_VERSION%
echo.

REM 升级pip
echo [3/4] 升级pip到最新版本...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo ⚠️  pip升级失败，但继续安装依赖包...
) else (
    echo ✅ pip升级完成
)
echo.

REM 安装依赖包
echo [4/4] 安装依赖包...
echo.
echo 正在安装所需的Python包：
echo - pywin32 (Windows COM接口支持)
echo - python-pptx (PowerPoint文件操作)
echo - Pillow (图片处理)
echo.

python -m pip install pywin32 python-pptx Pillow
if errorlevel 1 (
    echo.
    echo ❌ 依赖包安装失败！
    echo.
    echo 可能的解决方案：
    echo 1. 确保网络连接正常
    echo 2. 尝试使用管理员权限运行此脚本
    echo 3. 手动安装：pip install pywin32 python-pptx Pillow
    echo.
    pause
    exit /b 1
)

echo.
echo ✅ 所有依赖包安装完成！
echo.

REM 运行依赖测试
echo 正在运行依赖测试...
echo.
python test_dependencies.py
if errorlevel 1 (
    echo.
    echo ⚠️  依赖测试发现问题，请检查上面的错误信息。
) else (
    echo.
    echo 🎉 恭喜！所有依赖都已正确安装和配置。
    echo.
    echo 现在可以：
    echo 1. 双击 start.bat 启动图形化界面
    echo 2. 或在命令行中运行：python ppt_to_image_slides_gui.py
)

echo.
echo 安装完成，按任意键退出...
pause >nul
