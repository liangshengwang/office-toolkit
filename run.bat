@echo off
chcp 65001 >nul
title 办公文件工具箱 v1.0

echo ==============================================
echo  办公文件工具箱 v1.0
echo  - 文件分类整理
echo  - 办公模板库（日报/周报/会议纪要等14个模板）
echo  - 文件格式转换
echo ==============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 未找到Python！请先安装 Python 3.8+
    pause
    exit /b 1
)

:: Check and install dependencies
echo [检查] Flask...
python -c "import flask" 2>nul || (
    echo [安装] Flask...
    pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org flask
)

echo [检查] Pillow (图片处理)...
python -c "from PIL import Image; print('  Pillow OK')" 2>nul || (
    echo [安装] Pillow...
    pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org pillow
)

echo [检查] python-docx (Word处理)...
python -c "import docx; print('  python-docx OK')" 2>nul || (
    echo [安装] python-docx...
    pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org python-docx
)

echo [检查] PyMuPDF (PDF处理)...
python -c "import fitz; print('  PyMuPDF OK')" 2>nul || (
    echo [安装] PyMuPDF...
    pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org PyMuPDF
)

echo [检查] python-pptx (PPT处理)...
python -c "from pptx import Presentation; print('  python-pptx OK')" 2>nul || (
    echo [安装] python-pptx...
    pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org python-pptx
)

echo.
echo [启动] 正在启动服务器...
echo.
echo 访问地址: http://127.0.0.1:18700
echo ==============================================
echo.

python server.py
if %errorlevel% neq 0 (
    echo [ERROR] 启动失败，请查看上方错误信息
    pause
)
