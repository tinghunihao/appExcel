@echo off
cd /d %~dp0

where py >nul 2>nul
if %errorlevel%==0 (
    set PYTHON_CMD=py
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set PYTHON_CMD=python
    ) else (
        echo 没有找到 Python。请先安装 Python 3.10 或以上，并勾选 Add Python to PATH。
        pause
        exit /b 1
    )
)

%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖安装失败，请检查网络或 Python/pip 是否正常。
    pause
    exit /b 1
)

%PYTHON_CMD% -m pip install pyinstaller
if errorlevel 1 (
    echo PyInstaller 安装失败，请检查网络或 Python/pip 是否正常。
    pause
    exit /b 1
)

%PYTHON_CMD% -m PyInstaller --onefile --windowed --name "门店Excel智能处理工具" app\main.py
if errorlevel 1 (
    echo 打包失败，请把上面的错误信息发给开发人员。
    pause
    exit /b 1
)

echo.
echo 打包完成后，exe 在 dist 目录。
pause
