@echo off
cd /d %~dp0

echo 这个脚本用于在 Windows 32 位系统上打包。
echo 请先安装 32 位 Python 3.8 或 3.9，并勾选 Add Python to PATH。
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PYTHON_CMD=py -3-32
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set PYTHON_CMD=python
    ) else (
        echo 没有找到 Python。请安装 32 位 Python 3.8 或 3.9。
        pause
        exit /b 1
    )
)

%PYTHON_CMD% -c "import platform; print(platform.architecture()[0])"
%PYTHON_CMD% -c "import platform, sys; sys.exit(0 if platform.architecture()[0] == '32bit' else 1)"
if errorlevel 1 (
    echo 当前 Python 不是 32 位。请安装 32 位 Python，或用 py -3-32 运行。
    pause
    exit /b 1
)

%PYTHON_CMD% -m pip install --upgrade "pip<25"
%PYTHON_CMD% -m pip install "openpyxl>=3.1.0" "pyinstaller==5.13.2"
if errorlevel 1 (
    echo 依赖安装失败，请检查网络或 pip。
    pause
    exit /b 1
)

%PYTHON_CMD% -m PyInstaller --onefile --windowed --name "门店Excel智能处理工具_x86" app\main.py
if errorlevel 1 (
    echo 窗口版打包失败，请把上面的错误信息发给开发人员。
    pause
    exit /b 1
)

%PYTHON_CMD% -m PyInstaller --onefile --console --name "门店Excel智能处理工具_x86_兼容版" app\main_cli.py
if errorlevel 1 (
    echo 兼容版打包失败，请把上面的错误信息发给开发人员。
    pause
    exit /b 1
)

echo @echo off> dist\运行x86兼容版.bat
echo cd /d %%~dp0>> dist\运行x86兼容版.bat
echo 门店Excel智能处理工具_x86_兼容版.exe>> dist\运行x86兼容版.bat
echo echo.>> dist\运行x86兼容版.bat
echo echo 如果程序报错，请把 compat_error.log 或窗口内容发给开发人员。>> dist\运行x86兼容版.bat
echo pause>> dist\运行x86兼容版.bat

echo.
echo 打包完成：dist\门店Excel智能处理工具_x86.exe
echo 兼容版：dist\门店Excel智能处理工具_x86_兼容版.exe
echo 兼容版启动器：dist\运行x86兼容版.bat
pause
