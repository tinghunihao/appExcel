@echo on
cd /d %~dp0
echo 当前目录：%cd% > build_log.txt
echo 开始时间：%date% %time% >> build_log.txt

where py >> build_log.txt 2>&1
if %errorlevel%==0 (
    set PYTHON_CMD=py
) else (
    where python >> build_log.txt 2>&1
    if %errorlevel%==0 (
        set PYTHON_CMD=python
    ) else (
        echo 没有找到 Python。请先安装 Python 3.10 或以上，并勾选 Add Python to PATH。>> build_log.txt
        type build_log.txt
        pause
        exit /b 1
    )
)

echo 使用 Python 命令：%PYTHON_CMD% >> build_log.txt
%PYTHON_CMD% --version >> build_log.txt 2>&1

%PYTHON_CMD% -m pip install -r requirements.txt >> build_log.txt 2>&1
if errorlevel 1 (
    echo 依赖安装失败。>> build_log.txt
    type build_log.txt
    pause
    exit /b 1
)

%PYTHON_CMD% -m pip install pyinstaller >> build_log.txt 2>&1
if errorlevel 1 (
    echo PyInstaller 安装失败。>> build_log.txt
    type build_log.txt
    pause
    exit /b 1
)

%PYTHON_CMD% -m PyInstaller --onefile --windowed --name "门店Excel智能处理工具" app\main.py >> build_log.txt 2>&1
if errorlevel 1 (
    echo 打包失败。>> build_log.txt
    type build_log.txt
    pause
    exit /b 1
)

echo 打包成功：dist\门店Excel智能处理工具.exe >> build_log.txt
type build_log.txt
pause
