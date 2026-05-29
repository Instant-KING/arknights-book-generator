@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================
echo   明日方舟剧情书籍生成器 - 打包为exe
echo ==========================================
echo.
echo 正在检查 PyInstaller ...
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 未安装 PyInstaller，正在安装...
    python -m pip install pyinstaller
)
echo.
echo 开始打包（这可能需要几分钟）...
python -m PyInstaller --onefile --windowed --name "明日方舟书籍生成器" --add-data "_cover_assets;_cover_assets" --hidden-import docx --hidden-import lxml book_gui.py
echo.
echo 复制 exe 到当前目录...
copy /y "dist\明日方舟书籍生成器.exe" "明日方舟书籍生成器.exe"
echo.
echo ==========================================
echo   打包完成！
echo   发给朋友的文件夹需要包含:
echo     - 明日方舟书籍生成器.exe
echo     - arknights_dialogue\  (数据)
echo     - _cover_assets\       (封面素材)
echo ==========================================
pause
