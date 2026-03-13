@echo off
chcp 65001 >nul
echo ========================================
echo 邮箱筛查工具 - 一键运行
echo ========================================
echo.

REM 检查 input 文件夹
if not exist "input\" (
    echo [错误] input 文件夹不存在，正在创建...
    mkdir input
    echo [提示] 请将 Import.xlsx 和 Export.xlsx 放入 input 文件夹
    pause
    exit
)

REM 检查必要文件
if not exist "input\Import.xlsx" (
    echo [错误] 缺少 Import.xlsx (导入版)
    echo [提示] 请将文件放入 input 文件夹
    pause
    exit
)

if not exist "input\Export.xlsx" (
    echo [错误] 缺少 Export.xlsx (导出版)
    echo [提示] 请将文件放入 input 文件夹
    pause
    exit
)

echo [检查] 文件准备就绪
echo.

REM 运行程序
邮箱筛查工具.exe

pause
