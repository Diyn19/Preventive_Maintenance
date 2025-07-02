@echo off
setlocal enabledelayedexpansion

cd /d D:\flask

echo [1/6] 執行 Excel_Edge.py...
python "Excel_Edge.py"
if errorlevel 1 goto error

echo [2/6] 執行 run_update2.py...
python "run_update2.py"
if errorlevel 1 goto error

echo [3/6] 執行 data_updw.py...
python "data_updw.py"
if errorlevel 1 goto error

echo [4/6] 啟動 save_excel.exe...
start /wait "" "save_excel.exe"
if errorlevel 1 goto error

echo [5/6] 進行 Git 操作...
cd /d D:\flask

git pull
if errorlevel 1 goto error

git add -A

:: 建立 timestamp（格式：YYYY-MM-DD_HH:MM:SS）
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
    set y=%%c
    set m=%%a
    set d=%%b
)
set timestamp=%y%-%m%-%d_%time%
git commit -m "Auto commit on %timestamp%"
if errorlevel 1 goto error

git push
if errorlevel 1 goto error

echo [6/6] 所有程序已完成！
pause
exit /b

:error
echo 錯誤：某個步驟執行失敗，流程中止。
pause
exit /b
