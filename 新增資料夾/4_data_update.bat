@echo off
cd /d D:\synologydrive\flask

:: 拉取遠端最新版本，確保是最新的
git pull

:: 加入所有變更（新增、修改、刪除）
git add -A

:: 自動產生提交訊息（使用時間）
set timestamp=%date% %time%
git commit -m "Auto commit on %timestamp%"

:: 推送本地變更到遠端
git push

pause
