@echo off
chcp 65001
setlocal enabledelayedexpansion

git rm -r --cache .git
git rm -r --cache .idea
git rm -r --cache .vscode
git rm -r --cache __pycache__
git rm -r --cache build
git rm -r --cache out
git rm -r --cache 一键安装所需第三方库

REM 第一步：输入需合并的commitID
set /p commit_id=请输入需合并的commitID:

REM 第二步：软重置到该commit，不影响工作区
git reset --soft %commit_id%
if errorlevel 1 (
    echo git reset --soft 失败，请检查commitID是否正确。
    pause
    exit /b
)

REM 第三步：add已跟踪文件
git add -u

REM 第四步：输入commit信息
set /p commit_msg=请输入本次提交的信息:
git commit -m "%commit_msg%"
if errorlevel 1 (
    echo git commit 失败，请检查是否有改动。
    pause
    exit /b
)

REM 第五步：强制推送
git push --force
if errorlevel 1 (
    echo git push --force 失败，请检查网络或权限。
    pause
    exit /b
)

echo 操作完成！
pause