@echo off
chcp 65001 >nul
echo 正在編譯專案...
call mvnw.cmd clean package -DskipTests -q
if %ERRORLEVEL% equ 0 (
    echo 編譯成功！
) else (
    echo 編譯失敗，請檢查錯誤訊息
)
pause
