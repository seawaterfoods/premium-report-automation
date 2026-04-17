@echo off
chcp 65001 >nul
echo ========================================
echo  產險業務保費統計自動化系統
echo ========================================
echo.

set JAR_FILE=target\premium-report-automation-0.0.1-SNAPSHOT.jar

if not exist "%JAR_FILE%" (
    echo [錯誤] 找不到 JAR 檔案: %JAR_FILE%
    echo 請先執行 build.bat 進行編譯
    pause
    exit /b 1
)

java -jar "%JAR_FILE%" --spring.config.additional-location=file:./config/

echo.
echo ========================================
echo  執行完成，請查看 output 資料夾
echo ========================================
pause
