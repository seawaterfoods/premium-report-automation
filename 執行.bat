@echo off
chcp 65001 >nul

echo ========================================
echo   自動執行 Build + Run
echo ========================================
echo.

REM ===== 指定 Java =====
set "JAVA_EXE=C:\Users\user\.jdks\corretto-17.0.18\bin\java.exe"

REM ===== 檢查 Java =====
if not exist "%JAVA_EXE%" (
    echo [錯誤] 找不到 Java：
    echo %JAVA_EXE%
    pause
    exit /b 1
)

REM ===== 設定 JAVA_HOME =====
for %%i in ("%JAVA_EXE%") do set "JAVA_HOME=%%~dpi"
set "JAVA_HOME=%JAVA_HOME:~0,-1%"
for %%i in ("%JAVA_HOME%") do set "JAVA_HOME=%%~dpi"
set "JAVA_HOME=%JAVA_HOME:~0,-1%"

set "PATH=%JAVA_HOME%\bin;%PATH%"

echo 使用 Java：
"%JAVA_EXE%" -version
echo.

REM ========================================
REM Step 1. Build
REM ========================================
echo [Step 1] 編譯中...
call mvnw.cmd clean package -DskipTests -q

if %ERRORLEVEL% neq 0 (
    echo ❌ 編譯失敗，停止執行
    pause
    exit /b 1
)

echo ✅ 編譯成功！
echo.

REM ========================================
REM Step 2. Run
REM ========================================
echo [Step 2] 執行程式...

set "JAR_FILE=target\premium-report-automation-0.0.1-SNAPSHOT.jar"

if not exist "%JAR_FILE%" (
    echo [錯誤] 找不到 JAR：
    echo %JAR_FILE%
    pause
    exit /b 1
)

echo 開始執行...
echo %date% %time%
echo.

"%JAVA_EXE%" -jar "%JAR_FILE%" ^
 --spring.config.additional-location=file:./config/ ^
 --logging.level.root=INFO ^
 > run.log 2>&1

echo.
echo 結束時間：
echo %date% %time%
echo.

REM ========================================
REM Step 3. 檢查輸出
REM ========================================
if exist output (
    echo ✅ 已產生 output 資料夾
) else (
    echo ⚠️ 沒看到 output，請查看 run.log
)

if exist run.log (
    echo 📄 log 檔：run.log
)

echo.

REM ========================================
REM Step 4. 解析 report.log ERROR
REM ========================================
echo ===== 所有 ERROR 詳細內容 =====

set "REPORT_LOG=D:\workspace\premium-report-automation\output\report.log"

if exist "%REPORT_LOG%" (
    echo 發現 report.log，開始解析...
    echo.

    setlocal enabledelayedexpansion
    set "printing="
    set "foundError="

    for /f "usebackq delims=" %%L in ("%REPORT_LOG%") do (
        set "line=%%L"

        REM 判斷是否為新 log 開頭 (時間戳)
        echo !line! | findstr /r "^[0-9][0-9][0-9][0-9]-" >nul
        if !errorlevel! == 0 (
            if defined printing (
                echo ----------------------------------------
                set "printing="
            )
        )

        REM 如果遇到 ERROR，開始印
        echo !line! | findstr /i "ERROR" >nul
        if !errorlevel! == 0 (
            set "printing=1"
            set "foundError=1"
            echo ----------------------------------------
        )

        REM 正在 ERROR 區塊就持續印
        if defined printing (
            echo !line!
        )
    )

    REM 檔案結尾補分隔
    if defined printing (
        echo ----------------------------------------
    )

    echo.
    if defined foundError (
        echo ❌ 執行完成，但有發現 ERROR
    ) else (
        echo ✅ 正確執行完畢，無發現錯誤
    )

    endlocal
) else (
    echo ⚠️ 找不到 report.log：
    echo %REPORT_LOG%
)

echo.
pause