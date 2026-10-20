@echo off
rem Change status: MMAN_*.txt + status list (.xls/.xlsx/.csv) in input\  ->  input\MMAN_EDI.txt
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (where python >nul 2>nul && set "PY=python")
if not defined PY (
  echo [ERROR] ไม่พบ Python - ติดตั้งจาก https://www.python.org/downloads/ และติ๊ก Add python.exe to PATH
  pause
  exit /b 1
)

%PY% -c "import pandas, xlrd, openpyxl" 2>nul || (
  echo กำลังติดตั้งไลบรารีที่ต้องใช้ ^(ครั้งแรกเท่านั้น^)...
  %PY% -m pip install --quiet --disable-pip-version-check -r requirements.txt || (
    echo [ERROR] ติดตั้งไลบรารีไม่สำเร็จ - ตรวจการเชื่อมต่ออินเทอร์เน็ต
    pause
    exit /b 1
  )
)

set "SHEDYES=0145"
set "SHEDNO=0122"
set /p "SHEDYES=SHED NO. สำหรับ B/L ที่มี Status [กด Enter = 0145]: "
set /p "SHEDNO=SHED NO. สำหรับ B/L อื่น [กด Enter = 0122 / พิมพ์ = ใช้ค่าเดียวกับข้างบน]: "
if "%SHEDNO%"=="=" set "SHEDNO=%SHEDYES%"
echo.

%PY% change-status\scripts\change_status.py "%~dp0." --shed-yes %SHEDYES% --shed-no %SHEDNO%
if errorlevel 1 (
  echo.
  echo [ERROR] ทำไม่สำเร็จ - อ่านข้อความด้านบน
) else (
  echo.
  echo เสร็จแล้ว: เปิดไฟล์ input\MMAN_EDI.txt
)
pause
