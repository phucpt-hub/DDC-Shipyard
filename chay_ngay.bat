@echo off
chcp 65001 >nul
echo Dang chay cap nhat ngay (quet nesting + day len GitHub)...
call "%~dp0auto_push.bat"
echo.
echo Xong. Xem chi tiet trong nhat_ky_day_len.txt
pause
