@echo off
chcp 65001 >nul
echo Tao lich tu dong day bao cao len GitHub moi ngay luc 11:45 va 17:15...
schtasks /create /f /tn "DDC Bao cao chuoi thep - trua" /sc daily /st 11:45 /tr "wscript.exe \"%~dp0auto_push_an.vbs\""
schtasks /create /f /tn "DDC Bao cao chuoi thep - chieu" /sc daily /st 17:15 /tr "wscript.exe \"%~dp0auto_push_an.vbs\""
echo.
echo Xong. Muon doi gio: mo Task Scheduler, tim "DDC Bao cao chuoi thep".
echo Muon go bo: chay go_lich_tu_dong.bat
pause
