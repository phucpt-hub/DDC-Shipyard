@echo off
chcp 65001 >nul
echo Cai tu dong cap nhat bao cao chuoi thep:
echo  - Moi lan mo may (sau 3 phut, cho o mang Z: ket noi)
echo  - Hang ngay luc 11:45 va 17:15
schtasks /create /f /tn "DDC Bao cao chuoi thep - trua" /sc daily /st 11:45 /tr "wscript.exe \"%~dp0auto_push_an.vbs\""
schtasks /create /f /tn "DDC Bao cao chuoi thep - chieu" /sc daily /st 17:15 /tr "wscript.exe \"%~dp0auto_push_an.vbs\""
copy /y "%~dp0khi_mo_may.vbs" "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\DDC_bao_cao_chuoi_thep.vbs" >nul
echo.
echo Da cai xong. Dang chay lan dau ngay bay gio...
call "%~dp0auto_push.bat"
echo.
echo Xong. Nhat ky: nhat_ky_day_len.txt
echo Muon go bo: chay go_lich_tu_dong.bat
pause
