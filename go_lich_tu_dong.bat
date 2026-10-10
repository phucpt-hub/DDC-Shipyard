@echo off
schtasks /delete /f /tn "DDC Bao cao chuoi thep - trua"
schtasks /delete /f /tn "DDC Bao cao chuoi thep - chieu"
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\DDC_bao_cao_chuoi_thep.vbs" 2>nul
echo Da go lich tu dong.
pause
