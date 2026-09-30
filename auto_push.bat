@echo off
cd /d "D:\DDC SHIPYARD\DDC-Shipyard_tu_dong"
git add -A
git diff --cached --quiet || git commit -m "Auto update %date% %time%"
git pull --rebase --autostash origin main
git push origin main
