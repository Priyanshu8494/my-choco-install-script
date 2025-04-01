@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0DownloadOffice2024.ps1" -Branch "ProPlus2024Volume" -Channel "PerpetualVL2024" -Components "Word", "Excel", "PowerPoint"
pause
