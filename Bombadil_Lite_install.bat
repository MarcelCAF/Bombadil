@echo off
set SCRIPT_DIR=%~dp0
set TARGET=%USERPROFILE%\Desktop\Bombadil Lite.lnk

powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%TARGET%'); $s.TargetPath = 'pythonw.exe'; $s.Arguments = '\"%SCRIPT_DIR%Bombadil_Lite.py\"'; $s.WorkingDirectory = '%SCRIPT_DIR%'; $s.IconLocation = '%SCRIPT_DIR%logo.png'; $s.Save()"

echo Verknuepfung "Bombadil Lite" wurde auf dem Desktop erstellt.
pause
