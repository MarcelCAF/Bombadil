@echo off
:: PDF-Sammler – Aufgabenplaner-Eintrag anlegen
:: Diese Datei einmalig als Administrator ausfuehren!

echo Lege Aufgabe "PDF-Sammler 21Uhr" im Windows-Aufgabenplaner an...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$action = New-ScheduledTaskAction -Execute 'C:\Users\Abfuellung 15\AppData\Local\Programs\Python\Python314\python.exe' -Argument '\"C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil\pdf_sammler.py\"'; $trigger = New-ScheduledTaskTrigger -Daily -At '21:00'; Register-ScheduledTask -TaskName 'PDF-Sammler 21Uhr' -Action $action -Trigger $trigger -RunLevel Highest -Force"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Erledigt! Das Skript laeuft ab jetzt taeglich um 21:00 Uhr.
) else (
    echo.
    echo FEHLER beim Anlegen der Aufgabe.
    echo Bitte als Administrator ausfuehren!
)

pause
