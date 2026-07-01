@echo off
REM ============================================================
REM  Automatischer Bombadil-Push (geplant via Task Scheduler)
REM  Committet Bombadil.py (Hook erhoeht Version) und pusht.
REM  Schreibt alles in push_auto.log zur Nachkontrolle.
REM ============================================================
cd /d "C:\Users\Abfuellung 15\Documents\Marcels Skripts\Bombadil"

set LOG=push_auto.log
echo. >> "%LOG%"
echo ===== Auto-Push %DATE% %TIME% ===== >> "%LOG%"

git add Bombadil.py >> "%LOG%" 2>&1
git commit -m "JSON-Dateien in Unterordner + Auto-Update-Backups in _alte_versionen" -m "Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>" >> "%LOG%" 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo FEHLER beim Commit ^(evtl. nichts zu committen^) - kein Push. >> "%LOG%"
    goto :ende
)

git push origin master >> "%LOG%" 2>&1
if %ERRORLEVEL% EQU 0 (
    echo ERFOLG: Push abgeschlossen. >> "%LOG%"
) else (
    echo FEHLER beim Push - bitte manuell pruefen. >> "%LOG%"
)

:ende
echo ===== Ende %DATE% %TIME% ===== >> "%LOG%"
