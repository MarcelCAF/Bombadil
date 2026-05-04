@echo off
echo ============================================================
echo  Bombadil - Pakete installieren
echo ============================================================
echo.

echo Benutze Python:
python --version
echo.

echo [1/5] pandas + openpyxl (Excel lesen/schreiben)...
python -m pip install pandas openpyxl
echo.

echo [2/5] Google Drive API...
python -m pip install google-api-python-client google-auth
echo.

echo [3/5] Google OAuth2 (fuer Drive-Upload mit eigenem Google-Konto)...
python -m pip install google-auth-oauthlib
echo.

echo [4/5] tzdata (Zeitzone Windows-Fix)...
python -m pip install tzdata
echo.

echo [5/5] tksheet (Excel-artige Tabellen in Bombadil)...
python -m pip install tksheet
echo.

echo [6/6] python-dotenv (Konfigurationsdatei .env)...
python -m pip install python-dotenv
echo.

echo ============================================================
echo  Fertig! Bombadil kann jetzt gestartet werden.
echo ============================================================
pause
