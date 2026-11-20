@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ============================================================
echo  Outlook Mail Arama ve Ek Indirme Uygulamasi v1.2
echo ============================================================
echo.

:: Python ve pywin32 kontrolu
python -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo pywin32 kutuphanesi bulunamadi.
    echo Kurulum yapiliyor...
    pip install pywin32
    if errorlevel 1 (
        echo.
        echo HATA: pywin32 kurulamadi. Lutfen manuel olarak kurun:
        echo   pip install pywin32
        echo.
        pause
        exit /b 1
    )
    echo.
    echo pywin32 basariyla kuruldu.
    echo.
)

:: Uygulamayi baslat
python outlook_mail_ekleri_gui.py

if errorlevel 1 (
    echo.
    echo Uygulama hata ile sonlandi.
    pause
)