@echo off
chcp 65001 >nul
title تثبيت متطلبات نظام تقييم فنون

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║        نظام تقييم الأداء - مجموعة شركات فنون        ║
echo ║              تثبيت المتطلبات - الإعداد الأول         ║
echo ╚══════════════════════════════════════════════════════╝
echo.

:: ── التحقق من وجود Python ──
echo [1/5] التحقق من Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ⚠  Python غير مثبت على هذا الجهاز!
    echo.
    echo  يرجى تحميل Python من:
    echo  https://www.python.org/downloads/
    echo.
    echo  تأكد من تفعيل خيار "Add Python to PATH" أثناء التثبيت
    echo.
    pause
    exit /b 1
)
python --version
echo  ✓ Python موجود

:: ── ترقية pip ──
echo.
echo [2/5] تحديث pip...
python -m pip install --upgrade pip --quiet
echo  ✓ pip محدّث

:: ── تثبيت المكتبات ──
echo.
echo [3/5] تثبيت المكتبات المطلوبة...
echo  (قد يستغرق هذا 2-5 دقائق في أول مرة)
echo.

python -m pip install streamlit==1.32.0       --quiet
echo  ✓ streamlit

python -m pip install pandas==2.2.1           --quiet
echo  ✓ pandas

python -m pip install openpyxl==3.1.2         --quiet
echo  ✓ openpyxl

python -m pip install plotly==5.20.0          --quiet
echo  ✓ plotly

python -m pip install reportlab==4.1.0        --quiet
echo  ✓ reportlab

python -m pip install Pillow==10.3.0          --quiet
echo  ✓ Pillow

:: ── إنشاء ملف الإعدادات ──
echo.
echo [4/5] إنشاء ملف الإعدادات...

if not exist ".streamlit" mkdir .streamlit

(
echo [server]
echo headless = true
echo port = 8501
echo address = "0.0.0.0"
echo enableCORS = false
echo enableXsrfProtection = false
echo.
echo [browser]
echo gatherUsageStats = false
echo serverAddress = "localhost"
echo serverPort = 8501
echo.
echo [theme]
echo base = "light"
echo primaryColor = "#1E3A8A"
echo backgroundColor = "#FFFFFF"
echo secondaryBackgroundColor = "#F8FAFF"
echo textColor = "#1E293B"
echo font = "sans serif"
) > .streamlit\config.toml

echo  ✓ ملف الإعدادات جاهز

:: ── التحقق النهائي ──
echo.
echo [5/5] التحقق من التثبيت...
python -c "import streamlit, pandas, openpyxl, plotly, reportlab; print('  ✓ جميع المكتبات مثبتة بنجاح')"
if errorlevel 1 (
    echo  ⚠ حدث خطأ في التثبيت - حاول مرة أخرى
    pause
    exit /b 1
)

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║           ✅ تم الإعداد بنجاح!                       ║
echo ║                                                      ║
echo ║  الآن شغّل البرنامج عبر:  launcher.bat              ║
echo ╚══════════════════════════════════════════════════════╝
echo.
pause
