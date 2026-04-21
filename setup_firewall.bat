@echo off
chcp 65001 >nul
title فتح البورت للشبكة المحلية - نظام فنون

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║    فتح البورت 8501 للشبكة المحلية                    ║
echo ║    يحتاج صلاحيات Administrator                       ║
echo ╚══════════════════════════════════════════════════════╝
echo.

:: التحقق من صلاحيات Admin
net session >nul 2>&1
if errorlevel 1 (
    echo  ⚠  يرجى تشغيل هذا الملف كـ "Run as Administrator"
    echo.
    pause
    exit /b 1
)

echo  [1/2] إضافة قاعدة Firewall للبورت 8501...
netsh advfirewall firewall delete rule name="Fanoun Appraisal System" >nul 2>&1
netsh advfirewall firewall add rule name="Fanoun Appraisal System" dir=in action=allow protocol=TCP localport=8501

if errorlevel 1 (
    echo  ⚠  فشل إضافة القاعدة
) else (
    echo  ✓ تم فتح البورت 8501
)

echo.
echo  [2/2] التحقق من القاعدة...
netsh advfirewall firewall show rule name="Fanoun Appraisal System" | findstr "Enabled"

echo.
echo ╔══════════════════════════════════════════════════════╗
echo ║  ✅ الآن يمكن للأجهزة على الشبكة الوصول للبرنامج    ║
echo ║                                                      ║
echo ║  شغّل launcher.bat وأرسل الرابط للمستخدمين          ║
echo ╚══════════════════════════════════════════════════════╝
echo.
pause
