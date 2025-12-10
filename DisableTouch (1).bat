@echo off
chcp 65001 >nul
if not "%1"=="admin" (powershell start -verb runas '%0' admin & exit /b)

echo Поиск сенсорного экрана...
powershell -Command "Get-PnpDevice -FriendlyName '*HID-совместимый сенсорный экран*' | Disable-PnpDevice -Confirm:$false" >nul 2>&1

if %errorlevel% equ 0 (
    echo Сенсорное управление успешно отключено
) else (
    echo ОШИБКА: Не удалось найти/отключить сенсорный экран
)
timeout /t 3