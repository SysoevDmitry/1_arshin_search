# Портативная версия ФГИС АРШИН для Windows

## 📦 Portable Executable

### Что такое портативная версия?

Портативная версия Windows не требует установки:
- ✅ Не требует прав администратора
- ✅ Все зависимости включены
- ✅ Работает с USB-накопителя
- ✅ Запускается двойным кликом

---

## 🚀 Создание портативной версии

### Требования для сборки

```powershell
# Установите Python 3.9+
# https://www.python.org/downloads/

# Установка зависимостей
pip install pyinstaller
pip install -r requirements.txt
```

### Шаг 1: Создание.spec файла

Создайте файл `arshin.spec`:

```python
from PyInstaller.utils.hooks import collect_submodules
import os

a = Analysis(
    ['arshin_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config', 'config'),
        ('requirements.txt', 'requirements.txt'),
    ],
    hiddenimports=[
        'tkinter',
        'pandas',
        'openpyxl',
        'aiohttp',
        'requests',
        'loguru',
    ] + collect_submodules('pandas'),
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ФГИС_АРШИН',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='arshin.ico',  # Добавьте иконку
)
```

### Шаг 2: Сборка EXE

```powershell
# Запуск PyInstaller
pyinstaller arshin.spec

# Или быстрая сборка (без спецификации)
pyinstaller --onefile --windowed --name "ФГИС_АРШИН" --icon="arshin.ico" arshin_app.py
```

### Шаг 3: Создание портативной папки

```powershell
# Создание структуры
mkdir App_Windows\ARSHIN_Portable
cd App_Windows\ARSHIN_Portable

# Копирование EXE
copy ..\..\dist\ФГИС_АРШИН.exe .

# Копирование данных
mkdir config
copy ..\..\config\*.csv config\

# Создание скрипта запуска
@echo off
chcp 65001 >nul
start "" "ФГИС_АРШИН.exe"
```

---

## 📝 Альтернатива: NSIS Installer

### Создание установщика NSIS

```nsis
; arshin_installer.nsi
!include "MUI2.nsh"

Name "ФГИС АРШИН v2.14"
OutFile "ARSHIN_Setup.exe"
InstallDir "$PROGRAMFILES\ARSHIN"
RequestExecutionLevel user

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "Russian"

Section "Install"
    SetOutPath "$INSTDIR"
    File "dist\ФГИС_АРШИН.exe"
    File "config\*.csv"
    File "requirements.txt"
    File "README.md"
    
    CreateDirectory "$SMPROGRAMS\ARSHIN"
    CreateShortCut "$SMPROGRAMS\ARSHIN\ФГИС АРШИН.lnk" "$INSTDIR\ФГИС_АРШИН.exe"
    CreateShortCut "$DESKTOP\ФГИС АРШИН.lnk" "$INSTDIR\ФГИС_АРШИН.exe"
    
    WriteUninstaller "$INSTDIR\uninstall.exe"
SectionEnd

Section "Uninstall"
    Delete "$INSTDIR\*.*"
    RMDir "$INSTDIR"
    Delete "$SMPROGRAMS\ARSHIN\*.*"
    RMDir "$SMPROGRAMS\ARSHIN"
    Delete "$DESKTOP\ФГИС АРШИН.lnk"
SectionEnd
```

### Компиляция NSIS

```powershell
# Установите NSIS: https://nsis.sourceforge.io/Download
"C:\Program Files (x86)\NSIS\makensis.exe" arshin_installer.nsi
```

---

## 🏃 Быстрый запуск (без компиляции)

### Скрипт run_arshin.bat

```batch
@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo ФГИС АРШИН - Запуск
echo ========================================

REM Проверка Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Ошибка: Python не найден
    echo Установите Python 3.9+ с https://python.org
    pause
    exit /b 1
)

REM Создание venv если нет
if not exist "venv" (
    echo Создание виртуального окружения...
    python -m venv venv
)

REM Активация
call venv\Scripts\activate.bat

REM Установка зависимостей
echo Проверка зависимостей...
pip install -q -r requirements.txt

REM Запуск
echo Запуск приложения...
echo ========================================
python arshin_app.py

pause
```

---

## 📊 Структура портативной версии

```
ARSHIN_Portable/
├── ФГИС_АРШИН.exe      # Скомпилированное приложение
├── run_arshin.bat      # Скрипт запуска
├── config/
│   ├── exact_queries.csv
│   └── manufacturers.csv
├── exports/            # Экспортированные файлы
├── logs/               # Логи приложения
└── arshin_data.db      # База данных
```

---

## 🔧 Характеристики

| Параметр | Значение |
|----------|----------|
| Размер EXE | ~30-50 МБ |
| Размер папки | ~100-150 МБ |
| Зависимости | Включены |
| Windows | 7/8/10/11 |
| Python | Не требуется |

---

## ⚠️ Примечания

1. **Антивирус:** PyInstaller может вызывать ложные срабатывания
   - Добавьте в исключения
   - Или подпишите цифровой подписью

2. **Обновление:** Замените EXE файл на новую версию

3. **Данные:** База данных и логи сохраняются в папке с приложением

---

## 📥 Готовые файлы

В этой папке находятся:
- `build_portable.bat` - Скрипт сборки
- `run_arshin.bat` - Скрипт запуска
- `arshin.spec` - Спецификация PyInstaller

---

**Версия:** 2.14
**Дата:** 2026-03-23
