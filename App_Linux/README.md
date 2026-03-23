# Портативная версия ФГИС АРШИН для Linux

## 📦 AppImage

### Что такое AppImage?

AppImage — формат портативных приложений для Linux. Преимущества:
- ✅ Не требует установки
- ✅ Работает на большинстве дистрибутивов
- ✅ Все зависимости включены
- ✅ Запускается двойным кликом

---

## 🚀 Сборка AppImage

### Требования для сборки

```bash
# Ubuntu/Debian
sudo apt install python3 python3-venv python3-pip libgtk-3-dev

# ALT Linux
sudo apt-get install python3 python3-venv python3-pip libgtk-3-devel
```

### Шаг 1: Подготовка

```bash
cd /home/dmitry/Документы/АРШИН/Интерфейс/1_arshin_search

# Создание виртуального окружения
python3 -m venv venv
source venv/bin/activate

# Установка зависимостей
pip install -r requirements.txt
```

### Шаг 2: Создание AppDir

```bash
# Создание структуры AppDir
mkdir -p AppImage/AppDir/usr/bin
mkdir -p AppImage/AppDir/usr/lib
mkdir -p AppImage/AppDir/usr/share/applications
mkdir -p AppImage/AppDir/usr/share/icons/hicolor/256x256/apps

# Копирование приложения
cp arshin_app.py AppImage/AppDir/usr/bin/
cp -r venv AppImage/AppDir/usr/
cp requirements.txt AppImage/AppDir/usr/

# Создание файла запуска
cat > AppImage/AppDir/AppRun << 'EOF'
#!/bin/bash
HERE="$(dirname "$(readlink -f "${0}")")"
cd "$HERE/usr/bin"
source "$HERE/usr/venv/bin/activate"
exec python3 arshin_app.py "$@"
EOF

chmod +x AppImage/AppDir/AppRun
```

### Шаг 3: Создание desktop файла

```bash
cat > AppImage/AppDir/usr/share/applications/arshin.desktop << 'EOF'
[Desktop Entry]
Type=Application
Name=ФГИС АРШИН
Comment=Приложение для работы с реестром поверок
Exec=arshin_app.py
Icon=arshin
Categories=Utility;
EOF
```

### Шаг 4: Создание иконки

Сохраните иконку 256x256 как `arshin.png` в:
`AppImage/AppDir/usr/share/icons/hicolor/256x256/apps/arshin.png`

### Шаг 5: Создание AppImage

```bash
# Загрузка linuxdeploy
wget https://github.com/linuxdeploy/linuxdeploy/releases/download/continuous/linuxdeploy-x86_64.AppImage
chmod +x linuxdeploy-x86_64.AppImage

# Сборка AppImage
./linuxdeploy-x86_64.AppImage \
    --appdir AppImage/AppDir \
    --output appimage \
    --desktop-file=AppImage/AppDir/usr/share/applications/arshin.desktop \
    --icon-file=AppImage/AppDir/usr/share/icons/hicolor/256x256/apps/arshin.png
```

### Готово!

После сборки будет создан файл `ФГИС_АРШИН-x86_64.AppImage`

---

## 📝 Быстрый запуск (без AppImage)

### Скрипт запуска

```bash
#!/bin/bash
# save as: run_arshin.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Проверка виртуального окружения
if [ ! -d "venv" ]; then
    echo "Создание виртуального окружения..."
    python3 -m venv venv
fi

# Активация и запуск
source venv/bin/activate
pip install -q -r requirements.txt
python arshin_app.py
```

### Использование

```bash
chmod +x run_arshin.sh
./run_arshin.sh
```

---

## 🔧 Структура AppImage

```
AppImage/
└── AppDir/
    ├── AppRun                 # Точка входа
    ├── arshin.desktop         # Desktop файл
    ├── arshin.png             # Иконка
    └── usr/
        ├── bin/
        │   └── arshin_app.py  # Приложение
        ├── lib/               # Библиотеки Python
        └── share/
            ├── applications/
            │   └── arshin.desktop
            └── icons/
                └── hicolor/
                    └── 256x256/
                        └── apps/
                            └── arshin.png
```

---

## 📊 Характеристики

| Параметр | Значение |
|----------|----------|
| Размер | ~50-80 МБ |
| Зависимости | Включены |
| Дистрибутивы | Ubuntu, Debian, ALT Linux, Fedora, etc. |
| Python | 3.9+ (включён) |

---

## ⚠️ Примечания

1. **Безопасность:** AppImage требует права на выполнение:
   ```bash
   chmod +x ФГИС_АРШИН-x86_64.AppImage
   ```

2. **Обновление:** Замените AppImage файл на новую версию

3. **Данные:** База данных и логи сохраняются в папке с приложением

---

**Версия:** 2.6  
**Дата:** 2026-03-02
