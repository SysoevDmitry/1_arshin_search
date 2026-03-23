#!/bin/bash
# Скрипт сборки AppImage для ФГИС АРШИН
# Использование: ./build_appimage.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_NAME="ARSHIN"
VERSION="2.6"

echo "========================================"
echo "Сборка AppImage: $APP_NAME v$VERSION"
echo "========================================"

# Переход в директорию проекта
cd "$SCRIPT_DIR/.."

# Создание директории AppImage
echo "[1/6] Создание структуры AppDir..."
rm -rf AppImage
mkdir -p AppImage/AppDir/usr/bin
mkdir -p AppImage/AppDir/usr/lib
mkdir -p AppImage/AppDir/usr/share/applications
mkdir -p AppImage/AppDir/usr/share/icons/hicolor/256x256/apps
mkdir -p AppImage/AppDir/usr/share/icons/hicolor/512x512/apps

# Копирование приложения
echo "[2/6] Копирование файлов приложения..."
cp arshin_app.py AppImage/AppDir/usr/bin/
cp requirements.txt AppImage/AppDir/usr/

# Создание виртуального окружения
echo "[3/6] Создание виртуального окружения..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    source venv/bin/activate
    pip install --upgrade pip
    pip install -r requirements.txt
else
    source venv/bin/activate
fi

# Копирование venv
cp -r venv AppImage/AppDir/usr/

# Создание AppRun
echo "[4/6] Создание AppRun..."
cat > AppImage/AppDir/AppRun << 'EOF'
#!/bin/bash
HERE="$(dirname "$(readlink -F "${0}")")"
cd "$HERE/usr/bin"

# Активация виртуального окружения
if [ -f "$HERE/usr/venv/bin/activate" ]; then
    source "$HERE/usr/venv/bin/activate"
fi

# Запуск приложения
exec python3 arshin_app.py "$@"
EOF

chmod +x AppImage/AppDir/AppRun

# Создание desktop файла
echo "[5/6] Создание desktop файла..."
cat > AppImage/AppDir/arshin.desktop << 'EOF'
[Desktop Entry]
Type=Application
Name=ФГИС АРШИН
GenericName=Поиск и сбор данных ФГИС АРШИН
Comment=Приложение для работы с реестром поверок средств измерений
Exec=arshin_app.py
Icon=arshin
Categories=Utility;Science;
Keywords=metrology;registry;verification;
StartupNotify=true
EOF

# Копирование desktop файла в стандартное место
cp AppImage/AppDir/arshin.desktop AppImage/AppDir/usr/share/applications/

# Создание иконки (заглушка)
echo "[6/6] Создание иконки..."
# Создаём простую иконку (замените на реальную иконку 256x256)
if [ -f "arshin_icon.png" ]; then
    cp arshin_icon.png AppImage/AppDir/usr/share/icons/hicolor/256x256/apps/arshin.png
    cp arshin_icon.png AppImage/AppDir/usr/share/icons/hicolor/512x512/apps/arshin.png
    cp arshin_icon.png AppImage/AppDir/arshin.png
else
    # Создаём цветной квадрат как заглушку
    convert -size 256x256 xc:#4472C4 AppImage/AppDir/usr/share/icons/hicolor/256x256/apps/arshin.png 2>/dev/null || \
    echo "Warning: ImageMagick not found, using placeholder"
fi

# Загрузка linuxdeploy если нет
if [ ! -f "linuxdeploy-x86_64.AppImage" ]; then
    echo "Загрузка linuxdeploy..."
    wget -q https://github.com/linuxdeploy/linuxdeploy/releases/download/continuous/linuxdeploy-x86_64.AppImage
    chmod +x linuxdeploy-x86_64.AppImage
fi

# Сборка AppImage
echo ""
echo "Сборка AppImage..."
./linuxdeploy-x86_64.AppImage \
    --appdir AppImage/AppDir \
    --output appimage \
    --desktop-file=AppImage/AppDir/arshin.desktop \
    --icon-file=AppImage/AppDir/arshin.png 2>/dev/null || \
{
    echo ""
    echo "⚠️  linuxdeploy не найден или не работает."
    echo "   Создаю упрощённую версию AppImage..."
    
    # Упрощённая сборка без linuxdeploy
    cd AppImage
    
    # Создание AppImage вручную
    cat > make_appimage.sh << 'INNER_EOF'
#!/bin/bash
ARCHIVE="app_archive.tar.gz"
RUNTIME="AppImage_runtime"

# Создание архива AppDir
tar -czf "$ARCHIVE" AppDir

# Загрузка runtime AppImage
if [ ! -f "$RUNTIME" ]; then
    wget -q https://github.com/AppImage/AppImageKit/releases/download/continuous/AppImageKit-x86_64 -O "$RUNTIME"
fi

# Создание финального AppImage
cat "$RUNTIME" "$ARCHIVE" > "../${APP_NAME}_${VERSION}-x86_64.AppImage"
chmod +x "../${APP_NAME}_${VERSION}-x86_64.AppImage"

# Очистка
rm -f "$ARCHIVE" "$RUNTIME"
INNER_EOF

    chmod +x make_appimage.sh
    ./make_appimage.sh
    
    cd ..
}

# Финальное сообщение
echo ""
echo "========================================"
echo "✅ Сборка завершена!"
echo "========================================"
echo ""
echo "Файл: $(ls -1 ${APP_NAME}_*.AppImage 2>/dev/null | head -1)"
echo "Размер: $(du -h ${APP_NAME}_*.AppImage 2>/dev/null | cut -f1 | head -1)"
echo ""
echo "Для запуска:"
echo "  chmod +x ${APP_NAME}_*.AppImage"
echo "  ./${APP_NAME}_*.AppImage"
echo ""
