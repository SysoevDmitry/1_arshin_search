# 📦 Портативные версии ФГИС АРШИН

## 🗂 Структура

```
arshin_search/
├── App_Linux/           # Портативная версия для Linux
│   ├── README.md        # Документация
│   ├── build_appimage.sh    # Скрипт сборки AppImage
│   └── run_arshin.sh    # Скрипт запуска
│
└── App_Windows/         # Портативная версия для Windows
    ├── README.md        # Документация
    ├── build_portable.bat   # Скрипт сборки EXE
    └── run_arshin.bat   # Скрипт запуска
```

---

## 🐧 Linux

### Быстрый запуск

```bash
cd App_Linux
./run_arshin.sh
```

### Сборка AppImage

```bash
cd App_Linux
./build_appimage.sh
```

**Результат:** `ARSHIN_2.6-x86_64.AppImage`

### Требования

- Python 3.9+
- tkinter (`python3-tk`)
- sqlite3

---

## 🪟 Windows

### Быстрый запуск

```cmd
cd App_Windows
run_arshin.bat
```

### Сборка EXE

```cmd
cd App_Windows
build_portable.bat
```

**Результат:** `App_Windows\ARSHIN_Portable\FGIS_ARSHIN.exe`

### Требования

- Python 3.9+
- Установлен в системе или портативный

---

## 📊 Сравнение версий

| Версия | Формат | Размер | Запуск |
|--------|--------|--------|--------|
| Linux | AppImage | ~50-80 МБ | Двойной клик |
| Linux | Script | ~1 МБ | `./run_arshin.sh` |
| Windows | EXE | ~30-50 МБ | Двойной клик |
| Windows | Batch | ~1 МБ | `run_arshin.bat` |

---

## 🚀 Использование

### Перенос на другой компьютер

1. Скопируйте папку `App_Linux` или `App_Windows`
2. Запустите соответствующий скрипт
3. Приложение автоматически создаст venv и установит зависимости

### Обновление

1. Замените файлы приложения
2. Запустите скрипт сборки заново

---

## ⚠️ Примечания

### Linux

- AppImage требует права на выполнение:
  ```bash
  chmod +x ARSHIN_*.AppImage
  ```

### Windows

- При первом запуске PyInstaller может вызвать срабатывание антивируса
- Добавьте папку в исключения или подпишите приложение

---

## 📝 Логи и данные

Все данные сохраняются в основной папке проекта:
- `arshin_data.db` - база данных
- `logs/` - логи приложения
- `exports/` - экспортированные файлы
- `config/` - файлы конфигурации

---

## 🔧 Решение проблем

### Linux: "Permission denied"
```bash
chmod +x run_arshin.sh
chmod +x build_appimage.sh
```

### Windows: "Python not found"
Установите Python 3.9+ с https://python.org

### Windows: "Module not found"
```cmd
call venv\Scripts\activate.bat
pip install -r requirements.txt
```

---

## 📞 Поддержка

При возникновении проблем:
1. Проверьте логи в папке `logs/`
2. Убедитесь, что Python установлен
3. Проверьте зависимости: `pip install -r requirements.txt`

---

**Версия:** 2.6  
**Дата:** 2026-03-02  
**Платформы:** Linux (AppImage), Windows (EXE)
