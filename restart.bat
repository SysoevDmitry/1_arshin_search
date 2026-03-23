@echo off
:start
echo Запуск приложения ФГИС Аршин...
python arshin_app.py
echo Приложение завершило работу. Перезапуск через 5 секунд...
timeout /t 5
goto start
