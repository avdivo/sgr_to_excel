# Конвертор SGR TO EXCEL
(это не для SimCity))

Программа для получения сводной таблицы
лабораторных определений физико-механических свойств грунтов

В разделе Импорт выбрать папку с исходными данными.

В списке отобразятся все файлы sgr из выбранной папки.

В разделе экспорт можно указать куда будут сохранены Excel файлы.

Клавиша Экспорт начнет обработку выбранных файлов.
Имена получаемых файлов будут такими же как у входящих.

Рядом с запускаемым файлом должен быть шаблон исходящих документов:
sample.xlsx. Если его нет программа восстановил исходный шаблон при запуске.
Изменения, вносимые в шаблон будут отражаться в исходящих документах.

Для получения exe файла нужно выполнить в консоли команду:

pyinstaller --onefile --windowed --add-data "sample.xlsx;." --add-data "stic
ker.ico;." -i "sticker.ico" -n "Sgr to Excel" interface.py

В папке dist будет exe файл.