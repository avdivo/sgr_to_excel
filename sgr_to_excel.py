# Прочитать файл 006-23(вед.18).txt.sgr как текст Windows-1251
# Данные в строках файла разделены пробелами, создать список списков
# Если элемент может быть преобразован в число, то преобразовать его в число
# Открыть файл Вывод лаборатории в Excel.xlsx, очистить в нем вкладку Ввод данных
# и вставить в нее данные из файла 006-23(вед.18).txt.sgr
# сохранить файл Вывод лаборатории в Excel.xlsx
import os
import openpyxl

# Открыть файл 006-23(вед.18).txt.sgr как текст Windows-1251
with open('006-23(вед.18).txt.sgr', 'r', encoding='cp1251') as f:
    lines = f.readlines()

dates = [i.split() for i in lines]

for line in dates:
    for i in range(len(line)):
        try:
            line[i] = float(line[i])
        except ValueError:
            pass

# Открыть файл Вывод лаборатории в Excel.xlsx
wb = openpyxl.load_workbook('Вывод лаборатории в Excel.xlsx')
sheet = wb['Ввод данных']

# Очистить в нем вкладку Ввод данных и сделать для листа текстовый формат
sheet.delete_rows(1, sheet.max_row)

# и вставить в нее данные из файла 006-23(вед.18).txt.sgr
for i in range(len(dates)):
    sheet.append(dates[i])

sheet = wb["Итог"]
wb.active = sheet  # Установить активным нужный лист

# сохранить файл Вывод лаборатории в Excel.xlsx
wb.save('Вывод лаборатории в Excel.xlsx')

# Открыть файл Вывод лаборатории в Excel.xlsx
os.startfile('Вывод лаборатории в Excel.xlsx')




