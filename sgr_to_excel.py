# Прочитать файл 006-23(вед.18).txt.sgr как текст Windows-1251
# Данные в строках файла разделены пробелами, создать список списков
# Если элемент может быть преобразован в число, то преобразовать его в число
# Если значение 0 то удалить его, если значение дробная часть числа 0 то удалить ее
# преобразовать все данные в текстовые
# открыть файл Вывод лаборатории в Excel.xlsx, очистить в нем вкладку Ввод данных
# и вставить в нее данные из файла 006-23(вед.18).txt.sgr
# сохранить файл Вывод лаборатории в Excel.xlsx

import openpyxl

# Открыть файл 006-23(вед.18).txt.sgr как текст Windows-1251
with open('006-23(вед.18).txt.sgr', 'r', encoding='cp1251') as f:
    lines = f.readlines()

dates = [i.split() for i in lines]

for line in dates:
    for i in range(len(line)):
        try:
            out = float(line[i])
            if out - int(out) == 0:
                out = int(out)
            line[i] = str(out).replace('.', ',')
        except ValueError:
            pass

# Открыть файл Вывод лаборатории в Excel.xlsx
wb = openpyxl.load_workbook('Вывод лаборатории в Excel.xlsx')
sheet = wb['Ввод данных']

# Очистить в нем вкладку Ввод данных
sheet.delete_rows(1, sheet.max_row)

# и вставить в нее данные из файла 006-23(вед.18).txt.sgr
for i in range(len(dates)):
    sheet.append(dates[i])

# сохранить файл Вывод лаборатории в Excel.xlsx
wb.save('Вывод лаборатории в Excel.xlsx')






