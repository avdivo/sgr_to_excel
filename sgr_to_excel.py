# Прочитать полученный файл как текст Windows-1251
# Данные в строках файла разделены пробелами, создать список списков
# Если элемент может быть преобразован в число, то преобразовать его в число.
# Открыть файл sample.xlsx, очистить в нем вкладку Ввод данных
# и вставить в нее сформированные данные
# сохранить файл в .xlsx
import openpyxl


def sgr_to_excel(import_file, export_file):

    # Открыть файл как текст Windows-1251
    try:
        with open(import_file, 'r', encoding='cp1251') as f:
            lines = f.readlines()
    except Exception as e:
        raise Exception("Ошибка при открытии файла: " + str(e))

    try:
        datas = [i.split() for i in lines]

        for line in datas:
            for i in range(len(line)):
                try:
                    line[i] = float(line[i])
                except ValueError:
                    pass
    except Exception as e:
        raise Exception("Ошибка при обработке файла: " + str(e))

    # Открыть файл Вывод лаборатории в Excel.xlsx
    try:
        wb = openpyxl.load_workbook('sample.xlsx')
        sheet = wb['Ввод данных']
    except Exception as e:
        raise Exception("Ошибка при открытии шаблона: " + str(e))

    try:
        # Очистить в нем вкладку Ввод данных и сделать для листа текстовый формат
        sheet.delete_rows(1, sheet.max_row)

        # и вставить в нее данные из файла 006-23(вед.18).txt.sgr
        for i in range(len(datas)):
            sheet.append(datas[i])

        sheet = wb["Итог"]
        wb.active = sheet  # Активация нужного листа
    except Exception as e:
        raise Exception("Ошибка при изменении шаблона: " + str(e))

    # сохранить файл Вывод лаборатории в Excel.xlsx
    try:
        wb.save(export_file)
    except Exception as e:
        return "Ошибка при сохранении файла: " + str(e)

    return "Успешно экспортирован!"
