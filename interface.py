import os
from tkinter import *
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from configparser import ConfigParser

from sgr_to_excel import sgr_to_excel


def config_file(action='get', **kwargs):
    """ Получение и изменение параметров файла конфигурации

    import_path - путь к каталогу импорта
    export_path - путь к каталогу экспорта

    get - возвращается словарь с параметрами,
    set - в файл конфигурации записываются параметры kwargs.
    """

    cast = {'import_path': 'import_path', 'export_path': 'export_path'}

    config = ConfigParser()
    config.read('config.ini')  # Получение файла конфигурации
    if action == 'set':
        for arg, key in cast.items():
            if arg in kwargs:
                config['DEFAULT'][key] = kwargs[arg]
    elif action == 'get':
        out = dict()
        for arg, key in cast.items():
            out[arg] = config['DEFAULT'].get(key, '')
        return out

    with open('config.ini', 'w') as configfile:
        config.write(configfile)


def select_folder(path, name):
    """ Диалоговое окно выбора папки """
    path = filedialog.askdirectory(initialdir=path, title=name)
    return path


def button_import():
    """ Обработка нажатия кнопки выбора каталога импорта """
    path = select_folder(import_var.get(), 'Выберите каталог для импорта')
    if path:
        import_var.set(path)
        config_file(action='set', import_path=path)
        update_cast()

def button_export():
    """ Обработка нажатия кнопки выбора каталога экспорта """
    path = select_folder(export_var.get(), 'Выберите каталог для экспорта')
    if path:
        export_var.set(path)
        config_file(action='set', export_path=path)


def button_do():
    """ Обработка нажатия кнопки Экспорт """
    if not listbox.curselection():
        messagebox.showerror("Ошибка", "Не выбраны файлы для импорта!")
        return
    if not os.path.exists(export_var.get()):
        messagebox.showerror("Ошибка", "Каталог для экспорта не существует!")
        return
    if not os.path.exists(import_var.get()):
        messagebox.showerror("Ошибка", "Каталог для импорта не существует!")
        return

    ok = {}
    message = ''
    progress = ttk.Progressbar(root, length=400, value=0, maximum=100)
    progress.place(x=50, y=560)  # Показать прогресс бар
    root.update_idletasks()
    selected = listbox.curselection()
    step_progress = 100 / len(selected)  # Шаг прогресс бара
    progress['value'] = 0  # Обнулить прогресс бар
    for index in selected:

        file = listbox.get(index)  # Получаем имя файла
        new_file = file

        try:
            file = os.path.join(import_var.get(), file)  # Получаем полный путь к файлу
            # Проверяем существование файла
            if not os.path.exists(file):
                raise Exception(f'Файл не найден!')
            # Создаем имя файла для экспорта по следующим правилам:
            # Расширение нового файла - xlsx, имя как у исходного файла
            # Если файл с таким именем уже существует, то добавить к имени файла (1), (2) и т.д.
            # Проверяем существование файла и создаем имя нового файла
            i = 0
            basename = os.path.basename(file)
            while True:
                ending = f'({i})' if i else ''
                new_file = os.path.join(export_var.get(), basename.replace('.sgr', f'{ending}.xlsx'))
                if not os.path.exists(new_file):
                    break
                i += 1

            # Прочитать, обработать и сохранить файл
            ok[os.path.basename(new_file)] = sgr_to_excel(file, new_file)

        except Exception as e:
            if not message:
                message = e
            ok[os.path.basename(new_file)] = message
            print(new_file)

        # Увеличить прогресс бар
        progress['value'] += step_progress
        root.update_idletasks()

    # Вывести результаты экспорта в диалоговом окне
    message = ''
    for file, result in ok.items():
        message += f'{file}: {result}\n'
    messagebox.showinfo("Результат", message)
    progress.place_forget()  # Скрыть прогресс бар






# Инициализация
config = config_file()
import_path = config['import_path']
export_path = config['export_path']
# Если путь не существует, то установить текущий каталог
if not os.path.exists(import_path):
    import_path = None
if not os.path.exists(export_path):
    export_path = None
if not import_path:
    import_path = os.path.dirname(os.path.abspath(__file__))
if not export_path:
    export_path = os.path.dirname(os.path.abspath(__file__))
config_file(action='set', import_path=import_path, export_path=export_path)

# Проверка существования образца выходного файла
if not os.path.exists('sample.xlsx'):
    # Вывести диалоговое окно об ошибке и выйти из программы
    messagebox.showerror("Ошибка", "Отсутствует образец исходящего документа 'sample.xlsx'!")
    exit()


# Создаем окно
root = tk.Tk()

# Размер экране
w = root.winfo_screenwidth()
h = root.winfo_screenheight()

# Рисуем окно
root.title("Sqr to Excel Converter")
root.geometry(f'500x600+{(w-500)//2}+{(h-600)//2}')

# Импорт
import_frame = LabelFrame(root, width=470, height=310, text='Импорт', foreground='#083863', font=('Arial', 12))
import_frame.place(x=15, y=40)

# Метка Выберите каталог
import_label1 = tk.Label(import_frame, text='Выберите каталог', font=('Arial', 12))
import_label1.place(x=10, y=10)

# Поле для ввода пути к каталогу
import_var = StringVar(value=import_path)
import_entry = tk.Entry(import_frame, textvariable=import_var, width=41, font=('Arial', 12), state='readonly')
import_entry.place(x=10, y=40)

# Кнопка со стрелкой вниз Выбор каталога
import_button = tk.Button(import_frame, text='▼', font=('Arial', 8), width=3, height=1, command=button_import)
import_button.place(x=390, y=40)

# Кнопка со значком папки Открыть каталог
import_button_folder = tk.Button(import_frame, text='📁', font=('Arial', 8), width=3, height=1,
                          command=lambda: os.startfile(import_var.get()))
import_button_folder.place(x=425, y=40)


# Метка Выберите файл(ы) для импорта
import_label2 = tk.Label(import_frame, text='Выберите файл(ы) для импорта', font=('Arial', 12))
import_label2.place(x=10, y=70)

# Фрейм для списка
frame = tk.Frame(import_frame, width=40, height=10)
frame.place(x=10, y=100)

scrollbar = tk.Scrollbar(frame)
scrollbar.grid(row=0, column=1, sticky='ns')

listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, font=('Arial', 14), highlightthickness=0, width=39, height=7)
listbox.grid(row=0, column=0, sticky='nsew')

listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox.yview)

# отключаем подчеркивание текста при выделении
listbox.config(activestyle='none')


def update_cast():
    """ Обновление списка файлов в listbox """
    listbox.delete(0, tk.END)
    for file in os.listdir(import_var.get()):
        if file.endswith('.sgr'):
            listbox.insert(tk.END, file)


update_cast()  # Обновление списка файлов в listbox


# Экспорт
export_frame = LabelFrame(root, width=470, height=110, text='Экспорт', foreground='#083863', font=('Arial', 12))
export_frame.place(x=15, y=380)

# Метка Выберите каталог
export_label = tk.Label(export_frame, text='Выберите каталог', font=('Arial', 12))
export_label.place(x=10, y=10)

# Поле для ввода пути к каталогу
export_var = StringVar(value=export_path)
export_entry = tk.Entry(export_frame, textvariable=export_var, width=41, font=('Arial', 12), state='readonly')
export_entry.place(x=10, y=40)

# Кнопка со стрелкой вниз
export_button = tk.Button(export_frame, text='▼', font=('Arial', 8), width=3, height=1, command=button_export)
export_button.place(x=390, y=40)

# Кнопка со значком папки Открыть каталог
export_button_folder = tk.Button(export_frame, text='📁', font=('Arial', 8), width=3, height=1,
                          command=lambda: os.startfile(export_var.get()))
export_button_folder.place(x=425, y=40)

# Кнопка Экспорт
do_button = tk.Button(root, text='Экспорт', font=('Arial', 12), width=10, height=1, command=button_do)
do_button.place(x=135, y=510)

# Кнопка Отмена
cancel_button = tk.Button(root, text='Отмена', font=('Arial', 12), width=10, height=1)
cancel_button.place(x=265, y=510)
cancel_button.bind('<Button-1>', lambda e: root.destroy())

# listbox.bind('<Button-1>', lambda e: listbox.selection_toggle(listbox.nearest(e.y)))




root.mainloop()