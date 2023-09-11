from tkinter import *
import tkinter as tk


root = tk.Tk()

# Размер экране
w = root.winfo_screenwidth()
h = root.winfo_screenheight()

# Рисуем окно
# Установить размер шрифта 12
root.title("Sqr to Excel Converter")
root.geometry(f'500x600+{(w-500)//2}+{(h-600)//2}')

# Импорт
import_frame = LabelFrame(root, width=470, height=310, text='Импорт', foreground='#083863', font=('Arial', 12))
import_frame.place(x=15, y=40)

# Метка Выберите каталог
import_label1 = tk.Label(import_frame, text='Выберите каталог', font=('Arial', 12))
import_label1.place(x=10, y=10)

# Поле для ввода пути к каталогу
import_entry = tk.Entry(import_frame, width=44, font=('Arial', 12))
import_entry.place(x=10, y=40)

# Кнопка со стрелкой вниз
import_button = tk.Button(import_frame, text='▼', font=('Arial', 8), width=3, height=1)
import_button.place(x=428, y=40)

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

listbox.insert(tk.END, "Line 1")
listbox.insert(tk.END, "Line 2")
listbox.insert(tk.END, "Line 3")
listbox.insert(tk.END, "Line 4")
listbox.insert(tk.END, "Line 5")
listbox.insert(tk.END, "Line 1")
listbox.insert(tk.END, "Line 2")
listbox.insert(tk.END, "Line 3")
listbox.insert(tk.END, "Line 4")
listbox.insert(tk.END, "Line 5")

listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox.yview)

# отключаем подчеркивание текста при выделении
listbox.config(activestyle='none')


# Экспорт
export_frame = LabelFrame(root, width=470, height=110, text='Экспорт', foreground='#083863', font=('Arial', 12))
export_frame.place(x=15, y=380)

# Метка Выберите каталог
export_label = tk.Label(export_frame, text='Выберите каталог', font=('Arial', 12))
export_label.place(x=10, y=10)

# Поле для ввода пути к каталогу
export_entry = tk.Entry(export_frame, width=44, font=('Arial', 12))
export_entry.place(x=10, y=40)

# Кнопка со стрелкой вниз
export_button = tk.Button(export_frame, text='▼', font=('Arial', 8), width=3, height=1)
export_button.place(x=428, y=40)

# Кнопка Экспорт
export_button = tk.Button(root, text='Экспорт', font=('Arial', 12), width=10, height=1)
export_button.place(x=135, y=510)

# Кнопка Отмена
cancel_button = tk.Button(root, text='Отмена', font=('Arial', 12), width=10, height=1)
cancel_button.place(x=265, y=510)
cancel_button.bind('<Button-1>', lambda e: root.destroy())

# listbox.bind('<Button-1>', lambda e: listbox.selection_toggle(listbox.nearest(e.y)))

root.mainloop()