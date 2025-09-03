import datetime
import tkinter as tk
import os
import sys
from tkinter import filedialog, messagebox, ttk
from assembly_WIN64 import lab_assembler
from full_parsing_WIN64 import LabParser
from tests_WIN64 import run_tests
from errors_handler_WIN64 import find_suspicious_blocks, log_errors, remove_suspicious_blocks

# Для хранения parser между шагами
parser = None

def select_input_dir():
    path = filedialog.askdirectory(title="Выбрать папку с исходниками")
    if path:
        input_dir_var.set(path)

def select_output_file():
    path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                        initialfile=f"{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}_Выгрузка_ЛИ.xlsx",
                                        filetypes=[("Excel files", "*.xlsx")],
                                        title="Сохранить итоговый файл как...")
    if path:
        output_file_var.set(path)

def get_basedir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

def on_process_click():
    input_dir = input_dir_var.get()
    output_file = output_file_var.get()
    reference_file = os.path.join(get_basedir(), "Формат_выгрузки.xlsx")
    raw_combined_file = os.path.join(get_basedir(), "combined.xlsx")

    if not input_dir or not output_file:
        messagebox.showerror("Ошибка", "Выберите папку исходников и итоговый файл!")
        return

    lab_assembler(input_dir, raw_combined_file)
    global parser
    parser = LabParser(
        input_file=raw_combined_file,
        reference_file=reference_file,
        output_file=output_file
    )

    parser.full_parse_and_format()

    if os.path.exists(raw_combined_file):
        os.remove(raw_combined_file)

    messagebox.showinfo("Готово!", f"Файл успешно создан:\n{output_file}")

    test_btn.config(state='normal')
    error_btn.config(state='disabled')
    remove_btn.config(state='disabled')


def on_test_click():
    global parser
    errors = run_tests(parser.df)
    if not errors:
        messagebox.showinfo("Внимание!", "Ошибок не найдено.")
        error_btn.config(state='disabled')
        remove_btn.config(state='disabled')
    else:
        messagebox.showinfo("Внимание!",
              f"Несовпадающих строк: {len(errors)}.\nДля подробного разбора нажмите \"Проверить ошибки\".")
        error_btn.config(state='normal')

def on_error_handler_click():
    global parser
    df_exploded = parser.df_exploded
    errors_df = find_suspicious_blocks(df_exploded)
    log_errors(errors_df)
    if not errors_df.empty:
        messagebox.showinfo("Готово!", f"Сохранён лог ошибок (Ошибки.txt). \nМожно удалить {len(errors_df)} битых строк.")
        remove_btn.config(state='normal')
    else:
        messagebox.showinfo("Внимание!","Подозрительных строк не найдено.")
        remove_btn.config(state='disabled')

def on_remove_click():
    global parser
    df_exploded = parser.df_exploded
    errors_df = find_suspicious_blocks(df_exploded)
    # Диалог выбора файла для сохранения
    out_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=f"{datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}_Финальная_выгрузка_ЛИ.xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Сохранить файл без битых строк как..."
    )
    if not out_path:
        return  # пользователь отменил

    df_exploded_clean = remove_suspicious_blocks(df_exploded, errors_df)
    df_exploded_clean.to_excel(out_path, index=False)
    LabParser.reorder_and_save_df(df_exploded_clean, out_path)
    LabParser.apply_formatting_to_file(out_path, reference_file="Формат_выгрузки.xlsx")
    messagebox.showinfo("Готово!", f"Файл успешно создан:\n{out_path}")

def show_about():
    messagebox.showinfo(
        "О программе",
        "Парсер лабораторных исследований (v1.0)\n"
        "Автоматизация для ГИС Меркурий.\n"
        "Кратко: собирает, форматирует, проверяет и сохраняет выгрузки ЛИ в один Excel.\n\n"
        "Подробнее — читайте README в папке с программой.\n\n"
        "© 2025 D.Agurin"
    )


# ---- GUI ----

root = tk.Tk()
root.title("Сборка и форматирование лабораторных исследований")

# прикручиваем иконки
if getattr(sys, 'frozen', False):
    # Путь к иконке — рядом с .exe
    icon_path = os.path.join(os.path.dirname(sys.executable), 'logo.ico')
else:
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.ico')

root.iconbitmap(icon_path)

input_dir_var = tk.StringVar()
output_file_var = tk.StringVar()

about_btn = tk.Button(root, text="❓", command=show_about, relief='flat')
about_btn.grid(row=5, column=0, sticky='sw', pady=(8, 6), padx=(8, 0))

tk.Label(root, text="Powered by D.Agurin", fg="gray").grid(
    row=5, column=2, sticky='se', pady=(8, 6), padx=(0, 8)
)

tk.Label(root, text="Папка с исходниками:").grid(row=0, column=0, sticky='e')
tk.Entry(root, textvariable=input_dir_var, width=50).grid(row=0, column=1)
tk.Button(root, text="Обзор...", command=select_input_dir).grid(row=0, column=2)

tk.Label(root, text="Итоговый Excel-файл:").grid(row=1, column=0, sticky='e')
tk.Entry(root, textvariable=output_file_var, width=50).grid(row=1, column=1)
tk.Button(root, text="Обзор...", command=select_output_file).grid(row=1, column=2)

process_btn = tk.Button(root, text="Обработать", command=on_process_click)
process_btn.grid(row=2, column=1, pady=8)

test_btn = tk.Button(root, text="Запустить тесты", command=on_test_click, state='disabled')
test_btn.grid(row=3, column=1, pady=8)

error_btn = tk.Button(root, text="Проверить ошибки", command=on_error_handler_click, state='disabled')
error_btn.grid(row=4, column=1, pady=8)

remove_btn = tk.Button(root, text="Удалить битые строки", command=on_remove_click, state='disabled')
remove_btn.grid(row=5, column=1, pady=8)

# центрируем
def center_window(win, width=600, height=400):
    win.update_idletasks()
    x = (win.winfo_screenwidth()  - width) // 2
    y = (win.winfo_screenheight() - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")

center_window(root, width=565, height=220)
root.resizable(False, False)

root.grid_columnconfigure(0, minsize=140)

root.mainloop()
