import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime

class App:
    def __init__(self, *,
                 on_parse, on_process, on_test, on_errors, on_remove,
                 get_hs_names, show_about):
        """
        Колбэки:
          on_parse(ttn_path:str, hs_name:str, since_date:str)
          on_process(input_dir:str, output_file:str)
          on_test()
          on_errors()
          on_remove()
          get_hs_names() -> list[str]
          show_about()
        """
        self.on_parse   = on_parse
        self.on_process = on_process
        self.on_test    = on_test
        self.on_errors  = on_errors
        self.on_remove  = on_remove
        self.get_hs_names = get_hs_names
        self.show_about = show_about

        self.root = tk.Tk()
        self.root.title("Сборка и форматирование лабораторных исследований")

        self._build_ui()

    # -------- утилиты, чтобы раннер мог управлять UI --------
    def post(self, fn): self.root.after(0, fn)

    def set_busy(self, busy: bool):
        self.parse_btn.config(state=("disabled" if busy else "normal"),
                              text=("Загружаю…" if busy else "Загрузить ЭВСД"))

    def reveal_step2(self):
        if not self.step2.winfo_ismapped():
            self.step2.grid(row=1, column=0, sticky="ew")
            self.ent_in.focus_set()

    def set_buttons(self, *, process=False, test=False, error=False, remove=False):
        self.process_btn.config(state=("normal" if process else "disabled"))
        self.test_btn.config(state=("normal" if test else "disabled"))
        self.error_btn.config(state=("normal" if error else "disabled"))
        self.remove_btn.config(state=("normal" if remove else "disabled"))

    def info(self, title, msg):  messagebox.showinfo(title, msg)
    def error(self, title, msg): messagebox.showerror(title, msg)
    def warn (self, title, msg): messagebox.showwarning(title, msg)

    def _validate_date_chars(self, proposed: str) -> bool:
        # разрешаем пусто (на всякий) и длину до 10
        if proposed == "":
            return True
        if len(proposed) > 10:
            return False
        # только цифры и точки
        return re.fullmatch(r"[0-9.]*", proposed) is not None

    def _validate_date_full(self, s: str) -> bool:
        s = (s or "").strip()
        if not s:
            return False
        try:
            datetime.datetime.strptime(s, "%d.%m.%Y")
            return True
        except ValueError:
            return False

    def get_values(self):
        return {
            "ttn_path": self.ttn_var.get().strip(),
            "hs_name":  self.hs_var.get().strip(),
            "input_dir": self.in_dir.get().strip(),
            "output_file": self.out_file.get().strip(),
            "since_date": self.date_var.get().strip()
        }

    def run(self): self.root.mainloop()

    # ------------------- UI -------------------
    def _build_ui(self):
        self.ttn_var  = tk.StringVar()
        self.hs_var   = tk.StringVar()
        self.in_dir   = tk.StringVar()
        self.out_file = tk.StringVar()
        self.date_var = tk.StringVar(value="01.08.2025")

        # STEP 1
        step1 = tk.Frame(self.root, padx=12, pady=8)
        step1.grid(row=0, column=0, sticky="ew")
        step1.grid_columnconfigure(0, minsize=180)
        step1.grid_columnconfigure(1, weight=1)
        step1.grid_columnconfigure(2, minsize=90)

        # 0: файл ТТН
        tk.Label(step1, text="Файл с номерами ТТН:").grid(row=0, column=0, sticky="e", pady=2)
        tk.Entry(step1, textvariable=self.ttn_var, width=50).grid(row=0, column=1, sticky="ew", padx=6, pady=2)
        tk.Button(step1, text="Обзор…", command=self._browse_ttn).grid(row=0, column=2, pady=2)

        # 1: ХС
        tk.Label(step1, text="ХС:").grid(row=1, column=0, sticky="e", pady=(4, 2))
        self.hs_combo = ttk.Combobox(step1, textvariable=self.hs_var, state="readonly")
        self.hs_combo.grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=(4, 2))
        hs_names = self.get_hs_names() or []
        self.hs_combo["values"] = hs_names
        if hs_names:
            self.hs_combo.current(0)

        # 2: Дата с
        tk.Label(step1, text="Дата с (dd.MM.yyyy):").grid(row=2, column=0, sticky="e", pady=(4, 2))

        vcmd = (self.root.register(self._validate_date_chars), "%P")
        date_entry = tk.Entry(step1, textvariable=self.date_var, width=12,
                              validate="key", validatecommand=vcmd)
        date_entry.grid(row=2, column=1, sticky="w", padx=6, pady=(4, 2))

        # разделитель
        ttk.Separator(step1, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="ew", pady=(6, 6))

        # кнопка
        self.parse_btn = tk.Button(step1, text="Загрузить ЭВСД", command=self._on_parse_click)
        self.parse_btn.grid(row=4, column=0, columnspan=3, sticky="ew")

        # STEP 2 (скрыт до parse)
        self.step2 = tk.Frame(self.root, padx=12, pady=4)
        self.step2.grid_columnconfigure(1, weight=1)

        tk.Label(self.step2, text="Папка с исходниками:").grid(row=0, column=0, sticky="e", pady=2)
        self.ent_in = tk.Entry(self.step2, textvariable=self.in_dir, width=50)
        self.ent_in.grid(row=0, column=1, sticky="ew", padx=6, pady=2)
        tk.Button(self.step2, text="Обзор…", command=self._browse_in_dir).grid(row=0, column=2, pady=2)

        tk.Label(self.step2, text="Итоговый Excel-файл:").grid(row=1, column=0, sticky="e", pady=2)
        tk.Entry(self.step2, textvariable=self.out_file, width=50).grid(row=1, column=1, sticky="ew", padx=6, pady=2)
        tk.Button(self.step2, text="Обзор…", command=self._browse_out_file).grid(row=1, column=2, pady=2)

        self.process_btn = tk.Button(self.step2, text="Обработать", command=self._on_process_click, state="disabled")
        self.process_btn.grid(row=2, column=1, pady=8)

        self.test_btn = tk.Button(self.step2, text="Запустить тесты", command=self.on_test, state="disabled")
        self.test_btn.grid(row=3, column=1, pady=8)

        self.error_btn = tk.Button(self.step2, text="Проверить ошибки", command=self.on_errors, state="disabled")
        self.error_btn.grid(row=4, column=1, pady=8)

        self.remove_btn = tk.Button(self.step2, text="Удалить битые строки", command=self.on_remove, state="disabled")
        self.remove_btn.grid(row=5, column=1, pady=8)

        # footer
        tk.Button(self.root, text="❓", command=self.show_about, relief='flat').grid(
            row=2, column=0, sticky='w', padx=12, pady=(6,8)
        )
        tk.Label(self.root, text="Powered by xdanthecoolest", fg="gray").grid(
            row=2, column=0, sticky='e', padx=12, pady=(6,8)
        )

    # ---------- локальные хендлеры ----------
    def _browse_ttn(self):
        path = filedialog.askopenfilename(
            title="Выбрать файл с номерами ТТН",
            filetypes=[("Текст/Excel", "*.txt *.csv *.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if path: self.ttn_var.set(path)

    def _browse_in_dir(self):
        path = filedialog.askdirectory(title="Выбрать папку с исходниками")
        if path: self.in_dir.set(path)

    def _browse_out_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"{datetime.datetime.now():%Y-%m-%d_%H-%M-%S}_Выгрузка_ЛИ.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить итоговый файл как..."
        )
        if path: self.out_file.set(path)

    def _on_parse_click(self):
        v = self.get_values()
        # строгая валидация перед запуском
        if not self._validate_date_full(v["since_date"]):
            self.warn("Некорректная дата", "Введите дату в формате dd.MM.yyyy, пример: 01.08.2025")
            return
        self.on_parse(v["ttn_path"], v["hs_name"], v["since_date"])

    def _on_process_click(self):
        v = self.get_values()
        self.on_process(v["input_dir"], v["output_file"])
