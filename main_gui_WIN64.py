from pathlib import Path
import tkinter as tk
from tkinter import messagebox
import threading, time, os, sys, datetime

from gui import App
from hs_choose import load_hs_mapping

from parse_attempt import (
    build_driver_with_downloads, download_dir,
    login_via_qr, choose_radio_and_submit, navigate_to_vetdocs_and_outgoing,
    load_ttn_list_from_file, process_ttn_list, save_no_ttn_to_excel
)
from assembly_WIN64 import lab_assembler
from full_parsing_WIN64 import LabParser
from tests_WIN64 import run_validators
from errors_handler_WIN64 import (
    find_suspicious_blocks, log_errors, remove_suspicious_blocks, export_rows_without_lab_results
)

# ---- глобальное состояние ----
parser = None
HS_MAP = load_hs_mapping()  # {'Имя': {'uuid': '...', 'schema': '...'}, ...}
app: App | None = None

# ---- утилиты ----
def get_basedir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

# ---- колбэки для GUI ----
def get_hs_names():
    return list(HS_MAP.keys())

def show_about():
    messagebox.showinfo(
        "О программе",
        "Парсер лабораторных исследований (v2.0)\n"
        "Автоматизация для ГИС Меркурий.\n"
        "Кратко: собирает, форматирует, проверяет и сохраняет выгрузки ЛИ в один Excel.\n\n"
        "Подробнее — читайте README в папке с программой.\n\n"
        "© 2025 D.Agurin"
    )

def on_parse(ttn_path: str, hs_name: str):
    global app
    if not ttn_path or not Path(ttn_path).is_file():
        app.warn("Файл с ТТН", "Выберите корректный файл с номерами ТТН.")
        return

    info = HS_MAP.get(hs_name, {})
    hs_id = (info.get("uuid") or "").strip()
    schema_value = (info.get("schema") or "").strip()
    if not hs_id:
        app.error("ХС", f"Для «{hs_name}» не указан UUID в hs_mapping.json.")
        return
    if not schema_value:
        app.error("Схема", f"Для «{hs_name}» не указана schema в hs_mapping.json.")
        return

    def worker():
        try:
            driver = build_driver_with_downloads(download_dir, headless=False)
            # 1) вход
            login_via_qr(driver)
            # 2) ХС
            choose_radio_and_submit(driver, hs_id)
            print("✅ ХС выбран.")
            time.sleep(1)
            # 3) предприятие = 'null'
            choose_radio_and_submit(driver, "null")
            print("✅ Предприятие (id='null') выбрано.")
            time.sleep(1)
            # 4) в исходящие
            navigate_to_vetdocs_and_outgoing(driver)
            time.sleep(2)
            # 5) список ТТН и прогон
            ttns = load_ttn_list_from_file(Path(ttn_path))
            print(f"🔎 В файле {Path(ttn_path).name}: {len(ttns)} ТТН")
            no_ttn = process_ttn_list(
                driver, ttns, schema_value=schema_value, pause_between=(0.8, 1.5)
            )
            save_no_ttn_to_excel(no_ttn)

            app.post(app.reveal_step2)
            app.post(lambda: app.set_buttons(process=True))
        except Exception as e:
            app.post(lambda: app.error("Ошибка загрузки", str(e)))
        finally:
            app.post(lambda: app.set_busy(False))

    app.set_busy(True)
    threading.Thread(target=worker, daemon=True).start()

def on_process(input_dir: str, output_file: str):
    global parser, app
    if not input_dir or not output_file:
        app.error("Ошибка", "Выберите папку исходников и итоговый файл!")
        return

    def worker():
        try:
            reference_file = os.path.join(get_basedir(), "Формат_выгрузки.xlsx")
            raw_combined_file = os.path.join(get_basedir(), "combined.xlsx")

            lab_assembler(input_dir, raw_combined_file)
            parser = LabParser(
                input_file=raw_combined_file,
                reference_file=reference_file,
                output_file=output_file
            )
            parser.full_parse_and_format()
            app.post(lambda: app.info("Готово!", f"Файл успешно создан:\n{output_file}"))
            app.post(lambda: app.set_buttons(process=True, test=True, error=False, remove=False))
        except Exception as e:
            app.post(lambda: app.error("Ошибка обработки", str(e)))

    threading.Thread(target=worker, daemon=True).start()

def on_test():
    global parser, app
    errors, empty_rows = run_validators(parser.df)
    raw_combined_file = os.path.join(get_basedir(), "combined.xlsx")

    if not errors and not empty_rows:
        app.info("Внимание!", "Ошибок не найдено.")
        app.set_buttons(process=True, test=True, error=False, remove=False)
        if os.path.exists(raw_combined_file):
            os.remove(raw_combined_file)

    if errors and empty_rows:
        app.info("Внимание!",
                 f"Несовпадающих строк: {len(errors)} из {len(parser.df)}."
                 f"\nПустых ЛИ (в исходнике): {len(empty_rows)}"
                 "\nДля подробного разбора нажмите \"Проверить ошибки\".")
        app.set_buttons(process=True, test=True, error=True, remove=False)

    if errors and not empty_rows:
        app.info("Внимание!",
                 f"Несовпадающих строк: {len(errors)} из {len(parser.df)}."
                 "\nДля подробного разбора нажмите \"Проверить ошибки\".")
        app.set_buttons(process=True, test=True, error=True, remove=False)
        if os.path.exists(raw_combined_file):
            os.remove(raw_combined_file)

    if empty_rows and not errors:
        app.info("Внимание!",
                 f"Пустых ЛИ (в исходнике): {len(empty_rows)}."
                 "\nДля подробного разбора нажмите \"Проверить ошибки\".")
        app.set_buttons(process=True, test=True, error=True, remove=False)

def on_errors():
    global parser, app
    raw_combined_file = os.path.join(get_basedir(), "combined.xlsx")
    out_path = app.get_values()["output_file"]
    out_dir = os.path.dirname(out_path) if out_path else get_basedir()
    now_str = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    err_log_path = os.path.join(out_dir, f"{now_str}_Ошибки.txt")
    no_li_file   = os.path.join(out_dir, f"{now_str}_ЭВСД_без_ЛИ.xlsx")

    errors, empty_rows = run_validators(parser.df)
    errors_df = find_suspicious_blocks(parser.df_exploded)

    combined_path = os.path.join(get_basedir(), "combined.xlsx")
    had_empty_li = len(empty_rows) > 0
    created_no_li = False
    export_error = None
    if had_empty_li:
        try:
            created_no_li = export_rows_without_lab_results(combined_path, no_li_file)
            LabParser.apply_formatting_to_file(
                no_li_file, reference_file=os.path.join(get_basedir(), "Формат_выгрузки.xlsx"))
        except Exception as e:
            export_error = e

    created_err_log = False
    if not errors_df.empty:
        try:
            log_errors(errors_df, path=err_log_path)
            created_err_log = True
        except Exception as e:
            print("Не удалось записать Ошибки.txt:", e)

    if created_err_log and had_empty_li:
        msg = (
            "Обнаружены ОДНОВРЕМЕННО два типа проблем:\n"
            f"• Подозрительные блоки — лог:\n{err_log_path}\n"
            "• ЭВСД без ЛИ — "
        )
        if created_no_li:
            msg += f"файл:\n{no_li_file}\n"
            if os.path.exists(raw_combined_file):
                os.remove(raw_combined_file)
        else:
            msg += "обнаружены (по тестам), но не удалось сформировать файл из combined.xlsx.\n"
            if export_error:
                msg += f"\nПричина: {export_error}\n"
        msg += (
            "\nВы можете удалить подозрительные строки кнопкой «Удалить битые строки».\n"
            "Строки без ЛИ проверьте вручную в ГИС Меркурий."
        )
        app.set_buttons(process=True, test=True, error=True, remove=True)
        app.info("Проверка завершена", msg)

    elif created_err_log:
        msg = (
            f"Найдены подозрительные блоки.\nЛог сохранён:\n{err_log_path}\n\n"
            "Нажмите «Удалить битые строки», чтобы очистить выгрузку."
        )
        if os.path.exists(raw_combined_file):
            os.remove(raw_combined_file)
        app.set_buttons(process=True, test=True, error=True, remove=True)
        app.info("Проверка завершена", msg)

    elif had_empty_li:
        if created_no_li:
            msg = (
                "Дубли/битых блоков не найдено.\n"
                f"Но есть ЭВСД без ЛИ — сформирован файл:\n{no_li_file}\n\n"
                "Проверьте эти строки вручную в ГИС Меркурий."
            )
            if os.path.exists(raw_combined_file):
                os.remove(raw_combined_file)
        else:
            msg = (
                "Есть ЭВСД без ЛИ (выявлено тестами), но не удалось сформировать файл из combined.xlsx.\n"
                "Проверьте исходники/права доступа и попробуйте снова."
            )
        app.set_buttons(process=True, test=True, error=True, remove=False)
        app.info("Проверка завершена", msg)

    else:
        if len(errors) == 0:
            msg = "Проверка завершена: всё ок! Несовпадений и пустых ЛИ не обнаружено."
            if os.path.exists(raw_combined_file):
                os.remove(raw_combined_file)
        else:
            msg = (
                f"Есть {len(errors)} несовпадающих строк (см. консоль тестов),\n"
                "но явных «битых блоков» и пустых ЛИ не найдено.\n"
                "Проверьте различия вручную."
            )
        app.set_buttons(process=True, test=True, error=False, remove=False)
        app.info("Проверка завершена", msg)

def on_remove():
    global parser, app
    df_exploded = parser.df_exploded
    errors_df = find_suspicious_blocks(df_exploded)
    out_path = tk.filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=f"{datetime.datetime.now():%Y-%m-%d_%H-%M-%S}_Финальная_выгрузка_ЛИ.xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Сохранить файл без битых строк как..."
    )
    if not out_path:
        return
    df_exploded_clean = remove_suspicious_blocks(df_exploded, errors_df)
    df_exploded_clean.to_excel(out_path, index=False)
    LabParser.reorder_and_save_df(df_exploded_clean, out_path)
    LabParser.apply_formatting_to_file(out_path, reference_file=os.path.join(get_basedir(), "Формат_выгрузки.xlsx"))
    app.info("Готово!", f"Файл успешно создан:\n{out_path}")

# ---- запуск ----
if __name__ == "__main__":
    app = App(
        on_parse=on_parse,
        on_process=on_process,
        on_test=on_test,
        on_errors=on_errors,
        on_remove=on_remove,
        get_hs_names=get_hs_names,
        show_about=show_about
    )
    app.run()
