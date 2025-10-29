import os
import time
import random
from pathlib import Path
import datetime
from typing import List

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select

# мерк
START_URL = (
    "https://mercury.vetrf.ru/hs/operatorui"
    "?_action=listInputVetDocument&pageList=1&all=true&status=1&request=false"
)

# время жизни QR
QR_WAIT_SEC = 60

# Базовые рамки «человеческой» паузы
HUMAN_MIN = 0.6
HUMAN_MAX = 1.8

# хелпер для ИНДЕПОТЕНТНОСТИ №1
def is_checked(el) -> bool:
    try:
        return bool(el.get_attribute("checked")) or el.is_selected()
    except Exception:
        return False

# хелпер для ИНДЕПОТЕНТНОСТИ №2
def selected_option_value(select_el) -> str:
    try:
        from selenium.webdriver.support.ui import Select
        return Select(select_el).first_selected_option.get_attribute("value") or ""
    except Exception:
        return ""

# грузим ттн из файла
def load_ttn_list_from_file(file_path: str | Path) -> List[str]:
    """
    Загружает список ТТН из txt/csv/xlsx/xls (для Excel нужен pandas).
    - игнорирует пустые строки
    - убирает BOM/кавычки/невидимые пробелы
    - удаляет дубликаты, сохраняя порядок
    """
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(f"Не найден файл со списком ТТН: {p}")

    def clean(s) -> str:
        if s is None:
            return ""
        s = str(s).strip()
        # прибрать BOM и редкие пробелы
        for junk in ("\ufeff", "\u200b", "\u00a0"):
            s = s.strip(junk)
        # снять внешние кавычки '...' или "..."
        if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
            s = s[1:-1].strip()
        return s

    suf = p.suffix.lower()
    rows: list[str] = []

    if suf in {".xlsx", ".xls"}:
        try:
            import pandas as pd  # type: ignore
        except Exception as e:
            raise RuntimeError(
                "Для чтения Excel-файлов нужен pandas. "
                "Установите его или сохраните файл как CSV/TXT."
            ) from e
        col0 = pd.read_excel(p, header=None).iloc[:, 0]
        rows = [clean(v) for v in col0]

    elif suf == ".csv":
        import csv
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(2048)
            f.seek(0)
            # попробуем угадать разделитель, иначе возьмём ';'
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=";,|\t,")
            except Exception:
                class _D: pass
                dialect = _D(); dialect.delimiter = ";"
            reader = csv.reader(f, dialect)
            rows = [clean(row[0]) for row in reader if row]

    else:
        # .txt и всё остальное — построчно
        with p.open("r", encoding="utf-8-sig", errors="ignore") as f:
            rows = [clean(line) for line in f]

    # выбросить пустые и удалить дубляжи, сохранив порядок
    seen = set()
    out: List[str] = []
    for s in rows:
        if not s:
            continue
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out

# рандомная пауза между кликами
def human_pause(min_s: float = HUMAN_MIN, max_s: float = HUMAN_MAX):
    """Небольшая случайная задержка как у живого пользователя."""
    time.sleep(random.uniform(min_s, max_s))

# билдер симуляции с кастомной папкой для скачивания
def build_driver_with_downloads(download_dir: Path, headless=False) -> webdriver.Firefox:
    download_dir.mkdir(parents=True, exist_ok=True)          # создадим папку, если её нет
    opts = Options()
    opts.headless = headless

    # === ключевые префы Firefox для автоскачивания ===
    opts.set_preference("browser.download.folderList", 2)    # 2 = использовать custom dir
    opts.set_preference("browser.download.dir", str(download_dir))
    opts.set_preference("browser.download.useDownloadDir", True)
    opts.set_preference("browser.download.alwaysOpenPanel", False)
    opts.set_preference("browser.download.manager.showWhenStarting", False)

    # Не спрашивать «Открыть/Сохранить» для нужных MIME-типов
    # (xlsx, xls; добавь другие, если нужно)
    opts.set_preference(
        "browser.helperApps.neverAsk.saveToDisk",
        ",".join([
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  # .xlsx
            "application/vnd.ms-excel",                                           # .xls
        ])
    )
    # на всякий случай отключим встроенные вьюеры
    opts.set_preference("pdfjs.disabled", True)
    opts.set_preference("browser.helperApps.neverAsk.openFile", "")  # не предлагать «открыть»

    driver = webdriver.Firefox(options=opts)
    driver.maximize_window()
    return driver

# исправляем баги с кликом по исходящим ЭВСД
def click_outgoing_safely(driver):
    """
    Кликаем «Исходящие» надёжно:
    1) обычный click
    2) JS-dispatch MouseEvent
    3) выполнить js из href (javascript:onEnterMenu(...))
    """
    # сам линк «Исходящие»
    link = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((
            By.XPATH,
            "//a[starts-with(@href,'javascript:onEnterMenu') and contains(@href,'VetDocumentAjax') and contains(.,'Исходящие')]"
        ))
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)

    # 1) обычный клик
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, ".")))
        link.click()
        if wait_ajax_quiet(driver, timeout=10):
            return
    except Exception:
        pass

    # 2) синтетическое событие мыши
    try:
        driver.execute_script("""
            const el = arguments[0];
            el.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window}));
        """, link)
        if wait_ajax_quiet(driver, timeout=10):
            return
    except Exception:
        pass

    # 3) выполняем JS из href напрямую
    try:
        href = link.get_attribute("href")  # "javascript:onEnterMenu(...);"
        js_code = href.replace("javascript:", "").rstrip(";")
        driver.execute_script(js_code)
        wait_ajax_quiet(driver, timeout=10)
    except Exception:
        pass

# еще пауза для формирования
def wait_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") in ("complete", "interactive")
    )

# еще задержка кликов
def safe_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    # небольшая «прицеливание» задержка перед кликом
    human_pause(0.2, 0.6)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    # пауза после клика — страница думает/рисует
    human_pause(0.7, 1.6)

# авторизация по куар
def login_via_qr(driver):
    """Кликаем «Госуслуги» → белая кнопка (QR) → ждём завершения авторизации (до 60с)."""
    driver.get(START_URL)
    wait_ready(driver, 30)
    time.sleep(1)

    try:
        gos_a = WebDriverWait(driver, 20, poll_frequency=0.5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='ESIARedirect']"))
        )
    except TimeoutException:
        # fallback: контейнер div.loginFormButtonGos → ближайший <a>
        container = WebDriverWait(driver, 20, poll_frequency=0.5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.loginFormButtonGos"))
        )
        gos_a = container.find_element(By.XPATH, "./ancestor::a[1]")
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "./ancestor::a[1]"))
        )

    # клик по ссылке Госуслуг + пауза, чтобы произошёл переход
    safe_click(driver, gos_a)
    # даём странице/редиректу стабилизироваться
    time.sleep(2)                # ← дополнительная задержка
    wait_ready(driver, 30)
    time.sleep(1)

    # 2) «Белая» кнопка, которая открывает QR
    #   иногда грузится чуть дольше — подождём до 30с с мягким polling
    white_btn = WebDriverWait(driver, 30, poll_frequency=0.5).until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//*[contains(@class,'plain-button_white') and (self::a or self::button)]"
        ))
    )
    safe_click(driver, white_btn)
    time.sleep(2)  # даём QR точно отрисоваться

    # 3) Ждём, пока пользователь отсканирует QR в приложении и произойдёт вход
    WebDriverWait(driver, QR_WAIT_SEC).until(
        EC.any_of(
            EC.presence_of_element_located((By.LINK_TEXT, "Ветеринарные документы")),
            EC.presence_of_element_located((By.ID, "submitButton")),   # экран выбора организации
            EC.presence_of_element_located((By.NAME, "firmGuid"))      # радиокнопки ХС
        )
    )
    time.sleep(1)
    print("✅ Авторизация через QR завершена.")

# спуск в футер за "выбрать"
def choose_radio_and_submit(driver, radio_id: str):
    """Кликает радиокнопку по id и жмёт кнопку «Выбрать» в футере (id=submitButton)."""
    # радиокнопка
    radio = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, radio_id))
    )
    safe_click(driver, radio)

    # кнопка «Выбрать»
    submit_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "submitButton"))
    )
    safe_click(driver, submit_btn)
    time.sleep(1)

# открыть исходящие ЭВСД
def navigate_to_vetdocs_and_outgoing(driver):
    """Открыть: Ветеринарные документы → Исходящие."""
    # 1) «Ветеринарные документы»
    try:
        vetdocs_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "a[href*='operatorui?_action=listVetDocument'][href*='request=true'][href*='preview=true']"
            ))
        )
        safe_click(driver, vetdocs_link)
    except TimeoutException:
        vetdocs_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[contains(@href,'listVetDocument') or contains(.,'Ветеринарные документы')]"
            ))
        )
        safe_click(driver, vetdocs_link)

    time.sleep(1)

    # 2) «Исходящие» — надёжный клик
    click_outgoing_safely(driver)

    # 3) ждём, что раздел реально открылся: появится «Печать» или заголовок списка
    try:
        WebDriverWait(driver, 20).until(
            EC.any_of(
                EC.presence_of_element_located((By.ID, "printSettingsFormTop")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(.,'Список всех исходящих') or contains(.,'исходящих ВСД')]"))
            )
        )
    except TimeoutException:
        # запасной чек — просто ждём, пока ajax завершится
        wait_ajax_quiet(driver, timeout=10)

    set_rows_per_page(driver, value="100")
    time.sleep(1)
    print("✅ Перешли: Ветеринарные документы → Исходящие")

# выбираем 100 строк ЭВСД
def set_rows_per_page(driver, value: str = "100"):
    """
    В разделе 'Исходящие' выбирает кол-во строк в селекте name='rows' → value='100'.
    Делает это один раз (если уже 100 — ничего не трогаем).
    """
    # Находим селект по name='rows' и классам small/blue
    dropdown = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((
            By.XPATH,
            "//select[@name='rows' and contains(@class,'small') and contains(@class,'blue')]"
        ))
    )
    # Если уже стоит 100 — выходим
    sel = Select(dropdown)
    current = sel.first_selected_option.get_attribute("value") or ""
    if current == value:
        return

    # Кликаем и выбираем 100
    safe_click(driver, dropdown)              # фокус на селекте (чуть «по-человечески»)
    sel.select_by_value(value)

    # Ждём, пока отработает onchange (ajax перерисовка списка)
    wait_ajax_quiet(driver, timeout=15)
    human_pause(0.6, 1.2)

    # Опционально перепроверим
    try:
        current = Select(WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "rows"))
        )).first_selected_option.get_attribute("value") or ""
        if current != value:
            # запасной клик, если первый не применился
            sel = Select(driver.find_element(By.NAME, "rows"))
            sel.select_by_value(value)
            wait_ajax_quiet(driver, timeout=10)
    except Exception:
        pass

# еще ожидание
def wait_ajax_quiet(driver, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        try:
            ready = driver.execute_script("return document.readyState")
            active = driver.execute_script("return (window.jQuery && jQuery.active) || 0")
            if ready in ("complete", "interactive") and int(active) == 0:
                return True
        except Exception:
            if driver.execute_script("return document.readyState") in ("complete", "interactive"):
                return True
        time.sleep(0.3)
    return False

# скрываем аннулированные
def _ensure_revoked_unchecked(driver):
    """
    Снимает галку 'аннулированные' ОДИН РАЗ.
    Если уже снята — ничего не делаем.
    Поддерживаем <input id="findStateRevoked"> и обёртку id="find-status-revoked".
    """
    # Пытаемся найти сам input
    cb = None
    try:
        cb = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "findStateRevoked"))
        )
    except TimeoutException:
        pass

    def _is_checked(inp):
        try:
            return (inp.get_attribute("checked") in ("true", "checked")) or inp.is_selected()
        except Exception:
            return False

    if cb is not None:
        if _is_checked(cb):
            safe_click(driver, cb)  # кликнём, чтобы СНЯТЬ галку
        return

    # fallback: кликаем по wrapper/label (если input не доступен)
    try:
        wrapper = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "find-status-revoked"))
        )
        try:
            inner_cb = wrapper.find_element(By.XPATH, ".//input[@type='checkbox' and @id='findStateRevoked']")
        except Exception:
            inner_cb = None

        if inner_cb is not None:
            if _is_checked(inner_cb):
                safe_click(driver, inner_cb)
        else:
            # если input скрыт, а wrapper переключает состояние — кликнем wrapper,
            # но только если текущее состояние действительно "включено".
            # Попробуем определить по aria-атрибутам или классу:
            state = (wrapper.get_attribute("aria-checked") or "").lower()
            if state in ("true", "checked"):
                safe_click(driver, wrapper)
    except TimeoutException:
        # чекбокс не найден — просто идём дальше
        pass

# выбрать дату, с какого числа ищем
def choose_date_since(driver, since_date: str):
    el = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "vetDocumentDateId"))
    )
    # если уже заполнено — ничего не делаем
    current = (el.get_attribute("value") or el.get_property("value") or "").strip()
    if current:
        return

    # иначе вводим дату
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)

    el.clear()
    el.send_keys(since_date)   # формат dd.MM.yyyy

# поиск
def open_search_panel(driver):
    """Клик по 'Поиск' (id=findFormTop) и ожидание модалки."""
    btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "findFormTop"))
    )
    safe_click(driver, btn)
    # ждём, что модалка открылась (ищем любой элемент формы)
    WebDriverWait(driver, 20).until(
        EC.any_of(
            EC.presence_of_element_located((By.XPATH, "//span[contains(@class,'ui-dialog-title') and contains(.,'Поиск')]")),
            EC.presence_of_element_located((By.XPATH, "//*[@id='waybillNumberId']"))
        )
    )
    time.sleep(0.5)
    # отмечаем «исключить аннулированные» (один раз, идемпотентно)
    _ensure_revoked_unchecked(driver)
    # выбираем дату, с которой будем осуществлять поиск
    choose_date_since(driver, "01.05.2025")

# сокрытие ненужной формы при поиске
def collapse_general_info(driver):
    """
    Сворачиваем 'Общая информация' только если поле ТТН ещё не видно.
    Если 'Номер ТТН' уже в DOM/видим — ничего не трогаем.
    """
    if driver.find_elements(By.ID, "waybillNumberId"):
        return
    # кнопка сворачивания
    try:
        toggle_btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "allInfoGroupBtn")))
        safe_click(driver, toggle_btn)
        return
    except TimeoutException:
        pass
    # заголовок секции
    try:
        header = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(@class,'ffGroupHeaderTitle') and contains(.,'Общая информация')]"))
        )
        safe_click(driver, header)
    except TimeoutException:
        pass

# поиск по конкретной ттн
def search_one_ttn(driver, ttn_number: str):
    """
    Поиск одного номера ТТН:
      Поиск → свернуть 'Общая информация' → ввести ТТН → Найти.
    """
    open_search_panel(driver)
    collapse_general_info(driver)

    # поле ввода ТТН (если секция транспорта ещё свёрнута, всё равно найдём — поле появится)
    try:
        ttn_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "waybillNumberId"))
        )
    except TimeoutException:
        # секция «Информация о транспорте» может быть свернута — раскроем её по заголовку/плюсу
        try:
            # плюс
            plus_btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.ID, "transportInfoGroupExpandBtn"))
            )
            safe_click(driver, plus_btn)
        except TimeoutException:
            # заголовок
            header = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((
                    By.XPATH, "//*[contains(@class,'ffGroupHeaderTitle') and contains(.,'Информация о транспорте')]"
                ))
            )
            safe_click(driver, header)

        ttn_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "waybillNumberId"))
        )

    ttn_input.clear()
    ttn_input.send_keys(ttn_number)
    time.sleep(0.3)

    # кнопка «Найти»
    try:
        find_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH, "//button[contains(@class,'positive') and contains(@onclick,'findDocuments')]"
            ))
        )
    except TimeoutException:
        find_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'positive') and contains(.,'Найти')]"))
        )
    safe_click(driver, find_btn)

    # ждём результаты
    wait_ajax_quiet(driver, timeout=20)
    WebDriverWait(driver, 20).until(
        EC.any_of(
            EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'documentList') or contains(@class,'documentList')]")),
            EC.presence_of_element_located((By.XPATH, "//table//tr"))
        )
    )
    time.sleep(1)
    print(f"✅ Поиск ТТН {ttn_number}: результаты загружены.")

# хелпер проверочки пустоты
def has_search_results(driver, timeout: int = 10) -> bool:
    """
    Возвращает True, если после поиска есть результаты (хотя бы один vetDocumentPk).
    Возвращает False, если видим 'Список пуст' или не нашли ни одной строки.
    """
    # дождёмся, чтобы ajax закончил рисовать список
    wait_ajax_quiet(driver, timeout=min(timeout, 6))

    # Быстрый чек на "Список пуст"
    try:
        empty_marker = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located(
                (By.XPATH, "//h4[normalize-space()='Список пуст']")
            )
        )
        if empty_marker.is_displayed():
            return False
    except TimeoutException:
        pass

    # Чек хотя бы одной строки со столбцом-галочкой (активной)
    try:
        _ = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@type='checkbox' and @name='vetDocumentPk']")
            )
        )
        return True
    except TimeoutException:
        return False

# хелпер для выбора 5 страниц для печати
def set_print_pages_range(driver, start:int=1, end:int=5):
    """
    В модалке печати:
      1) выбрать radio name='printScope' value='pages'
      2) ввести диапазон в input.pagesForPrintInput (без нажатия Enter)
    Вызываем КАЖДЫЙ РАЗ, т.к. модалка сбрасывается после нового поиска.
    """
    # 1) radio «Страницы»
    pages_radio = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='radio' and @name='printScope' and @value='pages']"))
    )
    # клик только если не выбран
    try:
        if not (pages_radio.get_attribute("checked") in ("true", "checked") or pages_radio.is_selected()):
            safe_click(driver, pages_radio)
        else:
            # чуть-чуть подождём для стабильности
            human_pause(0.2, 0.5)
    except Exception:
        safe_click(driver, pages_radio)

    # 2) поле диапазона
    pages_input = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "input.pagesForPrintInput[name='pagesForPrint']"))
    )
    pages_input.clear()
    pages_input.send_keys(f"{start}-{end}")
    human_pause(0.3, 0.7)

# открыть печать
def open_print_modal(driver):
    btn = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "printSettingsFormTop")))
    safe_click(driver, btn)
    WebDriverWait(driver, 20).until(
        EC.any_of(
            EC.presence_of_element_located((By.ID, "printFieldCommands")),
            EC.presence_of_element_located((By.ID, "printFormatXlsx"))
        )
    )

# выбор табличной формы ЭВСД
def choose_table_layout(driver):
    """
    Выбрать 'Печать журнала ВСД (табличная форма)'.
    Жмём раскрытие только если опции не видны.
    """
    layout_radio_loc = (By.ID, "printSchemaLayoutHeaderRadio")
    layout_text_loc  = (By.XPATH, "//td[contains(.,'Печать журнала ВСД (табличная форма)')]")
    expand_btn_loc   = (By.ID, "printSchemaLayoutBtn")

    def layout_option_visible():
        for by, sel in (layout_radio_loc, layout_text_loc):
            try:
                el = driver.find_element(by, sel)
                if el.is_displayed():
                    return el
            except Exception:
                pass
        return None

    el = layout_option_visible()
    if not el:
        try:
            expand_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(expand_btn_loc))
            safe_click(driver, expand_btn)
            WebDriverWait(driver, 10).until(lambda d: layout_option_visible())
        except TimeoutException:
            pass
        el = layout_option_visible()

    if not el:
        el = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Печать журнала ВСД (табличная форма)')]"))
        )
    safe_click(driver, el)

# .xlsx формат печати
def select_xlsx_format(driver):
    """Выбрать .xlsx — если уже выбран, пропускаем."""
    xlsx_radio = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "printFormatXlsx")))
    if not is_checked(xlsx_radio):
        safe_click(driver, xlsx_radio)

# выбираем кастомную схему печати
def select_print_schema(driver, schema_value: str = "130174", schema_text: str | None = None):
    """
    Выбрать схему из селекта (по value или по видимому тексту).
    Если уже выбрана — ничего не делаем.
    """
    sel_el = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "printSchemaSelect")))
    current = selected_option_value(sel_el)
    if (schema_text is None and current == schema_value):
        return
    sel = Select(sel_el)
    if schema_text:
        sel.select_by_visible_text(schema_text)
    else:
        sel.select_by_value(schema_value)
    wait_ajax_quiet(driver, timeout=10)

# клик по печати
def click_print(driver):
    """5) Нажать 'Сформировать'."""
    btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((
            By.XPATH, "//button[contains(@class,'positive') and contains(@onclick,'VetDocumentPrintForm.print')]"
        ))
    )
    safe_click(driver, btn)
    # даём форме стартовать генерацию
    time.sleep(2)

# функция печати
def print_xlsx_table(driver, schema_value: str = "130174", schema_text: str | None = None, pages=(1,5)):
    open_print_modal(driver)                # откроет, если закрыта
    # ❗ всегда задаём диапазон страниц, модалка сбрасывает настройки
    set_print_pages_range(driver, start=pages[0], end=pages[1])
    choose_table_layout(driver)             # кликнет только если надо
    select_print_schema(driver, schema_value=schema_value, schema_text=schema_text)
    select_xlsx_format(driver)              # не будет переключать, если уже .xlsx
    click_print(driver)                     # «Сформировать»
    print("✅ Отправили на формирование XLSX по выбранной схеме.")

# цикл для множественной выгрузки
def process_ttn_list(driver, ttn_list: list[str], schema_value: str = "130174",
                     pause_between: tuple[float, float] = (1.2, 2.5)):
    """
    Для каждого ТТН:
      - открыть Поиск (или переиспользовать открытый)
      - ввести ТТН и нажать Найти
      - открыть Печать, выбрать схему и .xlsx
      - нажать «Сформировать»
    """

    no_ttn_list = []

    for idx, ttn in enumerate(ttn_list, start=1):
        print(f"\n— [{idx}/{len(ttn_list)}] Обрабатываем ТТН: {ttn}")
        try:
            # Поиск
            search_one_ttn(driver, ttn)

            if not has_search_results(driver):
                print(f"⏭️  ТТН {ttn}: список пуст — переходим к следующей.")
                no_ttn_list.append(ttn)
                human_pause(*pause_between)
                continue

            # Печать (с твоей схемой)
            print_xlsx_table(driver, schema_value="130174", pages=(1, 5))

            # можно добавить ожидание завершения скачивания файла
            # wait_download_xlsx(download_dir)  # допишем позже (возможно)

        except Exception as e:
            print(f"⚠️ Ошибка на ТТН {ttn}: {e}")

        # «человеческая» пауза между ТТН (чтобы не палиться и дать серверу отдышаться)
        human_pause(*pause_between)
    print(f"\n📄 Ненайденные ТТН: {len(no_ttn_list)}")
    return no_ttn_list

# сохранение списка ненайденных ттн
def save_no_ttn_to_excel(no_ttn_list: list[str], out_path: Path | None = None) -> Path:
    if out_path is None:
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        out_path = Path(__file__).with_name(f"not_found_ttns_{ts}.xlsx")
    if not no_ttn_list:
        print("✅ Пусто: все ТТН нашлись, файл не создаю.")
        return out_path
    pd.DataFrame({"TTN": no_ttn_list}).to_excel(out_path, index=False)
    print(f"💾 Сохранил ненайденные ТТН: {out_path}")
    return out_path

# папочка для скачивания
if os.name == "nt":
    download_dir = Path(os.getenv("USERPROFILE")) / "Desktop" / "actual_lab_tests"
else:
    download_dir = Path(os.getenv("HOME")) / "Desktop" / "actual_lab_tests"
