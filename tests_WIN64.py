import re
import pandas as pd

def normalize(s):
    """Привести к единому виду: NaN -> '', убрать множественные пробелы/переводы строк и крайние ;"""
    if s is None:
        return ''
    # обработка NaN
    try:
        if pd.isna(s):
            return ''
    except Exception:
        pass
    s = str(s).replace('\n', ' ')
    s = re.sub(r'\s+', ' ', s).strip(' ;\t')
    return s

def run_validators(df, df_exploded=None):
    """
    df — DataFrame с колонками:
      - 'Результат лабораторного исследования' (оригинал)
      - 'lab_blocks' (список блоков после парсинга)
    Возвращает:
      errors: список dict с расхождениями (без пустых ЛИ)
      empty_rows: список индексов строк df, где ЛИ пустой/отсутствует
    """
    errors = []
    empty_rows = []

    for idx, row in df.iterrows():
        orig = normalize(row.get('Результат лабораторного исследования', ''))
        blocks = row.get('lab_blocks', []) or []
        rebuilt = normalize('; '.join(blocks))

        # Пустые ЛИ считаем отдельным кейсом, НЕ ошибкой парсинга
        if orig == '':
            empty_rows.append(idx)
            continue

        if orig != rebuilt:
            errors.append({
                'index': idx,                 # индекс в df
                'excel_row': idx + 2,         # человеко-читаемый номер строки в Excel
                'orig': orig,
                'rebuilt': rebuilt,
                'count_orig_blocks': (orig.count(';') + 1) if orig else 0,
                'count_rebuilt_blocks': len(blocks),
            })

    print(f"Несовпадающих строк (без пустых ЛИ): {len(errors)} из {len(df)}")
    if empty_rows:
        print(f"Пустых ЛИ (в исходнике): {len(empty_rows)}")

    # Для наглядности первые 10 расхождений
    if errors:
        for err in errors[:10]:
            print("\n---")
            print(f"Строка (Excel): {err['excel_row']}")
            print(f"Оригинал:   {err['orig']}")
            print(f"Выгрузка:   {err['rebuilt']}")
            print(f"Исходных блоков: {err['count_orig_blocks']}, собралось: {err['count_rebuilt_blocks']}")

    return errors, empty_rows
