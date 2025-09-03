import re

def normalize(s):
    """Удаляет лишние пробелы, \n, точки с запятой в начале/конце."""
    return re.sub(r'\s+', ' ', str(s)).replace('\n', '').strip(' ;')

def run_tests(df):
    """
    df — DataFrame, в котором должны быть столбцы:
      - 'Результат лабораторного исследования' (оригинал)
      - 'lab_blocks' (список блоков после парсинга)
    Функция возвращает список ошибок (различий).
    """
    errors = []
    for idx, row in df.iterrows():
        orig = normalize(row['Результат лабораторного исследования'])
        blocks = row['lab_blocks']
        rebuilt = normalize('; '.join(blocks))
        if orig != rebuilt:
            errors.append({
                'index': idx,
                'orig': orig,
                'rebuilt': rebuilt,
                'count_orig_blocks': orig.count(';') + 1 if orig else 0,
                'count_rebuilt_blocks': len(blocks)
            })

    print(f'Несовпадающих строк: {len(errors)} из {len(df)}')
    if errors:
        for err in errors[:10]:  # максимум 10 различий в messagebox'е
            print("\n---")
            print(f"Строка: {err['index'] + 2}")
            print(f"Оригинал:   {err['orig']}")
            print(f"Выгрузка:   {err['rebuilt']}")
            print(f"Исходных блоков: {err['count_orig_blocks']}, собралось: {err['count_rebuilt_blocks']}")
    return errors  # Возвращаем ошибки для GUI/скриптов

