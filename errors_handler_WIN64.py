import os
import pandas as pd

def find_suspicious_blocks(df_exploded):
    """
    Находит подозрительные блоки в DataFrame (битые строки после explode).
    Возвращает errors_df.
    """
    mask = (
        (df_exploded['Лаборатория'].isna() | (df_exploded['Лаборатория'] == '')) &
        (df_exploded['Номер и дата лабораторного исследования'].isna() | (df_exploded['Номер и дата лабораторного исследования'] == '')) &
        (
            df_exploded['Результат лабораторного исследования'].str.strip().str.lower().isin([', отрицательный)', 'отрицательный)']) |
            df_exploded['Результат лабораторного исследования'].str.strip().str.lower().str.startswith(', отрицательный')
        )
    )
    errors_df = df_exploded[mask]
    return errors_df

def log_errors(errors_df, path='Ошибки.txt'):
    main_columns = [
        '№ п/п', 'Наименование продукции', 'Производитель',
        'Результат лабораторного исследования', 'Номер ТТН', 'Дата ТТН'
    ]
    if errors_df.empty:
        print("Подозрительных блоков не найдено!")
        return
    with open(path, 'w', encoding='utf-8') as f:
        f.write(f"Всего подозрительных строк: {len(errors_df)}\n\n")
        for idx, row in errors_df.iterrows():
            f.write(f"Индекс в Excel: {idx+2}\n")
            for col in main_columns:
                if col in errors_df.columns:
                    val = str(row[col])
                    f.write(f"{col}: {val}\n")
            f.write("\n---\n")
    print(f"Список подозрительных строк сохранён в {path}")

def remove_suspicious_blocks(df_exploded, errors_df):
    """
    Удаляет все строки из df_exploded, которые есть в errors_df.
    Возвращает новый DataFrame без битых строк.
    """
    return df_exploded.drop(errors_df.index).reset_index(drop=True)


def export_rows_without_lab_results(combined_path: str, save_path: str) -> bool:
    """
    Сохраняет в save_path строки из combined_path, где
    'Результат лабораторного исследования' пустой/NaN.
    Возвращает True, если файл создан (строки были), иначе False.
    """
    if not os.path.exists(combined_path):
        return False

    df = pd.read_excel(combined_path, engine='openpyxl', dtype={'Номер ТТН': str})
    if 'Результат лабораторного исследования' not in df.columns:
        return False

    col = df['Результат лабораторного исследования']
    mask = col.isna() | (col.astype(str).str.strip() == '')
    df_empty = df[mask].copy()

    if df_empty.empty:
        return False

    with pd.ExcelWriter(save_path, engine='openpyxl') as w:
        df_empty.to_excel(w, index=False, sheet_name='Без ЛИ')
        ws = w.book.active
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = 'A2'

    return True



