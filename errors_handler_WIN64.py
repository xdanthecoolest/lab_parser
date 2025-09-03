
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

