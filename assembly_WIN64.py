import pandas as pd
import os

def lab_assembler(input_dir, raw_combined_file):
    files = os.listdir(input_dir)
    data_frames = []
    for file in files:
        if file.endswith('.xlsx'):
            df = pd.read_excel(os.path.join(input_dir, file), dtype={'Номер ТТН': str})
            data_frames.append(df)
    combined_df = pd.concat(data_frames, ignore_index=True)
    writer = pd.ExcelWriter(raw_combined_file, engine='xlsxwriter')
    combined_df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()



