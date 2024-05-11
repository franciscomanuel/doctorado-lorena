import pandas as pd
import re


def reverse_data(data):
    for sheet_name, df in data.items():
        data[sheet_name] = df.iloc[::-1].reset_index(drop=True)


def clean_column_values(data, column_names):
    for sheet_name, df in data.items():
        for column_name in column_names:
            df[column_name] = df[column_name].astype(str).apply(
                lambda x: re.sub(r'(\d+(?:\.\d+)?).*', r'\1', x) if isinstance(x, str) else x)
            df[column_name] = pd.to_numeric(df[column_name], errors='coerce')


def save_results_to_excel(data, output_file):
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, df in data.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)


class ReverseRows:

    def __init__(self, file_path):
        self.file_path = file_path

    def load_data(self):
        return pd.read_excel(self.file_path, sheet_name=None)

    def process_all_sheets(self, output_file, columns_to_clean=None):
        data = self.load_data()
        reverse_data(data)
        if columns_to_clean:
            clean_column_values(data, columns_to_clean)
        save_results_to_excel(data, output_file)
