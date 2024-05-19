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


def fill_dates(data, date_column='fecha'):
    for sheet_name, df in data.items():
        # Asegurar que la columna de fecha esté en formato datetime
        df[date_column] = pd.to_datetime(df[date_column], errors='coerce')

        # Obtener el rango completo de fechas para cada año presente en los datos
        all_dates = pd.DataFrame()
        for year in df[date_column].dt.year.unique().astype(int):
            full_date_range = pd.date_range(start=f'{year}-01-01', end=f'{year}-12-31', freq='D')
            full_df = pd.DataFrame(full_date_range, columns=[date_column])
            all_dates = pd.concat([all_dates, full_df], ignore_index=True)

        # Combinar el DataFrame original con el DataFrame de rango completo de fechas
        df = all_dates.merge(df, on=date_column, how='left')

        # Ordenar los datos por fecha y restablecer el índice
        df = df.sort_values(by=date_column).reset_index(drop=True)

        # Actualizar el DataFrame en el diccionario con los datos completos
        data[sheet_name] = df


def save_results_to_excel(data, output_file):
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, df in data.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)


class ReverseRows:

    def __init__(self, file_path):
        self.file_path = file_path

    def load_data(self):
        return pd.read_excel(self.file_path, sheet_name=None)

    def process_all_sheets(self, output_file, columns_to_clean=None, date_column='Fecha'):
        data = self.load_data()
        reverse_data(data)
        if columns_to_clean:
            clean_column_values(data, columns_to_clean)
        fill_dates(data, date_column=date_column)
        save_results_to_excel(data, output_file)
