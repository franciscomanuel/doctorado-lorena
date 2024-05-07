import pandas as pd

from constants import constants


class DataProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None
        self.numerical_data = None
        self.averages_by_date = None

    def load_data(self, sheet_name):
        self.data = pd.read_excel(self.file_path, sheet_name=sheet_name)

    def sort_data(self, columns):
        if not all(col in self.data.columns for col in columns):
            missing_columns = [col for col in columns if col not in self.data.columns]
            raise ValueError(f"Las columnas {missing_columns} no están presentes en los datos.")

        self.data = self.data.sort_values(by=columns)

    def select_numerical_data(self, columns_to_exclude):
        data = self.data.select_dtypes(include=['number', 'datetime64']);
        if columns_to_exclude in data.columns.values:
            self.numerical_data = data.drop(columns=columns_to_exclude)
        else:
            self.numerical_data = data

    def calculate_averages_by_date(self, date_column):
        self.averages_by_date = {}
        for date, group in self.numerical_data.groupby(date_column):
            self.averages_by_date[date] = group.drop(columns=[date_column]).sum() / 24  # Calculamos la media por hora

    def save_results_to_excel(self, output_file):
        medias_df = pd.DataFrame.from_dict(self.averages_by_date, orient='index')
        medias_df.reset_index(inplace=True)
        medias_df.rename(columns={'index': constants.DATE_COLUMN}, inplace=True)
        medias_df.to_excel(output_file, sheet_name=constants.FIRST_SHEET_NAME, index=False)
