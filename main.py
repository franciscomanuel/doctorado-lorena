import os

from data_processor import DataProcessor
from csv_to_xlsx_converter import CSVtoXLSXConverter
from constants import constants


def calculate_average():
    input_folder = "assets/la_albufera/input";
    output_folder = "assets/la_albufera/output";

    for filename in os.listdir(input_folder):
        input_file = os.path.join(input_folder, filename)
        output_file = os.path.join(output_folder, filename)

        data_processor = DataProcessor(input_file)
        data_processor.load_data(constants.FIRST_SHEET_NAME)

        try:
            data_processor.sort_data([constants.DATE_COLUMN, constants.PERIOD_COLUMN])
            data_processor.select_numerical_data([constants.PERIOD_COLUMN])
            data_processor.calculate_averages_by_date(constants.DATE_COLUMN)
            data_processor.save_results_to_excel(output_file)
        except ValueError as e:
            print(f"Error procesando el archivo {input_file}: {e}")
            continue


def transform_csv_to_xlsx():
    input_folder = "assets/csv/input"
    output_folder = "assets/csv/output"

    for filename in os.listdir(input_folder):
        input_file = os.path.join(input_folder, filename)
        output_file = os.path.join(output_folder, filename.split(".")[0] + ".xlsx")

        converter = CSVtoXLSXConverter(input_file, output_file)
        converter.convert()


if __name__ == '__main__':
    calculate_average();
    # transform_csv_to_xlsx()
