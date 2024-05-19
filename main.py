import os

from data_processor import DataProcessor
from csv_to_xlsx_converter import CSVtoXLSXConverter
from reverse_rows import ReverseRows
from constants import constants


def process_files(input_folder, output_folder, process_function, is_csv = False):
    for filename in os.listdir(input_folder):
        input_file = os.path.join(input_folder, filename)
        output_file = os.path.join(output_folder, filename.split(".")[0] + ".xlsx") if is_csv else os.path.join(
            output_folder, filename)

        try:
            process_function(input_file, output_file)
        except Exception as e:
            print(f"Error procesando el archivo {input_file}: {e}")


def calculate_average(input_file, output_file):
    data_processor = DataProcessor(input_file)
    data_processor.load_data(constants.FIRST_SHEET_NAME)

    data_processor.sort_data([constants.DATE_COLUMN, constants.PERIOD_COLUMN])
    data_processor.select_numerical_data([constants.PERIOD_COLUMN])
    data_processor.calculate_averages_by_date(constants.DATE_COLUMN)
    data_processor.save_results_to_excel(output_file)


def transform_csv_to_xlsx(input_file, output_file):
    converter = CSVtoXLSXConverter(input_file, output_file)
    converter.convert()


def apply_reverse_row(input_file, output_file):
    reverse_rows = ReverseRows(input_file)
    reverse_rows.process_all_sheets(output_file, columns_to_clean=['SO2', 'NO2', 'O3', 'CO', 'PM10', 'NH2'])


if __name__ == '__main__':
    # input_folder = "assets/foners/input"
    # output_folder = "assets/foners/output"
    # process_files(input_folder, output_folder, calculate_average)
    #
    # input_folder = "assets/csv/input"
    # output_folder = "assets/csv/output"
    # process_files(input_folder, output_folder, transform_csv_to_xlsx, True)

    input_folder = "assets/reverse/input"
    output_folder = "assets/reverse/output"
    process_files(input_folder, output_folder, apply_reverse_row)
