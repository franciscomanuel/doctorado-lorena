import pandas as pd

from constants import constants


class CSVtoXLSXConverter:
    def __init__(self, csv_file, xlsx_file):
        self.csv_file = csv_file
        self.xlsx_file = xlsx_file

    def convert(self):
        # Lista para almacenar los datos del CSV
        data = []

        # Leer el archivo CSV línea por línea
        with open(self.csv_file, 'r') as file:
            lines = file.readlines()
            for idx, line in enumerate(lines):
                # Dividir la línea en valores individuales utilizando la coma como delimitador
                values = [value.strip('"') for value in line.strip().split(',')]

                if idx == 0:
                    values = [value.upper() for value in values]
                else:
                    # Convertir los valores a enteros si es posible
                    for i, value in enumerate(values):
                        try:
                            values[i] = int(value)
                        except ValueError:
                            pass  # Ignorar si no se puede convertir a entero

                # Agregar los valores a la lista de datos
                data.append(values)

        # Crear un DataFrame a partir de los datos
        df = pd.DataFrame(data)

        # Guardar el DataFrame en un archivo Excel
        print(self.xlsx_file)
        df.to_excel(self.xlsx_file, sheet_name=constants.FIRST_SHEET_NAME, index=False, header=None)
