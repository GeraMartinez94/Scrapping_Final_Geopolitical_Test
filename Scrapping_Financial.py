import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --- Configuración  #
URL = "https://www.geopriskindex.com/results-final-risk-index/"
EXCEL_FILENAME = "formato_wide_geopriskinder.xlsx"
SHEET_NAME = 'Índice de Riesgo Global'
COUNTRY_COLUMN_NAME = '\ufeffCountry'

try:
    # Obtiene los datos de la web #
    response = requests.get(URL)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find('table')

    if table:
        # Extrae los encabezados #
        headers = [th.text.strip() for th in table.find_all('th')]

        # Extraer datos de filas #
        data = []
        tbody = table.find('tbody')
        if tbody:
            for row in tbody.find_all('tr')[1:]:
                cols = [td.text.strip() for td in row.find_all('td')]
                data.append(cols)
        else:
            print("No se encontró el cuerpo de la tabla (tbody).")
            exit()

        # Crea DataFrame inicial #
        df = pd.DataFrame(data, columns=headers)
        print("Encabezados del DataFrame:", headers)

        # Crear DataFrame en formato ancho #
        unique_years = df['Year'].unique()
        unique_countries = df[COUNTRY_COLUMN_NAME].unique()
        df_wide = pd.DataFrame()
        df_wide['Date'] = [f"{year}-mm-dd" for year in unique_years for _ in unique_countries]
        df_wide['Country'] = [country for _ in unique_years for country in unique_countries]

        for header in headers:
            if header not in [COUNTRY_COLUMN_NAME, 'Year', 'Region']:
                values = []
                for year in unique_years:
                    df_year = df[df['Year'] == year]
                    country_value_map = df_year.set_index(COUNTRY_COLUMN_NAME)[header].to_dict()
                    for country in unique_countries:
                        values.append(country_value_map.get(country, ''))
                df_wide[header] = values

        # Guarda el DataFrame en  archivo Excel #
        df_wide.to_excel(EXCEL_FILENAME, index=False, sheet_name=SHEET_NAME)
        print(f"Se creo el archivo Excel: {EXCEL_FILENAME}")

        # Ajusta ancho de columnas en Excel #
        workbook = load_workbook(EXCEL_FILENAME)
        sheet = workbook[SHEET_NAME]

        for column_cells in sheet.columns:
            max_len = 0
            column = [cell.value for cell in column_cells]
            try:
                max_len = max(len(str(value)) for value in column if value is not None)
            except ValueError:
                pass

            adjusted_width = max_len + 2
            sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = adjusted_width

        workbook.save(EXCEL_FILENAME)

    else:
        print("No se encontró ninguna tabla en la página.")

except requests.exceptions.RequestException as e:
    print(f"Error al acceder a la página: {e}")
except Exception as e:
    print(f"Ocurrió un error: {e}")