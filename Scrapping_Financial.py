import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

url = "https://www.geopriskindex.com/results-final-risk-index/"

try:
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find('table')

    if table:
        headers = [th.text.strip() for th in table.find_all('th')]
        data = []
        tbody = table.find('tbody')
        if tbody:
            for row in tbody.find_all('tr')[1:]:
                cols = [td.text.strip() for td in row.find_all('td')]
                data.append(cols)
        else:
            print("No se encontró el cuerpo de la tabla (tbody).")
            exit()

        df = pd.DataFrame(data, columns=headers)
        print("Encabezados del DataFrame:", headers)

        # Crear el DataFrame en formato ancho
        unique_years = df['Year'].unique()
        country_column_name = '\ufeffCountry'

        unique_countries = df[country_column_name].unique()
        df_wide = pd.DataFrame()
        df_wide['Date'] = [f"{year}-mm-dd" for year in unique_years for _ in unique_countries]
        df_wide['Country'] = [country for _ in unique_years for country in unique_countries]

        for header in headers:
            if header not in [country_column_name, 'Year', 'Region']:
                values = []
                for year in unique_years:
                    df_year = df[df['Year'] == year]
                    country_value_map = df_year.set_index(country_column_name)[header].to_dict()
                    for country in unique_countries:
                        values.append(country_value_map.get(country, ''))
                df_wide[header] = values

        excel_filename = "formato_wide_geopriskinder.xlsx"
        sheet_name = 'Índice de Riesgo Global'
        df_wide.to_excel(excel_filename, index=False, sheet_name=sheet_name)
        print(f"Se creo el archivo Excel: {excel_filename}")
        workbook = load_workbook(excel_filename)
        sheet = workbook[sheet_name]
        for column_cells in sheet.columns:
            max_len = 0
            column = [cell.value for cell in column_cells]
            try:
                max_len = max(len(str(value)) for value in column if value is not None)
            except ValueError:
                pass 

            adjusted_width = (max_len + 2) 
            sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = adjusted_width

        workbook.save(excel_filename)

    else:
        print("No se encontró ninguna tabla en la página.")

except requests.exceptions.RequestException as e:
    print(f"Error al acceder a la página: {e}")
except Exception as e:
    print(f"Ocurrió un error: {e}")