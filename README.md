**Geopolitical Risk Index Scraping Test**


This project is a web scraping test performed on the page: https://www.geopriskindex.com/results-final-risk-index/.

Objective:

The main objective of this test is to extract the data from the table containing the Geopolitical Risk Index (GPR) and transform it into a structured (wide) format for further analysis or export.

Script Description:

The Python script used for this test performs the following actions:

HTTP Request: Uses the requests library to obtain the HTML content of the specified webpage.
HTML Parsing: Employs the BeautifulSoup library to parse the downloaded HTML and facilitate navigation through its structure.
Table Extraction: Locates the data table within the HTML.
Header and Data Extraction: Identifies and extracts the column names (headers) and the data from each row of the table.
DataFrame Creation (Pandas): Organizes the extracted data into a Pandas DataFrame, a powerful tabular data structure for data analysis in Python.
Wide Format Transformation: The initial DataFrame is transformed into a "wide" format where each country and year has a single row, and the different index variables (Financial Index, Political Risk Index, etc.) become separate columns.
Excel Export: Finally, the wide-format DataFrame is saved to an Excel file (formato_wide_geopriskinder.xlsx) with a worksheet named "Índice de Riesgo Global".
Column Width Adjustment (Optional): The script includes functionality to automatically adjust the width of the columns in the Excel file to make the text readable.
Result:

Upon executing the script, an Excel file (formato_wide_geopriskinder.xlsx) will be generated in the project directory. This file will contain the Geopolitical Risk Index data in a wide format, ready for use in other analysis or visualization tools.

**Libraries Used:**

**requests**: For making HTTP requests.

**beautifulsoup4**: For parsing HTML.

**pandas**: For data manipulation and analysis (DataFrames).

**openpyxl**: For writing and modifying Excel files.
