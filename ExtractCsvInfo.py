import pandas as pd
from openpyxl import load_workbook

excel_file_location1 = r'/Users/auarangu/Downloads/Invitación Bautizo Elegante Dorado-3/extracted_info-CP.csv'
excel_file_location2 = r'/Users/auarangu/Downloads/Invitación Bautizo Elegante Dorado-3/extracted_info.csv'

wb1 = pd.read_csv(excel_file_location1)
wb2 = pd.read_csv(excel_file_location2)

df = pd.concat([wb1, wb2])

df.to_csv('Final-ExtractedSessions.csv', index=False)

