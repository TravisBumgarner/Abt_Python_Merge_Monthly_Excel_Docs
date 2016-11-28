import pandas as pd

def csv_from_excel(file_name):
	xls = pd.ExcelFile(file_name)
	df = xls.parse('Sheet1', index_col=None, na_values=['NA'])
	df.to_csv(file_name[:-5] + ".csv", index = False)

csv_from_excel("Exports/2016-07.xlsx")