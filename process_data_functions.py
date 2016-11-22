import os
import openpyxl

def generate_header_names(workbook_sheet, columns):
	header_names = []
	for column in columns:
		header_names.append(workbook_sheet[column + "1"].value)
	return header_names

def generate_exported_files_list(files_directory):
	exported_files_list = []
	for root, dirs, excel_workbooks, in os.walk("Exports"):
		for workbook in excel_workbooks:
			exported_files_list.append(workbook)
	return exported_files_list

def generate_column_index_from_string(list_of_columns):
	modified_list_of_columns = []
	for each in list_of_columns:
		modified_each = openpyxl.utils.column_index_from_string(each)
		modified_list_of_columns.append(modified_each)
	return modified_list_of_columns

def generated_modified_date_header_names(base_headers):
	modified_headers = []
	for each in base_headers:
		modified_headers.append(each + "-Modified")
	return modified_headers

def generated_modified_date_column_ints(base_columns_ints):
	modified_columns_ints =[]
	for each in base_columns_ints:
		modified_columns_ints.append(each+len(base_columns_ints))
	return modified_columns_ints