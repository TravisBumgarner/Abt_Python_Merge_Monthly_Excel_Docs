import openpyxl
from process_data_functions import generate_header_names, generate_exported_files_list
from openpyxl import Workbook #Interact with Excel
import os #Handle file names and paths


#Get list of current files
exported_files_directory = "Exports"
exported_files_list = generate_exported_files_list(exported_files_directory)
#Base is the starting file for matching all newer files. It is the oldest file in the folder. 
base_file = exported_files_directory + "/" + exported_files_list[0]
base_file_name = os.path.basename(base_file)[:-5]
base_workbook = openpyxl.load_workbook(filename = base_file)
base_sheet = base_workbook['Sheet1']

#Merged will be where all the information is combined together
merged_file = "MergedResults.xlsx"
merged_workbook = openpyxl.Workbook()
merged_sheet = merged_workbook.active 
merged_sheet.title = "Sheet1"

#Type in list of columns to use which will also generate a list of the names of the headers
base_columns = [1,2,3,4]
id_column = 1 #Column that has the unique identifier, typically an ID, is located to do comparisons
base_headers = generate_header_names(base_sheet,base_columns)
merged_date_column = "E"
merged_date_header = "Modified Date"

merged_sheet[1,1].value