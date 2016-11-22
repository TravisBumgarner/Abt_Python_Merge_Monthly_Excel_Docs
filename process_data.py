import openpyxl
from process_data_functions import generate_modified_date_header_names, generate_modified_date_column_ints, generate_string_from_column_index, generate_header_names, generate_exported_files_list, generate_column_index_from_string, generate_modified_date_header_names
from openpyxl import Workbook #Interact with Excel
import os #Handle file names and paths


#User Defined Settings:
directory_with_exported_files = "Exports"
filename_of_merged_excel_docs = "MergedResults.xlsx"
workbook_sheet_name_of_memrged_excel_doc = "Sheet1"
columns_to_merge_together = ["A","B","C","D"]
column_with_unique_id = "A"

#Get list of current files
exported_files_directory = directory_with_exported_files
exported_files_list = generate_exported_files_list(exported_files_directory)

#Base is the starting file for matching all newer files. It is the oldest file in the folder. 
base_file = exported_files_directory + "/" + exported_files_list[0]
base_file_name = os.path.basename(base_file)[:-5]
base_workbook = openpyxl.load_workbook(filename = base_file)
base_sheet = base_workbook['Sheet1']

#Merged will be where all the information is combined together
merged_file = filename_of_merged_excel_docs
merged_workbook = openpyxl.Workbook()
merged_sheet = merged_workbook.active 
merged_sheet.title = workbook_sheet_name_of_memrged_excel_doc

#Type in list of columns to use which will also generate a list of the names of the headers
base_columns = columns_to_merge_together 
base_columns_ints = generate_column_index_from_string(base_columns)
id_column = column_with_unique_id #Column that has the unique identifier, typically an ID, is located to do comparisons
base_columns_without_id = base_columns[:]
base_columns_without_id.remove(id_column)
base_columns_without_id_ints =  generate_column_index_from_string(base_columns_without_id)
#print(base_columns_without_id)
base_headers = generate_header_names(base_sheet,base_columns)

merged_date_modified_columns_ints = generate_modified_date_column_ints(base_columns_ints)
len_of_merged_columns = len(merged_date_modified_columns_ints)
merged_date_modified_columns = generate_string_from_column_index(merged_date_modified_columns_ints)
#print(merged_date_modified_columns)
merged_date_modified_headers = generate_modified_date_header_names(base_headers)

merged_date_created_column = generate_string_from_column_index(max(base_columns_ints)*2 + 1) #add this column after the last column of the modified file  (len(base_columns) + len(modified_columns)) gives the last column, +1 to get an empty column 
merged_date_created_header = "Created Date"

#print(merged_date_created_column)

#Copy the values of the headers in the base_sheet to the merged_sheet
for index in range(0,len(base_columns)):
	merged_sheet[base_columns[index] + "1"].value = base_headers[index]

#print(merged_date_modified_headers)
#print(merged_date_modified_columns)
for index in range(0,len(merged_date_modified_columns_ints)):
	#print(index)
	merged_sheet[merged_date_modified_columns[index]+ "1"].value = merged_date_modified_headers[index]


merged_sheet[merged_date_created_column + "1"] = merged_date_created_header

#Copy the values of the cells in the base_sheet to the merged_sheet
for index_row in range(2,base_sheet.max_row+1):
	for index_col in base_columns:
		merged_sheet[index_col+str(index_row)].value = base_sheet[index_col+str(index_row)].value
	merged_sheet[merged_date_created_column + str(index_row)] = base_file_name


def loop_through_non_base_files(exported_files_list): 
	#loop through all files not used to generate base file by using 1:
	for exported_file in exported_files_list[1:]:
		print(exported_files_list)
		#Loop will be where all the information for the current Excel doc will be stored
		loop_file = exported_files_directory + "/" + exported_file
		loop_file_name = os.path.basename(loop_file)[:-5]
		loop_workbook = openpyxl.load_workbook(filename = loop_file)
		loop_sheet = loop_workbook['Sheet1']

		#For each row in the loop file, compare it against the Merged file by using
		#The ID as a comparison. If a user ID from the loop file is found in the merged
		#file, comparisons of answers will be made. Otherwise Continue on.
		#print("Comparing " + loop_file + " and " + base_file)
		loop_rows_not_found = []
		for loop_row in range(2,loop_sheet.max_row+1):
			id_row_found = False #Once the id_row match is found, send this variable to true
			for merged_row in range(2,merged_sheet.max_row+1):
				#print("Comparing loop:" + str(loop_sheet[id_column + str(loop_row)].value) + " and merged:" + str(merged_sheet[id_column + str(merged_row)].value))
				if(loop_sheet[id_column + str(loop_row)].value == merged_sheet[id_column + str(merged_row)].value):
					#print(merged_sheet[id_column + str(merged_row)].value)
					id_row_found = True

					#In this space will be a function that compares each 

					for base_column in base_columns_without_id:
						base_column_int = generate_column_index_from_string(base_column)
						if(loop_sheet[base_column + str(loop_row)].value != merged_sheet[base_column + str(merged_row)].value ):
							#print(base_column_int)
							#print(len_of_merged_columns)
							merged_sheet[generate_string_from_column_index(base_column_int[0] + len_of_merged_columns) + str(merged_row)].value = loop_file_name
							merged_sheet[base_column + str(merged_row)].value = loop_sheet[base_column + str(loop_row)].value
					break
			if(id_row_found != True):
				loop_rows_not_found.append(loop_row)
		#print(loop_rows_not_found)

		# For the rows that were not found, append them to the end of the document. 
		merged_sheet_empty_last_row = merged_sheet.max_row+1
		for loop_row_not_found in loop_rows_not_found:
			for  index_col in base_columns:
				 merged_sheet[index_col+str(merged_sheet_empty_last_row)].value = loop_sheet[index_col+str(loop_row_not_found)].value
			merged_sheet[merged_date_created_column + str(merged_sheet_empty_last_row)].value = loop_file_name
			merged_sheet_empty_last_row += 1
		merged_workbook.save(merged_file)	

loop_through_non_base_files(exported_files_list)



#Check whether the above cells match the expected string contents
"""
error_count = 0
for cell_key, cell_value in header_validation.items():
	if(base_sheet[cell_key].value != header_validation[cell_key]):
		print("There was an error in sheet " + base_file + " at " + cell_key + 
			  "\n'" + base_sheet[cell_key].value + "' does not match '" + header_validation[cell_key] +"'\n")
		error_count +=1
if(error_count != 0):
	print("Fix " + str(error_count) +" error(s) before continuing.")
"""


