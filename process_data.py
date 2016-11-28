import pandas as pd # to convert xlsx to csv
import os
import csv

##################################################################
#User Defined Settings############################################
##################################################################
csv_directory = "Exports"
results = "results.csv"
template_file = "2016-07.csv" #Specify oldest file to act as the template for results.csv
columns_to_modify = [1,3]
column_with_unique_id = 0

#number_columns_to_modify = columns_to_modify[:]
#number_columns_to_modify.remove(column_with_unique_id)
#number_columns_to_modify = len(number_columns_to_modify)




#columns_without_id = columns_to_modify[:]
#columns_without_id.remove(column_with_unique_id)

##################################################################
#Setup Base File, Merged File, and Directory of Exported Files####
##################################################################

for root, dirs, xlsx_files, in os.walk(csv_directory):
	for xlsx_file in xlsx_files:
		if xlsx_file.endswith('.xlsx'):
			xls = pd.ExcelFile(csv_directory+ "/" + xlsx_file)
			df = xls.parse('Sheet1', index_col=None, na_values=['NA'])
			df.to_csv(csv_directory + "/" + xlsx_file[:-5] + ".csv", index=False)

def open_csv_as_list(file_dir):
	with open(file_dir) as csvfile:
		reader = csv.reader(csvfile,delimiter=",",quotechar='"')
		return list(reader)

def save_list_as_csv(csv_list,file_name):
	with open(results, 'w', newline ='') as csvfile:
		writer = csv.writer(csvfile, delimiter = ',', quotechar ="'")
		for row in merged:
			writer.writerow(row)
		csvfile.close()

base = open_csv_as_list(csv_directory + "/" + template_file)
"""
def extract_columns(csv_file,columns_to_modify):
	#Remove columns that are not needed
	csv_extracted = []
	for row in range(0,len(csv_file)):
		extracted_row = []
		for column in columns_to_modify:
			extracted_row.append(csv_file[row][column])
		csv_extracted.append(extracted_row)
	return csv_extracted
"""

#base_extract = extract_columns(base,columns_to_modify)
base_extract = base

#Created -Mod Date Columns
columns_without_id = [col for col in columns_to_modify if col != column_with_unique_id]
merged_columns = []
for column in columns_without_id:
	merged_columns.append(base[0][column] + "- Mod Date")

#Create Merged Document Headers
merged = [base_extract[0] + merged_columns + ["Created Date"]]

##################################################################
#Loop through the remaining .csv files ###########################
##################################################################
def merge_and_date_csv_files(csv_directory, merged_result,id_column, columns_to_modify):
	list_of_csv_files = []
	for root, dirs, csv_files, in os.walk(csv_directory):
		for csv_file in csv_files:
			if csv_file.endswith('.csv'):
				list_of_csv_files.append(csv_file)

	for csv_file in sorted(list_of_csv_files): #Sort by date
		print("Looping through " + csv_file)
		csv_as_a_list = open_csv_as_list(csv_directory + "/" + csv_file)
		#list_looping = extract_columns(csv_as_a_list,columns_to_modify)
		list_looping = csv_as_a_list

		number_of_base_columns = len(list_looping[0])
		empty_columns = []
		for i in range(0, number_of_base_columns - len(columns_to_modify)):
			empty_columns.append("Null")
		#Append Rows to merged doc
		#If the merged csv only has the headers, add entries from the first csv file
		if len(merged_result) == 1:
			for row_looping in range(1,len(list_looping)):
				merged_result.append(list_looping[row_looping] + empty_columns + [csv_file[:-4]]) #empty_columns is a generated list of "Null" with a len() of the # of date modified columns added to the csv. Then add the created date

		#Generate list of ids
		id_list = []
		for i in range(1,len(merged_result)):
			id_list.append(merged_result[i][id_column])

		else:
			for row_looping in range(1,len(list_looping)):
				for row_merged in range(1, len(merged_result)):
					#id column is the column that has the unique table ID
					merged_id = merged_result[row_merged][id_column]
					looping_id = list_looping[row_looping][id_column]

					#If the current id is not found in the list, append
					if looping_id not in id_list:
						merged_result.append(list_looping[row_looping] + empty_columns + [csv_file[:-4]])
						break

					#If the id in each of the two lists are the same, check each cell to see if anything has changed
					elif (merged_id == looping_id):
						for column in columns_without_id:
							if merged_result[row_merged][column] != list_looping[row_looping][column]:
								merged_result[row_merged][column] = csv_as_a_list[row_looping][column]
								merged_result[row_merged][column + len(columns_to_modify) + 1] = csv_file[:-4]
	save_list_as_csv(merged_result, results)f









merge_and_date_csv_files(csv_directory, merged, column_with_unique_id, columns_to_modify)




