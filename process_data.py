import csv
import os #Handle file names and paths

##################################################################
#User Defined Settings############################################
##################################################################
csv_directory = "Exports"
results = "results.csv"
template_file = "2016-07.csv" #Specify oldest file to act as the template for results.csv
columns_to_keep = [0,1,2]
column_with_unique_id = 0
qty_columns_to_keep = columns_to_keep[:]
qty_columns_to_keep.remove(column_with_unique_id)
qty_columns_to_keep = len(qty_columns_to_keep)


##################################################################
#Setup Base File, Merged File, and Directory of Exported Files####
##################################################################
def open_csv_as_list(file_dir):
	with open(file_dir) as csvfile:
		reader = csv.reader(csvfile,delimiter=",",quotechar='"')
		return list(reader)

base = open_csv_as_list(csv_directory + "/" + template_file)

def extract_columns(csv_file,columns_to_keep):
	#Remove columns that are not needed
	csv_extracted = []
	for row in range(0,len(csv_file)):
		extracted_row = []
		for column in columns_to_keep:
			extracted_row.append(csv_file[row][column])
		csv_extracted.append(extracted_row)
	return csv_extracted

base_extract = extract_columns(base,columns_to_keep)

#Created -Mod Date Columns
columns_without_id = [col for col in columns_to_keep if col != column_with_unique_id]
merged_columns = []
for column in columns_without_id:
	merged_columns.append(base[0][column] + "- Mod Date")

#Create Merged Document Headers
merged = [base_extract[0] + merged_columns]

##################################################################
#Loop through the remaining .csv files ###########################
##################################################################
def merge_and_date_csv_files(csv_directory, merged_result,id_column):
	list_of_csv_files = []
	for root, dirs, csv_files, in os.walk(csv_directory):
		for csv_file in csv_files:
			if csv_file.endswith('.csv'):
				list_of_csv_files.append(csv_file)

	for csv_file in list_of_csv_files:
		csv_as_a_list = open_csv_as_list(csv_directory + "/" + csv_file)
		list_looping = extract_columns(csv_as_a_list,columns_to_keep)
		#Append Rows to merged doc
		#If the merged csv only has the headers, add entries from the first csv file
		if len(merged_result) == 1:
			for row_looping in range(1,len(list_looping)):
				merged_result.append(list_looping[row_looping])
		else:
			for row_looping in range(1,len(list_looping)):
				for row_merged in range(1, len(merged_result)):
					#id column is the column that has the unique table ID
					merged_id = merged[row_merged][id_column]
					looping_id = list_looping[row_looping][id_column]
					#If the id in each of the two lists are the same, check each cell to see if anything has changed
					if (merged_id == looping_id):
						for cell in range(0, len(merged[row_merged])):
							if merged[row_merged][cell] != list_looping[row_looping][cell]:
								merged[row_merged][cell] = list_looping[row_looping][cell]
								merged[row_merged][cell + qty_columns_to_keep] = csv_file[:-5]


					#Else if the looping_id was not found in the merged csv, append it to the end
					elif ((merged_id != looping_id) and (row_looping == len(list_looping)+1)):
						list_looping.append(list_looping[row_looping])








merge_and_date_csv_files(csv_directory, merged, column_with_unique_id)


def save_list_as_csv(csv_list,file_name):
	with open(results, 'w', newline ='') as csvfile:
		writer = csv.writer(csvfile, delimiter = ',', quotechar ="'")
		for row in merged:
			writer.writerow(row)
		csvfile.close()

save_list_as_csv(merged, results)