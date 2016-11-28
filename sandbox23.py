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

#Created Merged Document Headers
merged = [base_extract[0] + merged_columns]

##################################################################
#Loop through the remaining .csv files ###########################
##################################################################
def merge_and_date_csv_files(csv_directory, merged_result):
	list_of_csv_files = []
	for root, dirs, csv_files, in os.walk(csv_directory):
		for csv_file in csv_files:
			if csv_file.endswith('.csv'):
				list_of_csv_files.append(csv_file)

	for csv_file in list_of_csv_files:
		csv_as_a_list = open_csv_as_list(csv_directory + "/" + csv_file)
		csv_as_a_list_extracted = extract_columns(csv_as_a_list,columns_to_keep)
		#Append Rows to merged doc
		for row in range(1,len(csv_as_a_list_extracted)):
			merged_result.append(csv_as_a_list_extracted[row])

merge_and_date_csv_files(csv_directory, merged)


def save_list_as_csv(csv_list,file_name):
	with open(results, 'w', newline ='') as csvfile:
		writer = csv.writer(csvfile, delimiter = ',', quotechar ="'")
		for row in merged:
			writer.writerow(row)
		csvfile.close()

save_list_as_csv(merged, results)