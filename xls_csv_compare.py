import csv
import openpyxl
from pathlib import Path as P

print("Made by PunkOkami", "Published under GNU GPLv3 licence",
		"GitHub repo: https://github.com/PunkOkami/xls_csv_comparer",
		"Version: 1.0", "email adress: okami.github@gmail.com", "\n", sep="\n")

# Program finds the cvs file and opens it then makes list of all rows
fcsv_path = list(P(P.cwd()).rglob("Raport*.csv"))[0]
fcsv = open(fcsv_path, newline="")
rcsv = csv.reader(fcsv, delimiter=";")
rows = [row for row in rcsv]
# The first row is taken, then 3 index values are extracted that let code later to put tuples made of entries from those
# columns for each row, also makes set of all PESEL values from cvs file
name_row = rows.pop(0)
names = ["pacjent_ext", "reg_aktywne", "pj_data_zapisania_baza_danych"]
id_col = name_row.index(names[0])
reg_col = name_row.index(names[1])
pj_col = name_row.index(names[2])
ids = set([row[id_col] for row in rows])
reg_rows = [(row[id_col], row[reg_col], row[pj_col]) for row in rows if row[reg_col] != ""]


# Function used later to filter the list of rows from xlsx file
def id_filter(row):
	if row[col_num].value in ids:
		return False
	else:
		return True


# Finds xlsx file end opens proper sheet
fxls_path = list(P(P.cwd()).rglob("eRej*[!-results].xlsx"))[0]
fxls = openpyxl.load_workbook(fxls_path)
sheet = fxls["Terminy i wizyty"]
# Grabs all rows from sheet where value of "Status wizyty" is "Zrealizowana", takes first row, finds index of column with
# PESEL values and based on those filters the list of rows to keep ones that are in cvs file
keep = [row for row in sheet.iter_rows() if row[3].value == "Zrealizowana"]
id_row = [[cell.value for cell in row] for row in sheet.iter_rows()][0]
col_num = id_row.index("Wartość identyfikatora pacjenta")
filtered_keep = list(filter(id_filter, keep))
filtered_keep = [[cell.value for cell in row] for row in filtered_keep] # Makes list of strings, then displays the result
print(f"Liczba rzędów z wartością 'Zrealizowana': {len(keep)}",
		f"Liczba rzędów, która różni się od pliku csv: {len(filtered_keep)}",
	  	f"Liczna rzędów z niepustym polem {names[1]}: {len(reg_rows)}", sep="\n")
# Makes new file, fills it with rows not in cvs file, then adds ones that have "reg_aktywne" not empty
result_xls_path = P(f"{fxls_path.stem}-results{fxls_path.suffix}")
result_xls = openpyxl.Workbook()
result_sheet = result_xls.active
result_sheet.append(id_row)
for row in filtered_keep:
	result_sheet.append(row)
result_sheet.append(["", "", ""])
result_sheet.append(names)
for row in reg_rows:
	result_sheet.append(row)
result_xls.save(result_xls_path)
inn = input("Naciśnij Enter by zakończyć")

# Made by PunkOkami
# Published under GNU GPLv3 licence
# GitHub repo: https://github.com/PunkOkami/xls_csv_comparer
# Version: 1.0
# e-mail adress: okami.github@gmail.com
