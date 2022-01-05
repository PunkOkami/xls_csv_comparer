import csv
import openpyxl
from pathlib import Path as P

print("Made by PunkOkami", "Published under GNU GPLv3 licence",
		"GitHub repo: https://github.com/PunkOkami/xls_csv_comparer",
		"Version: 1.0", "email adress: okami.github@gmail.com", "\n", sep="\n")

fcsv_path = list(P(P.cwd()).rglob("Raport*.csv"))[0]
fcsv = open(fcsv_path, newline="")
rcsv = csv.reader(fcsv, delimiter=";")
rows = [row for row in rcsv]
name_row = rows.pop(0)
id_col = name_row.index("pacjent_ext")
reg_col = name_row.index("reg_aktywne")
ids = set([row[id_col] for row in rows])
reg_rows = [row for row in rows if row[reg_col] != ""]


def id_filter(row):
	if row[col_num].value in ids:
		return False
	else:
		return True


fxls_path = list(P(P.cwd()).rglob("eRej*[!-results].xlsx"))[0]
fxls = openpyxl.load_workbook(fxls_path)
sheet = fxls["Terminy i wizyty"]
keep = [row for row in sheet.iter_rows() if row[3].value == "Zrealizowana"]
id_row = [[cell.value for cell in row] for row in sheet.iter_rows()][0]
col_num = id_row.index("Wartość identyfikatora pacjenta")
filtered_keep = list(filter(id_filter, keep))
filtered_keep = [[cell.value for cell in row] for row in filtered_keep]
print(f"Liczba rzędów z wartością 'Zrealizowana': {len(keep)}",
		f"Liczba rzędów, która różni się od pliku csv: {len(filtered_keep)}", sep="\n")
inn = input("Press ENTER to continue\n")
result_xls_path = P(fxls_path.stem + "-results" + fxls_path.suffix)
result_xls = openpyxl.Workbook()
result_sheet = result_xls.active
result_sheet.append(id_row)
for row in filtered_keep:
	result_sheet.append(row)
result_xls.save(result_xls_path)

# Made by PunkOkami
# Published under GNU GPLv3 licence
# GitHub repo: https://github.com/PunkOkami/xls_csv_comparer
# Version: 1.0
# e-mail adress: okami.github@gmail.com
