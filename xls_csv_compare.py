import datetime
import sys
import csv
import openpyxl
from pathlib import Path as P


# Funkcja wykorzystana do przefiltrowania rzędów z pliku xlsx
def id_filter(row):
	if row[col_num].value in ids:
		return False
	else:
		return True


print("Made by PunkOkami", "Published under GNU GPLv3 licence",
		"Kod żródłowy: https://github.com/PunkOkami/xls_csv_comparer",
		"Version: 1.2", "email adress: okami.github@gmail.com", "\n", sep="\n")
print("---------------------------------------------------")

# Program znajduje wszystkie pliki CSV po czym sprawdza ich ilość i zależnie od ilości albo kontynuje bez problemów,
# powidamia o braku plików albo pyta się co zrobić - gdy napotka kilka
fcsv_list = list(P(P.cwd()).rglob("Raport*.csv"))
if len(fcsv_list) != 1:
	if len(fcsv_list) == 0:
		inn = input("Program nie znalazł żadnego pliku CSV, naciśnij ENTER by zamknąć program")
		sys.exit()
	else:
		print("Program znalazł więcej niż jeden plik CSV. Co program ma zrobić? Wpisz numer opcji, którą wybierasz",
			  "1.Zamknij program, sam uporządkuję zawartość katalogu", "2.Weź najnowszy plik do analizy",
			  "3.Pokaż mi listę plików do wyboru", sep="\n")
		option = input(">>>")
		if int(option) == 1:
			sys.exit()
		elif int(option) == 2:
			mtimes = {file.stat().st_mtime:file for file in fcsv_list}
			newest = max(mtimes.keys())
			fcsv_path = mtimes[newest]
		elif int(option) == 3:
			print("Pliki CSV znajdujące się w katalogu")
			for file in fcsv_list:
				mtime = file.stat().st_mtime
				print(f"{fcsv_list.index(file) + 1}. {file.name} --- {datetime.datetime.fromtimestamp(mtime)}")
			fnum = int(input("Wpisz numer pliku CSV jaki program ma wybrać\n>>>"))
			fcsv_path = fcsv_list[fnum - 1]
		else:
			inn = input("Opcja niepoprawna, naciśnij ENTER by zamknąć program")
			sys.exit()
else:
	fcsv_path = fcsv_list[0]
print("----------------------------------------------")
print(f"Plik CSV wzięty do analizy to {fcsv_path.name}")
print("----------------------------------------------")

# Program wykonuje dokładnie ten sam zestaw testów i zapytań co przy pliku CSV, tylko, że dla plików XLS
fxls_list = list(P(P.cwd()).rglob("eRej*[!-results].xlsx"))
if len(fxls_list) != 1:
	if len(fxls_list) == 0:
		inn = input("Program nie znalazł żadnego pliku XLS, naciśnij ENTER by zamknąć program")
		sys.exit()
	else:
		print("Program znalazł więcej niż jeden plik XLS. Co program ma zrobić? Wpisz numer opcji, którą wybierasz",
			  "1.Zamknij program, sam uporządkuję zawartość katalogu", "2.Weź najnowszy plik do analizy",
			  "3.Pokaż mi listę plików do wyboru", sep="\n")
		option = input(">>>")
		if int(option) == 1:
			sys.exit()
		elif int(option) == 2:
			mtimes = {file.stat().st_mtime:file for file in fxls_list}
			newest = max(mtimes.keys())
			fxls_path = mtimes[newest]
		elif int(option) == 3:
			print("Pliki CSV znajdujące się w katalogu")
			for file in fxls_list:
				mtime = file.stat().st_mtime
				print(f"{fxls_list.index(file) + 1}. {file.name} --- {datetime.datetime.fromtimestamp(mtime)}")
			fnum = int(input("Wpisz numer pliku CSV jaki program ma wybrać\n>>>"))
			fxls_path = fxls_list[fnum - 1]
		else:
			inn = input("Opcja niepoprawna, naciśnij ENTER by zamknąć program")
			sys.exit()
else:
	fxls_path = fxls_list[0]
print("----------------------------------------------")
print(f"Plik XLS wzięty do analizy to {fxls_path.name}")
print("----------------------------------------------")

# Program otwiera wybrany plik csv i wyciąga z niego dane
fcsv = open(fcsv_path, newline="")
rcsv = csv.reader(fcsv, delimiter=";")
rows = [row for row in rcsv]
# Wyciągany jest pierwszy rząd, następnie pobierane są 3 wartości indeksów z 3 różnych kolumn, by potem stworzyć listę tupli
# z wartości kolumn z każdego rzędu. Powstaje też zbiór numerów PESEL z pliku csv
name_row = rows.pop(0)
names = ["pacjent_ext", "reg_aktywne", "pj_data_zapisania_baza_danych"]
id_col = name_row.index(names[0])
reg_col = name_row.index(names[1])
pj_col = name_row.index(names[2])
ids = set([row[id_col] for row in rows])
reg_rows = [(row[id_col], row[reg_col], row[pj_col]) for row in rows if row[reg_col] != ""]

# Otwiera odpowiedni arkusz w pliku XLS
fxls = openpyxl.load_workbook(fxls_path)
sheet = fxls["Terminy i wizyty"]
# Bierze wszystkie rzędy z arkusza, w krórych kolumna "Status wizyty" ma wartość "Zrealizowana",
# bierze pierwszy rząd, znajduje indeksy  kolumn z numerami PESEL i bazując na tym tworzy listę rzędów nie będących w
# pliku csv
keep = [row for row in sheet.iter_rows() if row[3].value == "Zrealizowana"]
id_row = [[cell.value for cell in row] for row in sheet.iter_rows()][0]
col_num = id_row.index("Wartość identyfikatora pacjenta")
filtered_keep = list(filter(id_filter, keep))
filtered_keep = [[cell.value for cell in row] for row in filtered_keep] # Tworzy listę stringów oraz wyświetla wynik
print(f"Liczba rzędów z wartością 'Zrealizowana': {len(keep)}",
		f"Liczba rzędów, która różni się od pliku csv: {len(filtered_keep)}",
	  	f"Liczna rzędów z niepustym polem {names[1]}: {len(reg_rows)}", sep="\n")
# Tworzy nowy plik, wypełnia go rzędami nie będącymi w pliku csv, oraz dodaje te, gdzie komlumna "reg_aktywne" nie
# jest pusta
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
print("----------------------------------------------")
inn = input("Naciśnij Enter by zakończyć program")

# Made by PunkOkami
# Published under GNU GPLv3 licence
# Kod żródłowy: https://github.com/PunkOkami/xls_csv_comparer
# Version: 1.2
# e-mail adress: okami.github@gmail.com
