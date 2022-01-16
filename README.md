# XLS-CSV Comparer

Program został stworzony, by porównać plik CSV wygenerowany przez 
gabinet.gov.pl i plik XLSX wygenerowny przez eRejestracja. Głównym zadaniem programu 
jest znalezienie oraz dodanie do wynikowego pliku XLSX tych wierszy pochodzących z 
analizowanego pliku XLSX, które nie znajdują się w pliku CSV. Program jest sworzony, 
by uznawać plik XLSX jako szablon. Nie ma możliwości zamiany plików bez zmiany całego kodu.

## Nota
W czasie pracy tego programu żadne dane wrażliwe nie opuszczają folderu, w którym się
znajdują pliki wejściowe. W czasie pracy nad tym programem korzystałem jedynie z 
danych losowych oraz sztucznie stworzonych plików.

## Use case
Porgram został stworzony z jedenego prostego powodu i w jednym celu, co powoduje jego
wyjątkowe wyspecjalizowanie. Mój bliski krewny wykonuje szczepienia i sprawdzał, jak oba pliki wyglądają
oraz czy ich zawartość się pokrywa. Odkrył, że ilość rzędów w jednym pliku nie równa się
liczbie rzędów z drugiego pliku. Dokładnie rzecz ujmując wiecej było ich w pliku CSV. Drugą 
rzeczą, którą program robi jest znajdywanie tych rzędów w pliku csv, gdzie kolumna "reg_aktywna"
zawiera coś innego niż pustą wartość. 

## Manual
Program nie wymaga podawania nazw żadnego z plików. Wyszuka je w katalogu, gdzie się
znajdują pliki CSV i XLSX. Może sięgnąć nawet do katalogów poniżej, ale nigdy powyżej swojej
lokalizacji. Weźmie tylko pliki, których nazwa zgadza się z formatem stosowanymi przez eRejestracja
oraz gabinet.gov.pl. Od wersji 1.2 program jest w stanie poradzić sobie z wiecej niż jednym plikiem
XLS i CSV w katalogu (w takim wypadku użytkownik może wybrać plik do analizy, wybrać najnowszy lub 
opuścić program) a nawet powiadomi, gdy nie znajdzie żadnego pliku. Na pewno nie zinterpretuje on 
pliku wynikowego XLSX jako wejściowego, więc nie trzeba usuwać go między sesjami korzystania z 
programu. Program znajduje różnice między danymi w plikach wejściowych oraz zapisuje je w nowym 
pliku typu XLSX, którego nazwa jest zmodyfikowaną nazwą startowego pliku XLSX. Program dodaje 
"-results" na koniec nazwy. Program bez problemu nadpisze istniejący wynikowy plik XLSX, jeśli 
nazwa się zgadza, ale w żadnym momencie użytkowania programu pliki dostarczone przez użytkownika 
nie zostaną zmodyfkowane.

### Instrukcja postępowania:
1. Rozpakuj pobrany plik "xls-csv-comparer.zip" - kliknij prawym klawiszem myszki na ikonę  i wybierz opcję "wyodrębnij 
wszystkie". W otwartym oknie możesz zmienić lokalizację miejsca, gdzie ma znaleźć się program po rozpakowaniu. Program 
nie będzie prawidłowo pracował, jeśli chcesz go uruchomić klikając w oknie eksploratora pliku spakowanego!
2. Po rozpakowaniu powstanie katalog o nazwie "xls-csv-comparer". W tym katalogu znajduje się podkatalog 
"Work_space", w któym jest właściwy program.
3. Wytwórz pliki CSV oraz XLSX i zapisz je w podkatalogu "Work_space".
4. Uruchom program xls_csv_compare.exe. 
5. Postepuj zgodnie z instrukcjami widocznymi  konsoli programu.
6. Po zakończeniu pracy programu w folderze powinien pojawić się nowy plik XLSX z dodatkiem "results" w 
nazwie. Plik należy otworzyć i przeanalizować. W górnej części tabeli są pozycje widoczne w rejestracji, 
a nie widoczne w gabinecie, poniżej mogą być dodatkowo pozycje, które w raporcie z e-gabinetu miały pewien
komentarz, też warto je sprawdzić, czy zostały prawidłowo wprowadzone w gabinet.gov.pl.
7. Gdyby był jakiekolwiek problemy z uruchomieniem programu, spróbuj uruchomić go w trybie administratora:
prawy klawisz myszki na ikonie programu i opcja "uruchom jako administrator".

W razie problemów lub błędów programu moje dane kontaktowe znajdują sie w komunikatach
programu oraz na końcu tego pliku.

Made by PunkOkami

Published under GNU GPLv3 licence

[GitHub repo](https://github.com/PunkOkami/xls_csv_comparer)

[Link do pobrania pliku exe](https://github.com/PunkOkami/xls_csv_comparer/releases)

Version: 1.2

e-mail adress: okami.github@gmail.com
