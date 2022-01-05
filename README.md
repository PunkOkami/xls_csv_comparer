# XLS-CSV Comparer

Program został stworzony by porównać plik CSV wygenerowany przez 
XXX i XLSX wygenerowny przez XXX. Głównym zadaniem programu jest znalezienie oraz 
dodanie do wynikowego pliku XLSX tych wierszy pochodzących z analizowanego pliku
XLSX, które nie znajdują się w pliku CSV. Program jest sworzony by uznawać plik XLSX
jako szablon i nie da się zmienić kolejności bez zmiany całego kodu.

## Nota
W czasie pracy tego programu żadne dane wrażliwe nie opuszczają foderu w którym się
znajdują pliki wejściowe. W czasie pracy nad tym programem korzystałem jedynie z 
danych losowych oraz sztucznie stworzonych plików.

## Use case
Porgram został swtorzony z jedenego prostego powodu i do jednego celu, co powoduje jego
wyjątkowe wyspecjalizowanie. Mój bliski szczepi i sprawdzał jak oba pliki wyglądają
oraz jak mają się do siebie. Odkrył, że ilość rzędów w jednym pliku nie równa się
liczbie z drugiego pliku. Dokładnie rzecz ujmując wiecej było ich w pliku CSV. Drugą 
rzeczą, która program robi to znajduje te rzędy w pliku csv gdzie koluman "reg_aktywna"
zawiera coś innego niż pustą wartość. 

## Manual
Oba pliki ustawione są jako domyślne pliki i program będzie je wyszukiwał w katalogu,
w którym się znajduje oraz w katalogach poniżej. Zalecam, by program oraz oba pliki
znajdowały się w osobnym katalogu, ponieważ mimo, że program ominie pliki nie będące
tymi wygenrowanymi przez XXX, to nie jestem w stanie zapewnić by wziął odpowiedni
jednocześnie ściągając z użytkownika problem wpisywania trudnych nazw plików lub ich
sprawdzania. Plik sam znajdzie odpowiedni plik CSV i XLSX jeśli w tym katalogu (oraz
poniżej) znajduje się tylko jeden plik XLSX oraz jeden plik CSV z odpowiednią nazwą. 
Jak na razie program jedyne co robi, to znajduje różnice oraz zapisuje je w nowym 
pliku typu XLSX, którego nazwa jest zmodyfikowaną nazwą startowego pliku XLSX, 
program dodaje "-results" na koniec nazwy. Program bez problemu nadpisze istniejący 
wynikowy plik XLSX, jeśli nazwa się zgadza, ale w żadnym momencie wykonywania programu 
pliki dostarczone przez użytkownika nie zostaną zmodyfkowane.

### Instrukcja postępowania:
1. Wytwórz plik CSV oraz XLSX
2. Umieść oba pliki w katologu z plikiem wykonywalnym
3. Uruchom program
4. Postepuj zgodnie z instrukcjami w konsoli
5. Po zakończeniu pracy programu w folderze powinien pojawić się nowy plik XLSX

W razie problemów lub błędów programu moje dane kontaktowe znajdują sie w komukatach
programu oraz na końcu tego pliku.

Made by PunkOkami

Published under GNU GPLv3 licence

[GitHub repo](https://github.com/PunkOkami/xls_csv_comparer)

Version: 1.0

e-mail adress: okami.github@gmail.com
