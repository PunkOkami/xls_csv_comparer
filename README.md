# XLS-CSV Comparer

Program został stworzony by porównać plik CSV wygenerowany przez 
XXX i XLS wygenerowny przez XXX. Głównym zadaniem programu jest znalezienie oraz 
dodanie do wynikowego pliku XLS tych wierszy pochodzących z analizowanego pliku
XLS, które nie znajdują się w pliku CSV. Program jest sworzony by uznawać plik CSV jako
szablon i nie da się zmienić kolejności bez zmiany całego kodu.

# Manual
Oba pliki ustawione jako domyślne pliki i program będzie je wyszukiwał w katalogu
w którym się znajduje oraz w katalogach poniżej. Zalecam by program oraz oba pliki
znajdowały się w osobnym katalogu, ponieważ mimo, że program ominie pliki nie będące
tymi wygenrowanymi przez XXX, to nie jestem w stanie zapewnić by wziął odpowiedni
jednocześnie ściągając z użytkownika problem wpisywania trudnych nazw plików lub ich
sprawdzania. Plik sam znajdzie odpowiedni plik CSV i XLS jeśli w tym katalogu (oraz
poniżej) znajduje się tylko jeden plik XLS z odpowiednią nazwą. 
