# XLS-CSV Comparer

Program został stworzony by porównać plik CSV wygenerowany przez 
XXX i XLSX wygenerowny przez XXX. Głównym zadaniem programu jest znalezienie oraz 
dodanie do wynikowego pliku XLSX tych wierszy pochodzących z analizowanego pliku
XLSX, które nie znajdują się w pliku CSV. Program jest sworzony by uznawać plik CSV jako
szablon i nie da się zmienić kolejności bez zmiany całego kodu.

# Manual
Oba pliki ustawione są jako domyślne pliki i program będzie je wyszukiwał w katalogu
w którym się znajduje oraz w katalogach poniżej. Zalecam by program oraz oba pliki
znajdowały się w osobnym katalogu, ponieważ mimo, że program ominie pliki nie będące
tymi wygenrowanymi przez XXX, to nie jestem w stanie zapewnić by wziął odpowiedni
jednocześnie ściągając z użytkownika problem wpisywania trudnych nazw plików lub ich
sprawdzania. Plik sam znajdzie odpowiedni plik CSV i XLSX jeśli w tym katalogu (oraz
poniżej) znajduje się tylko jeden plik XLSX z odpowiednią nazwą. Jak na razie program 
jedyne co robi, to znajduje różnice oraz zapisuje je w nowym pliku typu XLSX, którego
nazwa jest zmodyfikowaną nazwą startowego pliku XLSX, program dodaje "-results" na
koniec nazwy. Program bez problemu nadpisze istniejący wynikowy plik XLSX, jeśli
nazwa się zgadza, ale w żadnym momencie wykonywania programu pliki dostarczone przez
użytkownika nie zostaną zmodyfkowane.

# Use case
Porgram został swtorzony z jedenego prostego powodu i do jednego celu, co powoduje jego
wyjątkowe wyspecjalizowanie. Mój bliski szczepi i sprawdzał jak oba pliki wyglądają
oraz jak mają się do siebie. Odkrył, że ilość rzędów w jednym pliku nie równa się
liczbie z drugiego pliku. Dokładnie rzecz ujmując wiecej było ich w pliku CSV. Z tego
powodu zostałem poproszony, by 
