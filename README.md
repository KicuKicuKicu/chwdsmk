# chwdsmk

Wymagania do odpalenia:

* Zainstalowanie Pythona - https://www.python.org/downloads/
* Umieszczenie w folderze pliku źródłowego
* Włączenie wiersza polecenia - z menu 'Start' albo za pomocą Win+R [Windows] lub aplikacji Teminal [Mac]
* wklepanie w tym okienku:

> cd (adres folderu)

np:
> cd C:\Users\user\Desktop\chwdsmk

włączenie Pythona
> python

gdy python się włączy:
> import procedura

Pierwsze odpalenie będzie trwało chwilę dłużej, pobrane zostaną rozszerzenia wymagane do uruchomienia programu

#################

Wymagania co do pliku źródłowego:
* Wymagane kolumny = 'data', 'rok', 'kod', 'operator', 'miejsce', 'staz', 'inicjaly', 'plec'
* Opcjonalne kolumny = 'asysta', 'grupa'
* data: format yyyy-mm-dd
* rok: numer roku specki
* kod: A lub B
* operator: imię i nazwisko
* miejsce: numer na liście
* staż: numer na liście
* płeć: M lub K

Instrukcja do odpalania:
#############

--> procedura.procedurki(file_name, file_type,  stop_index, start_index, start_delay, loop_delay)

     file name = nazwa pliku. Ujmij nazwę w ""
     
     file_type = csv albo xlsx (Excel). Ujmij rozszerzenie w ""
     
     stop_index = indeks ostatniego wiersza który ma zostać dodany; Domyślnie ostatni
     
     start_index = indeks pierwszeo wiersza który ma zostać dodany; Domyślnie zerowy
     
     start_delay = czas pomiędzy uruchomieniem skryptu a rozpoczęciem programu; Domyślnie 3[s]
     
     loop_delay = czas pomiędzy kolejnymi akcjami skryptu w sekundach; Domyślnie 0.1[s]
     
--> przykładowo: procedura.procedurki("ekg", "xlsx", 10, 0, 4, 0.5)

--> przykładowo: procedura.procedurki("usg", "csv", stop_index = 100, start_delay = 2, loop_delay = 0.5)

Obsługa:
* po odpaleniu skryptu, i zakończeniu odliczania najedź na przycisk "Dodaj", i na klawiaturze wciśnij Q - usłyszysz krótki beep
* następnie najedź na *wszystkie pola* w formularzu (pomiń ikonkę kalendarza) i również wiśnij Q - usłyszysz krótki beep po każdym wciśnięciu
* sit back and relax

Aby przerwać przytrzymaj Escape albo zamknij okno z programem
