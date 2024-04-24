import script_setup
import pandas as pd
import pyautogui
import time
import keyboard
import pyperclip
import winsound
from typing import Union



print('\nWymagania co do pliku źródłowego:')
print("> Wymagane kolumny = 'data', 'rok', 'kod', 'operator', 'miejsce', 'staz', 'inicjaly', 'plec'")
print("> Opcjonalne kolumny = 'asysta', 'grupa'")
print("data: format yyyy-mm-dd\nrok: numer roku specki\nkod: A lub B\noperator: imię i nazwisko")
print("miejsce: numer na liście\nstaż: numer na liście\npłeć: M lub K")
print('\nInstrukcja do odpalania:\n#############')
print('--> procedurki(file_name, file_type,  stop_index, start_index, start_delay, loop_delay)')
print('     file name = nazwa pliku. Ujmij nazwę w \"\"')
print('     file_type = csv albo xlsx (Excel). Ujmij rozszerzenie w \"\"')
print('     stop_index = indeks ostatniego wiersza który ma zostać dodany; Domyślnie ostatni')
print('     start_index = indeks pierwszeo wiersza który ma zostać dodany; Domyślnie zerowy')
print('     start_delay = czas pomiędzy uruchomieniem skryptu a rozpoczęciem programu; Domyślnie 3[s]')
print('     start_delay = czas pomiędzy kolejnymi akcjami skryptu w sekundach; Domyślnie 0.1[s]')
print('--> przykładowo: procedura.procedurki("ekg", "xlsx", 10, 0, 4, 0.5)')
print('\nObsługa:')
print('     po odpaleniu skryptu, i zakończeniu odliczania najedź na przycisk "Dodaj", i na klawiaturze wciśnij Q')
print('     następnie najedź na *wszystkie pola* w formularzu (pomiń ikonkę kalendarza) i również wiśnij Q')
print('     sit back and relax')
print(' Aby przerwać przytrzymaj Escape albo zamknij okno z programem')



def procedurki(file_name: str, file_type: str,  stop_index: int = None, start_index: int = 0, start_delay: Union[int, float] = 3, loop_delay: Union[int, float] = 0.1):
    if file_type not in ['csv', 'xlsx']:
        raise Exception("Invalid type - choose from \'csv\' and \'xlsx\'")
    elif file_type == 'csv':
        df_full = pd.read_csv(file_name + '.csv')
    else:
        df_full = pd.read_excel(file_name + '.xlsx')

    required_columns = ['data', 'rok', 'kod', 'operator', 'miejsce', 'staz', 'inicjaly', 'plec']
    missing_columns = [col for col in required_columns if col not in df_full.columns]
    if missing_columns:
        raise ValueError(f"Brakujące kolumny: {', '.join(missing_columns)}")
    else:
        print("Niezbędne kolumny obecne")

    df_full['data'] = pd.to_datetime(df_full['data'])

    if stop_index is None:
        stop_index = len(df_full)

    if start_index > len(df_full):
        raise ValueError(f"Indeks poczatkowy wiekszy niż ilość danych ({start_index}>{len(df_full)})")
    elif stop_index > len(df_full):
        raise ValueError(f"Indeks końcowy wiekszy niż ilość danych ({stop_index}>{len(df_full)})")
    elif stop_index < start_index:
        raise ValueError(f"Indeks końcowy ({stop_index}) mniejszy niż początkowy ({start_index})")
    df_full = df_full.loc[start_index:stop_index]
    print(f'Ilość procedur: {len(df_full)}')

    if start_delay > 0:
        for n in range(start_delay-1):
            winsound.Beep(1000, 100)
            time.sleep(0.9)
            if keyboard.is_pressed('esc'):  # Check if Escape key is pressed
                winsound.Beep(3000, 50)
                winsound.Beep(3000, 50)
                winsound.Beep(3000, 50)
                print('Zatrzymano program')
                exit()
        winsound.Beep(1000, 1000)

    if keyboard.is_pressed('esc'):  # Check if Escape key is pressed
        print('Zatrzymano adresowanie pól')
        winsound.Beep(3000, 50)
        winsound.Beep(3000, 50)
        winsound.Beep(3000, 50)
        exit()
    else:
        keyboard.wait('q')
    dodaj_pos = pyautogui.position()
    pola_pos = []
    winsound.Beep(500, 50)
    pyautogui.click()
    print("Pozycja przycisku ", dodaj_pos)

    for _ in range(10):
        if keyboard.is_pressed('esc'):  # Check if Escape key is pressed
            print('Zatrzymano adresowanie pól. Wyłączono.')
            winsound.Beep(3000, 50)
            winsound.Beep(3000, 50)
            winsound.Beep(3000, 50)
            exit()
        keyboard.wait('q')
        pola_pos.append(pyautogui.position())
        winsound.Beep(500, 50)

    wiersz = 0

    for n in range(3):
        winsound.Beep(1000, 100)
        time.sleep(0.9)
        if keyboard.is_pressed('esc'):  # Check if Escape key is pressed
            winsound.Beep(3000, 50)
            winsound.Beep(3000, 50)
            winsound.Beep(3000, 50)
            print('Zatrzymano program')
            exit()
    winsound.Beep(1000, 1000)

    print('Pola:', pola_pos)
    loop_start_time = time.time()
    for index, row in df_full.iterrows():  # Iterate over each recorded position
        for field in pola_pos:
            if keyboard.is_pressed('esc'):  # Check if Escape key is pressed
                print(f'Przerwano wykonywanie programu. '
                      f'Dodano {wiersz} procedur w czasie {time.time()-loop_start_time}s')
                winsound.Beep(3000, 50)
                winsound.Beep(3000, 50)
                winsound.Beep(3000, 50)
                exit()

            pyautogui.click(field)
            if field == pola_pos[0]:
                pyperclip.copy(str(row['data']))
                pyautogui.hotkey('ctrl', 'v')
            elif field == pola_pos[1]:
                pyautogui.typewrite(str(row['rok']))
                pyautogui.press('enter')
            elif field == pola_pos[2]:
                pyautogui.typewrite(str(row['kod']))
                pyautogui.press('enter')
            elif field == pola_pos[3]:
                pyperclip.copy(str(row['operator']))
                pyautogui.hotkey('ctrl', 'v')
            elif field == pola_pos[4]:
                kroki = int(row['miejsce'])
                for _ in range(kroki):
                    pyautogui.press('down')
                pyautogui.press('enter')
            elif field == pola_pos[5]:
                kroki = int(row['staz'])
                for _ in range(kroki):
                    pyautogui.press('down')
                pyautogui.press('enter')
            elif field == pola_pos[6]:
                pyperclip.copy(str(row['inicjaly']))
                pyautogui.hotkey('ctrl', 'v')
            elif field == pola_pos[7]:
                pyautogui.typewrite(row['plec'])
                time.sleep(0.1)
                pyautogui.press('enter')
            elif field == pola_pos[8] and 'asysta' in df_full.columns:
                if pd.notna(row['asysta']) and row['asysta'].strip():
                    pyperclip.copy(str(row['asysta']))
                    pyautogui.hotkey('ctrl', 'v')
            elif field == pola_pos[9] and 'grupa' in df_full.columns:
                if pd.notna(row['grupa']) and row['grupa'].strip():
                    pyperclip.copy(str(row['grupa']))
                    pyautogui.hotkey('ctrl', 'v')
        time.sleep(loop_delay)
        wiersz += 1
        pyautogui.click(dodaj_pos)
    print(f'Dodano {wiersz} procedur w czasie {time.time()-loop_start_time}s')
    winsound.Beep(3000, 50)
    winsound.Beep(3000, 50)
    winsound.Beep(3000, 50)
