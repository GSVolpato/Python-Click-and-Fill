import pandas as pd
import pyautogui
import time
import threading
import keyboard
import os

# Get the current script's directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Load the Excel file into a DataFrame using the correct path
df = pd.read_excel(os.path.join(current_dir, 'mySpreadsheet.xlsx'))

# Global variable to control the pause state
pause_execution = False

def format_date(date):
    return str(date).replace('/', '-').replace(';', ':')

def save_mySpreadsheetindex(index):
    index_file = os.path.join(current_dir, 'mySpreadsheetindex.txt')
    with open(index_file, 'w') as file:
        file.write(str(index))

def load_mySpreadsheetindex():
    index_file = os.path.join(current_dir, 'mySpreadsheetindex.txt')
    try:
        with open(index_file, 'r') as file:
            return int(file.read())
    except FileNotFoundError:   
        return 0

def on_pause_press(e):
    global pause_execution
    if e.event_type == keyboard.KEY_DOWN:
        pause_execution = not pause_execution
        if pause_execution:
            print("\nPaused. Press 'Pause Break' again to resume.")
        else:
            print("\nResumed.")

def check_pause():
    while pause_execution:
        time.sleep(0.1)

def click_at_positions():
    global pause_execution
    mySpreadsheetindex = load_mySpreadsheetindex()
    print("Starting from row ", mySpreadsheetindex)
    print("Total rows: ", len(df))
    print("Index: ", df.iloc[mySpreadsheetindex]['Index'])
    input("Press Enter to continue...")
    time.sleep(0.25)
    pyautogui.hotkey('alt', 'tab')
    time.sleep(0.25)

    # Set up the keyboard hook for the Pause Break key
    keyboard.on_press_key('pause', on_pause_press)

    idx = mySpreadsheetindex
    
    for _, row in df.iloc[mySpreadsheetindex:].iterrows():
        print("_____________________________________________________")
        
        check_pause()
        pyautogui.doubleClick((573, 235))
        
        check_pause()
        print("Index: ", row['Index'])
        pyautogui.write(str(row['Index']), interval=0.01)
        
        check_pause()
        pyautogui.press('enter')
        
        check_pause()
        print("Placa: ", row['Placa'])
        pyautogui.write(str(row['Placa']), interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        time.sleep(0.25)
        
        check_pause()
        pyautogui.press(['space', 'backspace'])
        
        check_pause()
        print("Km: ", row['Km'])
        pyautogui.write(str(int(row['Km'])), interval=0.15)
        #pyautogui.write(str(row['Km']).replace(',', '.'), interval=0.15)
        
        check_pause()
        pyautogui.press('tab')
        
        check_pause()
        print("NCOMB: ", row['NCOMB'])
        pyautogui.write(str(row['NCOMB']).zfill(2), interval=0.01)
        
        check_pause()
        pyautogui.press(['tab', 'tab'])
        
        check_pause()
        print("Qtde (L): ", row['Qtde (L)'])
        pyautogui.write(str(row['Qtde (L)']), interval=0.01)
        
        check_pause()
        pyautogui.press(['tab', 'tab', 'tab', 'tab', 'tab', 'tab'])
        
        check_pause()
        print("Registro: ", row['Registro'])
        pyautogui.write(str(int(row['Registro'])).zfill(6), interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        time.sleep(0.25)
        
        check_pause()
        pyautogui.press(['space','space'])
        
        check_pause()
        print("Fornecedor: ", row['FORNECEDOR'])
        pyautogui.write(str(row['FORNECEDOR']))
        
        check_pause()
        pyautogui.press('tab')
        
        check_pause()
        print("Autorizador: ", row['AUTORIZADOR'])
        pyautogui.write(str(row['AUTORIZADOR']), interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write('     ', interval=0.01)
        check_pause()
        print("Data do abastecimento: ", row['Data/Hora'])
        pyautogui.write('Data do abastecimento: ' + format_date(row['Data/Hora']))
        
        check_pause()
        pyautogui.press(['tab', 'tab'])
        
        check_pause()
        print("Data/Hora: ", row['Data/Hora'])
        pyautogui.write(str(row['Data/Hora']), interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        
        check_pause()
        print("Preco Unitário: ", row['Preco Unitário'])
        pyautogui.write(str(row['Preco Unitário']), interval=0.01)
        
        check_pause()
        pyautogui.press(['tab', 'tab'])
        
        check_pause()
        print("33")
        pyautogui.write('33 - Cupom Fiscal', interval=0.01)
        
        check_pause()
        pyautogui.press(['down'])
        pyautogui.press('tab')
        
        check_pause()
        print("NF: ", row['NF'])
        pyautogui.write(str(row['NF']), interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        pyautogui.write('1', interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        
        check_pause()
        print("DATANF: ", row['DATANF'])
        pyautogui.write('{:0>8}'.format(row['DATANF']), interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        
        check_pause()
        print("ANOEMPENHO: ", row['ANOEMPENHO'])
        pyautogui.write(str(row['ANOEMPENHO']), interval=0.01)
        
        check_pause()
        pyautogui.press('tab')
        
        check_pause()
        print("EMPENHO: ", row['EMPENHO'])
        pyautogui.write(str(row['EMPENHO']), interval=0.01)
        
        check_pause()
        time.sleep(0.25)
        pyautogui.click((212, 155))
        pyautogui.click((212, 155))
        
        check_pause()
        time.sleep(0.25)
        save_mySpreadsheetindex(idx + 1)
        idx += 1
        
        check_pause()
        time.sleep(0.25)
        pyautogui.press(['space', 'space'])
        time.sleep(1)

if __name__ == "__main__":
    click_at_positions()
    pyautogui.hotkey('alt', 'tab')
    input("Press Enter to close...")
