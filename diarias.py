import pandas as pd
import pyautogui
import time


def confirm():
    pyautogui.hotkey('alt', 'tab')
    input("Press Enter to continue...")
    time.sleep(0.5)
    pyautogui.hotkey('alt', 'tab')

# Load the Excel file into a DataFrame
df = pd.read_excel('imp.xlsx', dtype=str)  # Ensure all data is read as strings

# Ensure specific columns are formatted correctly
df['Partida'] = df['Partida'].apply(lambda x: str(x).zfill(8))
df['Hora Partida'] = df['Hora Partida'].apply(lambda x: str(x).zfill(4))
df['Retorno'] = df['Retorno'].apply(lambda x: str(x).zfill(8))
df['Hora Retorno'] = df['Hora Retorno'].apply(lambda x: str(x).zfill(4))

def format_date(date):
    return str(date).replace('/', '-')  # Replace '/' with '-' in the date string

# Function to save and load the last processed index
def save_impIndex(index):
    with open('impIndex.txt', 'w') as file:
        file.write(str(index))

def load_impIndex():
    try:
        with open('impIndex.txt', 'r') as file:
            return int(file.read())
    except FileNotFoundError:
        return 0

def click_at_positions():
    impIndex = load_impIndex()
    print("Starting from index:", impIndex)
    input("Press Enter to continue...")
    time.sleep(0.5)
    print("Pressing Alt+Tab...")
    pyautogui.hotkey('alt', 'tab')
    time.sleep(0.5) 
    
    # Counter to track the index
    idx = impIndex
    
    # Iterate over the rows of the DataFrame
    for _, row in df.iloc[impIndex:].iterrows():
        pyautogui.click((962, 290)) # Nova Diaria
        pyautogui.click((153, 168)) # Nome
        pyautogui.write(str(row['Nome']), interval=0.01)
        print("Nome:", str(row['Nome']))
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.write('Giovane')
        time.sleep(0.5)
        pyautogui.press('tab')
        pyautogui.write(str(row['Destino']), interval=0.01)
        print("Destino: ", str(row['Destino']))
        pyautogui.press('tab')
        pyautogui.write('S') # Tp Hospedagem: Sem Hospedagem
        pyautogui.press(['tab', 'tab'])
        pyautogui.write(row['Partida'], interval=0.01)
        print("Partida: ", row['Partida'])
        pyautogui.press('tab')
        pyautogui.write(row['Hora Partida'], interval=0.01)
        print("Hora Partida: ", row['Hora Partida'])
        pyautogui.press('tab')
        pyautogui.write(row['Retorno'], interval=0.01)
        print("Retorno: ", row['Retorno'])
        pyautogui.press('tab')
        pyautogui.write(row['Hora Retorno'], interval=0.01)
        print("Hora Retorno: ", row['Hora Retorno'])
        pyautogui.press('tab')

        #pyautogui.write(str(row['Transporte']), interval=0.01) # Meio de Transporte
        #print("Transporte: ", str(row['Transporte']))

        pyautogui.write('V') # Veiculo Oficial
        pyautogui.press('tab')
        pyautogui.write(str(row['Veiculo']), interval=0.01)
        print("Veiculo: ", str(row['Veiculo']))
        pyautogui.press(['tab', 'tab'])
        pyautogui.write('S') # Tanque Cheio
        pyautogui.press(['tab', 'tab', 'tab'])

        #pyautogui.write(str(row['Passagem']), interval=0.01)
        #print("Passagem: ", str(row['Passagem']))

        pyautogui.write('0') # Passagem
        pyautogui.press(['tab', 'tab', 'tab', 'tab'])
        
        pyautogui.write(str(row['Pedagio']).replace('.', ','), interval=0.01)
        print("Pedagio: ", str(row['Pedagio']))
        pyautogui.press(['tab', 'tab', 'tab'])

        pyautogui.write('V') # Alimentação Inteira -> Valor Alimentação
        pyautogui.press('tab')

        if row['MeiaAlim'] == 1:
            pyautogui.write('V', interval=0.01) # Valor 1/2 Alimentação
            print("MeiaAlim: ", str(row['MeiaAlim']))
        else:
            pyautogui.write(str(int(row['MeiaAlim'])), interval=0.01)
            print("MeiaAlim: ", str(int(row['MeiaAlim'])))

        pyautogui.press('tab')

        if row['MeiaHosp'] == 1:
            pyautogui.write('V', interval=0.01) # Valor 1/2 Hospedagem
            print("Meia Hospedagem: ", str(row['MeiaHosp']))
        else:
            pyautogui.write(str(int(row['MeiaHosp'])), interval=0.01)
            print("Meia Hospedagem: ", str(int(row['MeiaHosp'])))

        pyautogui.press(['tab', 'tab'])
        pyautogui.write(str(int(row['Alim'])), interval=0.01)
        print("Alimentacao: ", str(row['Alim']))
        pyautogui.press('tab')
        pyautogui.write(str(int(row['Hosp'])), interval=0.01)
        print("Hospedagem: ", str(row['Hosp']))
        pyautogui.press('tab')
        pyautogui.write(str(row['Finalidade']), interval=0.01)
        print("Finalidade: ", str(row['Finalidade']))
        pyautogui.press('tab')
        
        obs_value = row['Obs']
        # Check if the value is not NaN and not an empty string
        if pd.notna(obs_value) and obs_value != '':
            pyautogui.write(str(obs_value))

        time.sleep(0.5)
        pyautogui.click((982, 241)) # SALVAR
        time.sleep(0.5)
        pyautogui.click((969, 424))  # FSD
        time.sleep(1)
        pyautogui.click((982, 118))  # PDF
        time.sleep(1)
        pyautogui.write(str(row['NomeFSD']), interval=0.01)
        print("NomeFSD: ", str(row['NomeFSD']))
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('esc')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click((1152, 116))  # FECHAR VISUALIZAÇÃO
        time.sleep(0.5)
        pyautogui.click((971, 477))  # RELATORIO
        time.sleep(0.5)
        pyautogui.click((442, 467))  # DESCRIÇAO
        pyautogui.write(str(row['Finalidade']), interval=0.01)
        print("Finalidade: ", str(row['Finalidade']))
        pyautogui.click((952, 475))  # DOCUMENTOS
        #pyautogui.write(str(row['DOCS RV']), interval=0.01)
        #print("DOCS RV: ", str(row['DOCS RV']))
        pyautogui.write('Anexados ao relatorio.')
        pyautogui.click((390, 566)) # OUTRAS CONSIDERAÇOES
        #pyautogui.write(str(row['Obs']), interval=0.01)
        
        obs_value = row['Obs']
        if obs_value == 'nan':
            pyautogui.write('')

        pyautogui.click((1047, 262))  # SALVAR RV
        time.sleep(0.5)
        pyautogui.click((1053, 638))  # IMPRIMIR RV
        time.sleep(0.5)
        pyautogui.click((1102, 48))  # FECHAR RV
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.click((982, 118))  # PDF
        time.sleep(0.5)
        pyautogui.write(str(row['NomeRV']), interval=0.01)
        print("NomeRV: ", str(row['NomeRV']))
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('esc')
        time.sleep(0.5)
        pyautogui.click((1152, 116))  # FECHAR VISUALIZAÇÃO
        pyautogui.click((1152, 116))  # FECHAR VISUALIZAÇÃO
        save_impIndex(idx + 1)  # Save the index of the last processed row
        idx += 1
        time.sleep(1)

if __name__ == "__main__":
    click_at_positions()
    pyautogui.hotkey('alt', 'tab')    
    input("Press Enter to close...")  # Wait for user input to close the prompt
