
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


#ler os dados da planilha e guardar informações sobre nome e telefone 
workbook = openpyxl.load_workbook('Pasta 1.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    print(nome)
    print(telefone)

#criar os links personalizados e enviar mensagem para cada cliente com base nos dados da planilha
    
    mensagem = f'olá {nome} essa é uma mensagem automatica de teste'
    link_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_wpp)
    sleep(15)
    try:
        seta =  pyautogui.locateCenterOnScreen('seta.png')
        sleep(7)
        pyautogui.click(seta[0], seta[1])
        sleep(7)
        pyautogui.hotkey('ctrl', 'w')
        sleep(7)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')