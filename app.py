import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
from pathlib import Path

#Abre o Whatssap Web
webbrowser.open('https://web.whatsapp.com/')
sleep(30) # tempo para o usuário escanear o Qr Code

#Carrega planilha do excel
workbook = openpyxl.load_workbook('lista_numeros.xlsx')
pagina_clientes = workbook['Sheet1']

#Caminho da imagem
imagem_seta = Path('seta.png').resolve()
print(imagem_seta)
caminho_str = str(imagem_seta)

#Iterar sobre as linhas da planilha, a partir da segunda linha
for linha in pagina_clientes.iter_rows(min_row= 2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá {nome}, Esta mensagem é para lembrá-lo da sua data de vencimento dia {vencimento.strftime("%d/%m/%Y")}. Pague o quanto antes.'


    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)

    sleep(10)
    try:
        # Localizar e clicar na seta para enviar a mensagem
        seta = pyautogui.locateCenterOnScreen(caminho_str)
        if seta:
            sleep(10)
            pyautogui.click(seta[0],seta[1])
            sleep(10)
            pyautogui.hotkey('ctrl', 'w')
            sleep(10)
        else:
            raise FileNotFoundError('Seta não encontrada na tela')
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone} \n')