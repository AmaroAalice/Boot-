

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 



webbrowser.open('https://web.whatsapp.com/')
sleep(30) #esperar 30 segundos para o usuario escanear o QR code

woorkbook = openpyxl.load_workbook('10- Outubro Rating -.xlsx')
woorkbook['RANTING']
pagina_clientes = woorkbook['RANTING']
for linha in pagina_clientes.iter_rows(min_row=2):
 
 #Pdv,Nome Pdv ,contato,data chamada,data atendimento,motivo 

    Pdv = linha[0].value
    Nome_Pdv = linha[1].value
    contato = linha[2].value
    data_chamada = linha[5].value
    data_atendimento = linha[6].value
    motivo = linha[8].value

   #criar link persolizado do whatsapp e enviar msg para cada cliente com base nos dados da planilha
    mensagem = f"Olá {Nome_Pdv}, tudo bem? Gostaria de entender o motivo da sua Avaliação do dia {data_chamada.strftime('%d/%m/%y')} ser {motivo}. Podemos agendar uma conversa?"
   
 
    try:
        link_mensagem_whatsapp = f"https://api.whatsapp.com/send?phone={contato}&text={quote(mensagem)}"
        webbrowser.open(link_mensagem_whatsapp)  #abrir o link no navegador padrao
        sleep(10)
        seta = pyautogui.locateOnScreen('Seta.jpeg')
        (863, 417, 70, 13) #localizar o botao enviar na tela
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w') #fechar a aba do navegador
        sleep(5)
    except:
        print(f'Não foi possivel enviar a mensagem para {Nome_Pdv}, verifique se o número {contato} está correto.')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{Pdv},{Nome_Pdv},{contato}\n')
        

