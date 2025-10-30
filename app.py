import os
import re
import openpyxl
from urllib.parse import quote
from time import sleep
import datetime
import platform

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# ------------------------------
# CONFIGURAÇÃO
# ------------------------------
from dotenv import load_dotenv
load_dotenv()
CHROME_PROFILE_PATH = os.getenv("CHROME_PROFILE_PATH")  # Perfil do Chrome para manter sessão do WhatsApp
PLANILHAS_DIR = os.getenv("PLANILHAS_DIR")  # Diretório onde estão as planilhas Excel
os.makedirs(CHROME_PROFILE_PATH, exist_ok=True) # Garante que o diretório do perfil do Chrome exista
os.makedirs(PLANILHAS_DIR, exist_ok=True) # Garante que o diretório das planilhas exista

# Permite ao usuário alterar o diretório das planilhas
print(f"O programa, por padrão, utiliza o diretório '{PLANILHAS_DIR}' para ler as planilhas e enviar mensagens.\n")
pasta_input = input("Pressione Enter para manter ou digite o caminho absoluto do outro diretório a ser utilizado: ").strip()
if pasta_input:
    PLANILHAS_DIR = pasta_input

# ------------------------------
# CARREGA PLANILHA MAIS RECENTE
# ------------------------------
arquivos = [os.path.join(PLANILHAS_DIR, f) for f in os.listdir(PLANILHAS_DIR) if f.endswith('.xlsx')]
if not arquivos:
    raise FileNotFoundError(f"Nenhum arquivo .xlsx encontrado em {PLANILHAS_DIR}")

arquivo_mais_recente = max(arquivos, key=os.path.getctime)
print(f"📂 Arquivo mais recente encontrado: {arquivo_mais_recente}")

workbook = openpyxl.load_workbook(arquivo_mais_recente)
pagina_clientes = workbook['RANTING'] if 'RANTING' in workbook.sheetnames else workbook.active

# ------------------------------
# CONFIGURA CHROME / SELENIUM
# ------------------------------
chrome_options = Options()
chrome_options.add_argument(f"--user-data-dir={CHROME_PROFILE_PATH}")
chrome_options.add_argument("--profile-directory=Default")

SO = platform.system().lower()

if getattr(sys, 'frozen', False):
    if SO.startswith("win"):
        chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver-windows.exe')
    else:
        chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver-linux')
else:
    if SO.startswith("win"):
        chromedriver_path = "chrome_drivers/chromedriver-windows.exe"
    else:
        chromedriver_path = "chrome_drivers/chromedriver-linux"

service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Abre WhatsApp Web
driver.get("https://web.whatsapp.com/")
input("📌 Escaneie o QR Code ou confirme login e pressione Enter...")

# ------------------------------
# LOOP PARA ENVIAR MENSAGENS
# ------------------------------
for linha in pagina_clientes.iter_rows(min_row=2, values_only=True):
    Pdv = linha[0]
    Nome_Pdv = linha[1]
    contato = linha[2]
    data_chamada = linha[5] if len(linha) > 5 else None
    motivo = linha[8] if len(linha) > 8 else None

    # pula se não tiver contato
    if not contato:
        continue

    # formata contato
    contato_str = str(contato or "").strip()
    contato_digits = re.sub(r'[^\d+]', '', contato_str)
    if contato_digits.startswith('0'):
        contato_digits = contato_digits.lstrip('0')

    # formata data
    if isinstance(data_chamada, (datetime.date, datetime.datetime)):
        data_str = data_chamada.strftime('%d/%m/%y')
    else:
        data_str = str(data_chamada or "")

    # mensagem personalizada
    mensagem = (
        f"Olá {Nome_Pdv}, tudo bem?\n"
        f"Gostaria de entender o motivo da sua Avaliação do dia {data_str} ser '{motivo}'.\n"
        "Podemos agendar uma conversa para melhorarmos sua experiência?"
    )

    # abre conversa
    link_whatsapp = f"https://web.whatsapp.com/send?phone={contato_digits}"
    driver.get(link_whatsapp)

    try:
        # espera a caixa de mensagem carregar
        caixa_msg = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='10']"))
        )

        sleep(1)  # pequena pausa para garantir que está ativa
        caixa_msg.click()
        caixa_msg.send_keys(mensagem + Keys.ENTER)
        sleep(2)
        print(f"✅ Mensagem enviada para {Nome_Pdv}")

    except Exception as e:
        print(f"⚠️ Erro ao enviar mensagem para {Nome_Pdv}: {e}")
        with open('erros.csv', 'a', encoding='utf-8') as arquivo_erro:
            arquivo_erro.write(f"{e} - {Pdv},{Nome_Pdv},{contato}\n")

# ------------------------------
# FINALIZA
# ------------------------------
driver.quit() # fecha o navegador
print("🏁 Processo finalizado com sucesso!")
