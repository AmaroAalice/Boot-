import os
import re
import sys
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
# CONFIGURAÇÃO (variáveis de ambiente e diretórios)
# ------------------------------
# Usa python-dotenv para carregar variáveis do arquivo .env
from dotenv import load_dotenv
load_dotenv()

# PLANILHAS_DIR deve apontar para a pasta onde estão os arquivos .xlsx
# Se a variável não existir no .env, será None — o os.makedirs abaixo falhará.
# Portanto garanta que PLANILHAS_DIR esteja definido no .env antes de rodar.
PLANILHAS_DIR = os.getenv("PLANILHAS_DIR")
os.makedirs(PLANILHAS_DIR, exist_ok=True)  # cria a pasta se não existir

# Permite que o usuário altere o diretório em tempo de execução.
# Isso é útil para compartilhar o script com amigos que tenham pastas diferentes.
print(f"O programa, por padrão, utiliza o diretório '{PLANILHAS_DIR}' para ler as planilhas e enviar mensagens.\n")
pasta_input = input("Pressione Enter para manter ou digite o caminho absoluto do outro diretório a ser utilizado: ").strip()
if pasta_input:
    PLANILHAS_DIR = pasta_input

# ------------------------------
# CARREGA A PLANILHA MAIS RECENTE
# ------------------------------
# Lista todos os arquivos .xlsx na pasta e pega o mais recente (por data de criação)
arquivos = [os.path.join(PLANILHAS_DIR, f) for f in os.listdir(PLANILHAS_DIR) if f.endswith('.xlsx')]
if not arquivos:
    raise FileNotFoundError(f"Nenhum arquivo .xlsx encontrado em {PLANILHAS_DIR}")

arquivo_mais_recente = max(arquivos, key=os.path.getctime)
print(f"📂 Arquivo mais recente encontrado: {arquivo_mais_recente}")

# Abre a planilha com openpyxl e escolhe a aba 'RANTING' se existir,
# caso contrário usa a aba ativa. Isso evita erro se a aba tiver nome diferente.
workbook = openpyxl.load_workbook(arquivo_mais_recente)
pagina_clientes = workbook['RANTING'] if 'RANTING' in workbook.sheetnames else workbook.active

# ------------------------------
# CONFIGURA CHROME / SELENIUM
# ------------------------------
# Prepara opções do Chrome que ajudam em ambientes sem display (servidores) e estabilidade.
chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-software-rasterizer")
chrome_options.add_argument("--remote-debugging-port=9222")

# Detecta sistema operacional para escolher o chromedriver correto
SO = platform.system().lower()

# Quando o script é empacotado com PyInstaller, os arquivos extras (como chromedriver)
# são extraídos para sys._MEIPASS; por isso tratamos de forma diferente se está "frozen".
if getattr(sys, 'frozen', False):
    # caminho dentro do bundle criado pelo PyInstaller
    if SO.startswith("win"):
        chromedriver_path = os.path.join(sys._MEIPASS, 'chrome_drivers', 'chromedriver-windows.exe')
    else:
        chromedriver_path = os.path.join(sys._MEIPASS, 'chrome_drivers', 'chromedriver-linux')
else:
    # durante o desenvolvimento, usa a pasta local chrome_drivers
    if SO.startswith("win"):
        chromedriver_path = os.path.join('chrome_drivers', 'chromedriver-windows.exe')
    else:
        chromedriver_path = os.path.join('chrome_drivers', 'chromedriver-linux')

# Cria o Service do chromedriver e inicializa o webdriver com as opções definidas acima.
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Opção opcional: usar perfil do Chrome para persistir sessão (evitar QR toda vez).
# Para habilitar defina USE_CHROME_PROFILE=1 e opcionalmente PROFILE_DIR no .env.
use_profile = os.getenv("USE_CHROME_PROFILE") == "1"
if use_profile:
    profile_dir = os.getenv("PROFILE_DIR")
    if not profile_dir:
        # caminho padrão dentro do home do usuário (Linux)
        profile_dir = os.path.join(os.path.expanduser("~"), ".config", "chrome-whatsapp")
    os.makedirs(profile_dir, exist_ok=True)
    chrome_options.add_argument(f"--user-data-dir={profile_dir}")
    # Nota: use_profile precisa ser definido antes de criar o driver para surtir efeito.
    # Se quiser ativar o perfil, mova esse bloco acima da criação do webdriver.

# ------------------------------
# ABRE WHATSAPP WEB E PAUSA PARA LOGIN
# ------------------------------
# Abre a página do WhatsApp Web. O usuário precisa escanear o QR Code (na primeira vez)
# ou confirmar o login no navegador se já tiver sessão ativa.
driver.get("https://web.whatsapp.com/")
input("📌 Escaneie o QR Code ou confirme login e pressione Enter...")

# ------------------------------
# LOOP PARA ENVIAR MENSAGENS (LINHA A LINHA)
# ------------------------------
# Itera sobre as linhas da planilha (a partir da segunda, assumindo cabeçalho na primeira)
for linha in pagina_clientes.iter_rows(min_row=2, values_only=True):
    Pdv = linha[0]
    Nome_Pdv = linha[1]
    contato = linha[2]
    data_chamada = linha[4] if len(linha) >= 5 else None
    motivo = linha[7] if len(linha) >= 8 else None

    # Se não houver contato, pula essa linha
    if not contato:
        continue

    # FORMATAÇÃO DO NÚMERO:
    # - transforma em string
    # - remove tudo que não for dígito ou '+' (ex: espaços, parênteses, traços)
    # - remove zero inicial se houver (ajusta números locais)
    contato_str = str(contato or "").strip()
    contato_digits = re.sub(r'[^\d+]', '', contato_str)
    if contato_digits.startswith('0'):
        contato_digits = contato_digits.lstrip('0')

    # FORMATAÇÃO DA DATA: se a célula for do tipo date, converte para dd/mm/aa
    if isinstance(data_chamada, (datetime.date, datetime.datetime)):
        data_str = data_chamada.strftime('%d/%m/%y')
    else:
        data_str = str(data_chamada or "")

    # Mensagem personalizada - pode ser alterada conforme necessidade.
    mensagem = (
        f"Olá {Nome_Pdv}, tudo bem?\n"
        f"Gostaria de entender o motivo da sua avaliação do dia {data_str} ser '{motivo}'.\n"
        "Podemos agendar uma conversa para melhorarmos sua experiência?"
    )

    # Monta o link para abrir a conversa no WhatsApp Web.
    # Observação: se o número não estiver no formato internacional, o link pode falhar.
    link_whatsapp = f"https://web.whatsapp.com/send?phone={contato_digits}"
    driver.get(link_whatsapp)

    try:
        # Espera até que a caixa de mensagem esteja presente na página.
        # O XPath e o atributo data-tab podem mudar com atualizações do WhatsApp Web,
        # então pode ser necessário ajustar esse seletor no futuro.
        caixa_msg = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab='10']"))
        )

        sleep(1)  # pequena espera para garantir foco
        caixa_msg.click()
        caixa_msg.send_keys(mensagem + Keys.ENTER)  # envia a mensagem pressionando Enter
        sleep(2)  # espera curta após envio
        print(f"✅ Mensagem enviada para {Nome_Pdv}")

    except Exception as e:
        # Em caso de erro, registra no console e em um arquivo de erros para análise posterior.
        print(f"⚠️ Erro ao enviar mensagem para {Nome_Pdv}: {e}")
        with open('erros.csv', 'a', encoding='utf-8') as arquivo_erro:
            arquivo_erro.write(f"@{e} - {Pdv}, {Nome_Pdv}, {contato}\n")

# ------------------------------
# FINALIZAÇÃO
# ------------------------------
driver.quit()  # fecha o navegador controlado pelo Selenium
print("🏁 Processo finalizado com sucesso!")
