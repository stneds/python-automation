import gspread
import os
import time
import random
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES ---
CONFIG = {
    "JSON_CREDS": "credenciais.json.json",
    "PLANILHA": "Teste de Script-Preço TRRs 01/26",
    "ABA": "Preço 12-01",
    "PERFIL_AVULSO": os.path.join(os.getcwd(), "Perfil_Robo"),
    "USUARIO": "cacique01",
    "SENHA": "DERIVADOS02"
}

IDS = {
    "CAMPO_USER": "viewns_Z7_LA04H4G0POLN00QRBOJ72420P5_:j_id_6:login",
    "CAMPO_SENHA": "viewns_Z7_LA04H4G0POLN00QRBOJ72420P5_:j_id_6:senha",
    "BTN_ENTRAR": "viewns_Z7_LA04H4G0POLN00QRBOJ72420P5_:j_id_6:btnEntrar"
}

# --- FUNÇÕES DE HUMANIZAÇÃO ---

def pausa_aleatoria(minimo=1, maximo=3):
    time.sleep(random.uniform(minimo, maximo))

def digitar_como_humano(elemento, texto):
    """Digita o texto caractere por caractere com atrasos variáveis."""
    for caractere in texto:
        elemento.send_keys(caractere)
        time.sleep(random.uniform(0.05, 0.2))

def mover_e_clicar(driver, elemento):
    """Move o mouse até o elemento antes de clicar."""
    action = ActionChains(driver)
    action.move_to_element(elemento).pause(random.uniform(0.5, 1.0)).click().perform()

def clean_and_parse(text, discount=0.0):
    try:
        val = text.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.').strip()
        return round(float(val) - discount, 4)
    except: return 0.0

# --- INÍCIO ---
options = Options()
options.add_argument(f"--user-data-dir={CONFIG['PERFIL_AVULSO']}")
options.add_argument("--start-maximized")

# Esconde a flag "navigator.webdriver" para o site não saber que é automação
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Remove rastro de automação via execução de script
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
})

wait = WebDriverWait(driver, 25)

try:
    print("🚀 Acessando o Portal...")
    driver.get("https://www.redeipiranga.com.br/wps/portal/redeipiranga/login/")
    pausa_aleatoria(3, 5) # Espera o carregamento parecer natural

    # 1. Digita Usuário
    print("✍️ Preenchendo login...")
    user_el = wait.until(EC.element_to_be_clickable((By.ID, IDS["CAMPO_USER"])))
    user_el.clear()
    digitar_como_humano(user_el, CONFIG["USUARIO"])
    pausa_aleatoria(0.5, 1.5)

    # 2. Digita Senha
    print("✍️ Preenchendo senha...")
    pass_el = driver.find_element(By.ID, IDS["CAMPO_SENHA"])
    pass_el.clear()
    digitar_como_humano(pass_el, CONFIG["SENHA"])
    pausa_aleatoria(1, 2)

    # 3. Clica no Botão Entrar de forma humana
    print("🖱️ Movendo mouse e clicando em Entrar...")
    btn_entrar = driver.find_element(By.ID, IDS["BTN_ENTRAR"])
    mover_e_clicar(driver, btn_entrar)

    # 4. Navegação para Preços
    print("⌛ Aguardando login...")
    btn_pedidos = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'registrarpedido')]")))
    pausa_aleatoria(2, 4)
    mover_e_clicar(driver, btn_pedidos)

    # 5. Captura dos Preços
    print("📊 Coletando dados...")
    wait.until(EC.visibility_of_element_located((By.ID, "idPrecoRetira15310000")))
    pausa_aleatoria(2, 3) # Simula tempo de leitura
    
    s10_raw = driver.find_element(By.ID, "idPrecoRetira15310000").text
    gas_raw = driver.find_element(By.ID, "idPrecoRetira11100001").text

    resultados = {
        's10': clean_and_parse(s10_raw, 0.20),
        'gasolina': clean_and_parse(gas_raw, 0.20)
    }

    # 6. Grava na Planilha
    print("📝 Gravando no Google Sheets...")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CONFIG["JSON_CREDS"], scope)
    client = gspread.authorize(creds)
    sheet = client.open(CONFIG["PLANILHA"]).worksheet(CONFIG["ABA"])
    
    updates = [
        {'range': 'E12', 'values': [[resultados['s10']]]},
        {'range': 'M12', 'values': [[resultados['gasolina']]]}
    ]
    sheet.batch_update(updates)
    print(f"✅ SUCESSO! S10: {resultados['s10']} | Gas: {resultados['gasolina']}")

except Exception as e:
    print(f"💥 Erro: {e}")
finally:
    pausa_aleatoria(2, 4)
    driver.quit()