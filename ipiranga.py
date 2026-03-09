import re
import time
from dataclasses import dataclass
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================================================
# CONFIGURAÇÕES
# =========================================================
URL_HOME = "https://www.redeipiranga.com.br/wps/myportal/redeipiranga/homeoutros/"

ARQUIVO_JSON_GOOGLE = "dados-google.json"
NOME_ARQUIVO_MODELO = "Preço teste TRRs %m/%y"
INTERVALO_LEITURA = "A1:AD1000"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# TESTE COM UMA BASE SÓ
@dataclass
class BaseConfig:
    nome_base: str
    cnpj: str
    uf_cidade: str
    ponto_entrega: str
    titulo_bloco_planilha: str


BASE_TESTE = BaseConfig(
    nome_base="TERESINA",
    cnpj="61243090000136",
    uf_cidade="PI - Urucui",
    ponto_entrega="01 | Zona Rural",
    titulo_bloco_planilha="FOB - BASE TERESINA - PI - NEOAGRO",
)


# =========================================================
# SELENIUM
# =========================================================
def iniciar_driver():
    options = Options()
    options.add_argument("--start-maximized")
    return webdriver.Chrome(options=options)


def wait(driver, timeout=20):
    return WebDriverWait(driver, timeout)


def aguardar_login_manual(driver):
    driver.get(URL_HOME)
    print("Faça o login manual no portal da Ipiranga.")
    print("Depois que entrar e estiver na home, volte no terminal e aperte ENTER.")
    input()


def abrir_popup_cliente(driver):
    w = wait(driver, 20)

    tentativas = [
        (By.XPATH, "//*[contains(., 'CNPJ:')]/ancestor::*[self::div or self::button][1]"),
        (By.XPATH, "//*[contains(., 'Neoagro Diesel Ltda')]"),
    ]

    for by, xp in tentativas:
        try:
            el = w.until(EC.element_to_be_clickable((by, xp)))
            driver.execute_script("arguments[0].click();", el)
            time.sleep(2)
            return
        except Exception:
            continue

    raise RuntimeError("Não consegui abrir o seletor de cliente/CNPJ.")


def escolher_cnpj(driver, base: BaseConfig):
    abrir_popup_cliente(driver)

    w = wait(driver, 20)

    campo = w.until(
        EC.presence_of_element_located(
            (By.XPATH, "//input[contains(@placeholder,'Razão Social') or contains(@placeholder,'CNPJ')]")
        )
    )
    campo.clear()
    campo.send_keys(base.cnpj)
    time.sleep(2)

    xpath_card = (
        f"//*[contains(., '{base.cnpj}') and contains(., '{base.uf_cidade}') and contains(., '{base.ponto_entrega}')]"
    )
    card = w.until(EC.presence_of_element_located((By.XPATH, xpath_card)))

    botao = card.find_element(By.XPATH, ".//button[contains(., 'Trocar')]")
    driver.execute_script("arguments[0].click();", botao)
    time.sleep(4)


def abrir_pedido_combustivel(driver):
    w = wait(driver, 20)

    tentativas = [
        "//button[contains(., 'Pedido de Combustível')]",
        "//*[contains(., 'Pedido de Combustível')]",
        "//a[contains(., 'Pedido de Combustível')]",
    ]

    for xp in tentativas:
        try:
            el = w.until(EC.element_to_be_clickable((By.XPATH, xp)))
            driver.execute_script("arguments[0].click();", el)
            time.sleep(5)
            return
        except Exception:
            continue

    raise RuntimeError("Não consegui abrir 'Pedido de Combustível'.")


def capturar_precos_retira(driver):
    produtos_alvo = {
        "Diesel S10 Bb": "S10",
        "Diesel S500 Bb": "S500",
        "Gasolina Comum Bb": "GASOLINA",
    }

    resultados = {
        "S10": None,
        "S500": None,
        "GASOLINA": None,
    }

    for nome_tela, chave in produtos_alvo.items():
        try:
            el_produto = driver.find_element(By.XPATH, f"//*[contains(., '{nome_tela}')]")
            linha = el_produto.find_element(By.XPATH, "./ancestor::div[2]")
            texto = " ".join(linha.text.split())

            precos = re.findall(r"\d+,\d+", texto)
            retira = precos[1] if len(precos) >= 2 else None

            resultados[chave] = retira
        except Exception:
            resultados[chave] = None

    return resultados


# =========================================================
# GOOGLE SHEETS
# =========================================================
def conectar_planilha_mes():
    creds = Credentials.from_service_account_file(
        ARQUIVO_JSON_GOOGLE,
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    nome_arquivo = datetime.now().strftime(NOME_ARQUIVO_MODELO)
    print("Abrindo planilha:", nome_arquivo)
    return client.open(nome_arquivo)


def achar_aba_do_dia(ss):
    nome_aba = datetime.now().strftime("Preço %d-%m")
    try:
        return ss.worksheet(nome_aba)
    except Exception:
        raise RuntimeError(f"Aba do dia não encontrada: {nome_aba}")


def encontrar_bloco_e_linha_ipiranga(valores, titulo_bloco):
    linha_titulo = None

    for i, linha in enumerate(valores):
        for cel in linha:
            if str(cel).strip().upper() == titulo_bloco.upper():
                linha_titulo = i
                break
        if linha_titulo is not None:
            break

    if linha_titulo is None:
        raise RuntimeError(f"Bloco não encontrado: {titulo_bloco}")

    linha_companhia = None
    col_empresa = None

    for r in range(linha_titulo, min(linha_titulo + 8, len(valores))):
        for c, cel in enumerate(valores[r]):
            if str(cel).strip().lower() == "companhia":
                linha_companhia = r
                col_empresa = c
                break
        if linha_companhia is not None:
            break

    if linha_companhia is None:
        raise RuntimeError(f"'Companhia' não encontrada no bloco: {titulo_bloco}")

    linha_ipiranga = None
    for r in range(linha_companhia + 1, min(linha_companhia + 20, len(valores))):
        if col_empresa < len(valores[r]) and str(valores[r][col_empresa]).strip().upper() == "IPIRANGA":
            linha_ipiranga = r
            break

    if linha_ipiranga is None:
        raise RuntimeError(f"Linha IPIRANGA não encontrada no bloco: {titulo_bloco}")

    return linha_titulo, linha_companhia, col_empresa, linha_ipiranga


def encontrar_colunas_dia_novo(valores, linha_titulo):
    padrao_data = re.compile(r"^\d{2}/\d{2}/\d{4}$")

    for r in range(linha_titulo, min(linha_titulo + 8, len(valores))):
        linha = [str(x).strip() for x in valores[r]]
        grupos = []

        for j in range(len(linha) - 2):
            if (
                padrao_data.match(linha[j])
                and padrao_data.match(linha[j + 1])
                and "Dif" in linha[j + 2]
            ):
                grupos.append((j, j + 1, j + 2))

        if len(grupos) >= 3:
            return {
                "S10": grupos[0][1],
                "S500": grupos[1][1],
                "GASOLINA": grupos[2][1],
            }

    raise RuntimeError("Colunas do dia novo não encontradas no bloco.")


def escrever_precos_ipiranga(ss, base: BaseConfig, precos):
    aba = achar_aba_do_dia(ss)
    valores = aba.get(INTERVALO_LEITURA, value_render_option="FORMATTED_VALUE")

    linha_titulo, linha_companhia, col_empresa, linha_ipiranga = encontrar_bloco_e_linha_ipiranga(
        valores,
        base.titulo_bloco_planilha
    )

    colunas = encontrar_colunas_dia_novo(valores, linha_titulo)

    updates = []

    for produto, valor in precos.items():
        if valor:
            col = colunas[produto]
            a1 = gspread.utils.rowcol_to_a1(linha_ipiranga + 1, col + 1)
            print(f"Gravando {produto} = {valor} em {a1}")
            updates.append({
                "range": a1,
                "values": [[valor]]
            })

    if not updates:
        raise RuntimeError("Nenhum preço foi capturado para gravar.")

    aba.batch_update(updates, value_input_option="USER_ENTERED")
    print("✅ Preços gravados na planilha com sucesso.")


# =========================================================
# FLUXO PRINCIPAL
# =========================================================
def main():
    driver = iniciar_driver()

    try:
        ss = conectar_planilha_mes()

        aguardar_login_manual(driver)

        print(f"\nProcessando base: {BASE_TESTE.nome_base}")
        escolher_cnpj(driver, BASE_TESTE)
        abrir_pedido_combustivel(driver)

        precos = capturar_precos_retira(driver)
        print("Preços capturados:", precos)

        escrever_precos_ipiranga(ss, BASE_TESTE, precos)

        print("✅ Teste finalizado com sucesso.")

    finally:
        input("Pressione ENTER para fechar o navegador...")
        driver.quit()


if __name__ == "__main__":
    main()