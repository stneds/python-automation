import re
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1




# =========================================================
# CONFIGURAÇÕES
# =========================================================
ARQUIVO_JSON_GOOGLE = "dados-google.json"
NOME_ARQUIVO_MODELO = "Preço TRRs %m/%y"
ID_PLANILHA_OFICIAL = "1Va1byiasuU-k9dCDmY9mcsUzvlFXmVTCBfuL9IaJq9Y" 
INTERVALO_LEITURA = "A1:U1000"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TITULOS_PROIBIDOS = [
    "PREÇO DO PRODUTO NA BASE",
    "PRECO DO PRODUTO NA BASE",
]


# =========================================================
# CONEXÃO
# =========================================================
def conectar_cliente():
    creds = Credentials.from_service_account_file(
        ARQUIVO_JSON_GOOGLE,
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client, creds.service_account_email


def abrir_planilha_do_mes():
    client, email_servico = conectar_cliente()

    print("🔐 Conta de serviço:", email_servico)
    print(f"🔎 Abrindo planilha oficial por ID: {ID_PLANILHA_OFICIAL}")

    try:
        ss = client.open_by_key(ID_PLANILHA_OFICIAL)
        print("✅ Arquivo encontrado:", ss.title)
        return ss
    except Exception as e:
        print(f"❌ Não foi possível abrir o arquivo do mês: {ID_PLANILHA_OFICIAL}")
        print("Detalhe:", repr(e))
        return None


# =========================================================
# DATAS
# =========================================================
def proximo_dia_util(data_base):
    data = data_base + timedelta(days=1)
    while data.weekday() >= 5:
        data += timedelta(days=1)
    return data


def dia_util_anterior(data_base):
    data = data_base - timedelta(days=1)
    while data.weekday() >= 5:
        data -= timedelta(days=1)
    return data


# =========================================================
# ABAS
# =========================================================
def encontrar_aba_base(lista_abas, data_alvo_dt):
    for i in range(1, 15):
        busca = (data_alvo_dt - timedelta(days=i)).strftime("%d-%m")
        for aba in lista_abas:
            if aba.title.strip() == f"Preço {busca}":
                print("✅ Aba base encontrada:", aba.title)
                return aba

    for i in range(1, 15):
        busca = (data_alvo_dt - timedelta(days=i)).strftime("%d-%m")
        for aba in lista_abas:
            if busca in aba.title:
                print("✅ Aba base encontrada:", aba.title)
                return aba

    print("⚠️ Nenhuma aba anterior encontrada. Usando a primeira aba.")
    return lista_abas[0]


# =========================================================
# UTILITÁRIOS
# =========================================================
def normalizar(txt):
    return str(txt).strip().upper()


def linha_para_texto(linha):
    return " | ".join(str(c).strip() for c in linha if str(c).strip())


def eh_titulo_proibido(valor):
    texto = normalizar(valor)
    return texto in [normalizar(x) for x in TITULOS_PROIBIDOS]


def expandir_linha(linha, tamanho):
    while len(linha) < tamanho:
        linha.append("")
    return linha


def indice_coluna_para_letra(indice_zero_based):
    numero = indice_zero_based + 1
    letras = ""

    while numero > 0:
        numero, resto = divmod(numero - 1, 26)
        letras = chr(65 + resto) + letras

    return letras


def construir_remocoes_validacao_total(sheet_id):
    # Limpa todas as validações do intervalo de trabalho (A1:AD1000).
    return [{
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1000,
                "startColumnIndex": 0,
                "endColumnIndex": 30,
            }
        }
    }]





# =========================================================
# LOCALIZAÇÃO DOS BLOCOS FOB
# =========================================================
def localizar_blocos_fob(dados):
    blocos = []

    for i, linha in enumerate(dados):
        for c, valor in enumerate(linha):
            texto = normalizar(valor)

            if not texto:
                continue

            if texto.startswith("FOB -"):
                blocos.append((texto, i, c))

    return blocos


def achar_fim_bloco_fob(dados, linha_inicio, col_titulo):
    r = linha_inicio + 1

    while r < len(dados):
        for c, valor in enumerate(dados[r]):
            texto = normalizar(valor)

            if not texto:
                continue

            if texto.startswith("FOB -") and abs(c - col_titulo) <= 8:
                return r - 1

        r += 1

    return len(dados) - 1


def encontrar_linha_companhia_e_coluna(dados, linha_titulo, fim, col_titulo):
    for r in range(linha_titulo, min(fim + 1, linha_titulo + 8)):
        for c, valor in enumerate(dados[r]):
            if str(valor).strip().lower() == "companhia" and abs(c - col_titulo) <= 5:
                return r, c
    return None, None


def encontrar_linha_datas_por_offset(dados, linha_titulo, fim, col_titulo):
    padrao_data = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    candidatos = [linha_titulo + 2, linha_titulo + 3, linha_titulo + 4, linha_titulo + 5]

    for r in candidatos:
        if r >= len(dados) or r > fim:
            continue

        linha = [str(x).strip() for x in dados[r]]
        grupos = []

        inicio_busca = max(0, col_titulo - 1)
        fim_busca = min(len(linha) - 2, col_titulo + 20)

        for j in range(inicio_busca, fim_busca):
            a = linha[j]
            b = linha[j + 1]
            c = linha[j + 2]

            if padrao_data.match(a) and padrao_data.match(b) and "Dif" in c:
                grupos.append((j, j + 1, j + 2))

        if grupos:
            return r, grupos

    return None, []


def linha_parece_dado_empresa(linha, col_empresa):
    if col_empresa is None or col_empresa >= len(linha):
        return False

    nome = str(linha[col_empresa]).strip()
    if not nome:
        return False

    nome_up = nome.upper()

    if nome_up == "COMPANHIA":
        return False

    if nome_up.startswith("OBS."):
        return False

    if nome_up.startswith("FOB -"):
        return False

    return True


# =========================================================
# ALTERAÇÕES SEGURAS
# =========================================================
def construir_alteracoes_seguras_fob(dados_formatados, dados_formulas, data_ontem, data_hoje):
    alteracoes = []
    blocos_formatacao = []
    blocos = localizar_blocos_fob(dados_formatados)

    if not blocos:
        print("⚠️ Nenhum bloco FOB foi encontrado.")
        return alteracoes, blocos_formatacao

    print(f"📦 Blocos FOB encontrados: {len(blocos)}")

    for titulo, linha_titulo, col_titulo in blocos:
        if eh_titulo_proibido(titulo):
            continue

        fim = achar_fim_bloco_fob(dados_formatados, linha_titulo, col_titulo)

        linha_companhia, col_empresa = encontrar_linha_companhia_e_coluna(
            dados_formatados,
            linha_titulo,
            fim,
            col_titulo
        )

        linha_datas, grupos = encontrar_linha_datas_por_offset(
            dados_formatados,
            linha_titulo,
            fim,
            col_titulo
        )

        if linha_companhia is None or col_empresa is None:
            print(f"⚠️ Não achei 'Companhia' no bloco: {titulo}")
            continue

        if linha_datas is None or not grupos:
            print(f"⚠️ Não achei os pares de data no bloco: {titulo}")
            print("----- DEBUG BLOCO -----")
            for rr in range(linha_titulo, min(fim + 1, linha_titulo + 8)):
                print(rr + 1, dados_formatados[rr])
            print("-----------------------")
            continue

        inicio_dados = linha_datas + 1
        max_col = max(g[2] for g in grupos)

        blocos_formatacao.append({
            "inicio_dados": inicio_dados,
            "fim": fim,
            "grupos": grupos,
        })

        print(f"✅ Bloco FOB pronto para edição: {titulo}")

        # Atualiza o cabeçalho das datas
        for col_ontem, col_hoje, _ in grupos:
            alteracoes.append({
                "range": rowcol_to_a1(linha_datas + 1, col_ontem + 1),
                "values": [[data_ontem]]
            })
            alteracoes.append({
                "range": rowcol_to_a1(linha_datas + 1, col_hoje + 1),
                "values": [[data_hoje]]
            })

        # Move hoje -> ontem e limpa hoje somente se não for fórmula
        for r in range(inicio_dados, fim + 1):
            dados_formatados[r] = expandir_linha(dados_formatados[r], max_col + 1)
            dados_formulas[r] = expandir_linha(dados_formulas[r], max_col + 1)

            if not linha_parece_dado_empresa(dados_formatados[r], col_empresa):
                continue

            for col_ontem, col_hoje, col_dif in grupos:
                valor_hoje_fmt = dados_formatados[r][col_hoje]
                valor_hoje_formula = dados_formulas[r][col_hoje]

                # copia o valor visível da coluna "hoje" para a coluna "ontem"
                alteracoes.append({
                    "range": rowcol_to_a1(r + 1, col_ontem + 1),
                    "values": [[valor_hoje_fmt]]
                })

                # limpa "hoje" apenas se não for fórmula
                if not str(valor_hoje_formula).startswith("="):
                    alteracoes.append({
                        "range": rowcol_to_a1(r + 1, col_hoje + 1),
                        "values": [[""]]
                    })

                # NÃO mexe na coluna Dif. (R$)

    return alteracoes, blocos_formatacao


# =========================================================
# PRINCIPAL
# =========================================================
def preparar_aba():
    ss = abrir_planilha_do_mes()
    if not ss:
        return

    try:
        agora = datetime.now()
        nomes_abas = [a.title for a in ss.worksheets()]
        hoje_str = agora.strftime("%d-%m")

        if any(nome.strip() == f"Preço {hoje_str}" for nome in nomes_abas):
            data_alvo = proximo_dia_util(agora)
        else:
            data_alvo = agora
            while data_alvo.weekday() >= 5:
                data_alvo += timedelta(days=1)

        nome_novo = data_alvo.strftime("Preço %d-%m")
        data_hoje = data_alvo.strftime("%d/%m/%Y")
        data_ontem = dia_util_anterior(data_alvo).strftime("%d/%m/%Y")

        print("📅 Aba alvo:", nome_novo)
        print("📆 Data anterior:", data_ontem)
        print("📆 Data atual:", data_hoje)

        if nome_novo in nomes_abas:
            print(f"⚠️ A aba {nome_novo} já existe.")
            return

        aba_base = encontrar_aba_base(ss.worksheets(), data_alvo)

        print("🔄 Duplicando aba base...")
        nova = ss.duplicate_sheet(
            source_sheet_id=aba_base.id,
            new_sheet_name=nome_novo
        )

        print("📥 Lendo conteúdo formatado...")
        dados_formatados = nova.get(INTERVALO_LEITURA, value_render_option="FORMATTED_VALUE")

        print("📥 Lendo conteúdo com fórmulas...")
        dados_formulas = nova.get(INTERVALO_LEITURA, value_render_option="FORMULA")

        print("🛠 Montando alterações seguras apenas nos blocos FOB...")
        alteracoes, blocos_formatacao = construir_alteracoes_seguras_fob(
            dados_formatados=dados_formatados,
            dados_formulas=dados_formulas,
            data_ontem=data_ontem,
            data_hoje=data_hoje
        )

        remocoes_validacao = construir_remocoes_validacao_total(sheet_id=nova.id)

        if alteracoes:
            print(f"💾 Aplicando {len(alteracoes)} alterações...")
            nova.batch_update(alteracoes, value_input_option="USER_ENTERED")

            ss.batch_update({"requests": remocoes_validacao })
            print("🧹 Todas as validações de dados foram removidas (A1:U1000).")

            print(f"✅ Aba {nome_novo} criada e ajustada com segurança.")
        else:
            ss.batch_update({"requests": remocoes_validacao })
            print("🧹 Todas as validações de dados foram removidas (A1:U1000).")
            print(" Formatação dos blocos FOB foi renovada para o novo dia.")
            print(" Nenhuma alteração foi montada. A aba foi criada, mas nada foi editado.")

    except Exception as e:
        print(f"❌ Erro crítico: {repr(e)}")


if __name__ == "__main__":
    preparar_aba()

    # Executa a coleta da Vibra em seguida, usando o mesmo interpretador Python.
    caminho_vibra = Path(__file__).resolve().parent / "vibra.py"

    if not caminho_vibra.exists():
        print(f"❌ vibra.py não encontrado em: {caminho_vibra}")
    else:
        print("\n🚀 Iniciando processo da Vibra...")
        resultado = subprocess.run([sys.executable, str(caminho_vibra)])

        if resultado.returncode == 0:
            print("✅ Processo da Vibra finalizado com sucesso.")
        else:
            print(f"❌ Processo da Vibra finalizou com erro (código {resultado.returncode}).")
