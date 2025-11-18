import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import locale
import unicodedata
import time

# Selenium imports e ChromeDriver automático
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# -----------------------
# Configuração de locale
# -----------------------
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, "pt_BR")
    except locale.Error:
        locale.setlocale(locale.LC_TIME, "")

# ==============================
# 🔤 Normalização de colunas
# ==============================
def normalizar_colunas(df):
    def remover_acentos(s):
        if isinstance(s, str):
            return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        return s

    df.columns = [remover_acentos(col).strip().upper() for col in df.columns]
    renomear = {
        "MUNICIPIO ORIGEM": "MUNICIPIO_ORIGEM",
        "PRESIDENTE DE BANCA": "PRESIDENTE_DE_BANCA",
        "INICIO INDISPONIBILIDADE": "INICIO_INDISPONIBILIDADE",
        "FIM INDISPONIBILIDADE": "FIM_INDISPONIBILIDADE",
        "DIAS INDISPONIBILIDADE": "DIAS_INDISPONIBILIDADE",
        "TELEFONE": "TELEFONE"
    }
    for antiga, nova in renomear.items():
        if antiga in df.columns:
            df.rename(columns={antiga: nova}, inplace=True)
    for col in ["MUNICIPIO", "MUNICIPIO_ORIGEM", "CATEGORIA", "NOME", "TELEFONE"]:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(lambda s: remover_acentos(s).strip().upper() if isinstance(s, str) else s)
    return df

# ==============================
# 🔍 Contagem de categorias compatíveis (fallback E)
# ==============================
def matching_count_fallback(categorias_pessoa, categorias_operacao):
    if not isinstance(categorias_pessoa, str) or not isinstance(categorias_operacao, str):
        return 0
    pessoa_set = set(x.strip().upper() for x in categorias_pessoa.split(",") if x.strip())
    oper_list = [x.strip().upper() for x in categorias_operacao.split(",") if x.strip()]
    if not oper_list:
        return 0
    primeira = oper_list[0]
    ultima = oper_list[-1]
    precisa_E = "E" in oper_list
    if precisa_E:
        if "E" in pessoa_set:
            return 1 if ultima in pessoa_set and primeira in pessoa_set else 0
        else:
            return 1 if ultima in pessoa_set and primeira in pessoa_set else 1
    else:
        return 1 if ultima in pessoa_set and primeira in pessoa_set else 0

# ==============================
# 🚫 Verificação de indisponibilidade / férias
# ==============================
dias_map = {"SEGUNDA":0, "TERCA":1, "QUARTA":2, "QUINTA":3, "SEXTA":4, "SABADO":5, "DOMINGO":6}
def esta_indisponivel(nome, dias_indisponiveis, inicio, fim, data):
    if pd.notna(dias_indisponiveis) and str(dias_indisponiveis).strip() != "":
        dias = [d.strip().upper().replace("Ç","C").replace("Á","A") for d in str(dias_indisponiveis).split(",")]
        dias_num = [dias_map[d] for d in dias if d in dias_map]
        if data.weekday() in dias_num:
            return True
        return False
    if pd.notna(inicio) and pd.notna(fim):
        try:
            if inicio.date() <= data.date() <= fim.date():
                return True
        except Exception:
            pass
    return False

# ==============================
# 🌟 Regra Vanessa
# ==============================
def aplicar_regra_vanessa(df_candidatos, categoria_oper, data):
    if isinstance(categoria_oper, str) and categoria_oper.strip().upper() == "B":
        nome_vanessa = "VANESSA APARECIDA CARVALHO DE ASSIS"
        vanessa = df_candidatos[df_candidatos["NOME"].str.upper() == nome_vanessa]
        if not vanessa.empty:
            r = vanessa.iloc[0]
            if not esta_indisponivel(r["NOME"], r.get("DIAS_INDISPONIBILIDADE", ""), r.get("INICIO_INDISPONIBILIDADE"), r.get("FIM_INDISPONIBILIDADE"), data):
                resto = df_candidatos[df_candidatos["NOME"].str.upper() != nome_vanessa]
                df_candidatos = pd.concat([vanessa, resto]).reset_index(drop=True)
                return df_candidatos, True
    return df_candidatos, False

# ==============================
# 🔧 Evita repetição no mesmo dia e evita município de origem
# ==============================
def filtrar_candidatos(df_candidatos, municipio, data, convocados):
    nomes_no_dia = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
    candidatos_filtrados = df_candidatos[
        (~df_candidatos["NOME"].isin(nomes_no_dia)) &
        (df_candidatos["MUNICIPIO_ORIGEM"] != municipio)
    ].copy()
    return candidatos_filtrados

# ==============================
# 🔧 Peso de frequência semanal
# ==============================
def aplicar_regra_frequencia(df_candidatos, data, categoria_oper, conv_semana_global):
    semana = data.isocalendar()[1]
    def calcular_peso(r):
        nome = r["NOME"]
        conv_semana = conv_semana_global.get((nome, semana), 0)
        match = matching_count_fallback(r.get("CATEGORIA",""), categoria_oper)
        return (match * 10) + max(0, 5 - conv_semana)
    if df_candidatos.empty:
        return df_candidatos
    df_candidatos = df_candidatos.copy()
    df_candidatos["PESO"] = df_candidatos.apply(calcular_peso, axis=1)
    df_candidatos = df_candidatos.sort_values(by=["PESO"], ascending=False)
    return df_candidatos

# ==============================
# 🧠 Processamento principal
# ==============================
def processar_distribuicao(arquivo):
    df = pd.read_excel(arquivo)
    df = normalizar_colunas(df)
    
    if "PRESIDENTE_DE_BANCA" not in df.columns:
        df["PRESIDENTE_DE_BANCA"] = "NAO"
    if "TELEFONE" not in df.columns:
        df["TELEFONE"] = ""

    for col in ["DATA", "INICIO_INDISPONIBILIDADE", "FIM_INDISPONIBILIDADE"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if "DIA" not in df.columns and "DATA" in df.columns:
        df["DIA"] = df["DATA"].dt.day_name().str.upper()
    df["DATA"].fillna(method="ffill", inplace=True)
    df["DIA"].fillna(method="ffill", inplace=True)
    if "QUANTIDADE" not in df.columns:
        df["QUANTIDADE"] = 1
    df["QUANTIDADE"] = df["QUANTIDADE"].fillna(0).astype(int)

    nomes_unicos = df["NOME"].unique()
    cont_pres = {n: 0 for n in nomes_unicos}
    conv_semana_global = {}

    convocados = []
    mensagens_vanessa = []

    operacoes = df.groupby(["DIA", "DATA", "MUNICIPIO", "CATEGORIA", "QUANTIDADE"], dropna=False)
    for (dia, data, municipio, categoria_oper, qtd), _ in operacoes:
        qtd = int(qtd)
        if qtd <= 0:
            continue
        data = pd.to_datetime(data)

        candidatos = df.copy()
        candidatos = candidatos.loc[~candidatos.apply(lambda r: esta_indisponivel(
            r["NOME"], r.get("DIAS_INDISPONIBILIDADE",""), r.get("INICIO_INDISPONIBILIDADE"), r.get("FIM_INDISPONIBILIDADE"), data
        ), axis=1)].reset_index(drop=True)

        candidatos = filtrar_candidatos(candidatos, municipio, data, convocados)
        candidatos["MATCH_COUNT"] = candidatos["CATEGORIA"].apply(lambda c: matching_count_fallback(c, categoria_oper))
        candidatos, vanessa_ativa = aplicar_regra_vanessa(candidatos, categoria_oper, data)
        if vanessa_ativa:
            mensagens_vanessa.append(f"✨ Vanessa priorizada em {municipio} ({data.date()})")

        candidatos_pesados = aplicar_regra_frequencia(candidatos, data, categoria_oper, conv_semana_global)

        pres_cand = candidatos_pesados[candidatos_pesados["PRESIDENTE_DE_BANCA"].astype(str).str.upper() == "SIM"]
        pres_cand = pres_cand[~pres_cand["NOME"].isin(
            [c["NOME"] for c in convocados if c["DATA"] == data.date() and c["PRESIDENTE"] == "SIM"]
        )]
        presidente = None
        if not pres_cand.empty:
            nome_pres = sorted(pres_cand["NOME"].unique(), key=lambda n: cont_pres.get(n,0))[0]
            presidente = pres_cand[pres_cand["NOME"] == nome_pres].iloc[0]
            cont_pres[presidente["NOME"]] = cont_pres.get(presidente["NOME"], 0) + 1
            semana_pres = data.isocalendar()[1]
            conv_semana_global[(presidente["NOME"], semana_pres)] = conv_semana_global.get((presidente["NOME"], semana_pres), 0) + 1

        pool = candidatos_pesados.copy()
        if presidente is not None:
            pool = pool[pool["NOME"] != presidente["NOME"]]
        pool = pool.sort_values(by=["PESO"], ascending=False)

        nomes_ja_convocados_no_dia = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
        selecionados = []
        for _, r in pool.iterrows():
            if len(selecionados) >= (qtd - (1 if presidente is not None else 0)):
                break
            nome = r["NOME"]
            if nome not in nomes_ja_convocados_no_dia:
                selecionados.append(nome)
                nomes_ja_convocados_no_dia.append(nome)
                semana_sel = data.isocalendar()[1]
                conv_semana_global[(nome, semana_sel)] = conv_semana_global.get((nome, semana_sel), 0) + 1

        total_selecionados = ([presidente["NOME"]] if presidente is not None else []) + selecionados
        for i, nome in enumerate(total_selecionados):
            linha = df.loc[df["NOME"] == nome].iloc[0]
            cat = linha.get("CATEGORIA", "")
            telefone = linha.get("TELEFONE", "")
            presidente_flag = "SIM" if presidente is not None and nome == presidente["NOME"] else "NAO"
            convocados.append({
                "DIA": dia, "DATA": data.date(), "MUNICIPIO": municipio,
                "NOME": nome, "CATEGORIA": cat, "PRESIDENTE": presidente_flag, "TELEFONE": telefone
            })

    df_conv = pd.DataFrame(convocados).drop_duplicates()

    # --- Aba de não convocados ---
    dias_validos = df["DIA"].unique()
    combinacoes = []
    for nome in df["NOME"].unique():
        for dia in dias_validos:
            combinacoes.append({"NOME": nome, "DIA": dia})
    base_completa = pd.DataFrame(combinacoes)
    df_base = df[["NOME", "CATEGORIA", "PRESIDENTE_DE_BANCA",
                  "DIAS_INDISPONIBILIDADE", "INICIO_INDISPONIBILIDADE", "FIM_INDISPONIBILIDADE", "TELEFONE"]].drop_duplicates(subset=["NOME"])

    # Filtra quem está indisponível durante todo o período
    data_inicio_op = df['DATA'].min()
    data_fim_op = df['DATA'].max()

    def esta_totalmente_indisp(row):
        dias = row.get("DIAS_INDISPONIBILIDADE","")
        if pd.notna(dias) and str(dias).strip() != "":
            dias_list = [d.strip().upper() for d in str(dias).split(",")]
            if set(dias_list) == set(dias_map.keys()):
                return True
        inicio = row.get("INICIO_INDISPONIBILIDADE")
        fim = row.get("FIM_INDISPONIBILIDADE")
        if pd.notna(inicio) and pd.notna(fim):
            if inicio.date() <= data_inicio_op.date() and fim.date() >= data_fim_op.date():
                return True
        return False

    df_base = df_base[~df_base.apply(esta_totalmente_indisp, axis=1)]
    df_nao = base_completa.merge(df_base, on="NOME", how="left")
    convocados_dias = df_conv.groupby(["NOME", "DIA"]).size().reset_index(name="FOI_CONVOCADO")
    df_nao = df_nao.merge(convocados_dias, on=["NOME", "DIA"], how="left")
    df_nao["FOI_CONVOCADO"].fillna(0, inplace=True)

    def em_ferias_ou_indisp(row):
        dia_num = dias_map.get(row["DIA"], None)
        if pd.notna(row["DIAS_INDISPONIBILIDADE"]) and str(row["DIAS_INDISPONIBILIDADE"]).strip():
            dias = [d.strip().upper().replace("Ç", "C").replace("Á", "A") for d in str(row["DIAS_INDISPONIBILIDADE"]).split(",")]
            dias_num_list = [dias_map[d] for d in dias if d in dias_map]
            if dia_num in dias_num_list:
                return True
        if pd.notna(row["INICIO_INDISPONIBILIDADE"]) and pd.notna(row["FIM_INDISPONIBILIDADE"]):
            try:
                data_exemplo = datetime.strptime("2025-01-01", "%Y-%m-%d")
                if row["INICIO_INDISPONIBILIDADE"].date() <= data_exemplo.date() <= row["FIM_INDISPONIBILIDADE"].date():
                    return True
            except:
                pass
        return False

    df_nao_final = df_nao[(df_nao["FOI_CONVOCADO"] == 0) & (~df_nao.apply(em_ferias_ou_indisp, axis=1))]
    df_nao_final = (
        df_nao_final.groupby(["NOME", "CATEGORIA", "PRESIDENTE_DE_BANCA"])["DIA"]
        .apply(lambda x: ", ".join(sorted(x)))
        .reset_index()
        .rename(columns={"PRESIDENTE_DE_BANCA": "PRESIDENTE", "DIA": "DIAS_NAO_CONVOCADOS"})
    )
    df_nao_final = df_nao_final[["NOME", "CATEGORIA", "PRESIDENTE", "DIAS_NAO_CONVOCADOS"]]

    # --- Criando planilha Excel ---
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Convocados"
    for r in dataframe_to_rows(df_conv, index=False, header=True):
        ws1.append(r)
    ws2 = wb.create_sheet("Nao Convocados")
    for r in dataframe_to_rows(df_nao_final, index=False, header=True):
        ws2.append(r)

    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return "Distribuicao_Completa.xlsx", df_conv, df_nao_final, buf, mensagens_vanessa

# ==============================
# 💻 Interface Streamlit
# ==============================
st.set_page_config(page_title="Distribuição 100% Completa", page_icon="⚖️", layout="centered")
st.markdown("""
<style>
.stApp {background: linear-gradient(135deg,#003c63,#015e78,#027b91);background-attachment:fixed;color:white;font-family:'Segoe UI',sans-serif;}
.main-card {background:rgba(255,255,255,0.08);border-radius:20px;padding:40px;box-shadow:0 8px 25px rgba(0,0,0,0.4);text-align:center;margin-top:40px;}
.main-card h1 {font-size:2.2rem;font-weight:700;color:#fff;}
.main-card p {font-size:1.1rem;color:#dcdcdc;}
.stButton button {background:linear-gradient(90deg,#00c6ff,#0072ff);color:white;border:none;border-radius:12px;padding:12px 25px;font-size:1rem;font-weight:bold;transition:0.3s;}
.stButton button:hover {transform:scale(1.05);}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-card"><h1>Distribuição de Pessoas</h1><p>Carregue seu Excel e gere a distribuição automática</p></div>', unsafe_allow_html=True)

arquivo = st.file_uploader("📁 Escolha o arquivo Excel", type=["xlsx", "xls"])

if arquivo:
    nome_arquivo, df_conv, df_nao, buf_excel, msgs_vanessa = processar_distribuicao(arquivo)
    st.success("✅ Distribuição gerada com sucesso!")
    st.markdown(f"📥 [Baixar Excel completo]({st.download_button(label='Download', data=buf_excel, file_name=nome_arquivo, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')})", unsafe_allow_html=True)
    if msgs_vanessa:
        for msg in msgs_vanessa:
            st.info(msg)

# ==============================
# 📲 Botão para enviar WhatsApp
# ==============================
def send_whatsapp_messages(df, headless=False, delay_between=3):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--user-data-dir=./user_data")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")
    st.info("⌛ Escaneie o QR Code se necessário e aguarde carregar a tela do WhatsApp Web...")
    time.sleep(15)  # Tempo para logar e carregar chats

    total_enviadas = 0
    for idx, row in df.iterrows():
        try:
            telefone = row.get("TELEFONE", "")
            mensagem = f"Olá {row.get('NOME','')}, você foi convocado(a) para {row.get('MUNICIPIO','')} em {row.get('DATA','')}"
            if telefone and telefone.isdigit():
                url = f"https://web.whatsapp.com/send?phone={telefone}&text={mensagem}"
                driver.get(url)
                time.sleep(delay_between)
                try:
                    btn = driver.find_element(By.XPATH, "//button[@data-testid='compose-btn-send']")
                    btn.click()
                    total_enviadas += 1
                except:
                    pass
            time.sleep(delay_between)
        except Exception as e:
            print(f"Erro ao enviar para {row.get('NOME')}: {e}")
    driver.quit()
    return total_enviadas

if arquivo and not df_conv.empty:
    st.subheader("📩 Enviar mensagens via WhatsApp")
    headless_mode = st.checkbox("Rodar em segundo plano (Headless)", value=False)
    delay_sec = st.number_input("Intervalo entre mensagens (segundos)", min_value=1, max_value=10, value=3, step=1)

    if st.button("📤 Enviar mensagens"):
        with st.spinner("⌛ Enviando mensagens... abra o WhatsApp Web se não estiver logado"):
            try:
                sent_total = send_whatsapp_messages(df_conv, headless=headless_mode, delay_between=delay_sec)
                st.success(f"✅ Mensagens enviadas com sucesso! Total: {sent_total}")
            except Exception as e:
                st.error(f"❌ Ocorreu um erro: {e}")
