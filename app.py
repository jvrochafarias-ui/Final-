import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import locale
import unicodedata

# ==============================
# ⚙️ Configuração de locale
# ==============================
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
            return ''.join(c for c in unicodedata.normalize('NFD', s)
                           if unicodedata.category(c) != 'Mn')
        return s

    df.columns = [remover_acentos(col).strip().upper() for col in df.columns]

    renomear = {
        "MUNICIPIO ORIGEM": "MUNICIPIO_ORIGEM",
        "PRESIDENTE DE BANCA": "PRESIDENTE_DE_BANCA",
        "INICIO INDISPONIBILIDADE": "INICIO_INDISPONIBILIDADE",
        "FIM INDISPONIBILIDADE": "FIM_INDISPONIBILIDADE",
        "DIAS INDISPONIBILIDADE": "DIAS_INDISPONIBILIDADE"
    }
    for antiga, nova in renomear.items():
        if antiga in df.columns:
            df.rename(columns={antiga: nova}, inplace=True)

    for col in ["MUNICIPIO", "MUNICIPIO_ORIGEM", "CATEGORIA", "NOME"]:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(lambda s: remover_acentos(s).strip().upper())

    return df

# ==============================
# 🔍 Contagem de categorias compatíveis
# ==============================
def matching_count(categorias_pessoa, categorias_operacao):
    if not isinstance(categorias_pessoa, str) or not isinstance(categorias_operacao, str):
        return 0
    pessoa = [x.strip().upper() for x in categorias_pessoa.split(",") if x.strip()]
    oper = [x.strip().upper() for x in categorias_operacao.split(",") if x.strip()]
    return sum(1 for c in oper if c in pessoa)

# ==============================
# 🚫 Verificações de indisponibilidade
# ==============================
dias_map = {"SEGUNDA":0, "TERCA":1, "QUARTA":2, "QUINTA":3,
            "SEXTA":4, "SABADO":5, "DOMINGO":6}

def esta_indisponivel_fixa(nome, dias_indisponiveis, data):
    """Retorna True se a pessoa não pode ser convocada no dia fixo."""
    if pd.isna(dias_indisponiveis) or dias_indisponiveis.strip() == "":
        return False
    dias = [d.strip().upper().replace("Ç","C").replace("Á","A") for d in dias_indisponiveis.split(",")]
    dias_num = [dias_map[d] for d in dias if d in dias_map]
    return data.weekday() in dias_num


def esta_em_ferias(inicio, fim, data):
    """Retorna True se a data estiver dentro do período de férias."""
    if pd.isna(inicio) or pd.isna(fim):
        return False
    try:
        inicio = pd.to_datetime(inicio, errors="coerce")
        fim = pd.to_datetime(fim, errors="coerce")
        if pd.isna(inicio) or pd.isna(fim):
            return False
        return inicio.date() <= data.date() <= fim.date()
    except Exception:
        return False

# ==============================
# 🧠 Processamento principal
# ==============================
def processar_distribuicao(arquivo):
    df = pd.read_excel(arquivo)
    df = normalizar_colunas(df)

    for col in ["DATA", "INICIO_INDISPONIBILIDADE", "FIM_INDISPONIBILIDADE"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    df["DATA"].fillna(method="ffill", inplace=True)
    df["DIA"].fillna(method="ffill", inplace=True)
    if "QUANTIDADE" not in df.columns:
        df["QUANTIDADE"] = 1
    df["QUANTIDADE"] = df["QUANTIDADE"].fillna(0).astype(int)

    cont_conv = {n:0 for n in df["NOME"].unique()}
    cont_pres = {n:0 for n in df["NOME"].unique()}
    convocados, nao_conv = [], []

    # Agrupar por operação
    operacoes = df.groupby(["DIA","DATA","MUNICIPIO","CATEGORIA","QUANTIDADE"], dropna=False)

    for (dia, data, municipio, categoria_oper, qtd), _ in operacoes:
        qtd = int(qtd)
        if qtd <= 0:
            continue

        # Candidatos de outros municípios
        candidatos = df[df["MUNICIPIO_ORIGEM"] != municipio].copy().reset_index(drop=True)

        # ✅ Aplicar regras de indisponibilidade (fixa e por férias)
        candidatos = candidatos.loc[~candidatos.apply(
            lambda r: (
                esta_indisponivel_fixa(r["NOME"], r.get("DIAS_INDISPONIBILIDADE",""), data)
                or esta_em_ferias(r.get("INICIO_INDISPONIBILIDADE"), r.get("FIM_INDISPONIBILIDADE"), data)
            ),
            axis=1
        )].reset_index(drop=True)

        if candidatos.empty:
            continue

        # Compatibilidade de categorias
        candidatos["MATCH_COUNT"] = candidatos["CATEGORIA"].apply(lambda c: matching_count(c, categoria_oper))

        # ✅ Regra especial: se a operação precisa de apenas 1 categoria, aceitar quem tiver 1
        if len(categoria_oper.split(",")) == 1:
            candidatos_validos = candidatos[candidatos["MATCH_COUNT"] >= 1].reset_index(drop=True)
        else:
            candidatos_validos = candidatos[candidatos["MATCH_COUNT"] >= 2].reset_index(drop=True)

        if candidatos_validos.empty:
            continue

        candidatos = candidatos_validos

        # Seleção presidente
        pres_cand = candidatos[candidatos["PRESIDENTE_DE_BANCA"].astype(str).str.upper() == "SIM"]
        if not pres_cand.empty:
            nome_pres = sorted(pres_cand["NOME"].unique(), key=lambda n: cont_pres[n])[0]
            presidente = pres_cand[pres_cand["NOME"] == nome_pres].iloc[0]
        else:
            nome_pres = sorted(candidatos["NOME"].unique(), key=lambda n: cont_pres[n])[0]
            presidente = candidatos[candidatos["NOME"] == nome_pres].iloc[0]

        # ❌ Verifica se já foi convocado no mesmo dia
        if any(c["NOME"] == presidente["NOME"] and c["DATA"] == data.date() for c in convocados):
            candidatos_restantes = candidatos[candidatos["NOME"] != presidente["NOME"]].reset_index(drop=True)
            if candidatos_restantes.empty:
                continue
            presidente = candidatos_restantes.iloc[0]

        cont_conv[presidente["NOME"]] += 1
        cont_pres[presidente["NOME"]] += 1
        convocados.append({
            "DIA": dia, "DATA": data.date(), "MUNICIPIO": municipio,
            "NOME": presidente["NOME"], "CATEGORIA": presidente["CATEGORIA"],
            "PRESIDENTE": "SIM"
        })

        # Auxiliares
        pool_aux = candidatos[candidatos["NOME"] != presidente["NOME"]].copy()
        pool_aux["CONV_COUNT"] = pool_aux["NOME"].map(cont_conv)
        pool_aux = pool_aux.sort_values(by="CONV_COUNT").reset_index(drop=True)

        selecionados = []
        semana = data.isocalendar()[1]
        for _, r in pool_aux.iterrows():
            nome = r["NOME"]
            # ❌ Verifica se já foi convocado no mesmo dia
            if any(c["NOME"]==nome and c["DATA"]==data.date() for c in convocados):
                continue
            ja_mesmo_mun = any(
                c["NOME"] == nome and c["MUNICIPIO"] == municipio
                and c["DATA"].isocalendar()[1] == semana
                and c["PRESIDENTE"] == "NAO"
                for c in convocados
            )
            if ja_mesmo_mun:
                continue
            if cont_conv[nome] < 3 or len(selecionados) < (qtd-1):
                cont_conv[nome] += 1
                selecionados.append(nome)
            if len(selecionados) >= (qtd-1):
                break

        for nome in selecionados:
            linha = pool_aux[pool_aux["NOME"]==nome].iloc[0]
            convocados.append({
                "DIA": dia, "DATA": data.date(), "MUNICIPIO": municipio,
                "NOME": nome, "CATEGORIA": linha["CATEGORIA"], "PRESIDENTE": "NAO"
            })

        # Registro não convocados
        todos_nomes = [c["NOME"] for c in convocados]
        eliminados = df[~df["NOME"].isin(todos_nomes)]
        for _, r in eliminados.iterrows():
            motivo = ""
            if matching_count(r["CATEGORIA"], categoria_oper) < 2 and len(categoria_oper.split(",")) > 1:
                motivo = "Incompativel"
            elif r["MUNICIPIO_ORIGEM"] == municipio:
                motivo = "Mesmo municipio"
            else:
                motivo = "Indisponivel"
            nao_conv.append({
                "NOME": r["NOME"], "DIA": dia, "CATEGORIA": r["CATEGORIA"],
                "MUNICIPIO_ORIGEM": r["MUNICIPIO_ORIGEM"],
                "PRESIDENTE_DE_BANCA": r.get("PRESIDENTE_DE_BANCA",""),
                "DATA": data.date(), "MOTIVO": motivo
            })

    df_conv = pd.DataFrame(convocados).drop_duplicates()
    df_nao = pd.DataFrame(nao_conv).drop_duplicates(subset=["NOME","DIA","CATEGORIA"])

    # Excel
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Convocados"
    for r in dataframe_to_rows(df_conv, index=False, header=True):
        ws1.append(r)
    ws2 = wb.create_sheet("Nao Convocados")
    for r in dataframe_to_rows(df_nao, index=False, header=True):
        ws2.append(r)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return "Distribuicao_Convocacoes.xlsx", df_conv, df_nao, buf

# ==============================
# Interface Streamlit
# ==============================
st.set_page_config(page_title="Distribuição Equilibrada", page_icon="⚖️", layout="centered")

st.markdown("""
<style>
.stApp {background: linear-gradient(135deg, #003c63, #015e78, #027b91);
background-attachment: fixed; color: white; font-family: 'Segoe UI', sans-serif;}
.main-card {background: rgba(255,255,255,0.08); border-radius:20px; padding:40px;
box-shadow: 0 8px 25px rgba(0,0,0,0.4); text-align:center; margin-top:40px;}
.main-card h1 {font-size:2.2rem; font-weight:700; color:#fff;}
.main-card p {font-size:1.1rem; color:#dcdcdc;}
.stButton button {background: linear-gradient(90deg,#00c6ff,#0072ff);
color:white; border:none; border-radius:12px; padding:12px 25px;
font-size:1rem; font-weight:bold; transition:0.3s;}
.stButton button:hover {transform:scale(1.05); background:linear-gradient(90deg,#0072ff,#00c6ff);}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-card">
<h1>⚖️ Distribuição Equilibrada de Convocações</h1>
<p>O sistema aplica todas as regras de convocação, respeitando indisponibilidades absolutas de dias, períodos de férias, limite de convocações, compatibilidade mínima de categorias e equilíbrio semanal.</p>
</div>
""", unsafe_allow_html=True)

arquivo = st.file_uploader("📁 Envie a planilha (.xlsx)", type="xlsx")

if arquivo:
    st.markdown("### ⚙️ Processamento")
    if st.button("🔄 Gerar Distribuição"):
        with st.spinner("Processando..."):
            try:
                nome_saida, df_conv, df_nao, buf = processar_distribuicao(arquivo)
                st.success("✅ Distribuição gerada com sucesso!")

                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("### 👥 Convocados")
                    st.dataframe(df_conv, use_container_width=True)
                with col2:
                    st.markdown("### 🚫 Não Convocados")
                    st.dataframe(df_nao, use_container_width=True)

                b64 = base64.b64encode(buf.read()).decode()
                st.markdown(f"""
                <div style="text-align:center; margin-top:30px;">
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
                download="{nome_saida}" target="_blank"
                style="background:linear-gradient(90deg,#00c6ff,#0072ff); padding:12px 25px;
                color:white; text-decoration:none; border-radius:12px; font-size:16px; font-weight:bold;">
                ⬇️ Baixar Excel
                </a></div>""", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"❌ Erro ao processar: {e}")
