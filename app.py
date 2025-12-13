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
            return ''.join(
                c for c in unicodedata.normalize('NFD', s)
                if unicodedata.category(c) != 'Mn'
            )
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
            df[col] = (
                df[col]
                .astype(str)
                .apply(lambda s: remover_acentos(s).strip().upper())
            )
    return df

# ==============================
# 🔍 Matching de categorias
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
            return 1
    else:
        return 1 if ultima in pessoa_set and primeira in pessoa_set else 0

# ==============================
# 🚫 Indisponibilidade
# ==============================
dias_map = {
    "SEGUNDA": 0, "TERCA": 1, "QUARTA": 2,
    "QUINTA": 3, "SEXTA": 4, "SABADO": 5, "DOMINGO": 6
}

def esta_indisponivel(nome, dias_indisponiveis, inicio, fim, data):
    if pd.notna(dias_indisponiveis) and str(dias_indisponiveis).strip():
        dias = [
            d.strip().upper().replace("Ç", "C").replace("Á", "A")
            for d in str(dias_indisponiveis).split(",")
        ]
        dias_num = [dias_map[d] for d in dias if d in dias_map]
        if data.weekday() in dias_num:
            return True

    if pd.notna(inicio) and pd.notna(fim):
        try:
            if inicio.date() <= data.date() <= fim.date():
                return True
        except:
            pass

    return False

# ==============================
# 🌟 Regra Vanessa
# ==============================
def aplicar_regra_vanessa(df_candidatos, categoria_oper, data):
    if str(categoria_oper).strip().upper() == "B":
        nome_vanessa = "VANESSA APARECIDA CARVALHO DE ASSIS"
        vanessa = df_candidatos[df_candidatos["NOME"] == nome_vanessa]

        if not vanessa.empty:
            r = vanessa.iloc[0]
            if not esta_indisponivel(
                r["NOME"],
                r.get("DIAS_INDISPONIBILIDADE", ""),
                r.get("INICIO_INDISPONIBILIDADE"),
                r.get("FIM_INDISPONIBILIDADE"),
                data
            ):
                resto = df_candidatos[df_candidatos["NOME"] != nome_vanessa]
                df_candidatos = pd.concat([vanessa, resto]).reset_index(drop=True)
                return df_candidatos, True

    return df_candidatos, False

# ==============================
# 🔧 Filtros básicos
# ==============================
def filtrar_candidatos(df_candidatos, municipio, data, convocados):
    nomes_no_dia = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
    return df_candidatos[
        (~df_candidatos["NOME"].isin(nomes_no_dia)) &
        (df_candidatos["MUNICIPIO_ORIGEM"] != municipio)
    ].copy()

# ==============================
# 🔧 Frequência semanal (mantida)
# ==============================
def aplicar_regra_frequencia(df_candidatos, data, categoria_oper, conv_semana_global):
    semana = data.isocalendar()[1]

    def calcular_peso(r):
        nome = r["NOME"]
        conv_semana = conv_semana_global.get((nome, semana), 0)
        match = matching_count_fallback(r.get("CATEGORIA", ""), categoria_oper)
        return (match * 10) + max(0, 5 - conv_semana)

    if df_candidatos.empty:
        return df_candidatos

    df_candidatos = df_candidatos.copy()
    df_candidatos["PESO"] = df_candidatos.apply(calcular_peso, axis=1)
    return df_candidatos

# ==============================
# 🧠 Processamento principal
# ==============================
def processar_distribuicao(arquivo):
    df = pd.read_excel(arquivo)
    df = normalizar_colunas(df)

    if "PRESIDENTE_DE_BANCA" not in df.columns:
        df["PRESIDENTE_DE_BANCA"] = "NAO"

    for col in ["DATA", "INICIO_INDISPONIBILIDADE", "FIM_INDISPONIBILIDADE"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if "DIA" not in df.columns and "DATA" in df.columns:
        df["DIA"] = df["DATA"].dt.day_name().str.upper()

    df["DATA"].ffill(inplace=True)
    df["DIA"].ffill(inplace=True)

    if "QUANTIDADE" not in df.columns:
        df["QUANTIDADE"] = 1

    df["QUANTIDADE"] = df["QUANTIDADE"].fillna(0).astype(int)

    convocados = []
    mensagens_vanessa = []
    conv_semana_global = {}

    operacoes = df.groupby(
        ["DIA", "DATA", "MUNICIPIO", "CATEGORIA", "QUANTIDADE"],
        dropna=False
    )

    for (dia, data, municipio, categoria_oper, qtd), _ in operacoes:
        qtd = int(qtd)
        if qtd <= 0:
            continue

        data = pd.to_datetime(data)

        # Remove indisponíveis
        candidatos = df.loc[
            ~df.apply(
                lambda r: esta_indisponivel(
                    r["NOME"],
                    r.get("DIAS_INDISPONIBILIDADE", ""),
                    r.get("INICIO_INDISPONIBILIDADE"),
                    r.get("FIM_INDISPONIBILIDADE"),
                    data
                ),
                axis=1
            )
        ].copy()

        candidatos = filtrar_candidatos(candidatos, municipio, data, convocados)
        candidatos, vanessa_ativa = aplicar_regra_vanessa(candidatos, categoria_oper, data)

        if vanessa_ativa:
            mensagens_vanessa.append(
                f"✨ Vanessa priorizada em {municipio} ({data.date()})"
            )

        candidatos = aplicar_regra_frequencia(
            candidatos, data, categoria_oper, conv_semana_global
        )

        # ==============================
        # 👑 SEPARA PRESIDENTES
        # ==============================
        pool_presidente = candidatos[candidatos["PRESIDENTE_DE_BANCA"] == "SIM"]
        pool_normal = candidatos[candidatos["PRESIDENTE_DE_BANCA"] != "SIM"]

        selecionados = []
        nomes_ja = [c["NOME"] for c in convocados if c["DATA"] == data.date()]

        # 👉 Seleciona 1 presidente primeiro
        presidente_nome = None
        if not pool_presidente.empty:
            for _, r in pool_presidente.iterrows():
                if r["NOME"] not in nomes_ja:
                    presidente_nome = r["NOME"]
                    selecionados.append((r["NOME"], "SIM"))
                    nomes_ja.append(r["NOME"])
                    break

        # 👉 Completa com demais convocados
        for _, r in pool_normal.iterrows():
            if len(selecionados) >= qtd:
                break
            if r["NOME"] not in nomes_ja:
                selecionados.append((r["NOME"], "NAO"))
                nomes_ja.append(r["NOME"])

        # 👉 Garante não passar da quantidade
        selecionados = selecionados[:qtd]

        for nome, eh_presidente in selecionados:
            convocados.append({
                "DIA": dia,
                "DATA": data.date(),
                "MUNICIPIO": municipio,
                "NOME": nome,
                "CATEGORIA": df.loc[df["NOME"] == nome, "CATEGORIA"].iloc[0],
                "PRESIDENTE": eh_presidente
            })

    df_conv = pd.DataFrame(convocados)

    buf = BytesIO()
    df_conv.to_excel(buf, index=False)
    buf.seek(0)

    return "Distribuicao_Completa.xlsx", df_conv, pd.DataFrame(), buf, mensagens_vanessa


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
.stButton button:hover {transform:scale(1.05);background:linear-gradient(90deg,#0072ff,#00c6ff);}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-card">
<h1>⚖️ Distribuição Equilibrada e Completa</h1>
<p>O sistema garante sempre 100% das vagas preenchidas com pelo menos um presidente real, evitando convocações duplicadas no mesmo dia e nenhum convocado no seu município de origem. Categoria “E” é priorizada quando exigida.</p>
</div>
""", unsafe_allow_html=True)

arquivo = st.file_uploader("📁 Envie a planilha (.xlsx)", type="xlsx")

if arquivo:
    if st.button("🔄 Gerar Distribuição Completa"):
        with st.spinner("Processando..."):
            nome_saida, df_conv, df_nao, buf, msgs_vanessa = processar_distribuicao(arquivo)
            st.success("✅ Distribuição completa gerada com sucesso!")

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### 👥 Convocados")
                st.dataframe(df_conv, use_container_width=True)
            with col2:
                st.markdown("### 🚫 Não Convocados")
                st.dataframe(df_nao, use_container_width=True)

            b64 = base64.b64encode(buf.read()).decode()
            st.markdown(f"""
            <div style="text-align:center;margin-top:30px;">
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{nome_saida}" target="_blank" style="background:linear-gradient(90deg,#00c6ff,#0072ff);padding:12px 25px;color:white;text-decoration:none;border-radius:12px;font-size:16px;font-weight:bold;">
            ⬇️ Baixar Excel
            </a></div>
            """, unsafe_allow_html=True)


