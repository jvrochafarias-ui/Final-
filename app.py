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
# 🔍 Compatibilidade de categorias
# ==============================
def matching_count_fallback(cat_pessoa, cat_oper):
    if not isinstance(cat_pessoa, str) or not isinstance(cat_oper, str):
        return 0

    pessoa = {x.strip().upper() for x in cat_pessoa.split(",") if x.strip()}
    oper = [x.strip().upper() for x in cat_oper.split(",") if x.strip()]

    if not oper:
        return 0

    primeira = oper[0]
    ultima = oper[-1]
    exige_e = "E" in oper

    if exige_e:
        if "E" in pessoa:
            return int(primeira in pessoa and ultima in pessoa)
        return 1
    return int(primeira in pessoa and ultima in pessoa)

# ==============================
# 🚫 Indisponibilidade / férias
# ==============================
dias_map = {
    "SEGUNDA": 0, "TERCA": 1, "QUARTA": 2,
    "QUINTA": 3, "SEXTA": 4, "SABADO": 5, "DOMINGO": 6
}

def esta_indisponivel(nome, dias, inicio, fim, data):
    if pd.notna(dias) and str(dias).strip():
        dias_list = [
            d.strip().upper().replace("Ç", "C").replace("Á", "A")
            for d in str(dias).split(",")
        ]
        if data.weekday() in [dias_map[d] for d in dias_list if d in dias_map]:
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
def aplicar_regra_vanessa(df, categoria, data):
    if str(categoria).strip().upper() == "B":
        nome = "VANESSA APARECIDA CARVALHO DE ASSIS"
        v = df[df["NOME"] == nome]
        if not v.empty:
            r = v.iloc[0]
            if not esta_indisponivel(
                r["NOME"],
                r.get("DIAS_INDISPONIBILIDADE", ""),
                r.get("INICIO_INDISPONIBILIDADE"),
                r.get("FIM_INDISPONIBILIDADE"),
                data
            ):
                resto = df[df["NOME"] != nome]
                return pd.concat([v, resto]), True
    return df, False

# ==============================
# 🔁 Evita repetição e município
# ==============================
def filtrar_candidatos(df, municipio, data, convocados):
    nomes_dia = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
    return df[
        (~df["NOME"].isin(nomes_dia)) &
        (df["MUNICIPIO_ORIGEM"] != municipio)
    ].copy()

# ==============================
# ⚖️ Regra de peso (BLINDADA)
# ==============================
def aplicar_regra_frequencia(df, data, categoria, conv_semana):
    if df.empty:
        df["PESO"] = []
        return df

    semana = data.isocalendar()[1]

    def peso(r):
        nome = r["NOME"]
        freq = conv_semana.get((nome, semana), 0)
        match = matching_count_fallback(r.get("CATEGORIA", ""), categoria)
        return (match * 10) + max(0, 5 - freq)

    df = df.copy()
    df["PESO"] = df.apply(peso, axis=1)

    # 🔒 BLINDAGEM DEFINITIVA
    if "PESO" not in df.columns:
        df["PESO"] = 0

    return df.sort_values("PESO", ascending=False)

# ==============================
# 🧠 PROCESSAMENTO PRINCIPAL
# ==============================
def processar_distribuicao(arquivo):
    df = normalizar_colunas(pd.read_excel(arquivo))

    if "PRESIDENTE_DE_BANCA" not in df.columns:
        df["PRESIDENTE_DE_BANCA"] = "NAO"

    for c in ["DATA", "INICIO_INDISPONIBILIDADE", "FIM_INDISPONIBILIDADE"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    df["DATA"].ffill(inplace=True)
    df["DIA"] = df["DATA"].dt.day_name().str.upper()
    df["QUANTIDADE"] = df.get("QUANTIDADE", 1).fillna(0).astype(int)

    convocados = []
    cont_pres = {}
    conv_semana = {}
    msgs_vanessa = []

    grupos = df.groupby(["DIA", "DATA", "MUNICIPIO", "CATEGORIA", "QUANTIDADE"])

    for (dia, data, municipio, categoria, qtd), _ in grupos:
        if qtd <= 0:
            continue

        candidatos = df[
            ~df.apply(
                lambda r: esta_indisponivel(
                    r["NOME"],
                    r.get("DIAS_INDISPONIBILIDADE", ""),
                    r.get("INICIO_INDISPONIBILIDADE"),
                    r.get("FIM_INDISPONIBILIDADE"),
                    data
                ), axis=1
            )
        ]

        candidatos = filtrar_candidatos(candidatos, municipio, data, convocados)
        candidatos, v_ok = aplicar_regra_vanessa(candidatos, categoria, data)

        if v_ok:
            msgs_vanessa.append(f"✨ Vanessa priorizada em {municipio} ({data.date()})")

        candidatos = aplicar_regra_frequencia(candidatos, data, categoria, conv_semana)

        if candidatos.empty:
            continue

        pool = candidatos.copy()

        if "PESO" not in pool.columns:
            pool["PESO"] = 0

        pool = pool.sort_values("PESO", ascending=False)

        nomes_dia = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
        selecionados = []

        for _, r in pool.iterrows():
            if len(selecionados) >= qtd:
                break
            if r["NOME"] not in nomes_dia:
                selecionados.append(r["NOME"])
                nomes_dia.append(r["NOME"])

        for nome in selecionados:
            convocados.append({
                "DIA": dia,
                "DATA": data.date(),
                "MUNICIPIO": municipio,
                "NOME": nome,
                "CATEGORIA": categoria,
                "PRESIDENTE": "NAO"
            })

    df_conv = pd.DataFrame(convocados)

    wb = Workbook()
    ws = wb.active
    ws.title = "Convocados"
    for r in dataframe_to_rows(df_conv, index=False, header=True):
        ws.append(r)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    return "Distribuicao_Completa.xlsx", df_conv, df_conv, buf, msgs_vanessa

# ==============================
# 💻 INTERFACE STREAMLIT
# ==============================

st.set_page_config("Distribuição 100% Completa", "⚖️")

st.title("⚖️ Distribuição Equilibrada e Completa")
arquivo = st.file_uploader("📁 Envie a planilha (.xlsx)", type="xlsx")

if arquivo and st.button("🔄 Gerar Distribuição Completa"):
    with st.spinner("Processando..."):
        nome, dfc, _, buf, msgs = processar_distribuicao(arquivo)
        st.success("✅ Distribuição gerada")
        st.dataframe(dfc, use_container_width=True)

        b64 = base64.b64encode(buf.read()).decode()
        st.markdown(
            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{nome}">⬇️ Baixar Excel</a>',
            unsafe_allow_html=True
        )
