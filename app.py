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
            return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
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
# 🔍 Contagem de categorias compatíveis (atualizado com fallback E)
# ==============================
def matching_count(categorias_pessoa, categorias_operacao):
    """
    Retorna 1 se o candidato tiver pelo menos a última categoria da operação
    e a primeira categoria também; caso contrário, retorna 0.
    """
    if not isinstance(categorias_pessoa, str) or not isinstance(categorias_operacao, str):
        return 0

    pessoa_set = set(x.strip().upper() for x in categorias_pessoa.split(",") if x.strip())
    oper_list = [x.strip().upper() for x in categorias_operacao.split(",") if x.strip()]

    if not oper_list:
        return 0

    primeira = oper_list[0]
    ultima = oper_list[-1]

    # Verifica última e depois primeira categoria
    if ultima in pessoa_set and primeira in pessoa_set:
        return 1
    return 0

def matching_count_fallback(categorias_pessoa, categorias_operacao):
    """
    Igual a matching_count, mas com fallback para E.
    """
    if not isinstance(categorias_pessoa, str) or not isinstance(categorias_operacao, str):
        return 0

    pessoa_set = set(x.strip().upper() for x in categorias_pessoa.split(",") if x.strip())
    oper_list = [x.strip().upper() for x in categorias_operacao.split(",") if x.strip()]

    if not oper_list:
        return 0

    primeira = oper_list[0]
    ultima = oper_list[-1]

    # Categoria E obrigatória
    precisa_E = "E" in oper_list

    if precisa_E:
        # Prioriza candidatos com E
        if "E" in pessoa_set:
            return 1 if ultima in pessoa_set and primeira in pessoa_set else 0
        else:
            # fallback: aceita mesmo sem E
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

        # --- Filtra indisponíveis / férias ---
        candidatos = df.copy()
        candidatos = candidatos.loc[~candidatos.apply(lambda r: esta_indisponivel(
            r["NOME"], r.get("DIAS_INDISPONIBILIDADE",""), r.get("INICIO_INDISPONIBILIDADE"), r.get("FIM_INDISPONIBILIDADE"), data
        ), axis=1)].reset_index(drop=True)

        # --- Filtra repetição no dia e município de origem ---
        candidatos = filtrar_candidatos(candidatos, municipio, data, convocados)

        # --- Calcula MATCH_COUNT e aplica regra Vanessa ---
        candidatos["MATCH_COUNT"] = candidatos["CATEGORIA"].apply(lambda c: matching_count_fallback(c, categoria_oper))
        candidatos, vanessa_ativa = aplicar_regra_vanessa(candidatos, categoria_oper, data)
        if vanessa_ativa:
            mensagens_vanessa.append(f"✨ Vanessa priorizada em {municipio} ({data.date()})")

        # --- Mantém apenas candidatos compatíveis com a operação ---
        candidatos_pesados = aplicar_regra_frequencia(candidatos, data, categoria_oper, conv_semana_global)

        # --- Seleciona presidente ---
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

        # --- Seleciona demais participantes ---
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

        # --- Adiciona ao resultado final ---
        total_selecionados = ([presidente["NOME"]] if presidente is not None else []) + selecionados
        for i, nome in enumerate(total_selecionados):
            cat = df.loc[df["NOME"] == nome, "CATEGORIA"].iloc[0]
            presidente_flag = "SIM" if presidente is not None and nome == presidente["NOME"] else "NAO"
            convocados.append({
                "DIA": dia, "DATA": data.date(), "MUNICIPIO": municipio,
                "NOME": nome, "CATEGORIA": cat, "PRESIDENTE": presidente_flag
            })

    df_conv = pd.DataFrame(convocados).drop_duplicates()

    # --- Aba de não convocados ---
    df_nao = df.copy()
    df_nao = df_nao[~df_nao["NOME"].isin(df_conv["NOME"])]
    df_nao = df_nao[["NOME", "DIA", "CATEGORIA", "PRESIDENTE_DE_BANCA"]].rename(columns={"PRESIDENTE_DE_BANCA":"PRESIDENTE"})

    # --- Criando planilha Excel ---
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Convocados"
    for r in dataframe_to_rows(df_conv, index=False, header=True):
        ws1.append(r)
    ws2 = wb.create_sheet("Nao Convocados")
    for r in dataframe_to_rows(df_nao, index=False, header=True):
        ws2.append(r)

    # Ajuste automático de largura das colunas
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
    return "Distribuicao_Completa.xlsx", df_conv, df_nao, buf, mensagens_vanessa

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
