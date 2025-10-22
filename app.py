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
# ⚙️ Configuração interna
# ==============================
META_CONVOCACOES_TENTATIVA = 3
META_CONVOCACOES_MINIMA = 2

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
# 🔍 Contagem de categorias compatíveis
# ==============================
def matching_count(categorias_pessoa, categorias_operacao):
    if not isinstance(categorias_pessoa, str) or not isinstance(categorias_operacao, str):
        return 0
    pessoa = [x.strip().upper() for x in categorias_pessoa.split(",") if x.strip()]
    oper = [x.strip().upper() for x in categorias_operacao.split(",") if x.strip()]
    return sum(1 for c in oper if c in pessoa)

# ==============================
# 🚫 Verificação de indisponibilidade / férias
# ==============================
dias_map = {"SEGUNDA":0, "TERCA":1, "QUARTA":2, "QUINTA":3, "SEXTA":4, "SABADO":5, "DOMINGO":6}
def esta_indisponivel(nome, dias_indisponiveis, inicio, fim, data):
    """
    ⚠️ Prioridade: DIAS_INDISPONIBILIDADE > INICIO/FIM.
    - Se o dia atual estiver na lista de dias específicos, a pessoa não pode ser chamada.
    - Caso contrário, ela pode ser chamada, mesmo dentro do período de indisponibilidade.
    """
    # Verifica dias específicos primeiro
    if pd.notna(dias_indisponiveis) and str(dias_indisponiveis).strip() != "":
        dias = [d.strip().upper().replace("Ç","C").replace("Á","A") for d in str(dias_indisponiveis).split(",")]
        dias_num = [dias_map[d] for d in dias if d in dias_map]
        if data.weekday() in dias_num:
            return True  # bloqueado
        else:
            return False  # liberado, mesmo dentro do período
    
    # Se não houver dias específicos, considera período completo
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
# 🔧 Evita repetir município na mesma semana (R2) ignorando presidentes
# ==============================
def filtrar_repeticao_municipio(df_candidatos, municipio, data, convocados):
    semana_atual = data.isocalendar()[1]
    municipio_semana = {(c["NOME"], c["MUNICIPIO"], c["DATA"].isocalendar()[1])
                        for c in convocados if c["PRESIDENTE"] != "SIM"}
    candidatos_filtrados = df_candidatos[~df_candidatos["NOME"].apply(lambda n: (n, municipio, semana_atual) in municipio_semana)].copy()
    return candidatos_filtrados

# ==============================
# 🔧 Peso de frequência semanal (R1)
# ==============================
def aplicar_regra_frequencia(df_candidatos, data, categoria_oper, conv_semana_global):
    semana = data.isocalendar()[1]
    def calcular_peso(r):
        nome = r["NOME"]
        conv_semana = conv_semana_global.get((nome, semana), 0)
        match = matching_count(r.get("CATEGORIA",""), categoria_oper)
        return (match * 10) + max(0, (META_CONVOCACOES_TENTATIVA - conv_semana))
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
    cont_conv = {n: 0 for n in nomes_unicos}
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

        # --- Evita duplicação no mesmo dia ---
        nomes_no_dia = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
        candidatos = candidatos[~candidatos["NOME"].isin(nomes_no_dia)].reset_index(drop=True)

        # --- Filtra repetição de município na mesma semana (R2) ---
        candidatos_filtrados_muni = filtrar_repeticao_municipio(candidatos, municipio, data, convocados)
        candidatos_para_usar = candidatos_filtrados_muni.copy()
        if len(candidatos_para_usar) < max(qtd, 1):
            candidatos_para_usar = candidatos.copy()

        # --- Calcula MATCH_COUNT e aplica regra Vanessa ---
        candidatos_para_usar["MATCH_COUNT"] = candidatos_para_usar["CATEGORIA"].apply(lambda c: matching_count(c, categoria_oper))
        candidatos_para_usar, vanessa_ativa = aplicar_regra_vanessa(candidatos_para_usar, categoria_oper, data)
        if vanessa_ativa:
            mensagens_vanessa.append(f"✨ Vanessa priorizada em {municipio} ({data.date()})")

        # --- Aplica regra de frequência (PESO) ---
        candidatos_pesados = aplicar_regra_frequencia(candidatos_para_usar, data, categoria_oper, conv_semana_global)

        # --- Seleciona presidente ---
        pres_cand = candidatos_pesados[candidatos_pesados["PRESIDENTE_DE_BANCA"].astype(str).str.upper() == "SIM"]
        pres_cand = pres_cand[~pres_cand["NOME"].isin(
            [c["NOME"] for c in convocados if c["DATA"].isocalendar()[1] == data.isocalendar()[1] and c["MUNICIPIO"] == municipio and c["PRESIDENTE"] == "SIM"]
        )]
        if not pres_cand.empty:
            nome_pres = sorted(pres_cand["NOME"].unique(), key=lambda n: cont_pres.get(n,0))[0]
            presidente = pres_cand[pres_cand["NOME"] == nome_pres].iloc[0]
        else:
            nome_pres = sorted(candidatos_pesados["NOME"].unique(), key=lambda n: cont_pres.get(n,0))[0]
            presidente = candidatos_pesados[candidatos_pesados["NOME"] == nome_pres].iloc[0]

        cont_pres[presidente["NOME"]] = cont_pres.get(presidente["NOME"], 0) + 1
        cont_conv[presidente["NOME"]] = cont_conv.get(presidente["NOME"], 0) + 1
        semana_pres = data.isocalendar()[1]
        conv_semana_global[(presidente["NOME"], semana_pres)] = conv_semana_global.get((presidente["NOME"], semana_pres), 0) + 1

        # --- Seleciona demais participantes ---
        pool = candidatos_pesados[candidatos_pesados["NOME"] != presidente["NOME"]].copy()
        pool["MATCH_COUNT"] = pool["MATCH_COUNT"].fillna(0)
        pool = pool.sort_values(by=["PESO"], ascending=False)

        selecionados = []
        for _, r in pool.iterrows():
            if len(selecionados) >= (qtd - 1):
                break
            nome = r["NOME"]
            if nome not in [c["NOME"] for c in convocados if c["DATA"] == data.date()]:
                selecionados.append(nome)
                cont_conv[nome] = cont_conv.get(nome, 0) + 1
                semana_sel = data.isocalendar()[1]
                conv_semana_global[(nome, semana_sel)] = conv_semana_global.get((nome, semana_sel), 0) + 1

        # --- FORÇAR PREENCHIMENTO CASO FALTEM VAGAS ---
        total_selecionados = [presidente["NOME"]] + selecionados
        faltantes = qtd - len(total_selecionados)
        if faltantes > 0:
            extras = candidatos[~candidatos["NOME"].isin(total_selecionados)].copy()
            extras = aplicar_regra_frequencia(extras, data, categoria_oper, conv_semana_global)
            for _, r in extras.iterrows():
                if faltantes == 0:
                    break
                total_selecionados.append(r["NOME"])
                cont_conv[r["NOME"]] = cont_conv.get(r["NOME"], 0) + 1
                semana_sel = data.isocalendar()[1]
                conv_semana_global[(r["NOME"], semana_sel)] = conv_semana_global.get((r["NOME"], semana_sel), 0) + 1
                faltantes -= 1

        # --- Adiciona ao resultado final ---
        for i, nome in enumerate(total_selecionados):
            cat = df.loc[df["NOME"] == nome, "CATEGORIA"].iloc[0]
            presidente_flag = "SIM" if nome == presidente["NOME"] else "NAO"
            convocados.append({
                "DIA": dia, "DATA": data.date(), "MUNICIPIO": municipio,
                "NOME": nome, "CATEGORIA": cat, "PRESIDENTE": presidente_flag
            })

    df_conv = pd.DataFrame(convocados).drop_duplicates()

    # --- Aba de não convocados (novo padrão) ---
    dias_nao_chamados = []
    for _, r in df.iterrows():
        nome = r["NOME"]
        df_chamados = df_conv[df_conv["NOME"] == nome]
        if df_chamados.empty:
            dias_nao_chamados.append({
                "NOME": nome,
                "DIA": r.get("DIA",""),
                "CATEGORIA": r.get("CATEGORIA",""),
                "MUNICIPIO_ORIGEM": r.get("MUNICIPIO_ORIGEM",""),
                "PRESIDENTE_DE_BANCA": r.get("PRESIDENTE_DE_BANCA","")
            })
        else:
            dias_faltantes = sorted(set(df["DIA"]) - set(df_chamados["DIA"]))
            for dia_falt in dias_faltantes:
                dias_nao_chamados.append({
                    "NOME": nome,
                    "DIA": dia_falt,
                    "CATEGORIA": r.get("CATEGORIA",""),
                    "MUNICIPIO_ORIGEM": r.get("MUNICIPIO_ORIGEM",""),
                    "PRESIDENTE_DE_BANCA": r.get("PRESIDENTE_DE_BANCA","")
                })

    df_nao = pd.DataFrame(dias_nao_chamados)

    # --- Criando planilha Excel ---
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
<p>O sistema garante sempre 100% das vagas preenchidas e evita convocações duplicadas no mesmo dia.</p>
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

            if msgs_vanessa:
                st.markdown("### 🌟 Regras Especiais")
                for m in msgs_vanessa:
                    st.markdown(f"- {m}")

            b64 = base64.b64encode(buf.read()).decode()
            st.markdown(f"""
            <div style="text-align:center;margin-top:30px;">
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{nome_saida}" target="_blank" style="background:linear-gradient(90deg,#00c6ff,#0072ff);padding:12px 25px;color:white;text-decoration:none;border-radius:12px;font-size:16px;font-weight:bold;">
            ⬇️ Baixar Excel
            </a></div>
            """, unsafe_allow_html=True)

