import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import locale
import unicodedata

# ----------------------- Configuração de locale -----------------------
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.utf8")
except:
    locale.setlocale(locale.LC_TIME, "Portuguese_Brazil.1252")

# ----------------------- Função para normalizar colunas -----------------------
def normalizar_colunas(df):
    def remover_acentos(s):
        if isinstance(s, str):
            return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        return s

    df.columns = [remover_acentos(col).strip().upper() for col in df.columns]

    col_renomear = {
        "MUNICIPIO ORIGEM": "MUNICIPIO_ORIGEM",
        "PRESIDENTE DE BANCA": "PRESIDENTE_DE_BANCA",
        "INICIO INDISPONIBILIDADE": "INICIO_INDISPONIBILIDADE",
        "FIM INDISPONIBILIDADE": "FIM_INDISPONIBILIDADE"
    }
    for antiga, nova in col_renomear.items():
        if antiga in df.columns:
            df.rename(columns={antiga: nova}, inplace=True)

    def limpar_texto(s):
        if isinstance(s, str):
            s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
            s = s.strip().upper()
            return s
        return s

    if "MUNICIPIO" in df.columns:
        df["MUNICIPIO"] = df["MUNICIPIO"].apply(limpar_texto)
    if "MUNICIPIO_ORIGEM" in df.columns:
        df["MUNICIPIO_ORIGEM"] = df["MUNICIPIO_ORIGEM"].apply(limpar_texto)

    return df

# ----------------------- Função principal -----------------------
def processar_distribuicao(arquivo):
    df = pd.read_excel(arquivo)
    df = normalizar_colunas(df)

    # Colunas obrigatórias
    col_obrigatorias = ["NOME", "DIA", "DATA", "MUNICIPIO", "CATEGORIA", "QUANTIDADE",
                        "MUNICIPIO_ORIGEM", "PRESIDENTE_DE_BANCA"]
    col_faltando = [c for c in col_obrigatorias if c not in df.columns]
    if col_faltando:
        raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(col_faltando)}")

    # Tratar datas
    for col in ["DATA", "INICIO_INDISPONIBILIDADE", "FIM_INDISPONIBILIDADE"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    df["DATA"].fillna(method="ffill", inplace=True)
    df["DIA"].fillna(method="ffill", inplace=True)

    # Tratar QUANTIDADE
    df["QUANTIDADE"] = df["QUANTIDADE"].fillna(0).astype(int)

    operacoes = df.groupby(["DIA", "DATA", "MUNICIPIO", "CATEGORIA", "QUANTIDADE"], dropna=False)

    contagem_convocacoes = {nome: 0 for nome in df["NOME"].unique()}
    contagem_presidente = {nome: 0 for nome in df["NOME"].unique()}

    convocados = []
    nao_convocados = []

    # Função para verificar compatibilidade de categorias
    def possui_duas_categorias(categorias_pessoa, categorias_operacao):
        lista_pessoa = [x.strip() for x in categorias_pessoa.split(",")]
        lista_operacao = [x.strip() for x in categorias_operacao.split(",")]
        count = sum(1 for cat in lista_operacao if cat in lista_pessoa)
        return count >= 2

    for (dia, data, municipio, categoria_oper, qtd) in operacoes.groups:
        subset = df.copy()

        # ------------------ Regra 1: Não convocar mais de uma vez por dia ------------------
        nomes_convocados_no_dia = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
        subset = subset[~subset["NOME"].isin(nomes_convocados_no_dia)]
        # -----------------------------------------------------------------------------------

        # Respeitar indisponibilidade (Regra 4)
        subset = subset[~(
            (subset["INICIO_INDISPONIBILIDADE"].notna())
            & (subset["FIM_INDISPONIBILIDADE"].notna())
            & (subset["INICIO_INDISPONIBILIDADE"] <= data)
            & (subset["FIM_INDISPONIBILIDADE"] >= data)
        )]

        # Evitar pessoas do mesmo município de origem (Regra 5)
        subset = subset[subset["MUNICIPIO_ORIGEM"] != municipio]

        if qtd == 0 or subset.empty:
            for _, row in subset.iterrows():
                nao_convocados.append({
                    "NOME": row["NOME"],
                    "DIA": dia,
                    "CATEGORIA": row["CATEGORIA"],
                    "MUNICIPIO_ORIGEM": row["MUNICIPIO_ORIGEM"],
                    "PRESIDENTE_DE_BANCA": row["PRESIDENTE_DE_BANCA"],
                    "DATA": data.date()
                })
            continue

        # Compatibilidade mínima de categorias (Regra 3)
        subset = subset[subset["CATEGORIA"].apply(lambda c: possui_duas_categorias(c, categoria_oper))]

        # ------------------ Regra 9: Poá sexta-feira ------------------
        if municipio == "POA" and dia.upper() == "SEXTA" and qtd == 3:
            ops = categoria_oper.split(",")
            for op in ops:
                subset_op = subset[subset["CATEGORIA"].str.contains(op)]
                candidatos_pres = subset_op[subset_op["PRESIDENTE_DE_BANCA"].str.upper() == "SIM"]
                presidente = None
                if not candidatos_pres.empty:
                    presidente_nome = sorted(candidatos_pres["NOME"].unique(), key=lambda n: contagem_presidente[n])[0]
                    presidente = candidatos_pres[candidatos_pres["NOME"] == presidente_nome].iloc[0]

                if presidente is None:
                    menos_convocado = sorted(contagem_presidente.items(), key=lambda x: x[1])[0][0]
                    presidente_df = subset_op[subset_op["NOME"] == menos_convocado]
                    if not presidente_df.empty:
                        presidente = presidente_df.iloc[0]
                if presidente is None:
                    continue

                contagem_convocacoes[presidente["NOME"]] += 1
                contagem_presidente[presidente["NOME"]] += 1
                convocados.append({
                    "DIA": dia,
                    "DATA": data.date(),
                    "MUNICIPIO": municipio,
                    "NOME": presidente["NOME"],
                    "CATEGORIA": presidente["CATEGORIA"],
                    "PRESIDENTE": "SIM"
                })

                subset_rest = subset_op[subset_op["NOME"] != presidente["NOME"]].copy()
                subset_rest["CONV_COUNT"] = subset_rest["NOME"].map(lambda x: contagem_convocacoes.get(x, 0))
                subset_rest = subset_rest.sort_values(by="CONV_COUNT")
                selecionados = []
                for _, row in subset_rest.iterrows():
                    if contagem_convocacoes[row["NOME"]] < 3 or len(selecionados) < (qtd - 1):
                        contagem_convocacoes[row["NOME"]] += 1
                        selecionados.append(row["NOME"])
                    if len(selecionados) >= (qtd - 1):
                        break
                for nome in selecionados:
                    row_df = subset_rest[subset_rest["NOME"] == nome].iloc[0]
                    convocados.append({
                        "DIA": dia,
                        "DATA": data.date(),
                        "MUNICIPIO": municipio,
                        "NOME": nome,
                        "CATEGORIA": row_df["CATEGORIA"],
                        "PRESIDENTE": "NÃO"
                    })
            continue
        # -----------------------------------------------------------------------------------

        # ------------------ Seleção presidente padrão (Regra 2) ------------------
        candidatos_pres = subset[subset["PRESIDENTE_DE_BANCA"].str.upper() == "SIM"]
        presidente = None
        if not candidatos_pres.empty:
            presidente_nome = sorted(candidatos_pres["NOME"].unique(), key=lambda n: contagem_presidente[n])[0]
            presidente = candidatos_pres[candidatos_pres["NOME"] == presidente_nome].iloc[0]

        if presidente is None:
            menos_convocado = sorted(contagem_presidente.items(), key=lambda x: x[1])[0][0]
            presidente_df = subset[subset["NOME"] == menos_convocado]
            if not presidente_df.empty:
                presidente = presidente_df.iloc[0]
        if presidente is None:
            continue

        contagem_convocacoes[presidente["NOME"]] += 1
        contagem_presidente[presidente["NOME"]] += 1
        convocados.append({
            "DIA": dia,
            "DATA": data.date(),
            "MUNICIPIO": municipio,
            "NOME": presidente["NOME"],
            "CATEGORIA": presidente["CATEGORIA"],
            "PRESIDENTE": "SIM"
        })

        # Seleção dos demais convocados (Regra 7 e 8)
        subset = subset[subset["NOME"] != presidente["NOME"]].copy()
        subset["CONV_COUNT"] = subset["NOME"].map(lambda x: contagem_convocacoes.get(x, 0))
        subset = subset.sort_values(by="CONV_COUNT")

        selecionados = []
        semana_atual = data.isocalendar()[1]

        for _, row in subset.iterrows():
            nome = row["NOME"]
            if any(c["NOME"] == nome and c["MUNICIPIO"] == municipio and datetime.strptime(str(c["DATA"]), "%Y-%m-%d").isocalendar()[1] == semana_atual for c in convocados):
                continue
            if contagem_convocacoes[nome] < 3 or len(selecionados) < (qtd - 1):
                contagem_convocacoes[nome] += 1
                selecionados.append(nome)
            if len(selecionados) >= (qtd - 1):
                break

        for nome in selecionados:
            row_df = subset[subset["NOME"] == nome].iloc[0]
            convocados.append({
                "DIA": dia,
                "DATA": data.date(),
                "MUNICIPIO": municipio,
                "NOME": nome,
                "CATEGORIA": row_df["CATEGORIA"],
                "PRESIDENTE": "NÃO"
            })

        # Garantir número exato de convocados (Regra 6)
        total_previsto = int(qtd)
        convocados_no_dia_mun = [c for c in convocados if c["DATA"] == data.date() and c["MUNICIPIO"] == municipio]
        total_atual = len(convocados_no_dia_mun)

        if total_atual < total_previsto:
            faltam = total_previsto - total_atual
            subset_extra = subset[~subset["NOME"].isin([c["NOME"] for c in convocados_no_dia_mun])].copy()
            subset_extra["CONV_COUNT"] = subset_extra["NOME"].map(lambda x: contagem_convocacoes.get(x, 0))
            subset_extra = subset_extra.sort_values(by="CONV_COUNT")
            for _, row in subset_extra.head(faltam).iterrows():
                convocados.append({
                    "DIA": dia,
                    "DATA": data.date(),
                    "MUNICIPIO": municipio,
                    "NOME": row["NOME"],
                    "CATEGORIA": row["CATEGORIA"],
                    "PRESIDENTE": "NÃO"
                })
                contagem_convocacoes[row["NOME"]] += 1
        elif total_atual > total_previsto:
            excedente = total_atual - total_previsto
            for _ in range(excedente):
                convocados.pop()

        # Não convocados restantes (Regra 10)
        nomes_ja_nao_convocados = [x["NOME"] for x in nao_convocados if x["DATA"] == data.date()]
        disponiveis_no_dia = df[(df["MUNICIPIO_ORIGEM"] != municipio)
                                & (df["CATEGORIA"].apply(lambda c: possui_duas_categorias(c, categoria_oper)))]
        nomes_convocados = [c["NOME"] for c in convocados if c["DATA"] == data.date()]
        for _, row in disponiveis_no_dia.iterrows():
            if row["NOME"] not in nomes_convocados and row["NOME"] not in nomes_ja_nao_convocados:
                nao_convocados.append({
                    "NOME": row["NOME"],
                    "DIA": dia,
                    "CATEGORIA": row["CATEGORIA"],
                    "MUNICIPIO_ORIGEM": row["MUNICIPIO_ORIGEM"],
                    "PRESIDENTE_DE_BANCA": row["PRESIDENTE_DE_BANCA"],
                    "DATA": data.date()
                })

    # DataFrames finais
    df_convocados = pd.DataFrame(convocados).drop_duplicates()
    df_nao_convocados = pd.DataFrame(nao_convocados).drop_duplicates(subset=["NOME", "DIA", "CATEGORIA"])

    # Exportação
    mes_nome = datetime.now().strftime("%B").lower()
    nome_saida = f"distribuicao_{mes_nome}.xlsx"

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Convocados"
    for r in dataframe_to_rows(df_convocados, index=False, header=True):
        ws1.append(r)

    ws2 = wb.create_sheet("Nao Convocados")
    for r in dataframe_to_rows(df_nao_convocados, index=False, header=True):
        ws2.append(r)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return nome_saida, df_convocados, df_nao_convocados, buffer

# ----------------------- Interface Streamlit -----------------------
st.set_page_config(page_title="Distribuição Equilibrada", page_icon="📊", layout="centered")

# Fundo e estilo
page_bg = """
<style>
.stApp {
    background: linear-gradient(135deg, #002b45, #014d63, #028090);
    background-attachment: fixed;
    color: white;
    font-family: 'Segoe UI', sans-serif;
}
.main-card {
    background: rgba(255, 255, 255, 0.08);
    border-radius: 20px;
    padding: 40px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.4);
    text-align: center;
    margin-top: 40px;
}
.main-card h1 {
    font-size: 2.2rem;
    font-weight: 700;
    color: #ffffff;
    margin-bottom: 15px;
}
.main-card p {
    font-size: 1.1rem;
    color: #dcdcdc;
    margin-bottom: 30px;
}
.stButton button {
    background: linear-gradient(90deg, #00c6ff, #0072ff);
    color: white;
    border: none;
    border-radius: 12px;
    padding: 12px 25px;
    font-size: 1rem;
    font-weight: bold;
    transition: 0.3s;
}
.stButton button:hover {
    transform: scale(1.05);
    background: linear-gradient(90deg, #0072ff, #00c6ff);
}
</style>
"""
st.markdown(page_bg, unsafe_allow_html=True)

# Cartão de introdução
st.markdown("""
<div class="main-card">
<h1>📊 Distribuição Equilibrada de Convocações</h1>
<p>O sistema distribui as convocações respeitando todas as regras, garantindo 1 presidente por operação, compatibilidade mínima de categorias, limite de convocações, indisponibilidade e municípios de origem.</p>
</div>
""", unsafe_allow_html=True)

# Upload do arquivo
arquivo = st.file_uploader("📁 Envie a planilha (.xlsx)", type="xlsx")

if arquivo:
    try:
        st.markdown("### ⚙️ Processamento")
        st.info("Clique no botão abaixo para gerar a distribuição equilibrada.")
        if st.button("🔄 Gerar Distribuição"):
            with st.spinner("Processando..."):
                nome_saida, df_convocados, df_nao_convocados, arquivo_excel = processar_distribuicao(arquivo)

                if df_convocados.empty and df_nao_convocados.empty:
                    st.error("⚠️ Não foi possível gerar a distribuição. Verifique a planilha enviada.")
                else:
                    st.success("✅ Distribuição gerada com sucesso!")

                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("### 👥 Convocados")
                        st.dataframe(df_convocados, use_container_width=True)
                    with col2:
                        st.markdown("### 🚫 Não Convocados")
                        st.dataframe(df_nao_convocados, use_container_width=True)

                    b64 = base64.b64encode(arquivo_excel.read()).decode()
                    st.markdown(f"""
                    <div style="text-align:center; margin-top:30px;">
                    <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
                    download="{nome_saida}" target="_blank"
                    style="background:linear-gradient(90deg, #00c6ff, #0072ff); padding:12px 25px;
                    color:white; text-decoration:none; border-radius:12px; font-size:16px; font-weight:bold;">
                    ⬇️ Baixar Excel
                    </a></div>""", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
