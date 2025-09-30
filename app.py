import streamlit as st
import pandas as pd
from datetime import datetime
import unicodedata
import io
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import random

# ------------------------
# Funções auxiliares
# ------------------------
def normalizar_coluna(col):
    col = str(col).strip().upper()
    col = unicodedata.normalize('NFKD', col).encode('ASCII', 'ignore').decode('ASCII')
    col = col.replace(" ", "_")
    return col

def normalizar_texto(txt):
    if pd.isna(txt):
        return ""
    txt = str(txt).strip().upper()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')
    return txt

def esta_disponivel(row, data):
    """
    Verifica se a pessoa está disponível na data desejada.
    Bloqueia qualquer pessoa que tenha INICIO e FIM de indisponibilidade que inclua a data.
    """
    if pd.isna(data):
        return True

    data = pd.to_datetime(data).normalize()

    inicio = row.get('INICIO_INDISPONIBILIDADE', pd.NaT)
    fim = row.get('FIM_INDISPONIBILIDADE', pd.NaT)

    if str(inicio).strip().upper() == 'SIM':
        return False

    try:
        if pd.notna(inicio):
            inicio = pd.to_datetime(inicio, dayfirst=True).normalize()
        if pd.notna(fim):
            fim = pd.to_datetime(fim, dayfirst=True).normalize()
    except Exception:
        return True

    if pd.notna(inicio) and pd.notna(fim):
        if inicio <= data <= fim:
            return False

    return True

# ------------------------
# Função para alocar candidatos em uma operação
# ------------------------
def alocar_operacao(candidatos_op, quantidade, presidentes_ja_convocados):
    """
    Seleciona candidatos para a operação garantindo:
      - exatamente 'quantidade' pessoas
      - pelo menos 1 presidente
    """
    if candidatos_op.empty:
        return pd.DataFrame(), None

    presidentes_disponiveis = candidatos_op[candidatos_op['PRESIDENTE_DE_BANCA'].str.upper() == 'SIM']

    if not presidentes_disponiveis.empty:
        presidente_selecionado = presidentes_disponiveis.sample(
            1, random_state=random.randint(0, 10000)
        )
        restantes = candidatos_op[~candidatos_op['NOME'].isin(presidente_selecionado['NOME'])] \
            .sample(max(0, quantidade - 1), random_state=random.randint(0, 10000))
        selecionados = pd.concat([presidente_selecionado, restantes])
    else:
        selecionados = candidatos_op.sample(
            min(quantidade, len(candidatos_op)), random_state=random.randint(0, 10000)
        )

    if len(selecionados) > quantidade:
        selecionados = selecionados.sample(quantidade, random_state=random.randint(0, 10000))

    presidentes = selecionados[selecionados['PRESIDENTE_DE_BANCA'].str.upper() == 'SIM']
    presidente_nome = None
    for p in presidentes['NOME']:
        if p not in presidentes_ja_convocados:
            presidente_nome = p
            break
    if presidente_nome is None and not presidentes.empty:
        presidente_nome = presidentes.iloc[0]['NOME']
    if presidente_nome:
        presidentes_ja_convocados.add(presidente_nome)

    return selecionados, presidente_nome

# ------------------------
# Processamento da distribuição
# ------------------------
def processar_distribuicao(arquivo_excel):
    xls = pd.ExcelFile(arquivo_excel)
    sheet_name = 'Planilha1' if 'Planilha1' in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)

    df.columns = [normalizar_coluna(col) for col in df.columns]

    colunas_possiveis_nome = ['NOME', 'NOME_COMPLETO', 'NOME_PESSOA']
    for col in colunas_possiveis_nome:
        if col in df.columns:
            df['NOME'] = df[col]
            break
    if 'NOME' not in df.columns:
        st.error(f"❌ Erro: não foi possível localizar a coluna de nomes. Colunas disponíveis: {df.columns.tolist()}")
        return None, pd.DataFrame(), pd.DataFrame(), io.BytesIO()

    df['INDISPONIBILIDADE'] = df.get('INDISPONIBILIDADE', pd.Series("NAO")).fillna("NAO")
    df['PRESIDENTE_DE_BANCA'] = df.get('PRESIDENTE_DE_BANCA', pd.Series("NAO")).fillna("NAO")
    df['MUNICIPIO_ORIGEM'] = df.get('MUNICIPIO_ORIGEM', pd.Series("")).fillna("")
    df['INICIO_INDISPONIBILIDADE'] = df.get('INICIO_INDISPONIBILIDADE', pd.NaT)
    df['FIM_INDISPONIBILIDADE'] = df.get('FIM_INDISPONIBILIDADE', pd.NaT)

    distribuicoes = []
    contador_convocacoes = {nome: 0 for nome in df['NOME'].unique()}
    presidentes_ja_convocados = set()

    dias_distribuicao = df[['DIA', 'DATA', 'MUNICIPIO', 'CATEGORIA', 'QUANTIDADE']].dropna(subset=['DIA'])
    candidatos_df = df[['NOME', 'CATEGORIA', 'INDISPONIBILIDADE', 'PRESIDENTE_DE_BANCA',
                        'MUNICIPIO_ORIGEM', 'INICIO_INDISPONIBILIDADE', 'FIM_INDISPONIBILIDADE']].dropna(subset=['NOME'])

    traducao_dias_eng = {'MONDAY':'SEGUNDA','TUESDAY':'TERCA','WEDNESDAY':'QUARTA','THURSDAY':'QUINTA','FRIDAY':'SEXTA'}

    for (dia_raw, municipio, data_municipio), grupo in dias_distribuicao.groupby(['DIA','MUNICIPIO','DATA']):
        data_municipio = pd.to_datetime(data_municipio, dayfirst=True, errors='coerce')
        dia_semana_pt = traducao_dias_eng.get(
            pd.to_datetime(data_municipio).strftime('%A').upper(), str(dia_raw).upper()
        ) if pd.notna(data_municipio) else str(dia_raw).upper()

        candidatos = candidatos_df[
            (candidatos_df['MUNICIPIO_ORIGEM'].apply(normalizar_texto) != normalizar_texto(municipio))
        ].copy()
        candidatos = candidatos[candidatos.apply(lambda x: esta_disponivel(x, data_municipio), axis=1)]

        if candidatos.empty:
            continue

        candidatos['CONVOCACOES'] = candidatos['NOME'].map(contador_convocacoes)
        candidatos = candidatos.sort_values('CONVOCACOES')
        pessoas_disponiveis = candidatos.copy()

        for _, op in grupo.iterrows():
            categorias_necessarias = [cat.strip() for cat in str(op['CATEGORIA']).split(',')]
            quantidade = int(op['QUANTIDADE'])

            candidatos_op = pessoas_disponiveis[
                pessoas_disponiveis['CATEGORIA'].apply(lambda x: any(cat in str(x) for cat in categorias_necessarias))
            ]

            if candidatos_op.empty:
                continue

            selecionados, presidente_nome = alocar_operacao(
                candidatos_op, quantidade, presidentes_ja_convocados
            )

            for _, pessoa in selecionados.iterrows():
                distribuicoes.append({
                    "DIA": dia_semana_pt,
                    "DATA": data_municipio.strftime("%d/%m/%y") if pd.notna(data_municipio) else "",
                    "MUNICIPIO": municipio,
                    "NOME": pessoa['NOME'],
                    "CATEGORIA": pessoa['CATEGORIA'],
                    "PRESIDENTE": "SIM" if pessoa['NOME'] == presidente_nome else "NAO"
                })
                contador_convocacoes[pessoa['NOME']] += 1

            pessoas_disponiveis = pessoas_disponiveis[~pessoas_disponiveis['NOME'].isin(selecionados['NOME'])]

    df_convocados = pd.DataFrame(distribuicoes)

    # ------------------------
    # Montagem dos não convocados (ajustado: sem DATA nem CATEGORIAS)
    # ------------------------
    nao_convocados_lista = []

    for (dia_raw, municipio, data_municipio), grupo in dias_distribuicao.groupby(['DIA','MUNICIPIO','DATA']):
        data_municipio = pd.to_datetime(data_municipio, dayfirst=True, errors='coerce')
        dia_semana_pt = traducao_dias_eng.get(
            pd.to_datetime(data_municipio).strftime('%A').upper(), str(dia_raw).upper()
        ) if pd.notna(data_municipio) else str(dia_raw).upper()

        for _, op in grupo.iterrows():
            categorias_necessarias = [cat.strip() for cat in str(op['CATEGORIA']).split(',')]
            quantidade = int(op['QUANTIDADE'])

            candidatos_op = candidatos_df[
                candidatos_df['CATEGORIA'].apply(lambda x: any(cat in str(x) for cat in categorias_necessarias))
            ]
            candidatos_op = candidatos_op[candidatos_op.apply(lambda x: esta_disponivel(x, data_municipio), axis=1)]

            convocados_op = df_convocados[
                (df_convocados["MUNICIPIO"] == municipio) &
                (df_convocados["DIA"] == dia_semana_pt)
            ]["NOME"].tolist()

            for _, row in candidatos_op.iterrows():
                if row["NOME"] not in convocados_op:
                    nao_convocados_lista.append({
                        "NOME": row["NOME"],
                        "DIA": dia_semana_pt,
                        "MUNICIPIO": municipio,
                        "PRESIDENTE": row["PRESIDENTE_DE_BANCA"]
                    })

    df_nao_convocados = pd.DataFrame(nao_convocados_lista).drop_duplicates(
        subset=["NOME", "DIA", "MUNICIPIO"]
    )

    # ------------------------
    # Exportação para Excel
    # ------------------------
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Convocados"
    for r_idx, row in enumerate(dataframe_to_rows(df_convocados, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws1.cell(row=r_idx, column=c_idx, value=value)

    ws2 = wb.create_sheet("Nao Convocados")
    for r_idx, row in enumerate(dataframe_to_rows(df_nao_convocados, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            ws2.cell(row=r_idx, column=c_idx, value=value)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    nome_arquivo_saida = f'distribuicao_{datetime.now().strftime("%B").upper()}.xlsx'
    return nome_arquivo_saida, df_convocados, df_nao_convocados, output

# ------------------------
# Interface Streamlit
# ------------------------
st.set_page_config(page_title="Distribuição Aleatória", page_icon="📊", layout="centered")

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

st.markdown(
    """
    <div class="main-card">
        <h1>📊 Distribuição Aleatória de Pessoas</h1>
        <p>Envie sua planilha Excel e gere automaticamente uma distribuição de convocados e não convocados de forma rápida e organizada.</p>
    </div>
    """,
    unsafe_allow_html=True
)

arquivo = st.file_uploader("📁 Envie a planilha (.xlsx)", type="xlsx")

if arquivo:
    st.markdown("### ⚙️ Processamento")
    st.info("Clique no botão abaixo para gerar a distribuição.")

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
            st.markdown(
                f"""
                <div style="text-align:center; margin-top:30px;">
                    <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
                       download="{nome_saida}"
                       target="_blank"
                       style="background:linear-gradient(90deg, #00c6ff, #0072ff); padding:12px 25px; color:white; text-decoration:none; border-radius:12px; font-size:16px; font-weight:bold;">
                        ⬇️ Baixar Excel
                    </a>
                </div>
                """,
                unsafe_allow_html=True
            )



