import streamlit as st
import pandas as pd
from datetime import datetime
import unicodedata
import io
import base64
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import random
import locale
from collections import defaultdict

# -----------------------
# Configuração de locale
# -----------------------
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'portuguese')
    except:
        pass

# -----------------------
# Funções auxiliares
# -----------------------
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
    inicio = row.get('INICIO_INDISPONIBILIDADE', pd.NaT)
    fim = row.get('FIM_INDISPONIBILIDADE', pd.NaT)

    if pd.notna(inicio):
        inicio = pd.to_datetime(inicio, dayfirst=True, errors='coerce')
    if pd.notna(fim):
        fim = pd.to_datetime(fim, dayfirst=True, errors='coerce')

    if pd.notna(inicio) and pd.notna(fim):
        if inicio <= data <= fim:
            return False
    return True

def traduzir_dia(data_item_dt):
    dias_traducao = {
        'MONDAY': 'SEGUNDA',
        'TUESDAY': 'TERCA',
        'WEDNESDAY': 'QUARTA',
        'THURSDAY': 'QUINTA',
        'FRIDAY': 'SEXTA',
        'SATURDAY': 'SABADO',
        'SUNDAY': 'DOMINGO'
    }
    dia_semana = data_item_dt.strftime("%A").upper()
    return dias_traducao.get(dia_semana, dia_semana)

def alocar_operacao(candidatos_op, quantidade, presidentes_ja_convocados):
    if candidatos_op.empty:
        return pd.DataFrame(), None

    candidatos_op = candidatos_op.sort_values('CONVOCACOES')
    presidentes_disponiveis = candidatos_op[candidatos_op['PRESIDENTE_DE_BANCA'].str.upper() == 'SIM']
    presidente_selecionado = None
    if not presidentes_disponiveis.empty:
        presidente_selecionado = presidentes_disponiveis.sample(1, random_state=random.randint(0, 10000))
        restantes = candidatos_op[~candidatos_op['NOME'].isin(presidente_selecionado['NOME'])]
        restantes = restantes.sample(min(len(restantes), max(0, quantidade - 1)), random_state=random.randint(0, 10000))
        selecionados = pd.concat([presidente_selecionado, restantes])
    else:
        selecionados = candidatos_op.sample(min(quantidade, len(candidatos_op)), random_state=random.randint(0, 10000))

    if len(selecionados) > quantidade:
        selecionados = selecionados.sample(quantidade, random_state=random.randint(0, 10000))

    presidente_nome = None
    if presidente_selecionado is not None:
        presidente_nome = presidente_selecionado.iloc[0]['NOME']
        presidentes_ja_convocados.add(presidente_nome)

    return selecionados, presidente_nome

# -----------------------
# Processamento com regras avançadas
# -----------------------
def processar_distribuicao(arquivo_excel):
    xls = pd.ExcelFile(arquivo_excel)
    sheet_name = 'Planilha1' if 'Planilha1' in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # Normalizar colunas
    df.columns = [normalizar_coluna(col) for col in df.columns]

    colunas_possiveis_nome = ['NOME', 'NOME_COMPLETO', 'NOME_PESSOA']
    for col in colunas_possiveis_nome:
        if col in df.columns:
            df['NOME'] = df[col]
            break
    if 'NOME' not in df.columns:
        st.error("❌ Erro: coluna de nomes não encontrada")
        return None, pd.DataFrame(), pd.DataFrame(), io.BytesIO()

    # Preencher colunas obrigatórias
    for col in ['INDISPONIBILIDADE', 'PRESIDENTE_DE_BANCA', 'MUNICIPIO_ORIGEM']:
        if col not in df.columns:
            df[col] = 'NAO' if col != 'MUNICIPIO_ORIGEM' else ''
        df[col] = df[col].fillna('NAO' if col != 'MUNICIPIO_ORIGEM' else '').astype(str).str.strip().str.upper()
    for col in ['INICIO_INDISPONIBILIDADE', 'FIM_INDISPONIBILIDADE']:
        if col not in df.columns:
            df[col] = pd.NaT
        else:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')

    distribuicoes = []
    contador_convocacoes = {nome: 0 for nome in df['NOME'].unique()}
    presidentes_ja_convocados = set()
    datas_convocados = {}
    municipios_por_pessoa = defaultdict(set)

    dias_distribuicao = df[['DIA', 'DATA', 'MUNICIPIO', 'CATEGORIA', 'QUANTIDADE']].dropna(subset=['DIA'])
    candidatos_df = df[['NOME', 'CATEGORIA', 'INDISPONIBILIDADE', 'PRESIDENTE_DE_BANCA',
                        'MUNICIPIO_ORIGEM', 'INICIO_INDISPONIBILIDADE', 'FIM_INDISPONIBILIDADE']].dropna(subset=['NOME'])

    rodadas = 0
    max_rodadas = 15
    min_convocacoes = 3

    while rodadas < max_rodadas:
        rodadas += 1
        alteracao = False
        for (dia_raw, municipio, data_municipio), grupo in dias_distribuicao.groupby(['DIA', 'MUNICIPIO', 'DATA']):
            data_municipio_dt = pd.to_datetime(data_municipio, dayfirst=True, errors='coerce')
            if pd.isna(data_municipio_dt):
                continue
            dia_semana_pt = traduzir_dia(data_municipio_dt)
            data_str_pt = data_municipio_dt.strftime("%d/%m/%Y")

            candidatos = candidatos_df.copy()
            candidatos = candidatos[candidatos.apply(lambda x: esta_disponivel(x, data_municipio_dt), axis=1)]
            candidatos = candidatos[candidatos['MUNICIPIO_ORIGEM'].apply(normalizar_texto) != normalizar_texto(municipio)]
            candidatos['CONVOCACOES'] = candidatos['NOME'].map(contador_convocacoes)

            for _, op in grupo.iterrows():
                categorias_necessarias = [cat.strip() for cat in str(op['CATEGORIA']).split(',')]
                quantidade = int(op['QUANTIDADE'])
                candidatos_op = candidatos[candidatos['CATEGORIA'].apply(
                    lambda x: any(cat in str(x) for cat in categorias_necessarias))]

                data_key = data_municipio_dt.strftime("%Y-%m-%d")
                convocados_na_data = datas_convocados.get(data_key, set())
                candidatos_op = candidatos_op[~candidatos_op['NOME'].isin(convocados_na_data)]
                candidatos_op = candidatos_op[~candidatos_op['NOME'].isin([n for n in municipios_por_pessoa if municipio in municipios_por_pessoa[n]])]

                candidatos_op = candidatos_op.sort_values('CONVOCACOES', ascending=True)
                if candidatos_op.empty:
                    continue

                selecionados, presidente_nome = alocar_operacao(candidatos_op, quantidade, presidentes_ja_convocados)

                # ---------- Garantir pelo menos 1 presidente ----------
                if not any(selecionados['PRESIDENTE_DE_BANCA'] == 'SIM'):
                    presidentes_disponiveis = candidatos_op[candidatos_op['PRESIDENTE_DE_BANCA'] == 'SIM']
                    presidentes_disponiveis = presidentes_disponiveis[~presidentes_disponiveis['NOME'].isin(selecionados['NOME'])]
                    if not presidentes_disponiveis.empty:
                        pres_subs = presidentes_disponiveis.sample(1, random_state=random.randint(0,10000)).iloc[0]
                        # Substitui o primeiro não-presidente
                        for idx, row_sel in selecionados.iterrows():
                            if row_sel['PRESIDENTE_DE_BANCA'] != 'SIM':
                                selecionados.loc[idx] = pres_subs
                                presidente_nome = pres_subs['NOME']
                                break

                for _, pessoa in selecionados.iterrows():
                    nome = pessoa['NOME']
                    presidente = pessoa['PRESIDENTE_DE_BANCA'] == 'SIM'

                    if municipio in municipios_por_pessoa[nome]:
                        continue

                    if presidente or contador_convocacoes[nome] < min_convocacoes:
                        distribuicoes.append({
                            "DIA": dia_semana_pt,
                            "DATA": data_str_pt,
                            "MUNICIPIO": municipio,
                            "NOME": nome,
                            "CATEGORIA": pessoa['CATEGORIA'],
                            "PRESIDENTE": "SIM" if nome == presidente_nome else "NAO"
                        })
                        contador_convocacoes[nome] += 1
                        datas_convocados.setdefault(data_key, set()).add(nome)
                        municipios_por_pessoa[nome].add(municipio)
                        alteracao = True

        if all(v >= min_convocacoes or df.loc[df['NOME'] == n, 'PRESIDENTE_DE_BANCA'].iloc[0] == 'SIM'
               for n, v in contador_convocacoes.items()):
            break
        if not alteracao:
            break

    # ----------------------- Não convocados
    nao_convocados_lista = []
    for _, row in candidatos_df.iterrows():
        datas_validas = [d for d in dias_distribuicao['DATA'].dropna().unique()
                         if esta_disponivel(row, pd.to_datetime(d, dayfirst=True, errors='coerce'))]
        for data_item in datas_validas:
            data_item_dt = pd.to_datetime(data_item, dayfirst=True, errors='coerce')
            data_item_str = data_item_dt.strftime("%d/%m/%Y")
            convocado_datas = [x['DATA'] for x in distribuicoes if x['NOME'] == row['NOME']]
            if data_item_str not in convocado_datas:
                dia_semana_pt = traduzir_dia(data_item_dt)
                nao_convocados_lista.append({
                    "NOME": row['NOME'],
                    "DIA": dia_semana_pt,
                    "CATEGORIA": row['CATEGORIA']
                })

    df_nao_convocados = pd.DataFrame(nao_convocados_lista).drop_duplicates(subset=["NOME", "DIA", "CATEGORIA"])
    df_convocados = pd.DataFrame(distribuicoes)

    # ----------------------- Exportação Excel
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
    nome_arquivo_saida = f'distribuicao_{datetime.now().strftime("%B").lower()}.xlsx'

    return nome_arquivo_saida, df_convocados, df_nao_convocados, output

# -----------------------
# Interface Streamlit
# -----------------------
st.set_page_config(page_title="Distribuição Equilibrada", page_icon="📊", layout="centered")

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

st.markdown("""
    <div class="main-card">
        <h1>📊 Distribuição Equilibrada de Convocações</h1>
        <p>O sistema garante pelo menos 3 convocações por pessoa, 1 presidente por município/dia, respeitando disponibilidade, categorias e municípios/dia/semana.</p>
    </div>
    """, unsafe_allow_html=True)

arquivo = st.file_uploader("📁 Envie a planilha (.xlsx)", type="xlsx")

if arquivo:
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
                       download="{nome_saida}"
                       target="_blank"
                       style="background:linear-gradient(90deg, #00c6ff, #0072ff); padding:12px 25px; color:white; text-decoration:none; border-radius:12px; font-size:16px; font-weight:bold;">
                        ⬇️ Baixar Excel
                    </a>
                </div>
            """, unsafe_allow_html=True)
