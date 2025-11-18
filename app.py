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

# -----------------------
# Selenium (WhatsApp) try import
# -----------------------
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    SELENIUM_AVAILABLE = True
except Exception:
    SELENIUM_AVAILABLE = False

# If you need to specify chromedriver path, set here (or leave None if chromedriver is in PATH)
CHROMEDRIVER_PATH = None  # ex: r"C:\tools\chromedriver.exe" or None

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
# 🧠 Processamento principal (mantive suas regras)
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

    # --- Aba de não convocados (corrigida) ---
    dias_validos = df["DIA"].unique()
    combinacoes = []
    for nome in df["NOME"].unique():
        for dia in dias_validos:
            combinacoes.append({"NOME": nome, "DIA": dia})
    base_completa = pd.DataFrame(combinacoes)
    df_base = df[["NOME", "CATEGORIA", "PRESIDENTE_DE_BANCA",
                  "DIAS_INDISPONIBILIDADE", "INICIO_INDISPONIBILIDADE", "FIM_INDISPONIBILIDADE", "TELEFONE"]].drop_duplicates(subset=["NOME"])

    # --- FILTRAR quem está indisponível durante todo o período da operação ---
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
# 💻 Interface Streamlit (UI)
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

arquivo = st.file_uploader("📁 Envie a planilha (.xlsx) — precisa ter coluna TELEFONE", type="xlsx")

# -----------------------------
# Função auxiliar: formata telefone internacional para WhatsApp (apenas dígitos)
# Exemplo: transforma "(11) 9 9123-4567" -> "5511991234567" (você pode ajustar para seu país)
# -----------------------------
def formatar_telefone_para_whatsapp(tel_raw):
    if not isinstance(tel_raw, str):
        tel_raw = str(tel_raw)
    digits = "".join(ch for ch in tel_raw if ch.isdigit())
    # Se o número já tem código do país (ex: 55...), devolve direto
    if digits.startswith("55") and len(digits) >= 12:
        return digits
    # Caso comum BR: se começa com 0 ou 9? Aqui tentamos adicionar '55' (Brasil) se faltar.
    # Ajuste conforme seu país/regra
    if len(digits) >= 10:
        return "55" + digits
    return digits

# -----------------------------
# Função que envia mensagens via WhatsApp Web usando Selenium
# -----------------------------
def send_whatsapp_messages(resumo_df, headless=False, delay_between=3):
    """
    resumo_df deve ter colunas: NOME, TELEFONE, ITENS (mensagem já formatada).
    Requisitos: Selenium + Chrome + chromedriver. Roda localmente.
    """
    if not SELENIUM_AVAILABLE:
        raise RuntimeError("Selenium não está disponível. Instale selenium e chromedriver.")

    # Configura driver
    options = webdriver.ChromeOptions()
    # mantemos perfil do usuário para já estar logado no WhatsApp Web (evita QR)
    # OBS: caminho do perfil precisa ser ajustado se usar
    # options.add_argument(r"user-data-dir=C:\Users\<seu_usuario>\AppData\Local\Google\Chrome\User Data")
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1200,800")
    # Inicia o driver
    if CHROMEDRIVER_PATH:
        driver = webdriver.Chrome(executable_path=CHROMEDRIVER_PATH, options=options)
    else:
        driver = webdriver.Chrome(options=options)

    wait = WebDriverWait(driver, 60)
    sent = 0
    try:
        for _, row in resumo_df.iterrows():
            nome = row.get("NOME", "")
            telefone = row.get("TELEFONE", "")
            itens = row.get("ITENS", "")
            if not telefone or str(telefone).strip() == "":
                continue
            phone = formatar_telefone_para_whatsapp(telefone)
            if not phone or len(phone) < 10:
                continue
            # Monta mensagem
            mensagem = f"Bom Dia, Prezado {nome}\\n\\nSegue abaixo os dias e os municípios que você foi convocado para fazer a operação:\\n\\n{itens}\\n\\nAtenciosamente,"
            url = f"https://web.whatsapp.com/send?phone={phone}&text={mensagem}"
            driver.get(url)
            # espera carregar o botão de enviar ou textarea
            try:
                # Aguarda até que o botão de enviar (xpath por exemplo) esteja presente
                # alternativa: esperar textarea e dar ENTER
                txt_xpath = "//div[@contenteditable='true' and @data-tab='10']"
                wait.until(EC.presence_of_element_located((By.XPATH, txt_xpath)))
                time.sleep(1)  # pequena espera para estabilidade
                # Foca no campo e envia ENTER
                input_box = driver.find_element(By.XPATH, txt_xpath)
                input_box.click()
                # já vem com o texto do parâmetro, apenas pressionar ENTER
                input_box.send_keys(Keys.ENTER)
                sent += 1
                time.sleep(delay_between)  # espera entre envios
            except Exception as e:
                # caso não consiga enviar automaticamente, deixa a aba aberta para envio manual
                st.warning(f"Não foi possível enviar automaticamente para {nome} ({telefone}). Aba aberta para você enviar manualmente.")
                # mantém a aba aberta alguns segundos para você ver; ou continue
                time.sleep(2)
                continue
    finally:
        # não fechamos o driver imediatamente; fechar é responsabilidade do fluxo
        driver.quit()
    return sent

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
            <div style="text-align:center;margin-top:20px;">
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{nome_saida}" target="_blank" style="background:linear-gradient(90deg,#00c6ff,#0072ff);padding:12px 25px;color:white;text-decoration:none;border-radius:12px;font-size:16px;font-weight:bold;">
            ⬇️ Baixar Excel
            </a></div>
            """, unsafe_allow_html=True)

            # --- Preparar mensagens por convocado (para WhatsApp) ---
            resumo = df_conv.groupby("NOME").apply(
                lambda g: pd.Series({
                    "TELEFONE": g["TELEFONE"].iloc[0] if "TELEFONE" in g.columns else "",
                    "ITENS": "\n".join([f"{r['DATA'].strftime('%d/%m/%Y')} - {r['MUNICIPIO']}" for _, r in g.sort_values(by="DATA").iterrows()])
                })
            ).reset_index()

            st.markdown("### ✳️ Enviar convocação pelo WhatsApp Web (automático)")
            st.write("Requisitos: WhatsApp Web autenticado no navegador, Chrome + chromedriver instalados, app rodando no seu PC.")
            if not SELENIUM_AVAILABLE:
                st.error("Selenium não está instalado neste ambiente. Instale `selenium` e o ChromeDriver no seu PC para ativar o envio automático.")
            else:
                # Mostrar resumo simples
                st.dataframe(resumo, use_container_width=True)

                col_a, col_b = st.columns([1,2])
                with col_a:
                    headless = st.checkbox("Rodar headless (não mostra janela)", value=False)
                with col_b:
                    delay_between = st.number_input("Segundos entre cada envio", min_value=1, max_value=20, value=3)

                enviar_whatsapp = st.button("📱 Enviar via WhatsApp Web (automático)")
                if enviar_whatsapp:
                    st.info("Iniciando envio via WhatsApp Web... atenção ao navegador que vai abrir. Se aparecer QR, escaneie-o no seu celular.")
                    try:
                        sent_count = send_whatsapp_messages(resumo, headless=headless, delay_between=delay_between)
                        st.success(f"Mensagens enviadas: {sent_count}")
                    except Exception as e:
                        st.error(f"Erro no envio via Selenium: {e}")
                        st.info("Verifique se o ChromeDriver está instalado e compatível com sua versão do Chrome, e se o Selenium está instalado (pip install selenium).")

            # Mostrar mensagens da regra Vanessa, se houver
            if msgs_vanessa:
                st.markdown("### Mensagens especiais")
                for m in msgs_vanessa:
                    st.info(m)
