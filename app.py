import streamlit as st
import pandas as pd
import win32com.client
import pythoncom
import re
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime

st.set_page_config(layout="wide", page_title="IA Soccer PRO", page_icon="⚽")

FILE_PATH = r"C:\Users\vagse\Desktop\Projetos\IA Soccer\Base1.xlsm"


# ==============================
# PDF
# ==============================

def gerar_pdf(row, aposta_recomendada):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer)
    styles = getSampleStyleSheet()

    conteudo = []

    conteudo.append(Paragraph(f"<b>{row.get('Mandante')} vs {row.get('Visitante')}</b>", styles["Title"]))
    conteudo.append(Paragraph(f"Local: {row.get('Local')}", styles["Normal"]))
    conteudo.append(Paragraph(f"Mercado: {row.get('Mercado')}", styles["Normal"]))
    conteudo.append(Paragraph(f"Aposta Recomendada: {aposta_recomendada}", styles["Normal"]))
    conteudo.append(Paragraph(f"Confiança: {row.get('Confianca')}%", styles["Normal"]))
    conteudo.append(Paragraph("", styles["Normal"]))
    conteudo.append(Paragraph("<b>Análise:</b>", styles["Normal"]))

    if row.get("Dica 1"):
        conteudo.append(Paragraph(f"• {row.get('Dica 1')}", styles["Normal"]))
    if row.get("Dica 2"):
        conteudo.append(Paragraph(f"• {row.get('Dica 2')}", styles["Normal"]))

    conteudo.append(Paragraph("", styles["Normal"]))
    conteudo.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))

    doc.build(conteudo)
    buffer.seek(0)
    return buffer


# ==============================
# FUNÇÕES
# ==============================

def atualizar_excel():
    try:
        # Inicializa o COM para a thread atual
        pythoncom.CoInitialize()

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(FILE_PATH)
        wb.RefreshAll()

        # Aguarda as queries terminarem
        excel.CalculateUntilAsyncQueriesDone()

        # Salva e fecha
        wb.Save()
        wb.Close()
        excel.Quit()

        # Limpa o COM
        pythoncom.CoUninitialize()

        return True
    except Exception as e:
        st.error(f"Erro: {str(e)}")
        return False


@st.cache_data(ttl=60)
def load_data():
    return pd.read_excel(FILE_PATH, engine="openpyxl", dtype=str)


def extrair_periodo_mercado(mercado):
    """Extrai se o mercado é HT ou FT baseado no nome"""
    if not isinstance(mercado, str):
        return ""

    mercado_lower = mercado.lower()

    if "first half" in mercado_lower or "1st half" in mercado_lower or "ht" in mercado_lower:
        return " (1º Tempo)"
    return ""


def renomear_mercado(mercado):
    if not isinstance(mercado, str):
        return mercado

    m = mercado.lower()

    if "corners" in m:
        if "first half" in m or "1st half" in m or "ht" in m:
            return "Escanteios (1º Tempo)"
        return "Escanteios"
    if "goal" in m:
        if "first half" in m or "1st half" in m or "ht" in m:
            return "Gols (1º Tempo)"
        return "Gols"
    if "btts" in m:
        if "first half" in m or "1st half" in m or "ht" in m:
            return "Ambas Marcam (1º Tempo)"
        return "Ambas Marcam"

    return mercado


def extrair_percentual(texto):
    if not isinstance(texto, str):
        return None
    match = re.search(r'\((\d+)%\)', texto)
    return int(match.group(1)) if match else None


def extrair_tipo_e_valor(texto):
    """Extrai o tipo (over/under) e valor da dica"""
    if not isinstance(texto, str):
        return None, None

    tipo = None
    valor = None

    # Verifica o tipo (Over/Under)
    if "over" in texto.lower():
        tipo = "Over"
    elif "under" in texto.lower():
        tipo = "Under"

    # Extrai o valor numérico
    valor_match = re.search(r'(\d+\.?\d*)', texto)
    if valor_match:
        valor = valor_match.group(1)

    return tipo, valor


def extrair_periodo_dica(texto):
    """Extrai se a dica é para HT (first half) ou FT"""
    if not isinstance(texto, str):
        return ""

    texto_lower = texto.lower()

    # Verifica se é first half / HT
    if "first half" in texto_lower or "1st half" in texto_lower or "ht" in texto_lower:
        return " no primeiro tempo"

    # Se não tem indicação, é jogo completo (FT)
    return ""


def extrair_mercado_da_dica(texto):
    """Extrai se a dica é sobre corners, goals, etc"""
    if not isinstance(texto, str):
        return ""

    texto_lower = texto.lower()

    if "corner" in texto_lower:
        return "Escanteios"
    elif "goal" in texto_lower:
        return "Gols"
    elif "btts" in texto_lower:
        return "Ambas Marcam"

    return ""


def gerar_aposta_recomendada(row):
    """Gera a aposta recomendada baseada nas dicas"""
    mercado = row.get('Mercado', '')
    dica1 = row.get('Dica 1', '')
    dica2 = row.get('Dica 2', '')

    # Extrai o período da dica (HT ou FT)
    periodo = extrair_periodo_dica(dica1)
    if not periodo:
        periodo = extrair_periodo_dica(dica2)

    # Tenta extrair da primeira dica
    tipo1, valor1 = extrair_tipo_e_valor(dica1)
    mercado_dica1 = extrair_mercado_da_dica(dica1)

    # Se não encontrou na primeira, tenta na segunda
    if not tipo1 or not valor1:
        tipo2, valor2 = extrair_tipo_e_valor(dica2)
        mercado_dica2 = extrair_mercado_da_dica(dica2)

        if tipo2 and valor2:
            mercado_usar = mercado_dica2 if mercado_dica2 else mercado
            return f"{tipo2} {valor2} {mercado_usar}{periodo}"
        elif tipo2:
            mercado_usar = mercado_dica2 if mercado_dica2 else mercado
            return f"{tipo2} {mercado_usar}{periodo}"

    # Se encontrou na primeira
    if tipo1 and valor1:
        mercado_usar = mercado_dica1 if mercado_dica1 else mercado
        return f"{tipo1} {valor1} {mercado_usar}{periodo}"
    elif tipo1:
        mercado_usar = mercado_dica1 if mercado_dica1 else mercado
        return f"{tipo1} {mercado_usar}{periodo}"

    return f"{mercado}{periodo}"


def traduzir_dica(texto, mercado):
    if not isinstance(texto, str):
        return texto

    texto_original = texto
    texto = texto.replace("\n", " ").strip()

    # Verifica se é first half / HT
    periodo = extrair_periodo_dica(texto)

    # Extrai o nome do time
    time_match = re.match(r"^(.*?) have", texto, re.IGNORECASE)
    time = time_match.group(1) if time_match else ""

    perc = re.search(r"\((\d+%)\)", texto)
    perc_text = perc.group(1) if perc else ""

    jogos = re.search(r"(\d+) of their last (\d+)", texto)
    jogos_text = f"{jogos.group(1)} dos últimos {jogos.group(2)} jogos" if jogos else ""

    local = " em casa" if "home" in texto.lower() else " fora" if "away" in texto.lower() else ""

    # Verifica o tipo e valor
    tipo = ""
    valor = ""

    if "over" in texto.lower():
        tipo = "mais de"
    elif "under" in texto.lower():
        tipo = "menos de"

    valor_match = re.search(r'(\d+\.?\d*)', texto)
    if valor_match:
        valor = valor_match.group(1)

    # Monta a frase baseada no conteúdo
    if "corner" in texto.lower():
        if tipo and valor:
            frase = f"{time} teve {tipo} {valor} escanteios"
        else:
            frase = f"{time} teve tendência de escanteios"
    elif "goal" in texto.lower():
        if tipo and valor:
            frase = f"{time} teve {tipo} {valor} gols"
        else:
            frase = f"{time} teve tendência de gols"
    elif "btts" in texto.lower():
        frase = f"{time} teve jogos com ambas marcam"
    else:
        frase = texto

    # Adiciona informações complementares
    if jogos_text:
        frase += f" em {jogos_text}"
    if local:
        frase += local
    if periodo:
        frase += periodo
    if perc_text:
        frase += f" ({perc_text})"

    return frase


def processar_dados(df):
    df = df.fillna("")

    df["Mercado"] = df["Mercado"].apply(renomear_mercado)

    df["Dica 1"] = df.apply(lambda r: traduzir_dica(r["Dica 1"], r["Mercado"]), axis=1)
    df["Dica 2"] = df.apply(lambda r: traduzir_dica(r["Dica 2"], r["Mercado"]), axis=1)

    df["Perc_Dica1"] = df["Dica 1"].apply(extrair_percentual)
    df["Perc_Dica2"] = df["Dica 2"].apply(extrair_percentual)

    def calc(row):
        vals = [row["Perc_Dica1"], row["Perc_Dica2"]]
        vals = [v for v in vals if v is not None]
        return round(sum(vals) / len(vals)) if vals else 0

    df["Confianca"] = df.apply(calc, axis=1)

    # Adiciona a aposta recomendada
    df["Aposta_Recomendada"] = df.apply(gerar_aposta_recomendada, axis=1)

    return df


# ==============================
# HEADER
# ==============================

st.title("⚽ IA Soccer PRO")
st.caption("Desenvolvido por **Vagner S**")

# ==============================
# SIDEBAR
# ==============================

with st.sidebar:
    st.header("📊 Controles")

    if st.button("🔄 Atualizar Dados", use_container_width=True):
        with st.spinner("Atualizando dados do Excel..."):
            if atualizar_excel():
                st.success("✅ Dados atualizados com sucesso!")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error("❌ Falha ao atualizar os dados")

    st.divider()

    # Carrega os dados (já com cache)
    with st.spinner("Carregando dados..."):
        df = processar_dados(load_data())

    st.header("🔍 Filtros")

    locais = ["Todos"] + sorted(df["Local"].unique().tolist())
    local_sel = st.selectbox("🌍 Local", locais)

    mercados = ["Todos"] + sorted(df["Mercado"].unique().tolist())
    mercado_sel = st.selectbox("🎯 Mercado", mercados)

# ==============================
# FILTROS
# ==============================

df_filtrado = df.copy()

if local_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Local"] == local_sel]

if mercado_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Mercado"] == mercado_sel]

df_filtrado = df_filtrado.sort_values("Confianca", ascending=False)

# ==============================
# MÉTRICAS
# ==============================

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("📊 Total de Jogos", len(df_filtrado))

with col2:
    if len(df_filtrado) > 0:
        st.metric("⭐ Confiança Média", f"{df_filtrado['Confianca'].mean():.0f}%")
    else:
        st.metric("⭐ Confiança Média", "0%")

with col3:
    if len(df_filtrado) > 0:
        st.metric("🚀 Alta Confiança (70%+)", len(df_filtrado[df_filtrado['Confianca'] >= 70]))
    else:
        st.metric("🚀 Alta Confiança (70%+)", "0")

st.divider()

# ==============================
# CARDS
# ==============================

if len(df_filtrado) == 0:
    st.warning("⚠️ Nenhum jogo encontrado com os filtros selecionados.")
else:
    for idx, row in df_filtrado.iterrows():
        conf = row.get("Confianca", 0)
        aposta = row.get("Aposta_Recomendada", row.get("Mercado", "Apostar"))

        # Definir cor baseada na confiança
        if conf >= 70:
            cor = "green"
        elif conf >= 50:
            cor = "orange"
        else:
            cor = "red"

        # Container do card
        with st.container():
            # Linha 1: Times e Confiança
            col1, col2 = st.columns([3, 1])
            with col1:
                st.subheader(f"{row.get('Mandante', '?')} vs {row.get('Visitante', '?')}")
            with col2:
                st.markdown(f"**⭐ {conf}%**")

            # Linha 2: Local e Mercado
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"🌍 **Local:** {row.get('Local', '?')}")
            with col2:
                st.write(f"🎯 **Mercado:** {row.get('Mercado', '?')}")

            # Linha 3: Aposta Recomendada (destacada)
            st.markdown(f"""
            <div style="background-color:{cor}20; padding:10px; border-radius:10px; border-left:4px solid {cor}; margin:10px 0;">
                <b>🎲 APOSTA RECOMENDADA:</b> <span style="font-size:1.1rem;">{aposta}</span>
            </div>
            """, unsafe_allow_html=True)

            # Linha 4: Análise
            st.write("📈 **Análise:**")
            st.write(f"• {row.get('Dica 1', '')}")
            st.write(f"• {row.get('Dica 2', '')}")

            # Linha 5: Botão PDF
            pdf = gerar_pdf(row, aposta)
            nome_arquivo = f"{row.get('Mandante', 'Time')}_vs_{row.get('Visitante', 'Time2')}.pdf"

            st.download_button(
                label="📄 Baixar PDF",
                data=pdf,
                file_name=nome_arquivo,
                mime="application/pdf",
                key=f"pdf_{idx}_{row.get('Mandante', '')}_{row.get('Visitante', '')}"
            )

            st.divider()