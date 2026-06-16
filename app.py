# -*- coding: utf-8 -*-
"""
=================================================================================
 SISTEMA DE CONTROLE DE METAS DE QUILOMETRAGEM - MOTORISTAS
 EPR VIA MINEIRA
=================================================================================
Aplicação desenvolvida em Streamlit para automatizar a geração de reportes
INDIVIDUAIS de motoristas (por data), a partir da leitura (OCR) de prints
de WhatsApp enviados pelo CCO, com cálculo automático de meta de KM e
histórico consultável.

Autor: Gerado para uso interno da EPR Via Mineira
=================================================================================
"""

import os
import io
import random
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image
from fpdf import FPDF


# =================================================================================
# 1. CONFIGURAÇÕES GERAIS DA APLICAÇÃO
# =================================================================================

DATA_FILE = "registros_motoristas.csv"   # "banco de dados" local em CSV
DEFAULT_META_KM = 150.0                   # meta padrão de KM (editável na tela)

# Colunas do "banco de dados" - cada LINHA = 1 plantão de 1 motorista em 1 data
COLUMNS = [
    "data_plantao",
    "motorista",
    "km_rodado",
    "meta_km",
    "status_meta",
    "numero_ocorrencia",
    "tempo_parado",
    "descricao_atraso",
    "registrado_em",
]

st.set_page_config(
    page_title="EPR Via Mineira | Controle de Metas de KM",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =================================================================================
# 2. IDENTIDADE VISUAL - CSS CUSTOMIZADO (EPR VIA MINEIRA)
# =================================================================================

def inject_custom_css():
    """Injeta o CSS customizado para dar a identidade visual corporativa da
    EPR Via Mineira (azul institucional, cartões brancos, status em verde/vermelho)."""
    st.markdown(
        """
        <style>
        :root {
            --epr-azul-escuro: #0B2D4E;
            --epr-azul: #154C79;
            --epr-azul-claro: #1F6FB2;
            --epr-cinza-fundo: #F2F4F7;
            --epr-cinza-borda: #D9DEE3;
            --epr-branco: #FFFFFF;
            --epr-verde: #1E8449;
            --epr-verde-claro: #E8F8F0;
            --epr-laranja: #E67E22;
            --epr-laranja-claro: #FDF2E3;
            --epr-vermelho: #C0392B;
            --epr-vermelho-claro: #FDEDEC;
            --epr-texto: #1B2631;
        }

        /* Fundo geral da aplicação */
        .stApp {
            background-color: var(--epr-cinza-fundo);
        }

        /* Esconde o menu padrão e o "Made with Streamlit" para visual mais corporativo */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}

        /* Tipografia e títulos */
        h1, h2, h3, h4 {
            color: var(--epr-azul-escuro) !important;
            font-family: "Segoe UI", Arial, sans-serif;
        }
        p, span, label, .stMarkdown {
            color: var(--epr-texto);
        }

        /* ===================== CABEÇALHO CORPORATIVO ===================== */
        .epr-header {
            background: linear-gradient(135deg, var(--epr-azul-escuro) 0%, var(--epr-azul) 100%);
            padding: 1.4rem 2rem;
            border-radius: 10px;
            margin-bottom: 1.6rem;
            display: flex;
            align-items: center;
            gap: 1.2rem;
            box-shadow: 0 4px 14px rgba(11, 45, 78, 0.25);
        }
        .epr-header img {
            height: 46px;
            background: white;
            padding: 6px 10px;
            border-radius: 8px;
        }
        .epr-header-titulo {
            color: white !important;
            font-size: 1.5rem;
            font-weight: 700;
            margin: 0;
            line-height: 1.2;
        }
        .epr-header-subtitulo {
            color: #D6E4F0 !important;
            font-size: 0.92rem;
            margin: 0;
        }

        /* ===================== CARTÕES (CONTAINERS) ===================== */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: var(--epr-branco);
            border: 1px solid var(--epr-cinza-borda);
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(11, 45, 78, 0.06);
        }

        /* ===================== BOTÕES ===================== */
        .stButton > button, .stDownloadButton > button {
            background-color: var(--epr-azul-escuro);
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            padding: 0.55rem 1.2rem;
            transition: background-color 0.2s ease-in-out;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            background-color: var(--epr-azul-claro);
            color: white;
        }
        .stButton > button[kind="primary"] {
            background-color: var(--epr-verde);
        }
        .stButton > button[kind="primary"]:hover {
            background-color: #239B56;
        }

        /* ===================== CAMPOS DE ENTRADA ===================== */
        .stTextInput > div > div > input,
        .stNumberInput > div > div > input,
        .stTextArea > div > div > textarea,
        .stDateInput input,
        .stSelectbox > div > div {
            border-radius: 6px !important;
            border: 1px solid var(--epr-cinza-borda) !important;
        }

        /* ===================== UPLOAD (DRAG AND DROP) ===================== */
        [data-testid="stFileUploaderDropzone"] {
            background-color: #FAFBFC;
            border: 2px dashed var(--epr-azul-claro) !important;
            border-radius: 10px;
        }

        /* ===================== BADGES DE STATUS DA META ===================== */
        .epr-badge {
            display: block;
            text-align: center;
            font-size: 1.05rem;
            font-weight: 700;
            letter-spacing: 0.3px;
            padding: 0.75rem 1rem;
            border-radius: 8px;
            margin: 0.6rem 0 1rem 0;
        }
        .epr-badge-success {
            background-color: var(--epr-verde-claro);
            color: var(--epr-verde);
            border: 1.5px solid var(--epr-verde);
        }
        .epr-badge-danger {
            background-color: var(--epr-vermelho-claro);
            color: var(--epr-vermelho);
            border: 1.5px solid var(--epr-vermelho);
        }

        /* ===================== CARTÃO DE HISTÓRICO ===================== */
        .epr-card-historico {
            background-color: var(--epr-branco);
            border: 1px solid var(--epr-cinza-borda);
            border-left: 6px solid var(--epr-azul-escuro);
            border-radius: 10px;
            padding: 1.2rem 1.4rem;
            margin-bottom: 1rem;
        }
        .epr-card-historico h4 {
            margin-top: 0;
        }
        .epr-linha-info {
            display: flex;
            justify-content: space-between;
            padding: 0.35rem 0;
            border-bottom: 1px dashed #E5E8EC;
            font-size: 0.95rem;
        }
        .epr-linha-info span:first-child {
            color: #5A6B7B;
            font-weight: 600;
        }
        .epr-linha-info span:last-child {
            color: var(--epr-texto);
            font-weight: 500;
            text-align: right;
        }

        /* Rodapé customizado */
        .epr-footer {
            text-align: center;
            color: #8C98A4;
            font-size: 0.8rem;
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid var(--epr-cinza-borda);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    """Renderiza o cabeçalho institucional da EPR Via Mineira."""
    logo_url = "https://eprviamineira.com.br/wp-content/uploads/2024/10/epr-via-mineira-logo.png"
    st.markdown(
        f"""
        <div class="epr-header">
            <img src="{logo_url}" onerror="this.style.display='none'" />
            <div>
                <p class="epr-header-titulo">EPR VIA MINEIRA</p>
                <p class="epr-header-subtitulo">
                    Sistema de Controle de Metas de Quilometragem &middot; Reportes Individuais de Motoristas
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =================================================================================
# 3. CAMADA DE DADOS (CSV como "banco de dados" local)
# =================================================================================

def load_data() -> pd.DataFrame:
    """Carrega o histórico de registros do CSV. Cria o arquivo se não existir."""
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE, dtype=str)
        # garante que todas as colunas esperadas existam (compatibilidade futura)
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[COLUMNS]
    return pd.DataFrame(columns=COLUMNS)


def save_record(record: dict) -> None:
    """Salva (ou atualiza, caso já exista) o registro de um motorista em uma
    determinada data. A combinação (motorista + data_plantao) é a chave única
    do reporte individual."""
    df = load_data()
    mask_existente = (df["motorista"] == record["motorista"]) & (
        df["data_plantao"] == record["data_plantao"]
    )
    df = df[~mask_existente]  # remove registro antigo da mesma data/motorista, se houver
    novo_df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
    novo_df.to_csv(DATA_FILE, index=False)


def format_data_br(iso_date_str: str) -> str:
    """Converte 'YYYY-MM-DD' para 'DD/MM/YYYY' para exibição."""
    try:
        return datetime.strptime(iso_date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return str(iso_date_str)


# =================================================================================
# 4. PROCESSAMENTO DE IMAGEM (OCR) - PRINT DO WHATSAPP ENVIADO PELO CCO
# =================================================================================

def process_ocr_image(image: Image.Image):
    """
    Extrai automaticamente o NOME do motorista e a QUILOMETRAGEM (KM) rodada
    a partir do print do WhatsApp enviado pelo CCO.

    -----------------------------------------------------------------------
    ATENÇÃO: esta função está em MODO SIMULAÇÃO (para demonstração/testes
    sem depender da instalação de engines de OCR). Para uso em produção,
    escolha UMA das opções reais abaixo e substitua o bloco de simulação.
    -----------------------------------------------------------------------

    OPÇÃO 1 - pytesseract
    (requer o binário "Tesseract-OCR" instalado no sistema operacional)
    -----------------------------------------------------------------------
        import pytesseract
        texto_extraido = pytesseract.image_to_string(image, lang="por")

    OPÇÃO 2 - easyocr
    (não depende de binário externo, porém é mais pesado e mais lento)
    -----------------------------------------------------------------------
        import easyocr
        reader = easyocr.Reader(["pt"], gpu=False)
        resultados = reader.readtext(np.array(image))
        texto_extraido = " ".join([item[1] for item in resultados])

    Após obter o texto bruto ("texto_extraido"), utilize expressões
    regulares para localizar o nome e o KM, por exemplo:
    -----------------------------------------------------------------------
        import re
        nome_match = re.search(
            r"(?:Nome|Motorista)[:\\-]\\s*([A-Za-zÀ-ÿ\\s]+)",
            texto_extraido, re.IGNORECASE,
        )
        km_match = re.search(
            r"(\\d{1,4}[.,]?\\d{0,2})\\s*km",
            texto_extraido, re.IGNORECASE,
        )
        nome_final = nome_match.group(1).strip() if nome_match else ""
        km_final = float(km_match.group(1).replace(",", ".")) if km_match else 0.0
        return nome_final, km_final
    -----------------------------------------------------------------------
    """
    # ============================ SIMULAÇÃO ============================
    nomes_exemplo = [
        "Carlos Eduardo Silva",
        "José Roberto Almeida",
        "Antônio Marcos Pereira",
        "Sebastião Ferreira Lima",
        "Wellington Souza Costa",
        "Marcos Vinícius Ribeiro",
        "Geraldo Henrique Souza",
    ]
    nome_simulado = random.choice(nomes_exemplo)
    km_simulado = round(random.uniform(70, 230), 1)
    return nome_simulado, km_simulado


# =================================================================================
# 5. EXPORTAÇÃO DE REPORTES INDIVIDUAIS (EXCEL / PDF)
# =================================================================================

def export_record_to_excel(record: dict) -> bytes:
    """Gera um arquivo .xlsx em memória com os dados do reporte individual."""
    status_label = "META BATIDA" if record["status_meta"] == "BATIDA" else "META NÃO BATIDA"
    df_export = pd.DataFrame(
        [
            {
                "Data do Plantão": format_data_br(record["data_plantao"]),
                "Motorista": record["motorista"],
                "KM Rodado": record["km_rodado"],
                "Meta de KM": record["meta_km"],
                "Status da Meta": status_label,
                "Nº Ocorrência CCO": record.get("numero_ocorrencia", ""),
                "Tempo Parado na Ocorrência": record.get("tempo_parado", ""),
                "Descrição/Observações do Atraso": record.get("descricao_atraso", ""),
                "Registrado em": record.get("registrado_em", ""),
            }
        ]
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Reporte Individual")
    buffer.seek(0)
    return buffer.getvalue()


def export_record_to_pdf(record: dict) -> bytes:
    """Gera um arquivo .pdf em memória com os dados do reporte individual,
    seguindo a identidade visual da EPR Via Mineira (cabeçalho azul)."""
    status_label = "META BATIDA" if record["status_meta"] == "BATIDA" else "META NAO BATIDA"

    pdf = FPDF(format="A4")
    pdf.add_page()

    # Cabeçalho azul institucional
    pdf.set_fill_color(11, 45, 78)
    pdf.rect(0, 0, 210, 28, style="F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 17)
    pdf.set_xy(10, 7)
    pdf.cell(0, 9, "EPR VIA MINEIRA")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_xy(10, 17)
    pdf.cell(0, 6, "Reporte Individual de Controle de Meta de Quilometragem")

    # Corpo do reporte
    pdf.set_text_color(27, 38, 49)
    pdf.set_xy(10, 38)

    linhas = [
        ("Data do Plantao", format_data_br(record["data_plantao"])),
        ("Motorista", record["motorista"]),
        ("KM Rodado", f"{float(record['km_rodado']):.1f} km"),
        ("Meta Estabelecida", f"{float(record['meta_km']):.1f} km"),
        ("Status da Meta", status_label),
    ]

    pdf.set_font("Helvetica", "B", 11)
    for rotulo, valor in linhas:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(55, 9, f"{rotulo}:", border=0)
        pdf.set_font("Helvetica", "", 11)
        pdf.cell(0, 9, str(valor), border=0, ln=1)

    # Bloco de justificativa, apenas se a meta não foi batida
    if record["status_meta"] != "BATIDA":
        pdf.ln(4)
        pdf.set_fill_color(253, 237, 236)
        pdf.set_draw_color(192, 57, 43)
        y_inicial = pdf.get_y()
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(192, 57, 43)
        pdf.cell(0, 8, "Justificativa do Nao Cumprimento da Meta", ln=1)
        pdf.set_text_color(27, 38, 49)

        pdf.set_font("Helvetica", "B", 10.5)
        pdf.cell(60, 8, "No. Ocorrencia (CCO):")
        pdf.set_font("Helvetica", "", 10.5)
        pdf.cell(0, 8, str(record.get("numero_ocorrencia", "")), ln=1)

        pdf.set_font("Helvetica", "B", 10.5)
        pdf.cell(60, 8, "Tempo Parado:")
        pdf.set_font("Helvetica", "", 10.5)
        pdf.cell(0, 8, str(record.get("tempo_parado", "")), ln=1)

        pdf.set_font("Helvetica", "B", 10.5)
        pdf.cell(0, 8, "Descricao/Observacoes:", ln=1)
        pdf.set_font("Helvetica", "", 10.5)
        pdf.multi_cell(0, 7, str(record.get("descricao_atraso", "")))

    pdf.set_y(270)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(140, 152, 164)
    pdf.cell(0, 6, f"Documento gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} - Uso interno EPR Via Mineira")

    pdf_output = pdf.output()
    if isinstance(pdf_output, str):
        return pdf_output.encode("latin-1")
    return bytes(pdf_output)


# =================================================================================
# 6. APLICAÇÃO PRINCIPAL (STREAMLIT)
# =================================================================================

def init_session_state():
    """Inicializa os valores padrão no session_state (uma única vez)."""
    defaults = {
        "nome_motorista": "",
        "km_rodado": 0.0,
        "meta_km_editavel": DEFAULT_META_KM,
        "numero_ocorrencia": "",
        "tempo_parado": "",
        "descricao_atraso": "",
    }
    for chave, valor in defaults.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def main():
    inject_custom_css()
    render_header()
    init_session_state()

    col_esquerda, col_direita = st.columns(2, gap="large")

    # ----------------------------------------------------------------------
    # COLUNA ESQUERDA: ENTRADA DE DADOS E OCR (CAPTURA)
    # ----------------------------------------------------------------------
    with col_esquerda:
        with st.container(border=True):
            st.markdown("### 📥 Captura de Dados")
            st.caption("Selecione a data do plantão e envie o print enviado pelo CCO no WhatsApp.")

            data_plantao = st.date_input(
                "Data do Plantão", value=date.today(), format="DD/MM/YYYY", key="data_plantao_input"
            )

            arquivo_print = st.file_uploader(
                "Arraste e solte o print do WhatsApp (CCO)",
                type=["png", "jpg", "jpeg"],
                key="arquivo_print",
                help="Print enviado pelo Centro de Controle Operacional (CCO) com Nome e KM do motorista.",
            )

            imagem_carregada = None
            if arquivo_print is not None:
                try:
                    imagem_carregada = Image.open(arquivo_print)
                    st.image(imagem_carregada, caption="Pré-visualização do print", use_container_width=True)
                except Exception:
                    st.error("Não foi possível abrir a imagem enviada. Verifique o arquivo e tente novamente.")

            if st.button(
                "🔍 Processar Imagem (OCR)",
                disabled=imagem_carregada is None,
                use_container_width=True,
            ):
                with st.spinner("Lendo informações da imagem..."):
                    nome_extraido, km_extraido = process_ocr_image(imagem_carregada)
                st.session_state.nome_motorista = nome_extraido
                st.session_state.km_rodado = float(km_extraido)
                st.success("Dados extraídos com sucesso! Confira na coluna de validação ao lado.")
                st.rerun()

            st.caption(
                "ℹ️ OCR em modo de simulação. Para extração real, configure 'pytesseract' "
                "ou 'easyocr' na função `process_ocr_image()` (estrutura já comentada no código)."
            )

    # ----------------------------------------------------------------------
    # COLUNA DIREITA: VALIDAÇÃO, CÁLCULO DE META E JUSTIFICATIVA
    # ----------------------------------------------------------------------
    with col_direita:
        with st.container(border=True):
            st.markdown("### ✅ Validação e Cálculo de Meta")
            st.caption("Confira os dados extraídos e ajuste manualmente se o OCR não tiver acertado.")

            nome_motorista = st.text_input(
                "Nome do Motorista", key="nome_motorista", placeholder="Será preenchido pelo OCR..."
            )

            col_km, col_meta = st.columns(2)
            with col_km:
                km_rodado = st.number_input(
                    "KM Rodado", key="km_rodado", min_value=0.0, step=0.5, format="%.1f"
                )
            with col_meta:
                meta_km = st.number_input(
                    "Meta de KM (padrão editável)",
                    key="meta_km_editavel",
                    min_value=0.0,
                    step=5.0,
                    format="%.1f",
                )

            meta_atingida = km_rodado >= meta_km

            if meta_atingida:
                st.markdown(
                    f"""<div class="epr-badge epr-badge-success">
                    ✅ META BATIDA &nbsp;|&nbsp; {km_rodado:.1f} km / {meta_km:.1f} km
                    </div>""",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"""<div class="epr-badge epr-badge-danger">
                    ❌ META NÃO BATIDA &nbsp;|&nbsp; {km_rodado:.1f} km / {meta_km:.1f} km
                    </div>""",
                    unsafe_allow_html=True,
                )

            numero_ocorrencia = ""
            tempo_parado = ""
            descricao_atraso = ""

            if not meta_atingida:
                st.markdown("**⚠️ Justificativa obrigatória do não cumprimento da meta:**")
                numero_ocorrencia = st.text_input(
                    "Justificativa (Nº da Ocorrência CCO) *",
                    key="numero_ocorrencia",
                    placeholder="Ex: 20260616-0042",
                )
                tempo_parado = st.text_input(
                    "Tempo Parado na Ocorrência *",
                    key="tempo_parado",
                    placeholder="Ex: 01:30",
                )
                descricao_atraso = st.text_area(
                    "Descrição/Observações do Atraso *",
                    key="descricao_atraso",
                    placeholder="Ex: Ficou parado na ocorrência XPTO e isso atrasou a rodar a meta de kms.",
                    height=110,
                )
            else:
                st.info("Meta cumprida — nenhuma justificativa é necessária.")

    # ----------------------------------------------------------------------
    # SALVAMENTO DO REGISTRO
    # ----------------------------------------------------------------------
    st.write("")
    with st.container(border=True):
        col_btn_a, col_btn_b, col_btn_c = st.columns([1, 2, 1])
        with col_btn_b:
            salvar = st.button(
                "💾 Salvar Registro do Motorista", type="primary", use_container_width=True
            )

        if salvar:
            erros = []
            if not st.session_state.nome_motorista.strip():
                erros.append("Informe o nome do motorista.")
            if not meta_atingida:
                if not st.session_state.numero_ocorrencia.strip():
                    erros.append("Informe o número da ocorrência (CCO).")
                if not st.session_state.tempo_parado.strip():
                    erros.append("Informe o tempo parado na ocorrência.")
                if not st.session_state.descricao_atraso.strip():
                    erros.append("Descreva o motivo do atraso.")

            if erros:
                for erro in erros:
                    st.error(erro)
            else:
                registro = {
                    "data_plantao": st.session_state.data_plantao_input.isoformat(),
                    "motorista": st.session_state.nome_motorista.strip(),
                    "km_rodado": round(float(st.session_state.km_rodado), 1),
                    "meta_km": round(float(st.session_state.meta_km_editavel), 1),
                    "status_meta": "BATIDA" if meta_atingida else "NAO_BATIDA",
                    "numero_ocorrencia": st.session_state.numero_ocorrencia.strip() if not meta_atingida else "",
                    "tempo_parado": st.session_state.tempo_parado.strip() if not meta_atingida else "",
                    "descricao_atraso": st.session_state.descricao_atraso.strip() if not meta_atingida else "",
                    "registrado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                save_record(registro)
                st.success(
                    f"Registro de **{registro['motorista']}** "
                    f"({format_data_br(registro['data_plantao'])}) salvo com sucesso!"
                )
                # limpa os campos para o próximo lançamento (mantém a data selecionada)
                st.session_state.nome_motorista = ""
                st.session_state.km_rodado = 0.0
                st.session_state.numero_ocorrencia = ""
                st.session_state.tempo_parado = ""
                st.session_state.descricao_atraso = ""
                st.rerun()

    st.divider()

    # ----------------------------------------------------------------------
    # CONSULTA DE HISTÓRICO INDIVIDUAL
    # ----------------------------------------------------------------------
    st.markdown("## 🔎 Consulta de Histórico Individual")

    df_hist = load_data()

    if df_hist.empty:
        st.info("Nenhum registro encontrado ainda. Salve o primeiro reporte para iniciar o histórico.")
    else:
        df_hist["km_rodado"] = pd.to_numeric(df_hist["km_rodado"], errors="coerce")
        df_hist["meta_km"] = pd.to_numeric(df_hist["meta_km"], errors="coerce")

        with st.container(border=True):
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                motoristas_disponiveis = sorted(df_hist["motorista"].dropna().unique().tolist())
                motorista_sel = st.selectbox(
                    "Motorista", options=motoristas_disponiveis, key="busca_motorista"
                )
            with col_f2:
                datas_disponiveis = sorted(
                    df_hist.loc[df_hist["motorista"] == motorista_sel, "data_plantao"]
                    .dropna()
                    .unique()
                    .tolist(),
                    reverse=True,
                )
                data_sel = st.selectbox(
                    "Data do Plantão",
                    options=datas_disponiveis,
                    format_func=format_data_br,
                    key="busca_data",
                )

            resultado = df_hist[
                (df_hist["motorista"] == motorista_sel) & (df_hist["data_plantao"] == data_sel)
            ]

            if resultado.empty:
                st.warning("Nenhum reporte encontrado para este motorista nesta data.")
            else:
                registro = resultado.iloc[0].to_dict()
                status_ok = registro["status_meta"] == "BATIDA"
                status_html = (
                    '<span style="color:#1E8449;font-weight:700;">✅ META BATIDA</span>'
                    if status_ok
                    else '<span style="color:#C0392B;font-weight:700;">❌ META NÃO BATIDA</span>'
                )

                linhas_extra = ""
                if not status_ok:
                    linhas_extra = f"""
                    <div class="epr-linha-info"><span>Nº Ocorrência CCO</span><span>{registro.get("numero_ocorrencia", "") or "-"}</span></div>
                    <div class="epr-linha-info"><span>Tempo Parado</span><span>{registro.get("tempo_parado", "") or "-"}</span></div>
                    <div class="epr-linha-info"><span>Descrição/Observações</span><span>{registro.get("descricao_atraso", "") or "-"}</span></div>
                    """

                st.markdown(
                    f"""
                    <div class="epr-card-historico">
                        <h4>📋 Reporte Individual — {registro['motorista']}</h4>
                        <div class="epr-linha-info"><span>Data do Plantão</span><span>{format_data_br(registro['data_plantao'])}</span></div>
                        <div class="epr-linha-info"><span>KM Rodado</span><span>{float(registro['km_rodado']):.1f} km</span></div>
                        <div class="epr-linha-info"><span>Meta Estabelecida</span><span>{float(registro['meta_km']):.1f} km</span></div>
                        <div class="epr-linha-info"><span>Status da Meta</span><span>{status_html}</span></div>
                        {linhas_extra}
                        <div class="epr-linha-info"><span>Registrado em</span><span>{registro.get("registrado_em", "") or "-"}</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                excel_bytes = export_record_to_excel(registro)
                pdf_bytes = export_record_to_pdf(registro)

                col_exp1, col_exp2 = st.columns(2)
                with col_exp1:
                    st.download_button(
                        "📊 Exportar Reporte para Excel",
                        data=excel_bytes,
                        file_name=f"reporte_{registro['motorista'].replace(' ', '_')}_{registro['data_plantao']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with col_exp2:
                    st.download_button(
                        "📄 Exportar Reporte para PDF",
                        data=pdf_bytes,
                        file_name=f"reporte_{registro['motorista'].replace(' ', '_')}_{registro['data_plantao']}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

        with st.expander("📁 Ver base completa de registros"):
            df_visual = df_hist.copy()
            df_visual["data_plantao"] = df_visual["data_plantao"].apply(format_data_br)
            df_visual["status_meta"] = df_visual["status_meta"].map(
                {"BATIDA": "✅ Meta Batida", "NAO_BATIDA": "❌ Meta Não Batida"}
            )
            st.dataframe(df_visual, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Baixar base completa (CSV)",
                data=df_hist.to_csv(index=False).encode("utf-8-sig"),
                file_name="registros_motoristas_completo.csv",
                mime="text/csv",
            )

    st.markdown(
        '<div class="epr-footer">Sistema interno de apoio operacional &middot; EPR Via Mineira</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
