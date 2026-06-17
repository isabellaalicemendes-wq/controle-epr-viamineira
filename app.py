# -*- coding: utf-8 -*-
"""
=================================================================================
 SISTEMA DE CONTROLE DE METAS DE QUILOMETRAGEM - MOTORISTAS  (v3.1)
 EPR VIA MINEIRA - VERSÃO COM FILTRO DE LAYOUT REAL DO CCO
=================================================================================
"""

import os
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =================================================================================
# 1. CONFIGURAÇÕES GERAIS
# =================================================================================

ARQUIVO_PLANTAO = "registros_plantao.csv"
ARQUIVO_OCORRENCIAS = "registros_ocorrencias.csv"

META_IDEAL_KM = 400.0     
META_MINIMA_KM = 380.0    

MOTOR_OCR = "pytesseract"

COLUNAS_PLANTAO = [
    "data_plantao", "turno", "base", "colaborador", "vtr", "km_rodados",
    "status_meta", "tempo_parado_total", "registrado_em",
]
COLUNAS_OCORRENCIAS = [
    "data_plantao", "turno", "colaborador", "numero_ocorrencia", "tempo_parado",
]

DIAS_SEMANA_PT = {
    0: "SEGUNDA-FEIRA", 1: "TERÇA-FEIRA", 2: "QUARTA-FEIRA",
    3: "QUINTA-FEIRA", 4: "SEXTA-FEIRA", 5: "SÁBADO", 6: "DOMINGO",
}

LOGO_URL = "https://eprviamineira.com.br/wp-content/uploads/2024/10/epr-via-mineira-logo.png"


st.set_page_config(
    page_title="EPR Via Mineira | Controle de Metas de KM",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =================================================================================
# 2. IDENTIDADE VISUAL - CSS CUSTOMIZADO (EPR VIA MINEIRA)
# =================================================================================

def inject_custom_css():
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

        .stApp { background-color: var(--epr-cinza-fundo); }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}

        h1, h2, h3, h4 {
            color: var(--epr-azul-escuro) !important;
            font-family: "Segoe UI", Arial, sans-serif;
        }
        p, span, label, .stMarkdown { color: var(--epr-texto); }

        [data-testid="stSidebar"] {
            background-color: var(--epr-azul-escuro);
            border-right: 1px solid var(--epr-cinza-borda);
        }
        [data-testid="stSidebar"] * { color: var(--epr-branco) !important; }
        .epr-sidebar-brand {
            text-align: center;
            padding: 0.4rem 0 1.1rem 0;
            border-bottom: 1px solid rgba(255,255,255,0.18);
            margin-bottom: 0.6rem;
        }
        .epr-sidebar-brand img { height: 36px; margin-bottom: 0.5rem; background: white; border-radius: 6px; padding: 4px 8px; }
        .epr-sidebar-brand p {
            color: var(--epr-branco) !important;
            font-weight: 800;
            font-size: 0.98rem;
            margin: 0;
            letter-spacing: 0.4px;
        }
        .epr-sidebar-brand span { color: #BFD2E3 !important; font-size: 0.74rem; }
        .epr-sidebar-footer {
            text-align: center;
            color: #8CA3B8 !important;
            font-size: 0.7rem;
            margin-top: 1.4rem;
            padding-top: 0.8rem;
            border-top: 1px solid rgba(255,255,255,0.18);
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: var(--epr-branco);
            border: 1px solid var(--epr-cinza-borda);
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(11, 45, 78, 0.06);
        }

        [data-testid="stDataFrame"] {
            border: 1px solid var(--epr-cinza-borda);
            border-radius: 10px;
            overflow: hidden;
        }

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
        .stButton > button[kind="primary"] { background-color: var(--epr-verde); }
        .stButton > button[kind="primary"]:hover { background-color: #239B56; }

        .epr-badge {
            display: block;
            text-align: center;
            font-size: 1.0rem;
            font-weight: 700;
            letter-spacing: 0.3px;
            padding: 0.65rem 1rem;
            border-radius: 8px;
            margin: 0.6rem 0 1rem 0;
        }
        .epr-badge-success { background-color: var(--epr-verde-claro); color: var(--epr-verde); border: 1.5px solid var(--epr-verde); }
        .epr-badge-warning { background-color: var(--epr-laranja-claro); color: var(--epr-laranja); border: 1.5px solid var(--epr-laranja); }
        .epr-badge-danger { background-color: var(--epr-vermelho-claro); color: var(--epr-vermelho); border: 1.5px solid var(--epr-vermelho); }

        .epr-tag {
            display: inline-block;
            padding: 0.2rem 0.65rem;
            border-radius: 999px;
            font-size: 0.74rem;
            font-weight: 700;
            letter-spacing: 0.2px;
        }
        .epr-tag-success { background-color: var(--epr-verde-claro); color: var(--epr-verde); }
        .epr-tag-warning { background-color: var(--epr-laranja-claro); color: var(--epr-laranja); }
        .epr-tag-danger { background-color: var(--epr-vermelho-claro); color: var(--epr-vermelho); }

        .epr-card-historico {
            background-color: var(--epr-branco);
            border: 1px solid var(--epr-cinza-borda);
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 1.3rem;
        }
        .epr-bloco-titulo {
            background: linear-gradient(135deg, var(--epr-azul-escuro) 0%, var(--epr-azul) 100%);
            color: white !important;
            font-weight: 700;
            font-size: 0.98rem;
            padding: 0.7rem 1.1rem;
        }
        .epr-table { width: 100%; border-collapse: collapse; }
        .epr-table th {
            background-color: #F7F9FA;
            color: var(--epr-azul-escuro);
            text-align: left;
            font-size: 0.74rem;
            font-weight: 700;
            padding: 0.55rem 1rem;
            border-bottom: 2px solid var(--epr-cinza-borda);
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        .epr-table td {
            padding: 0.55rem 1rem;
            font-size: 0.88rem;
            border-bottom: 1px solid #EDEFF2;
            color: var(--epr-texto);
        }

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


def render_sidebar_brand():
    st.sidebar.markdown(
        f"""
        <div class="epr-sidebar-brand">
            <img src="{LOGO_URL}" onerror="this.style.display='none'" />
            <p>EPR VIA MINEIRA</p>
            <span>Controle de Metas de KM</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_footer():
    st.sidebar.markdown(
        '<div class="epr-sidebar-footer">Sistema interno &middot; uso operacional</div>',
        unsafe_allow_html=True,
    )


# =================================================================================
# 3. FUNÇÕES UTILITÁRIAS (regras de negócio / formatação)
# =================================================================================

def calcular_status(km_rodados: float):
    km_rodados = float(km_rodados or 0)
    if km_rodados >= META_IDEAL_KM:
        return "BATIDA", "META BATIDA", "epr-badge-success"
    if km_rodados >= META_MINIMA_KM:
        return "ACEITAVEL", "META ACEITÁVEL", "epr-badge-warning"
    return "NAO_BATIDA", "META NÃO BATIDA", "epr-badge-danger"


def format_data_br(iso_date_str: str) -> str:
    try:
        return datetime.strptime(str(iso_date_str), "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return str(iso_date_str)


def nome_dia_semana_pt(d: date) -> str:
    return DIAS_SEMANA_PT.get(d.weekday(), "")


def hhmm_to_minutes(valor: str) -> int:
    try:
        partes = str(valor).strip().split(":")
        horas = int(partes[0])
        minutos = int(partes[1]) if len(partes) > 1 else 0
        return horas * 60 + minutos
    except (ValueError, IndexError):
        return 0


def minutes_to_hhmm(total_minutos: int) -> str:
    total_minutos = max(0, int(total_minutos))
    horas, minutos = divmod(total_minutos, 60)
    return f"{horas:02d}:{minutos:02d}"


def somar_tempos(lista_tempos) -> str:
    total = sum(hhmm_to_minutes(t) for t in lista_tempos)
    return minutes_to_hhmm(total)


# =================================================================================
# 4. CAMADA DE DADOS
# =================================================================================

def load_plantao_df() -> pd.DataFrame:
    if os.path.exists(ARQUIVO_PLANTAO):
        df = pd.read_csv(ARQUIVO_PLANTAO, dtype=str)
        for col in COLUNAS_PLANTAO:
            if col not in df.columns:
                df[col] = ""
        return df[COLUNAS_PLANTAO]
    return pd.DataFrame(columns=COLUNAS_PLANTAO)


def load_ocorrencias_df() -> pd.DataFrame:
    if os.path.exists(ARQUIVO_OCORRENCIAS):
        df = pd.read_csv(ARQUIVO_OCORRENCIAS, dtype=str)
        for col in COLUNAS_OCORRENCIAS:
            if col not in df.columns:
                df[col] = ""
        return df[COLUNAS_OCORRENCIAS]
    return pd.DataFrame(columns=COLUNAS_OCORRENCIAS)


def salvar_plantao_atual(data_plantao_iso: str, turno: str, lista_motoristas: list) -> list:
    if not lista_motoristas:
        return ["Não há motoristas confirmados para salvar."]

    erros = []
    for idx, motorista in enumerate(lista_motoristas, start=1):
        nome = str(motorista.get("colaborador", "")).strip()
        if not nome:
            erros.append(f"Linha {idx}: informe o nome do colaborador.")
            continue
        status_code, _, _ = calcular_status(motorista.get("km_rodados"))
        if status_code != "BATIDA" and not motorista.get("ocorrencias"):
            erros.append(f"{nome}: meta não batida/aceitável precisa de ao menos 1 ocorrência registrada.")

    if erros:
        return erros

    linhas_plantao, lines_ocorrencias = [], []
    for motorista in lista_motoristas:
        status_code, _, _ = calcular_status(motorista.get("km_rodados"))
        tempos = [oc["tempo"] for oc in motorista.get("ocorrencias", [])]
        linhas_plantao.append({
            "data_plantao": data_plantao_iso,
            "turno": turno,
            "base": str(motorista.get("base", "")).strip() or "-",
            "colaborador": str(motorista.get("colaborador", "")).strip(),
            "vtr": str(motorista.get("vtr", "")).strip() or "-",
            "km_rodados": round(float(motorista.get("km_rodados") or 0), 1),
            "status_meta": status_code,
            "tempo_parado_total": somar_tempos(tempos),
            "registrado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        for oc in motorista.get("ocorrencias", []):
            lines_ocorrencias.append({
                "data_plantao": data_plantao_iso,
                "turno": turno,
                "colaborador": str(motorista.get("colaborador", "")).strip(),
                "numero_ocorrencia": oc["numero"],
                "tempo_parado": oc["tempo"],
            })

    df_plantao = load_plantao_df()
    chave_atual = (df_plantao["data_plantao"] == data_plantao_iso) & (df_plantao["turno"] == turno)
    df_plantao = df_plantao[~chave_atual]
    df_plantao = pd.concat([df_plantao, pd.DataFrame(linhas_plantao, columns=COLUNAS_PLANTAO)], ignore_index=True)
    df_plantao.to_csv(ARQUIVO_PLANTAO, index=False)

    df_ocorrencias = load_ocorrencias_df()
    chave_oc_atual = (df_ocorrencias["data_plantao"] == data_plantao_iso) & (df_ocorrencias["turno"] == turno)
    df_ocorrencias = df_ocorrencias[~chave_oc_atual]
    df_ocorrencias = pd.concat(
        [df_ocorrencias, pd.DataFrame(lines_ocorrencias, columns=COLUNAS_OCORRENCIAS)], ignore_index=True
    )
    df_ocorrencias.to_csv(ARQUIVO_OCORRENCIAS, index=False)

    return []


# =================================================================================
# 5. INTELIGÊNCIA OCR ATUALIZADA (FILTRO DIRETO PARA PLANILHA REAL DO CCO)
# =================================================================================

def preprocessar_imagem_para_ocr(image: Image.Image) -> Image.Image:
    imagem = image.convert("L")
    imagem = ImageOps.autocontrast(imagem)
    largura, altura = imagem.size
    if largura < 1500:
        fator = 1500 / largura
        imagem = imagem.resize((int(largura * faktor), int(altura * fator)), Image.LANCZOS)
    return imagem


def interpretar_texto_completo_real(linhas_texto: list) -> list:
    """
    Motor inteligente feito sob medida para ler o print real do CCO da EPR:
    - Vasculha o
