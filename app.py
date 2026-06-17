import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import easyocr
from PIL import Image, ImageEnhance, ImageOps
import io

# ==========================================
# 🎨 CONFIGURAÇÃO DE TEMA E IDENTIDADE VISUAL (EPR)
# ==========================================
st.set_page_config(page_title="EPR Via Mineira - Controle de Metas", layout="wide")

# Injeção de CSS para customização premium da marca
st.markdown("""
    <style>
    /* Cor de fundo e fontes */
    .stApp { background-color: #F8F9FA; }
    
    /* Customização do Header Lateral e Principal */
    [data-testid="stSidebar"] { background-color: #002F6C !important; color: white; }
    [data-testid="stSidebar"] * { color: white !important; }
    
    /* Cartões de Conteúdo (Cards) */
    .epr-card {
        background-color: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border-top: 5px solid #002F6C;
    }
    
    /* Títulos */
    h1, h2, h3 { color: #002F6C !important; font-family: 'Segoe UI', sans-serif; }
    
    /* Badges de Status */
    .badge-batida { background-color: #2E7D32; color: white; padding: 6px 12px; border-radius: 20px; font-weight: bold; }
    .badge-aceitavel { background-color: #EF6C00; color: white; padding: 6px 12px; border-radius: 20px; font-weight: bold; }
    .badge-nao-batida { background-color: #C62828; color: white; padding: 6px 12px; border-radius: 20px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# Initialize Session States para guardar os dados entre cliques
if 'dados_plantao' not in s:
    s['dados_plantao'] = pd.DataFrame(columns=['Data', 'Turno', 'Motorista', 'VTR', 'KM_Rodados', 'Status'])
if 'ocorrencias' not in s:
    s['ocorrencias'] = {}

# ==========================================
# 📍 MENU LATERAL DE NAVEGAÇÃO
# ==========================================
with st.sidebar:
    st.image("https://eprviamineira.com.br/wp-content/themes/epr-via-mineira/assets/images/logo.svg", width=180) # Fallback se a URL mudar
    st.title("Painel Operacional")
    menu = st.radio("Navegação", ["📝 Lançar Plantão Diário", "📊 Consultar Histórico / Turnos"])
    st.write("---")
    st.caption("EPR Via Mineira © 2026")

# ==========================================
# 📝 TELA 1: LANÇAR PLANTÃO DIÁRIO
# ==========================================
if menu == "📝 Lançar Plantão Diário":
    st.title("📝 Controle de Metas de Quilometragem")
    
    # Inputs de Cabeçalho do Turno
    col_data, col_turno = st.columns(2)
    with col_data:
        data_plantao = st.date_input("Data do Plantão", datetime.today())
    with col_turno:
        turno_plantao = st.selectbox("Turno", ["DIURNO (06h às 18h)", "NOTURNO (18h às 06h)"])
        
    st.write("---")
    
    col_esq, col_dir = st.columns([1, 1.2])
    
    # --- COLUNA ESQUERDA: UPLOAD E MOTOR DE OCR REAL ---
    with col_esq:
        st.markdown('<div class="epr-card"><h3>📷 Captura de Dados (Print CCO)</h3>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Arraste e solte o print da tabela do CCO aqui", type=["jpg", "jpeg", "png"])
        
        df_ocr = pd.DataFrame(columns=['Motorista', 'VTR', 'KM_Rodados'])
        
        if uploaded_file is not None:
            st.image(uploaded_file, caption="Pré-visualização do Print", use_container_width=True)
            
            with st.spinner("🧠 Processando imagem com OCR Estruturado..."):
                try:
                    # Pré-processamento da imagem para melhorar a leitura do OCR
                    image = Image.open(uploaded_file)
                    image = ImageOps.grayscale(image) # Converte para cinza
                    image = ImageEnhance.Contrast(image).enhance(2.0) # Aumenta contraste
                    
                    # Converte para array que o EasyOCR aceita
                    img_byte_arr = io.BytesIO()
                    image.save(img_byte_arr, format='PNG')
                    img_bytes = img_byte_arr.getvalue()
                    
                    # Inicializa o leitor em Português
                    reader = easyocr.Reader(['pt'])
                    result = reader.readtext(img_bytes)
                    
                    # Lógica de extração baseada nas colunas aproximadas do print
                    # Se falhar ou misturar, criamos linhas padrão baseadas no formato do CCO
                    linhas_detectadas = []
                    
                    # Filtra os textos numéricos (KMs) e nomes
                    textos = [res[1] for res in result]
                    
                    # Como OCR de tabela complexa pode falhar na ordem, separamos o que parece nome/VTR/KM
                    kms = [t for t in textos if t.isdigit() and len(t) <= 3 and int(t) > 50]
                    vtrs = [t for t in textos if t.startswith('T') or t.startswith('t')]
                    nomes = [t for t in textos if len(t) > 4 and not t.isdigit() and 'BSO' not in t and '/' not in t]
                    
                    # Monta o DataFrame dinâmico com o máximo de linhas que conseguir parear
                    max_len = max(len(nomes), len(vtrs), len(kms), 5) # Garante pelo menos 5 linhas
                    
                    for i in range(max_len):
                        nome = nomes[i] if i < len(nomes) else f"Motorista Oculto {i+1}"
                        vtr = vtrs[i] if i < len(vtrs) else "T01"
                        km = int(kms[i]) if i < len(kms) else 350 # valor padrão para correção caso falhe
                        linhas_detectadas.append({'Motorista': nome, 'VTR': vtr, 'KM_Rodados': km})
                        
                    df_ocr = pd.DataFrame(linhas_detectadas)
                    st.success("✅ Imagem processada! Ajuste qualquer divergência na tabela ao lado.")
                    
                except Exception as e:
                    st.error("⚠️ Não foi possível ler todas as linhas automaticamente devido à qualidade da imagem.")
                    # Cria tabela limpa para preenchimento manual se der erro
                    df_ocr = pd.DataFrame([{'Motorista': '', 'VTR': '', 'KM_Rodados': 0} for _ in range(5)])
        else:
            # Caso não tenha subido arquivo, mostra estrutura vazia
            df_ocr = pd.DataFrame([{'Motorista': 'Ex: Rômulo', 'VTR': 'T30', 'KM_Rodados': 337},
                                   {'Motorista': 'Ex: Gilsimar', 'VTR': 'T16', 'KM_Rodados': 98}])
        st.markdown('</div>', unsafe_allow_html=True)
        
    # --- COLUNA DIREITA: VALIDAÇÃO, METAS E OCORRÊNCIAS ---
    with col_dir:
        st.markdown('<div class="epr-card"><h3>📊 Validação e Cálculo de Metas</h3>', unsafe_allow_html=True)
        st.write("Ajuste os dados se necessário. Status calculado com base nas metas (Ideal: 400km | Aceitável: 380km).")
        
        # Tabela interativa para o seu pai corrigir nomes e KMs na hora
        df_editado = st.data_editor(df_ocr, num_rows="dynamic", use_container_width=True)
        
        # Aplica as regras de negócio reais da EPR de forma automática
        if not df_editado.empty:
            status_list = []
            for km in df_editado['KM_Rodados']:
                try:
                    km_val = float(km)
                    if km_val >= 400: status_list.append("META BATIDA")
                    elif km_val >= 380: status_list.append("META ACEITÁVEL")
                    else: status_list.append("META NÃO BATIDA")
                except:
                    status_list.append("META NÃO BATIDA")
            df_editado['Status'] = status_list
            
            # Seção dinâmica para adicionar as ocorrências estilo "Planilha 2"
            st.markdown("### ⚠️ Justificativas de Ocorrências")
            
            for idx, row in df_editado.iterrows():
                if row['Status'] in ["META NÃO BATIDA", "META ACEITÁVEL"] and row['Motorista']:
                    key_id = f"{row['Motorista']}_{idx}"
                    st.markdown(f"**Motorista: {row['Motorista']} ({row['VTR']}) - Rodou {row['KM_Rodados']} KM**")
                    
                    col_oc, col_tempo = st.columns([2, 1])
                    with col_oc:
                        oc = st.text_input(f"Nº Ocorrência / Motivo", key=f"oc_{key_id}", placeholder="Ex: oc 170 tombamento")
                    with col_tempo:
                        tempo = st.text_input(f"Tempo Parado", key=f"tempo_{key_id}", placeholder="Ex: 01:04")
                        
                    s['ocorrencias'][key_id] = {"oc": oc, "tempo": tempo}
            
            if st.button("💾 Salvar Plantão e Gerar Reportes", use_container_width=True):
                # Processa e consolida tudo no histórico permanente
                linhas_finais = []
                data_str = data_plantao.strftime('%d/%m/%Y')
                
                for idx, row in df_editado.iterrows():
                    key_id = f"{row['Motorista']}_{idx}"
                    oc_info = s['ocorrencias'].get(key_id, {"oc": "", "tempo": ""})
                    
                    linhas_finais.append({
                        'Data': data_str,
                        'Turno': turno_plantao,
                        'Motorista': row['Motorista'],
                        'VTR': row['VTR'],
                        'KM_Rodados': row['KM_Rodados'],
                        'Status': row['Status'],
                        'Ocorrencia': oc_info['oc'],
                        'Tempo_Parado': oc_info['tempo']
                    })
                
                novos_dados = pd.DataFrame(linhas_finais)
                s['dados_plantao'] = pd.concat([s['dados_plantao'], novos_dados], ignore_index=True)
                st.success("✨ Plantão salvo com sucesso no banco histórico!")
                
        st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# 📊 TELA 2: CONSULTAR HISTÓRICO (FORMATO PLANILHA 2)
# ==========================================
elif menu == "📊 Consultar Histórico / Turnos":
    st.title("📊 Histórico de Reportes por Turno")
    
    if s['dados_plantao'].empty:
        st.info("Nenhum plantão lançado até o momento. Vá na aba de lançamentos.")
    else:
        # Agrupa os dados para exibir exatamente no formato empilhado vertical da sua "Planilha 2"
        df_historico = s['dados_plantao']
        
        # Filtros de busca
        datas_disponiveis = df_historico['Data'].unique()
        data_sel = st.selectbox("Filtrar por Data", datas_disponiveis)
        
        df_filtrado = df_historico[df_historico['Data'] == data_sel]
        
        for turno in df_filtrado['Turno'].unique():
            df_turno = df_filtrado[df_filtrado['Turno'] == turno]
            
            # Header azul escuro institucional igual ao seu modelo
            st.markdown(f"""
                <div style="background-color:#002F6C; color:white; padding:10px; font-weight:bold; text-align:center; border-radius:4px; margin-top:15px;">
                    {data_sel} - {turno.upper()}
                </div>
            """, unsafe_allow_html=True)
            
            # Montagem visual em blocos imitando a Planilha 2
            for idx, row in df_turno.iterrows():
                col1, col2, col3 = st.columns([2, 2, 1])
                with col1:
                    st.markdown(f"**{row['Motorista']} ({row['VTR']})** - {row['KM_Rodados']} KM")
                with col2:
                    if row['Ocorrencia']:
                        st.markdown(f"↳ *{row['Ocorrencia']}*")
                    else:
                        st.markdown("↳ *Sem alterações de meta*")
                with col3:
                    if row['Tempo_Parado']:
                        st.markdown(f"⏱️ {row['Tempo_Parado']}")
            st.write("---")
            
        # Botão de Exportação Completa
        st.markdown("### 📥 Exportar Dados Consolidados")
        
        # Converte para Excel usando openpyxl em background
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_historico.to_excel(writer, index=False, sheet_name='Histórico_Metas')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📊 Baixar Relatório Completo em Excel (.xlsx)",
            data=processed_data,
            file_name=f"EPR_ViaMineira_Relatorio_Metas_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
