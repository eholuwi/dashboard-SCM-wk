# Referência histórica de cálculo (fluxo atual é gerar_dashboard.py). Rodar com
# `streamlit run legacy/app.py` a partir da raiz do projeto (onde estão as planilhas).
import streamlit as st
import pandas as pd
from datetime import date
import plotly.express as px
import unicodedata
import os

# --- ARQUIVOS DE DADOS ---
ARQUIVO_SOLICITACOES = "Solicitações.xlsx"
ARQUIVO_COMPRAS = "Relatório de Compras.xlsx"

def normalize_str(s):
    if not isinstance(s, str): return ""
    nfkd_form = unicodedata.normalize('NFKD', s)
    only_ascii = nfkd_form.encode('ASCII', 'ignore').decode('utf-8')
    return only_ascii.lower().strip()

@st.cache_data(show_spinner=False)
def load_excel_compras():
    if not os.path.exists(ARQUIVO_COMPRAS):
        return None
    return pd.read_excel(ARQUIVO_COMPRAS)

@st.cache_data(show_spinner=False)
def load_excel_solicitacoes():
    if not os.path.exists(ARQUIVO_SOLICITACOES):
        return None
    return pd.read_excel(ARQUIVO_SOLICITACOES, header=0)

# --- CONFIGURAÇÃO DA PÁGINA E IDENTIDADE VISUAL (INVENTUS POWER) ---
st.set_page_config(page_title="Dashboard SCM", layout="wide", initial_sidebar_state="collapsed")

# CSS Customizado para Identidade Visual Inventus (Estilo Clean/Card - FULL WIDTH)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    :root {
        --primary: #0B5FFF;
        --secondary: #0A1B2B;
        --bg: #F5F7FA;
        --surface: #FFFFFF;
        --text: #111827;
        --muted: #6B7280;
        --border: #E5E7EB;
        --shadow: 0 18px 50px rgba(10, 27, 43, 0.08);
    }

    html, body, [data-testid="stAppViewContainer"] {
        background: var(--bg);
        color: var(--text);
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background-color: var(--bg);
    }

    /* MUDANÇA PRINCIPAL AQUI: Largura Máxima Aumentada */
    .css-1d391kg, .css-2trqyj, .css-1outpf7, .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 98%; /* Antes era 1200px, agora ocupa quase toda a tela */
        padding-left: 2rem;
        padding-right: 2rem;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0A1B2B 0%, #0D2D4B 100%);
        border-right: 1px solid rgba(255,255,255,0.05);
    }

    [data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }

    #MainMenu, footer, header {
        visibility: hidden;
    }

    h1, h2, h3, h4, h5 {
        font-family: 'Inter', sans-serif;
        font-weight: 700;
        color: var(--secondary);
        margin-bottom: 0.75rem;
    }

    p, span, label, .css-1p0v2v2 {
        font-family: 'Inter', sans-serif;
        color: var(--text);
        line-height: 1.6;
    }

    .stMetric {
        border: 1px solid var(--border) !important;
        border-radius: 20px !important;
        background: white !important;
        padding: 1.4rem !important;
        box-shadow: 0 18px 50px rgba(15, 23, 42, 0.05) !important;
    }

    div[data-testid="stMetricValue"] {
        color: var(--secondary) !important;
    }

    div[data-testid="stMetricLabel"] {
        color: var(--muted) !important;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        font-size: 0.8rem;
        margin-top: 0.4rem;
    }

    .stPlotlyChart > div {
        border-radius: 22px !important;
        box-shadow: 0 18px 40px rgba(15, 23, 42, 0.05);
        background: white !important;
    }

    .stDivider {
        margin: 2rem 0 !important;
    }

    /* INFO BUBBLE */
    .info-bubble {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 20px;
        height: 20px;
        border-radius: 50%;
        background: #F3F4F6;
        color: var(--secondary);
        font-weight: 700;
        font-size: 12px;
        margin-left: 8px;
        cursor: help;
        position: relative;
        border: 1px solid var(--border);
    }
    .info-bubble:hover::after {
        content: attr(data-tooltip);
        position: absolute;
        bottom: calc(100% + 8px);
        left: 50%;
        transform: translateX(-50%);
        background: rgba(17,24,39,0.95);
        color: #fff;
        padding: 8px 10px;
        border-radius: 6px;
        white-space: nowrap;
        font-size: 12px;
        z-index: 60;
    }

</style>
""", unsafe_allow_html=True)

# Helper: subheader with tooltip (HTML)
def subheader_with_info(text, tooltip=None):
    if tooltip:
        safe = str(tooltip).replace('"', '&quot;')
        st.markdown(f'<div style="display:flex;align-items:center;gap:8px;"><h3 style="margin:0;">{text}</h3><span class="info-bubble" data-tooltip="{safe}">?</span></div>', unsafe_allow_html=True)
    else:
        st.subheader(text)


# Helper: metric with tooltip (shows small info under the metric)
def metric_with_info(label, value, delta=None, tooltip=None):
    if delta is not None:
        st.metric(label=label, value=value, delta=delta)
    else:
        st.metric(label=label, value=value)
    if tooltip:
        safe = str(tooltip).replace('"', '&quot;')
        st.markdown(f'<div style="margin-top:6px;"><span class="info-bubble" data-tooltip="{safe}">?</span></div>', unsafe_allow_html=True)

# Título Principal
st.title(" Dashboard SCM", anchor=None)

# --- FILTRO DE PERÍODO ---
_fcol, _ = st.columns([2, 4])
with _fcol:
    _date_range = st.date_input(
        "Período de análise",
        value=(date(2026, 1, 1), date.today()),
        min_value=date(2020, 1, 1),
        max_value=date(2030, 12, 31),
        format="DD/MM/YYYY"
    )
start_date = _date_range[0] if len(_date_range) > 0 else date(2026, 1, 1)
end_date = _date_range[1] if len(_date_range) > 1 else date.today()


# --- LÓGICA DE CARREGAMENTO DOS KPIS MANUAIS E GRÁFICOS (DO RELATÓRIO DE COMPRAS) ---
VALOR_PO_QTD = 0
VALOR_ITENS_PO_QTD = 0
df_compras_processed = None
col_num_pc = None
col_dt_emissao_pc = None

with st.spinner("Carregando dados..."):
    _df_compras_orig = load_excel_compras()
    _df_sol_orig = load_excel_solicitacoes()

if _df_compras_orig is not None:
    try:
        df_compras_raw = _df_compras_orig.copy()

        cols_compras_norm = [normalize_str(c) for c in df_compras_raw.columns]
        col_pc_map = dict(zip(cols_compras_norm, df_compras_raw.columns))

        col_comprador_pc = None
        for norm_name, orig_name in col_pc_map.items():
            if 'numero pc' in norm_name or 'num pc' in norm_name or 'pedido' in norm_name:
                col_num_pc = orig_name
            if 'dt emissao' in norm_name or 'data emissao' in norm_name or 'emissao' in norm_name:
                col_dt_emissao_pc = orig_name
            if 'comprador' in norm_name or 'buyer' in norm_name:
                col_comprador_pc = orig_name

        if col_dt_emissao_pc:
            _dt_pc = pd.to_datetime(df_compras_raw[col_dt_emissao_pc], errors='coerce')
            _mask_pc = (_dt_pc.dt.date >= start_date) & (_dt_pc.dt.date <= end_date)
            df_compras_raw = df_compras_raw[_mask_pc].copy()

        _COMPRADORES_PC = ['miguel magalhaes do nascimento', 'davi rocha de oliveira']
        if col_comprador_pc:
            _comp_norm = df_compras_raw[col_comprador_pc].astype(str).apply(normalize_str)
            df_compras_raw = df_compras_raw[_comp_norm.isin(_COMPRADORES_PC)].copy()

        VALOR_ITENS_PO_QTD = len(df_compras_raw)

        if col_num_pc:
            pcs_unicos = df_compras_raw[col_num_pc].dropna().astype(str).str.strip()
            pcs_unicos = pcs_unicos[pcs_unicos != '']
            VALOR_PO_QTD = pcs_unicos.nunique()
            
        if col_dt_emissao_pc and col_num_pc:
            df_compras_processed = df_compras_raw.copy()
            df_compras_processed['Data_PC'] = pd.to_datetime(df_compras_processed[col_dt_emissao_pc], errors='coerce')
            df_compras_processed['Mes_PC'] = df_compras_processed['Data_PC'].dt.month_name(locale='pt_BR.utf8')
            
            pos_por_mes = df_compras_processed.groupby('Mes_PC')[col_num_pc].nunique().reset_index(name='Qtd_POs')
            itens_por_mes_pc = df_compras_processed.groupby('Mes_PC').size().reset_index(name='Qtd_Itens')
            
            ordem_meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
            
            pos_por_mes['Mes_PC'] = pd.Categorical(pos_por_mes['Mes_PC'], categories=ordem_meses, ordered=True)
            pos_por_mes = pos_por_mes.sort_values('Mes_PC').dropna(subset=['Mes_PC'])
            
            itens_por_mes_pc['Mes_PC'] = pd.Categorical(itens_por_mes_pc['Mes_PC'], categories=ordem_meses, ordered=True)
            itens_por_mes_pc = itens_por_mes_pc.sort_values('Mes_PC').dropna(subset=['Mes_PC'])
            
    except Exception as e:
        st.warning(f"Erro ao ler '{ARQUIVO_COMPRAS}' para KPIs/Gráficos: {e}")

# --- CARREGAMENTO AUTOMÁTICO DO ARQUIVO PRINCIPAL (SOLICITAÇÕES) ---
if _df_sol_orig is not None:
    try:
        df_raw = _df_sol_orig.copy()
        
        original_cols = [str(c).strip() for c in df_raw.columns]
        normalized_cols = [normalize_str(c) for c in original_cols]
        col_map = dict(zip(normalized_cols, original_cols))
        
        def find_col(keywords):
            for norm_name, orig_name in col_map.items():
                for keyword in keywords:
                     if keyword in norm_name: return orig_name
            return None

        col_num_sc = find_col(['numero da solicitacao', 'num sc', 'solicitacao', 'nr solicitacao'])
        col_status = find_col(['status', 'sta tus', 'situação'])
        col_depto = find_col(['departamento', 'depto', 'area', 'setor'])
        col_comprador = find_col(['comprador', 'buyer', 'responsavel'])
        col_emissao = find_col(['emissao', 'data emissao', 'dt emissao', 'data criacao']) 
        col_aprovacao = find_col(['aprovacao', 'data aprov', 'dt aprov', 'data liberacao'])
        col_entrega = find_col(['entrega nfe', 'previsao nfe', 'data entrega', 'previsao entrega'])
        col_data_necess = find_col(['data necessidade', 'necessidade'])
        valor_col = find_col(['vlr. total', 'valor total', 'preco total', 'total value'])

        df = df_raw.copy()
        df.columns = original_cols
        
        if col_num_sc:
            df = df[df[col_num_sc].astype(str) != col_num_sc]

        if col_aprovacao and not df.empty:
            _dt_aprov = pd.to_datetime(df[col_aprovacao], errors='coerce')
            _mask_sol = (_dt_aprov.dt.date >= start_date) & (_dt_aprov.dt.date <= end_date)
            df = df[_mask_sol].copy()

        if not df.empty:
            st.info(f"📅 Período selecionado: **{start_date.strftime('%d/%m/%Y')}** até **{end_date.strftime('%d/%m/%Y')}** — {len(df):,} registros no período")

        df['status_norm'] = df[col_status].astype(str).str.strip().apply(normalize_str) if col_status else ""
        df_cotacao = df[df['status_norm'] == 'cotacao'].copy() if col_status else pd.DataFrame()
        
        total_geral_itens = len(df)
        total_geral_scs = df[col_num_sc].nunique() if col_num_sc else 0

        if not df_cotacao.empty:
            if col_aprovacao:
                try:
                    df_cotacao['data_aprov_parsed'] = pd.to_datetime(df_cotacao[col_aprovacao], errors='coerce')
                    df_cotacao['aging_dias'] = (pd.Timestamp.now().normalize() - df_cotacao['data_aprov_parsed']).dt.days
                except: pass

        scs_por_mes = pd.DataFrame()
        itens_por_mes = pd.DataFrame()
        
        if col_emissao and col_num_sc and not df_cotacao.empty:
            try:
                df_cotacao['data_emissao_parsed'] = pd.to_datetime(df_cotacao[col_emissao], errors='coerce')
                df_cotacao['mes_emissao'] = df_cotacao['data_emissao_parsed'].dt.month_name(locale='pt_BR.utf8')
                ordem_meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
                
                scs_por_mes = df_cotacao.groupby('mes_emissao')[col_num_sc].nunique().reset_index(name='qtd_scs')
                scs_por_mes['mes_emissao'] = pd.Categorical(scs_por_mes['mes_emissao'], categories=ordem_meses, ordered=True)
                scs_por_mes = scs_por_mes.sort_values('mes_emissao').dropna(subset=['mes_emissao'])
                
                itens_por_mes = df_cotacao.groupby('mes_emissao').size().reset_index(name='qtd_itens')
                itens_por_mes['mes_emissao'] = pd.Categorical(itens_por_mes['mes_emissao'], categories=ordem_meses, ordered=True)
                itens_por_mes = itens_por_mes.sort_values('mes_emissao').dropna(subset=['mes_emissao'])
            except: pass

        qtd_itens_abertos = len(df_cotacao)
        qtd_scs_abertas = df_cotacao[col_num_sc].nunique() if col_num_sc else 0
        
        col_pedido = find_col(['pedido', 'num pedido', 'po', 'order number'])
        df_com_po = pd.DataFrame()
        if col_pedido and not df.empty:
            df_com_po = df[(df[col_pedido].astype(str).str.strip() != '') & (~df[col_pedido].astype(str).str.strip().str.lower().isin(['nan', 'none', 'null', '-']))]

        kpi_cols = st.columns(4) 
        
        with kpi_cols[0]:
            pct_cotacao = (qtd_itens_abertos / total_geral_itens * 100) if total_geral_itens > 0 else 0
            metric_with_info(label="📄 Qty Itens Aberto (CT)", value=f"{qtd_itens_abertos:,}", delta=f"{pct_cotacao:.1f}% do total", tooltip="Quantidade de itens abertos que estão com o status 'cotação' na planilha SCM")
            
        with kpi_cols[1]:
            pct_scs_cotacao = (qtd_scs_abertas / total_geral_scs * 100) if total_geral_scs > 0 else 0
            metric_with_info(label="📑 Qty SCs Abertas (CT)", value=f"{qtd_scs_abertas:,}", delta=f"{pct_scs_cotacao:.1f}% do total", tooltip="Quantidade de solicitações de compra (SC) com status 'cotação' na planilha SCM")
            
        with kpi_cols[2]:
            metric_with_info(label="🛒 Qty P.Os emitidos", value=f"{VALOR_PO_QTD:,}", tooltip="Número de Pedidos de Compra (PO) únicos emitidos no período — extraído do Relatório de Compras")
            
        with kpi_cols[3]:
            metric_with_info(label="💡 Qty Itens com P.Os emitidos", value=f"{VALOR_ITENS_PO_QTD:,}", tooltip="Quantidade de itens no Relatório de Compras que possuem P.O. emitido no período")

        # with kpi_cols[4]:
        #     if valor_col and not df_com_po.empty:
        #          total_valor = df_com_po[valor_col].sum()
        #          dispendio_real = f"R$ {total_valor / 1_000_000:.1f}M".replace('.', ',')
        #          st.metric(label=" Dispêndio Total", value=dispendio_real)
        #     else:
        #          st.metric(label="💸 Dispêndio Total", value="$27,0M") 

        st.divider()

        chart_cols_row1 = st.columns(2)
        
        with chart_cols_row1[0]:
            subheader_with_info("1. Qtd. Itens por Mês (SCs)", "Quantidade de itens em cotação por mês (base: planilha Solicitações, status 'cotação').")
            if not itens_por_mes.empty:
                fig_itens = px.bar(itens_por_mes, x='mes_emissao', y='qtd_itens', text='qtd_itens', labels={'mes_emissao': 'Mês'}, color_discrete_sequence=['#F59A23'])
                fig_itens.update_traces(textposition='outside', insidetextanchor='middle', textfont=dict(size=12, color='#1F2933', family="Arial Black"))
                fig_itens.update_layout(
                    height=300, 
                    margin=dict(l=20, r=20, t=20, b=20), 
                    showlegend=False, 
                    plot_bgcolor='white',
                    xaxis=dict(
                        tickfont=dict(family="Arial", size=12, weight="bold")
                    )
                )
                st.plotly_chart(fig_itens, width='stretch')
        
        with chart_cols_row1[1]:
            subheader_with_info("2. Qtd. SCs por Mês", "Quantidade de SCs únicas em cotação por mês (base: planilha Solicitações).")
            if not scs_por_mes.empty:
                fig_scs = px.bar(scs_por_mes, x='mes_emissao', y='qtd_scs', text='qtd_scs', labels={'mes_emissao': 'Mês'}, color_discrete_sequence=['#F59A23'])
                fig_scs.update_traces(textposition='outside', insidetextanchor='middle', textfont=dict(size=12, color='#1F2933', family="Arial Black"))
                fig_scs.update_layout(
                    height=300, 
                    margin=dict(l=20, r=20, t=20, b=20), 
                    showlegend=False, 
                    plot_bgcolor='white',
                    xaxis=dict(
                        tickfont=dict(family="Arial", size=12, weight="bold")
                    )
                )
                st.plotly_chart(fig_scs, width='stretch')

        chart_cols_row2 = st.columns(2)
        
        with chart_cols_row2[0]:
            subheader_with_info("3. Ranking Departamentos (Top 10)", "Top 10 departamentos por número de itens em cotação (planilha Solicitações).")
            if not df_cotacao.empty and col_depto:
                df_depto_itens = df_cotacao.groupby(col_depto).size().reset_index(name='Qtd_Itens')
                df_depto_itens = df_depto_itens.sort_values('Qtd_Itens', ascending=False).head(10)
                
                fig_depto = px.bar(df_depto_itens, x=col_depto, y='Qtd_Itens', text='Qtd_Itens', labels={col_depto: 'Departamento'}, color_discrete_sequence=['#F59A23'])
                fig_depto.update_traces(textposition='outside', insidetextanchor='middle', textfont=dict(size=14, color='#1F2933', family="Arial Black"))
                fig_depto.update_layout(height=300, margin=dict(l=20, r=20, t=20, b=20), showlegend=False, plot_bgcolor='white')
                st.plotly_chart(fig_depto, width='stretch')
            else:
                st.info("Sem dados de departamento.")

        with chart_cols_row2[1]:
            subheader_with_info("4. Itens por Comprador", "Distribuição de itens em cotação por comprador. 'Sem comprador' = registros sem comprador informado.")
            if not df_cotacao.empty and col_comprador:
                _cv = df_cotacao[col_comprador].astype(str).str.strip().str.lower()
                _total_cotacao = len(df_cotacao)
                _davi = int(_cv.str.contains('davi', na=False).sum())
                _miguel = int(_cv.str.contains('miguel', na=False).sum())
                _luis = int(_cv.str.contains('luis', na=False).sum())
                _adrya = int(_cv.str.contains('adrya', na=False).sum())
                _sem = max(0, _total_cotacao - (_davi + _miguel + _luis + _adrya))

                df_buyers = pd.DataFrame({
                    'Comprador': ['Davi', 'Miguel', 'Luis Gabriel', 'Adrya', 'Sem comprador'],
                    'Qtd': [_davi, _miguel, _luis, _adrya, _sem]
                })
                _cores_comprador = {
                    'Davi': '#F59A23', 'Miguel': '#F59A23',
                    'Luis Gabriel': '#F59A23', 'Adrya': '#F59A23',
                    'Sem comprador': '#6B7280'
                }
                fig_comprador = px.bar(
                    df_buyers, x='Comprador', y='Qtd', text='Qtd',
                    color='Comprador', color_discrete_map=_cores_comprador,
                    labels={'Qtd': 'Itens', 'Comprador': ''}
                )
                fig_comprador.update_traces(textposition='outside', textfont=dict(size=14, color='#1F2933', family="Arial Black"))
                fig_comprador.update_layout(height=300, margin=dict(l=20, r=20, t=40, b=20), showlegend=False, plot_bgcolor='white')
                st.plotly_chart(fig_comprador, width='stretch')
            else:
                st.info("Sem dados de comprador.")

        # with chart_cols_row2[2]:
        #     st.subheader("5. Dispêndio Mensal")
        #     data_dispendio = {
        #         'Mês': ['Jan/26', 'Fev/26', 'Mar/26', 'Abr/26'],
        #         'Valor': [5789074.14, 5672483.62, 8524970.01, 7009948.36]
        #     }
        #     df_dispendio = pd.DataFrame(data_dispendio)
        #     df_dispendio['Texto_M'] = [f"{v/1_000_000:.1f}M".replace('.', ',') for v in data_dispendio['Valor']]
            
        #     fig_dispendio = px.bar(df_dispendio, x='Mês', y='Valor', text='Texto_M', labels={'Mês': 'Mês', 'Valor': 'R$'}, color_discrete_sequence=['#F59A23'])
        #     fig_dispendio.update_traces(textposition='inside', insidetextanchor='end', textfont=dict(size=14, color='white', family="Arial Black"), width=0.6)
        #     fig_dispendio.update_yaxes(tickprefix='', ticksuffix='', tickformat=',.0f', title_text=None, range=[0, max(data_dispendio['Valor']) * 1.15])
        #     fig_dispendio.update_layout(
        #         height=300, 
        #         margin=dict(l=20, r=20, t=20, b=20), 
        #         showlegend=False, 
        #         plot_bgcolor='white',
        #         xaxis=dict(
        #             tickangle=0, 
        #             tickfont=dict(family="Arial", size=13, weight="bold")
        #         )
        #     )
        #     st.plotly_chart(fig_dispendio, width='stretch')

    
        st.divider()
        subheader_with_info("6. Status dos Pedidos de Compra (POs)", "Comparativo mensal entre POs emitidos, entregues e aguardando — dados extraídos do Relatório de Compras.")
        
        # Verifica se o relatório de compras foi processado e tem as colunas necessárias
        if df_compras_processed is not None and 'Mes_PC' in df_compras_processed.columns and 'Numero PC' in df_compras_processed.columns:
            
            chart_cols_row3 = st.columns(2)
            
            # --- FUNÇÃO AUXILIAR PARA CALCULAR MÉTRICAS MENSAIS (LÓGICA CORRETA) ---
            def calcular_status_pos_correto(df_base):
                """
                Lógica exata conforme solicitado:
                - POs: Conta Numero PC únicos
                - Itens: Conta número de LINHAS (registros)
                - Entregue: Qtd.Entregue > 0 (não é 0 nem vazio)
                - Aberto: Emitidos - Entregues
                """
                resultados = []
                ordem_meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                              'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
                
                for mes in ordem_meses:
                    # Filtra APENAS as linhas daquele mês baseado na DT Emissão do PC
                    df_mes = df_base[df_base['Mes_PC'] == mes].copy()
                    
                    if df_mes.empty:
                        continue
                    
                    # === POs (Contagem de PCs ÚNICOS) ===
                    # POs Emitidos: PCs únicos no mês
                    pos_emitidos = df_mes['Numero PC'].nunique()
                    
                    # === Itens (Contagem de LINHAS/Registros) ===
                    # Itens Emitidos: Total de LINHAS no mês (cada linha = 1 item)
                    itens_emitidos = len(df_mes)
                    
                    # === Lógica de Entrega ===
                    # Converte Qtd.Entregue para numérico (vazios/nulos viram 0)
                    df_mes['Qtd_Ent_Num'] = pd.to_numeric(df_mes['Qtd.Entregue'], errors='coerce').fillna(0)
                    
                    # Filtra linhas ENTREGUES (Qtd.Entregue > 0)
                    df_entregue = df_mes[df_mes['Qtd_Ent_Num'] > 0]
                    
                    # POs Entregues: PCs únicos que tiveram pelo menos uma linha com entrega > 0
                    pos_entregues = df_entregue['Numero PC'].nunique() if not df_entregue.empty else 0
                    
                    # Itens Entregues: Contar LINHAS entregues (não soma da coluna Quantidade!)
                    itens_entregues = len(df_entregue)
                    
                    # === Cálculo de Abertos (Subtração direta) ===
                    pos_abertos = max(0, pos_emitidos - pos_entregues)
                    itens_abertos = max(0, itens_emitidos - itens_entregues)
                    
                    resultados.append({
                        'Mês': mes,
                        'POs_Emitidos': int(pos_emitidos),
                        'POs_Entregues': int(pos_entregues),
                        'POs_Abertos': int(pos_abertos),
                        'Itens_Emitidos': int(itens_emitidos),
                        'Itens_Entregues': int(itens_entregues),
                        'Itens_Abertos': int(itens_abertos)
                    })
                    
                return pd.DataFrame(resultados)

            # Calcula tudo de uma vez
            df_status_completo = calcular_status_pos_correto(df_compras_processed)
            
            # --- GRÁFICO 1: VOLUME DE POs ---
            with chart_cols_row3[0]:
                subheader_with_info("Qty de Pedidos Emitidos", "POs: contagem de POs únicos por mês; Entregues: POs com pelo menos uma linha com entrega > 0 (Relatório de Compras).")
                
                if not df_status_completo.empty:
                    # Prepara dados para Plotly (formato longo)
                    df_plot_pos = df_status_completo.melt(
                        id_vars=['Mês'], 
                        value_vars=['POs_Emitidos', 'POs_Entregues', 'POs_Abertos'], 
                        var_name='Status', 
                        value_name='Quantidade'
                    )
                    
                    # Renomeia para ficar bonito na legenda
                    map_nomes = {'POs_Emitidos': 'Emitidos', 'POs_Entregues': 'Entregues', 'POs_Abertos': 'Aguardando entrega'}
                    df_plot_pos['Status'] = df_plot_pos['Status'].map(map_nomes)

                    ordem_status = ['Emitidos', 'Entregues', 'Aguardando entrega']
                    df_plot_pos['Status'] = pd.Categorical(df_plot_pos['Status'], categories=ordem_status, ordered=True)

                    cores_pos = {'Emitidos': '#F59A23', 'Entregues': '#16A34A', 'Aguardando entrega': '#DC2626'}
                    
                    fig_pos = px.bar(
                        df_plot_pos, 
                        x='Mês', 
                        y='Quantidade', 
                        color='Status',
                        barmode='group',
                        labels={'Quantidade': 'Qtd. POs', 'Mês': 'Mês'},
                        color_discrete_map=cores_pos,
                        text='Quantidade'
                    )
                    
                    fig_pos.update_traces(textposition='outside', textfont=dict(size=10, family="Arial Black"))
                    fig_pos.update_layout(
                        height=350,
                        margin=dict(l=20, r=20, t=20, b=20),
                        plot_bgcolor='white',
                        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                        xaxis=dict(
                            categoryorder='array',
                            categoryarray=ordem_meses,
                            tickfont=dict(family="Arial", size=12, weight="bold")
                        )
                    )
                    st.plotly_chart(fig_pos, width='stretch')
                else:
                    st.info("Sem dados de POs para gerar gráfico.")

            # --- GRÁFICO 2: VOLUME DE ITENS ---
            with chart_cols_row3[1]:
                subheader_with_info("Qty de Itens por Pedido Emitidos", "Itens: contagem de linhas (itens) por mês, categorizadas por status (Relatório de Compras).")
                
                if not df_status_completo.empty:
                    df_plot_itens = df_status_completo.melt(
                        id_vars=['Mês'], 
                        value_vars=['Itens_Emitidos', 'Itens_Entregues', 'Itens_Abertos'], 
                        var_name='Status', 
                        value_name='Quantidade'
                    )
                    
                    map_nomes_itens = {'Itens_Emitidos': 'Emitidos', 'Itens_Entregues': 'Entregues', 'Itens_Abertos': 'Aguardando entrega'}
                    df_plot_itens['Status'] = df_plot_itens['Status'].map(map_nomes_itens)

                    ordem_status = ['Emitidos', 'Entregues', 'Aguardando entrega']
                    df_plot_itens['Status'] = pd.Categorical(df_plot_itens['Status'], categories=ordem_status, ordered=True)

                    cores_itens = {'Emitidos': '#F59A23', 'Entregues': '#16A34A', 'Aguardando entrega': '#DC2626'}
                    
                    fig_itens = px.bar(
                        df_plot_itens, 
                        x='Mês', 
                        y='Quantidade', 
                        color='Status',
                        barmode='group',
                        labels={'Quantidade': 'Qtd. Itens', 'Mês': 'Mês'},
                        color_discrete_map=cores_itens,
                        text='Quantidade'
                    )
                    
                    fig_itens.update_traces(textposition='outside', textfont=dict(size=10, family="Arial Black"))
                    fig_itens.update_layout(
                        height=350,
                        margin=dict(l=20, r=20, t=20, b=20),
                        plot_bgcolor='white',
                        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                        xaxis=dict(
                            categoryorder='array',
                            categoryarray=ordem_meses,
                            tickfont=dict(family="Arial", size=12, weight="bold")
                        )
                    )
                    st.plotly_chart(fig_itens, width='stretch')
                else:
                    st.info("Sem dados de Itens para gerar gráfico.")
                    
        else:
            st.warning("Dados do Relatório de Compras não carregados corretamente para análise de Status de POs.")

        st.divider()
        # --- AGING: GRÁFICO + LEGENDA LATERAL SIMPLES ---
        subheader_with_info("7. Aging dos Itens em Cotação", "Aging dos itens em cotação calculado como diferença entre hoje e data de aprovação; mostra distribuição por faixa de dias.")

        if not df_cotacao.empty and 'aging_dias' in df_cotacao.columns:
            def categorizar_aging(dias):
                if pd.isna(dias) or dias <= 0: return None
                elif 1 <= dias <= 3: return '1-3 dias'
                elif 4 <= dias <= 7: return '4-7 dias'
                elif 8 <= dias <= 11: return '8-11 dias'
                elif 12 <= dias <= 15: return '12-15 dias'
                elif 16 <= dias <= 30: return '16-30 dias'
                elif 31 <= dias <= 60: return '31-60 dias'
                elif 61 <= dias <= 70: return '61-70 dias'
                elif 71 <= dias <= 80: return '71-80 dias'
                elif 81 <= dias <= 90: return '81-90 dias'
                elif 91 <= dias <= 100: return '91-100 dias'
                else: return '> 100 dias'

            df_cotacao['aging_categoria'] = df_cotacao['aging_dias'].apply(categorizar_aging)
            aging_counts = df_cotacao['aging_categoria'].value_counts().reset_index()
            aging_counts.columns = ['Categoria', 'Qtd']
            aging_counts = aging_counts.dropna(subset=['Categoria'])

            categorias_ordem = ['1-3 dias', '4-7 dias', '8-11 dias', '12-15 dias',
                                 '16-30 dias', '31-60 dias', '61-70 dias', '71-80 dias',
                                 '81-90 dias', '91-100 dias', '> 100 dias']
            aging_counts['Categoria'] = pd.Categorical(aging_counts['Categoria'], categories=categorias_ordem, ordered=True)
            aging_counts = aging_counts.sort_values('Categoria')

            # Cálculo da Porcentagem
            total_aging = aging_counts['Qtd'].sum()
            aging_counts['Pct'] = (aging_counts['Qtd'] / total_aging * 100).round(1)

            # Cores: verde → amarelo → laranja → vermelho → roxo
            cores_aging = {
                '1-3 dias': '#22C55E', '4-7 dias': '#A3E635', '8-11 dias': '#FDE047',
                '12-15 dias': '#FB923C', '16-30 dias': '#FFA500', '31-60 dias': '#FF8C00',
                '61-70 dias': '#FF4500', '71-80 dias': '#DC143C', '81-90 dias': '#B22222',
                '91-100 dias': '#8B0000', '> 100 dias': '#4B0082'
            }
            
            # Layout: 2 Colunas (Gráfico | Legenda)
            aging_cols = st.columns([2, 1], gap="large")
            
            with aging_cols[0]:
                # CARD DO GRÁFICO
                st.markdown('<div style="background-color: white; padding: 1.5rem; border-radius: 16px; border: 1px solid #E5E7EB; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); height: 100%;">', unsafe_allow_html=True)
                
                fig_aging = px.pie(
                    aging_counts, 
                    names='Categoria', 
                    values='Qtd', 
                    hole=0.6, 
                    color='Categoria', 
                    color_discrete_map=cores_aging
                )
                
                # Configuração do Gráfico: Apenas valores dentro das fatias
                fig_aging.update_traces(
                    textposition='inside', 
                    textinfo='value',     # Mostra apenas o número (ex: 281)
                    textfont_size=14,     
                    textfont_color='white', 
                    marker=dict(line=dict(color='#FFFFFF', width=2))
                )
                
                total_atrasados = int(aging_counts['Qtd'].sum())
                
                # Total no Centro
                fig_aging.add_annotation(
                    text=f"<b>{total_atrasados}</b><br><span style='font-size:14px'>em Atraso</span>",
                    x=0.5, y=0.5,
                    font_size=28,
                    showarrow=False,
                    font_family="Arial Black",
                    font_color="#0A1B2B"
                )
                
                fig_aging.update_layout(
                    height=400, 
                    margin=dict(l=20, r=20, t=20, b=20), 
                    showlegend=False, # Sem legenda nativa
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)'
                )
                
                st.plotly_chart(fig_aging, width='stretch')
                st.markdown('</div>', unsafe_allow_html=True) # Fecha Card Gráfico

            with aging_cols[1]:
                # CARD DA LEGENDA
                st.markdown('<div style="background-color: white; padding: 1.5rem; border-radius: 16px; border: 1px solid #E5E7EB; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); height: 100%; display: flex; flex-direction: column; justify-content: center;">', unsafe_allow_html=True)
                
                st.markdown("### 📋 Detalhamento")
                
                legend_html = ""
                for _, row in aging_counts.iterrows():
                    cat = row['Categoria']
                    pct = row['Pct']
                    color = cores_aging.get(cat, '#000000')
                    
                    # Formato: Bolinha + Nome da Faixa + Quantidade + Porcentagem
                    legend_item = f"""
                    <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 1px solid #F3F4F6;">
                        <div style="display: flex; align-items: center; gap: 10px;">
                            <div style="width: 12px; height: 12px; background-color: {color}; border-radius: 50%;"></div>
                            <span style="font-size: 14px; color: #374151; font-weight: 500;">{cat}</span>
                        </div>
                        <div style="display: flex; gap: 8px; align-items: center;">
                            <span style="font-size: 14px; color: {color}; font-weight: 700;">{int(row['Qtd'])}</span>
                            <span style="font-size: 12px; color: #6B7280;">({pct}%)</span>
                        </div>
                    </div>
                    """
                    legend_html += legend_item
                
                st.markdown(legend_html, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True) # Fecha Card Legenda

        else:
            st.info("Sem dados de aging disponíveis.")

        subheader_with_info(" Ranking do Aging por Departamento (> 15 dias)", "Top 10 departamentos com itens atrasados (>15 dias) — base: Solicitações em cotação.")
        
        if not df_cotacao.empty and col_depto and 'aging_dias' in df_cotacao.columns:
            # 1. Filtra apenas itens ATASADOS (mais de 15 dias)
            df_atrasados = df_cotacao[
                (df_cotacao['aging_dias'] > 15) & 
                (pd.notna(df_cotacao['aging_dias']))
            ].copy()
            
            if not df_atrasados.empty:
                # 2. Agrupa por Departamento e conta os itens
                df_depto_atraso = df_atrasados.groupby(col_depto).size().reset_index(name='Qtd_Atrasados')
                
                # 3. Ordena do maior para o menor e pega Top 10
                df_depto_atraso = df_depto_atraso.sort_values('Qtd_Atrasados', ascending=False).head(10)
                
                # 4. Cria o gráfico de Barras Verticais
                fig_depto_atraso = px.bar(
                    df_depto_atraso, 
                    x=col_depto, 
                    y='Qtd_Atrasados', 
                    text='Qtd_Atrasados',
                    labels={col_depto: 'Departamento', 'Qtd_Atrasados': 'Itens Atrasados'},
                    color_discrete_sequence=['#DC2626'] # Vermelho para indicar alerta/atraso
                )
                
                # Configura o texto para ficar EM CIMA da barra, bem grande e visível
                fig_depto_atraso.update_traces(
                    textposition='outside', 
                    insidetextanchor='middle', 
                    textfont=dict(size=16, color='#1F2933', family="Arial Black")
                )
                
                fig_depto_atraso.update_layout(
                    height=400, 
                    margin=dict(l=20, r=20, t=50, b=80), # Margem inferior aumentada para rótulos girados
                    showlegend=False, 
                    plot_bgcolor='white',
                    xaxis=dict(
                        tickangle=45, # Gira os nomes dos departamentos em 45 graus
                        tickfont=dict(family="Arial", size=13, weight="bold") # Fonte em negrito, tamanho 12
                    )
                )
                
                st.plotly_chart(fig_depto_atraso, width='stretch')
            else:
                st.info("✅ Nenhum item em Cotação está atrasado (todos dentro do SLA de 15 dias).")
        else:
            st.info("Sem dados de departamento ou aging para gerar gráfico.")

        # --- SLA: APROVAÇÃO → EMISSÃO PC ---
        st.divider()
        st.subheader("8. SLA: Aprovação → Emissão PC")

        if col_aprovacao and df_compras_processed is not None and col_num_pc:
            try:
                if col_pedido and not df.empty:
                    _df_sla = df[df[col_pedido].astype(str).str.strip().replace({'nan': '', 'None': '', 'NaN': ''}) != ''].copy()
                    _df_sla['_data_aprov_sla'] = pd.to_datetime(_df_sla[col_aprovacao], errors='coerce')
                    _df_sla['_pc_str'] = _df_sla[col_pedido].astype(str).str.strip()

                    _df_pc_dates = df_compras_processed[[col_num_pc, 'Data_PC']].copy()
                    _df_pc_dates[col_num_pc] = _df_pc_dates[col_num_pc].astype(str).str.strip()
                    _df_pc_dates = _df_pc_dates.drop_duplicates(subset=[col_num_pc])

                    _df_sla = _df_sla.merge(_df_pc_dates, left_on='_pc_str', right_on=col_num_pc, how='inner')
                    _df_sla['sla_dias'] = (_df_sla['Data_PC'] - _df_sla['_data_aprov_sla']).dt.days
                    _df_sla = _df_sla.dropna(subset=['sla_dias'])
                    _df_sla = _df_sla[_df_sla['sla_dias'] >= 0]

                    if not _df_sla.empty:
                        sla_kpi_cols = st.columns(3)
                        with sla_kpi_cols[0]:
                            st.metric("Média SLA (dias)", f"{_df_sla['sla_dias'].mean():.1f}")
                        with sla_kpi_cols[1]:
                            st.metric("Mediana SLA (dias)", f"{_df_sla['sla_dias'].median():.1f}")
                        with sla_kpi_cols[2]:
                            st.metric("Máximo SLA (dias)", f"{int(_df_sla['sla_dias'].max())}")

                        fig_sla = px.histogram(_df_sla, x='sla_dias', nbins=20,
                            labels={'sla_dias': 'Dias (Aprovação → Emissão PC)'},
                            color_discrete_sequence=['#F59A23'])
                        fig_sla.update_layout(height=300, margin=dict(l=20, r=20, t=20, b=20),
                            plot_bgcolor='white', showlegend=False)
                        st.plotly_chart(fig_sla, width='stretch')
                    else:
                        st.info("Sem dados suficientes para calcular SLA no período selecionado.")
                else:
                    st.info("Coluna de pedido não encontrada nas Solicitações.")
            except Exception as e_sla:
                st.info(f"SLA não disponível: {e_sla}")
        else:
            st.info("Dados insuficientes para SLA (necessário: coluna Aprovação nas Solicitações e Relatório de Compras).")

    except Exception as e:
        st.error(f"Erro ao carregar o arquivo '{ARQUIVO_SOLICITACOES}': {e}")
else:
    st.error(f"Arquivo '{ARQUIVO_SOLICITACOES}' não encontrado! Verifique se o arquivo está na pasta do app.")