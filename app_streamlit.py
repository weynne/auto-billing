# app_streamlit.py
import streamlit as st
import pandas as pd
import os
import logging
import sys
import shutil
from datetime import datetime
import plotly.express as px

# --- Configuração da Página ---
st.set_page_config(
    page_title="Dashboard de Cobrança",
    page_icon="📊",
    layout="wide"
)

# --- Ajuste de Path e Imports do Projeto ---
_project_root = os.path.dirname(os.path.abspath(__file__))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from modules import leitor_planilha, arquivo_txt_sender
import gerador_mensagens_cobranca

# --- Funções Auxiliares (Logging, Limpeza, etc.) ---
class StreamlitLogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        if 'log_records' not in st.session_state: st.session_state.log_records = []
    def emit(self, record):
        try:
            st.session_state.log_records.append(self.format(record))
        except Exception as e:
            print(f"ERRO_HANDLER_STREAMLIT: {e}", file=sys.stderr)

def setup_streamlit_logging():
    if 'logging_setup_done' not in st.session_state:
        root_logger = logging.getLogger()
        if root_logger.level == logging.NOTSET or root_logger.level > logging.INFO: root_logger.setLevel(logging.INFO)
        st_handler = StreamlitLogHandler()
        st_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s: %(message)s', datefmt='%H:%M:%S'))
        root_logger.addHandler(st_handler)
        st.session_state.logging_setup_done = True
        logging.info("Handler de log do Streamlit configurado.")

def limpar_diretorios_saida():
    logging.info("Limpando diretórios de saída...")
    output_dirs = {"mensagens_txt_geradas", "contatos_descartados"}
    for dir_path in output_dirs:
        if os.path.exists(dir_path):
            try:
                shutil.rmtree(dir_path)
            except Exception as e:
                logging.error(f"Erro ao remover o diretório '{dir_path}': {e}")
        try:
            os.makedirs(dir_path, exist_ok=True)
        except Exception as e:
            logging.error(f"Erro ao criar o diretório '{dir_path}': {e}")

def exibir_resultados_processamento(resultados):
    """
    Exibe os resultados do processamento principal DENTRO DA BARRA LATERAL,
    usando um expander para manter a interface organizada.
    """
    if not resultados:
        st.sidebar.error("O processamento não retornou resultados válidos.")
        logging.error("O dicionário de resultados do processamento principal estava vazio.")
        return

    with st.sidebar.expander("✔️ Ver Resultados do Processamento", expanded=True):
        st.success(f"🎉 **{resultados.get('message', 'Processamento concluído!')}**")
        tab_resumo, tab_arquivos, tab_logs = st.tabs(["📊 Resumo", "📂 Arquivos", "📝 Logs"])
        with tab_resumo:
            st.subheader("Visão Geral")
            resumo_data = {
                "Métrica": ["Planilha", "Linhas", "Descartados", "Válidos", "Telefones", "Sucessos", "Falhas"],
                "Valor": [
                    str(resultados.get('planilha_processada', 'N/A')),
                    str(resultados.get('total_registros_carregados', 'N/A')),
                    str(resultados.get('num_descartados_arquivo', 'N/A')),
                    str(resultados.get('total_com_telefone', 'N/A')),
                    str(resultados.get('total_telefones_unicos', 'N/A')),
                    str(resultados.get('mensagens_sucesso', 'N/A')),
                    str(resultados.get('falhas_envio', 'N/A'))
                ]
            }
            st.table(pd.DataFrame(resumo_data).set_index("Métrica"))
        with tab_arquivos:
            st.subheader("Downloads")
            caminho_descartados = resultados.get('caminho_arquivo_descartados')
            num_descartados = resultados.get('num_descartados_arquivo', 0)
            if caminho_descartados and os.path.exists(caminho_descartados) and num_descartados > 0:
                with open(caminho_descartados, "rb") as fp:
                    st.download_button(
                        label=f"⬇️ Baixar Descartados ({num_descartados})",
                        data=fp,
                        file_name=os.path.basename(caminho_descartados),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                st.info("✅ Nenhum contato foi descartado.")
        with tab_logs:
            st.subheader("Logs da Execução")
            log_content = "\n".join(st.session_state.get('log_records', ["N/A"]))
            st.text_area("Logs:", log_content, height=250, key="log_display_sidebar")

def formatar_milhar(valor):
    if valor >= 1_000_000: return f"R$ {valor/1_000_000:.2f} Mi"
    if valor >= 1_000: return f"R$ {valor/1_000:.1f} Mil"
    return f"R$ {valor:.2f}"

def main():
    setup_streamlit_logging()
    
    st.title("📊 Dashboard e Gerador de Cobranças")
    
    config_app = {k: os.getenv(v) for k, v in {
        'col_vencimento': 'COLUNA_VENCIMENTO', 'col_loteamento': 'COLUNA_LOTEAMENTO',
        'col_nome': 'COLUNA_NOME', 'col_telefone': 'COLUNA_TELEFONE', 'col_valor': 'COLUNA_VALOR'
    }.items()}

    with st.sidebar:
        st.header("⚙️ Configuração")
        uploaded_file = st.file_uploader("1. Escolha a planilha de cobrança", type=["xlsx", "xls"])
        
        st.markdown("---")
        st.header("🚀 Ação Principal")
        gerar_mensagens_btn = st.button("2. Gerar Mensagens de Cobrança", type="primary", use_container_width=True)

    if uploaded_file is not None:
        cache_key = f"{uploaded_file.name}-{uploaded_file.size}"
        if st.session_state.get('df_cache_key') != cache_key:
            st.session_state.df_cobranca = leitor_planilha.carregar_inadimplentes(uploaded_file, config_app)
            st.session_state.df_cache_key = cache_key
    
    if 'df_cobranca' not in st.session_state or st.session_state.df_cobranca is None:
        st.info("⬅️ Por favor, carregue uma planilha na barra lateral para começar a análise.")
        return
    
    df = st.session_state.df_cobranca
    if df.empty:
        st.error("A planilha carregada está vazia ou não foi lida corretamente.")
        return

    st.header("Análise de Inadimplência")
    
    lista_loteamentos = ["Todos"] + df['Nome Loteamento'].unique().tolist()
    
    col1, col2 = st.columns(2)
    with col1:
        loteamento_selecionado = st.selectbox("Filtrar por Empreendimento", options=lista_loteamentos)
    with col2:
        data_base = st.date_input("Data Base da Análise", value=datetime.now())

    dados_dashboard = leitor_planilha.gerar_dados_dashboard_aging(df, config_app, data_base, loteamento_selecionado)
    
    st.markdown("---")
    
    if not dados_dashboard or dados_dashboard['kpis']['total_vencido'] == 0:
        st.warning("Nenhum dado vencido encontrado para os filtros selecionados.")
    else:
        kpis = dados_dashboard['kpis']
        col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
        col_kpi1.metric("Valor Total Vencido", value=formatar_milhar(kpis['total_vencido']))
        col_kpi2.metric("PDD 120 dias (Exemplo)", "R$ 0") 
        col_kpi3.metric("120 dias 'Arrasto' (Exemplo)", "R$ 0")

        st.markdown("<hr>", unsafe_allow_html=True)
        
        col_chart1, col_chart2 = st.columns([0.55, 0.45])
        with col_chart1:
            st.subheader("Aging do Valor Vencido")
            dados_donut = dados_dashboard['dados_donut']
            
            fig = px.pie(
                dados_donut, 
                values='Valor', 
                names='Faixa de Atraso', 
                hole=0.4, 
                color_discrete_sequence=px.colors.sequential.Reds_r
            )
            
            # Atualiza a aparência dos textos e das fatias
            fig.update_traces(
                textposition='outside', 
                textinfo='percent+label', 
                rotation=90 # Ajustei a rotação para melhorar o layout
            )
            
            # MUDANÇA: Aumenta a margem superior (t) e inferior (b) para dar espaço aos rótulos
            fig.update_layout(
                showlegend=False, 
                height=500, # Aumentei um pouco a altura geral
                margin=dict(t=100, b=100, l=0, r=0) # Margens generosas no topo (t) e embaixo (b)
            )
            
            st.plotly_chart(fig, use_container_width=True)
        with col_chart2:
            st.subheader("Tabela de Apoio")
            tabela_formatada = dados_donut.copy()
            tabela_formatada['Valor'] = tabela_formatada['Valor'].apply(lambda x: f"R$ {x:,.2f}")
            tabela_formatada['Percentual'] = tabela_formatada['Percentual'].apply(lambda x: f"{x:.2f}%")
            st.dataframe(tabela_formatada, use_container_width=True, height=450, hide_index=True)

    if gerar_mensagens_btn:
        st.markdown("---")
        st.header("Processamento das Mensagens de Cobrança")
        if uploaded_file is not None:
            limpar_diretorios_saida()
            st.session_state.log_records = []
            with st.spinner("⏳ Processando..."):
                try:
                    resultados = gerador_mensagens_cobranca.processar_cobrancas(arquivo_planilha_input=uploaded_file)
                    exibir_resultados_processamento(resultados)
                except Exception as e:
                    logging.exception("Erro inesperado.")
                    st.error(f"Ocorreu um erro crítico: {e}")
        else:
            st.sidebar.error("É necessário carregar um arquivo para gerar as mensagens.")

if __name__ == "__main__":
    main()