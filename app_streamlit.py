# app_streamlit.py
import streamlit as st

# --- 1. st.set_page_config() DEVE SER O PRIMEIRO COMANDO STREAMLIT ---
st.set_page_config(
    page_title="Gerador de Mensagens de Cobrança",
    page_icon="🤖", 
    layout="wide"
)

# --- Imports Padrão ---
import pandas as pd
import os
import logging
import sys
import time 

# --- 2. AJUSTE DE CAMINHO E IMPORT DE MÓDULOS DO PROJETO ---
_project_root = os.path.dirname(os.path.abspath(__file__))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

# --- 3. Configuração do Handler de Log do Streamlit (CLASSE) ---
class StreamlitLogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        if 'streamlit_log_records_list' not in st.session_state:
            st.session_state.streamlit_log_records_list = []

    def emit(self, record):
        # O print de depuração para o console foi removido daqui.
        try:
            msg = self.format(record)
            # Garante que a lista exista no session_state antes de adicionar
            if 'streamlit_log_records_list' not in st.session_state: 
                st.session_state.streamlit_log_records_list = []
            st.session_state.streamlit_log_records_list.append(msg)
        except Exception as e_format:
            # Mantém um print para o console caso a formatação para o Streamlit falhe
            print(f"STREAMLIT_HANDLER_ERROR: Erro ao formatar/adicionar log para UI: {e_format}", file=sys.stderr)


# --- 4. Instanciar e Adicionar o Handler Globalmente ---
logger_raiz = logging.getLogger() 
if logger_raiz.level == logging.NOTSET or logger_raiz.level > logging.INFO:
    logger_raiz.setLevel(logging.INFO)

if 'streamlit_handler_globally_added_flag' not in st.session_state:
    s_handler_global = StreamlitLogHandler()
    s_handler_global.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s: %(message)s', datefmt='%H:%M:%S'))
    s_handler_global.setLevel(logging.DEBUG) 
    logger_raiz.addHandler(s_handler_global)
    st.session_state.streamlit_handler_globally_added_flag = True

# --- 5. Importa os módulos do projeto ---
try:
    from modules import arquivo_txt_sender 
    import gerador_mensagens_cobranca     
except ImportError as e:
    st.error(f" Falha crítica na importação de módulos necessários: {e}. O app não pode continuar.")
    logging.critical(f"Falha crítica na importação em app_streamlit.py: {e}", exc_info=True)
    st.stop()

# --- 6. Função principal da lógica da interface Streamlit ---
def main_streamlit_app_logic():
    # st.image("logo.png", width=100) # Se você tiver um logo, descomente e ajuste o caminho
    st.title("🤖 Gerador de Mensagens de Cobrança Automatizadas")
    st.markdown("Carregue sua planilha de inadimplentes e gere as mensagens de cobrança e arquivos de controle.")
    st.markdown("---")

    st.sidebar.header("⚙️ Configurações e Ações")
    uploaded_file = st.sidebar.file_uploader(
        "1. Escolha a planilha Excel (.xlsx, .xls)",
        type=["xlsx", "xls"],
        key="file_uploader_main_v5" 
    )
    
    st.sidebar.markdown("---") 
    st.sidebar.markdown("**Nomes das Colunas:**")
    st.sidebar.caption(f"Nome Cliente: `{os.getenv('COLUNA_NOME', 'NÃO DEFINIDO')}`")
    st.sidebar.caption(f"Telefone: `{os.getenv('COLUNA_TELEFONE', 'NÃO DEFINIDO')}`")
    st.sidebar.caption(f"Valor/Saldo: `{os.getenv('COLUNA_VALOR', 'NÃO DEFINIDO')}`")
    st.sidebar.caption(f"Vencimento: `{os.getenv('COLUNA_VENCIMENTO', 'NÃO DEFINIDO')}`")
    st.sidebar.caption(f"Lote (Referência): `{os.getenv('COLUNA_LOTE', 'NÃO DEFINIDO')}`")
    st.sidebar.caption(f"Cód. Loteamento (Duplicata): `{os.getenv('COLUNA_LOTEAMENTO', 'NÃO DEFINIDO')}`") 
    st.sidebar.markdown("---")
    
    if st.sidebar.button("🚀 Gerar Mensagens", key="generate_messages_button_v7", type="primary", use_container_width=True):
        if uploaded_file is not None:
            if 'streamlit_log_records_list' in st.session_state:
                st.session_state.streamlit_log_records_list = [] 
            
            # Define os placeholders para as abas aqui, para que possam ser preenchidos depois
            # A criação das abas em si ocorrerá após ter 'resultados'
            results_placeholder = st.empty() # Placeholder para toda a área de resultados (incluindo abas)

            st.info(f" Iniciando processamento do arquivo: **{uploaded_file.name}**...")
            logging.info("APP_STREAMLIT_LOG_BTN: Botão 'Gerar Mensagens' clicado.")
            
            resultados = None 
            try:
                with st.spinner("⏳ Aguarde, processando sua planilha..."):
                    output_dir_config_msg = arquivo_txt_sender.DIRETORIO_SAIDA 
                    DIR_DESCARTADOS = "contatos_descartados" 
                    for dir_path in [output_dir_config_msg, DIR_DESCARTADOS]:
                        if os.path.exists(dir_path):
                            for f_name in os.listdir(dir_path):
                                file_full_path = os.path.join(dir_path, f_name)
                                if (dir_path == output_dir_config_msg and f_name.endswith(".txt")) or \
                                   (dir_path == DIR_DESCARTADOS and f_name.endswith(".xlsx")):
                                    try: os.remove(file_full_path)
                                    except Exception as e_rm: logging.warning(f"AppStreamlit: Não removeu {file_full_path}: {e_rm}")
                        else:
                            os.makedirs(dir_path, exist_ok=True)

                    resultados = gerador_mensagens_cobranca.processar_cobrancas(
                        arquivo_planilha_input=uploaded_file
                    )
                
                logging.info("APP_STREAMLIT_LOG_BTN: Processamento principal concluído.")

                # Agora preenche o placeholder com as abas e os resultados
                with results_placeholder.container():
                    if resultados:
                        st.success(f"🎉 **{resultados.get('message', 'Processamento concluído!')}**")
                        
                        tab_resumo, tab_arquivos_gerados, tab_logs = st.tabs(["📊 Resumo", "📂 Arquivos de Saída", "📝 Logs Detalhados"])

                        with tab_resumo:
                            st.subheader("Visão Geral do Processamento")
                            resumo_df_data = {
                                "Métrica": ["Planilha Processada", "Total de Registros Carregados", 
                                            "Contatos Descartados (sem telefone)", 
                                            "Registros COM Telefone (para mensagens)", 
                                            "Telefones Únicos Processados", 
                                            "Mensagens Geradas com Sucesso (backend)", 
                                            "Grupos com Falha"],
                                "Valor": [resultados.get('planilha_processada', uploaded_file.name), 
                                          resultados.get('total_registros_carregados', 'N/A'), 
                                          resultados.get('num_descartados_arquivo', 'N/A'), 
                                          resultados.get('total_com_telefone', 'N/A'), 
                                          resultados.get('total_telefones_unicos', 'N/A'), 
                                          resultados.get('mensagens_sucesso', 'N/A'), 
                                          resultados.get('falhas_envio', 'N/A')]
                            }
                            resumo_df = pd.DataFrame(resumo_df_data)
                            resumo_df.set_index('Métrica', inplace=True)
                            st.table(resumo_df)

                        with tab_arquivos_gerados:
                            st.subheader("Arquivos Gerados")
                            if resultados.get('caminho_arquivo_descartados') and resultados.get('num_descartados_arquivo', 0) > 0:
                                caminho_desc = resultados['caminho_arquivo_descartados']
                                nome_arq_desc = os.path.basename(caminho_desc)
                                st.markdown("---")
                                st.write(f"📋 **Arquivo de Contatos Descartados:**")
                                st.info(f"{resultados['num_descartados_arquivo']} contatos foram descartados e salvos.")
                                try:
                                    with open(caminho_desc, "rb") as fp_desc:
                                        st.download_button(
                                            label=f"Baixar: {nome_arq_desc}",
                                            data=fp_desc,
                                            file_name=nome_arq_desc,
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key="download_descartados_button_final_v3" # Chave atualizada
                                        )
                                except Exception as e_dl_desc:
                                    st.error(f"Erro ao preparar download do arquivo de descartados: {e_dl_desc}")
                            elif resultados.get('status', '').startswith('sucesso'):
                                st.success("✅ Nenhum contato foi descartado por falta de telefone.")
                            
                            st.markdown("---")
                            st.write("📨 **Pasta das Mensagens de Texto (.txt):**")
                            output_dir_msg = resultados.get('output_dir', arquivo_txt_sender.DIRETORIO_SAIDA)
                            if resultados.get('mensagens_sucesso', 0) > 0 and os.path.exists(output_dir_msg):
                                generated_files_count = 0
                                try: generated_files_count = len([f for f in os.listdir(output_dir_msg) if f.endswith(".txt")])
                                except Exception: pass 

                                if generated_files_count > 0:
                                    st.success(f"{generated_files_count} arquivos de mensagem (.txt) foram salvos.")
                                    path_para_exibir = os.path.abspath(output_dir_msg)
                                    st.markdown(f"Acesse-os na pasta:")
                                    st.code(path_para_exibir, language=None)
                                    st.caption("*Copie o caminho acima e cole no seu explorador de arquivos.*")
                                elif resultados.get('mensagens_sucesso', 0) > 0 : 
                                    st.warning(f"O processamento indicou mensagens salvas, mas nenhum .txt foi encontrado em '{output_dir_msg}'.")
                            elif resultados.get('mensagens_sucesso', 0) > 0:
                                st.warning(f"Mensagens reportadas como salvas, mas diretório '{output_dir_msg}' não encontrado.")
                            else:
                                st.info("Nenhuma mensagem de texto (.txt) foi gerada nesta execução.")

                        with tab_logs: 
                            st.subheader("Registro Detalhado do Processamento")
                            log_content_final_tab = "\n".join(st.session_state.get('streamlit_log_records_list', []))
                            if not log_content_final_tab.strip():
                                log_content_final_tab = "Nenhuma mensagem de log foi capturada para exibição."
                            st.text_area("Log:", log_content_final_tab, height=350, key="log_display_in_tab_key_v3") 
                    else:
                        st.error("O processamento não retornou resultados válidos.")
                        logging.error("APP_STREAMLIT_LOG_BTN: Resultados do processamento principal foram None.")
                        # Exibir logs mesmo se 'resultados' for None
                        log_content_final_erro_sem_res = "\n".join(st.session_state.get('streamlit_log_records_list', []))
                        st.expander("Log de Erro (sem resultados)", expanded=True).text_area(
                            "Log:", log_content_final_erro_sem_res, height=300, key="log_display_no_results_key"
                        )


            except Exception as e:
                st.error(f"Ocorreu um erro geral durante o processamento: {e}")
                logging.exception("APP_STREAMLIT_LOG_BTN: Erro geral capturado:") 
                # Exibe logs também em caso de exceção não capturada pelo finally do try principal
                log_content_excecao = "\n".join(st.session_state.get('streamlit_log_records_list', []))
                st.expander("Log de Exceção Detalhado", expanded=True).text_area(
                    "Log (Exceção):", log_content_excecao, height=300, key="log_display_exception_key"
                )
        else:
            st.sidebar.warning("Por favor, faça o upload de uma planilha Excel para processar.")
    else:
        st.info("ℹ️ Carregue sua planilha na barra lateral e clique em 'Gerar Mensagens' para iniciar.")

if __name__ == "__main__":
    main_streamlit_app_logic()