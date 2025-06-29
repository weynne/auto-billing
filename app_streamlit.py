# app_streamlit.py
import streamlit as st

st.set_page_config(
    page_title="Gerador de Mensagens de Cobrança",
    page_icon="🤖",
    layout="wide"
)

import pandas as pd
import os
import logging
import sys
import shutil # Usado para remover diretórios

_project_root = os.path.dirname(os.path.abspath(__file__))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

class StreamlitLogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        if 'log_records' not in st.session_state:
            st.session_state.log_records = []

    def emit(self, record):
        try:
            msg = self.format(record)
            st.session_state.log_records.append(msg)
        except Exception as e:
            print(f"ERRO_HANDLER_STREAMLIT: Falha ao formatar log: {e}", file=sys.stderr)

def setup_streamlit_logging():
    if 'logging_setup_done' not in st.session_state:
        root_logger = logging.getLogger()
        if root_logger.level == logging.NOTSET or root_logger.level > logging.INFO:
            root_logger.setLevel(logging.INFO)

        st_handler = StreamlitLogHandler()
        st_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s: %(message)s', datefmt='%H:%M:%S'))
        root_logger.addHandler(st_handler)

        st.session_state.logging_setup_done = True
        logging.info("Handler de log do Streamlit configurado com sucesso.")

try:
    from modules import arquivo_txt_sender
    import gerador_mensagens_cobranca
except ImportError as e:
    st.error(f"Falha crítica na importação de módulos: {e}. O app não pode continuar.")
    logging.critical(f"Falha na importação em app_streamlit.py: {e}", exc_info=True)
    st.stop()

def limpar_diretorios_saida():
    logging.info("Limpando diretórios de saída de execuções anteriores...")
    output_dirs = {
        arquivo_txt_sender.DIRETORIO_SAIDA: ".txt",
        "contatos_descartados": ".xlsx"
    }

    for dir_path, extension in output_dirs.items():
        if os.path.exists(dir_path):
            try:
                shutil.rmtree(dir_path)
                logging.info(f"Diretório '{dir_path}' removido.")
            except Exception as e:
                logging.error(f"Erro ao remover o diretório '{dir_path}': {e}")
        
        try:
            os.makedirs(dir_path, exist_ok=True)
            logging.info(f"Diretório '{dir_path}' garantido.")
        except Exception as e:
            logging.error(f"Erro ao criar o diretório '{dir_path}': {e}")


def exibir_resultados(resultados):
    if not resultados:
        st.error("O processamento não retornou resultados válidos.")
        logging.error("O dicionário de resultados do processamento principal estava vazio.")
        return

    st.success(f"🎉 **{resultados.get('message', 'Processamento concluído!')}**")

    tab_resumo, tab_arquivos, tab_logs = st.tabs(["📊 Resumo", "📂 Arquivos Gerados", "📝 Logs"])

    with tab_resumo:
        st.subheader("Visão Geral do Processamento")
        
        resumo_data = {
            "Métrica": [
                "Planilha Processada", "Total de Linhas na Planilha",
                "Contatos Descartados (sem tel.)", "Contatos Válidos para Mensagem",
                "Telefones Únicos", "Mensagens Geradas", "Falhas"
            ],
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
        
        resumo_df = pd.DataFrame(resumo_data).set_index("Métrica")
        st.table(resumo_df)

    with tab_arquivos:
        st.subheader("Downloads")
        caminho_descartados = resultados.get('caminho_arquivo_descartados')
        if caminho_descartados and os.path.exists(caminho_descartados):
            nome_arq_desc = os.path.basename(caminho_descartados)
            with open(caminho_descartados, "rb") as fp:
                st.download_button(
                    label=f"⬇️ Baixar Planilha de Descartados ({resultados.get('num_descartados_arquivo', 0)} linhas)",
                    data=fp,
                    file_name=nome_arq_desc,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("✅ Nenhum contato foi descartado por falta de telefone.")

        st.markdown("---")
        output_dir_msg = resultados.get('output_dir')
        if resultados.get('mensagens_sucesso', 0) > 0 and output_dir_msg and os.path.exists(output_dir_msg):
            path_abs = os.path.abspath(output_dir_msg)
            st.write("📨 **Pasta com as Mensagens de Texto (.txt):**")
            st.info(f"As {resultados['mensagens_sucesso']} mensagens foram salvas no seu computador.")
            st.markdown(f"Acesse-as no seguinte caminho:")
            st.code(path_abs, language=None)
        else:
            st.info("Nenhuma mensagem de texto (.txt) foi gerada.")

    with tab_logs:
        st.subheader("Registro Detalhado da Execução")
        log_content = "\n".join(st.session_state.get('log_records', ["Nenhum log capturado."]))
        st.text_area("Logs:", log_content, height=400, key="log_display_area")

def main():
    setup_streamlit_logging()

    st.title("🤖 Gerador de Mensagens de Cobrança")
    st.markdown("Carregue sua planilha de inadimplentes para gerar as mensagens e arquivos de controle.")
    st.markdown("---")

    with st.sidebar:
        st.header("⚙️ 1. Configuração")
        uploaded_file = st.file_uploader(
            "Escolha a planilha Excel (.xlsx, .xls)",
            type=["xlsx", "xls"],
            key="file_uploader"
        )

        with st.expander("Verificar Nomes das Colunas Esperadas"):
            st.caption("Certifique-se de que sua planilha contenha colunas com estes nomes:")
            st.markdown(f"- **Cliente:** `{os.getenv('COLUNA_NOME', 'NÃO DEFINIDO')}`")
            st.markdown(f"- **Telefone:** `{os.getenv('COLUNA_TELEFONE', 'NÃO DEFINIDO')}`")
            st.markdown(f"- **Valor:** `{os.getenv('COLUNA_VALOR', 'NÃO DEFINIDO')}`")
            st.markdown(f"- **Vencimento:** `{os.getenv('COLUNA_VENCIMENTO', 'NÃO DEFINIDO')}`")
            st.markdown(f"- **Lote/Ref:** `{os.getenv('COLUNA_LOTE', 'NÃO DEFINIDO')}`")
            st.markdown(f"- **Cód. Loteamento:** `{os.getenv('COLUNA_LOTEAMENTO', 'NÃO DEFINIDO')}`")

        st.markdown("---")
        st.header("🚀 2. Ação")
        if st.button("Gerar Mensagens", type="primary", use_container_width=True, key="btn_gerar"):
            if uploaded_file:
                st.session_state.log_records = []
                limpar_diretorios_saida()
                
                results_placeholder = st.empty()
                
                with results_placeholder.container():
                    with st.spinner("⏳ Processando a planilha... Esta operação pode levar alguns instantes."):
                        try:
                            resultados = gerador_mensagens_cobranca.processar_cobrancas(
                                arquivo_planilha_input=uploaded_file
                            )
                            exibir_resultados(resultados)

                        except Exception as e:
                            logging.exception("Erro inesperado durante o processamento principal.")
                            st.error(f"Ocorreu um erro crítico: {e}")
                            log_content = "\n".join(st.session_state.get('log_records', []))
                            st.expander("Ver Logs do Erro").text_area("Logs:", log_content, height=300)
            else:
                st.warning("Por favor, carregue uma planilha para iniciar.", icon="⚠️")
        
    if 'log_records' not in st.session_state or not st.session_state.log_records:
         st.info("ℹ️ Para começar, carregue uma planilha na barra lateral e clique em 'Gerar Mensagens'.")

if __name__ == "__main__":
    main()