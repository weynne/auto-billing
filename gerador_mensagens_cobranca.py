# gerador_mensagens_cobranca.py
import time
import os
from dotenv import load_dotenv
import pandas as pd
import logging
import sys
import locale

# --- Configurações Iniciais ---
# MUDANÇA: Carregamento do .env e configuração do logging movidos para uma função.
def setup_environment_and_logging():
    """Carrega .env e configura handlers de log (console e arquivo) se não existirem."""
    if load_dotenv():
        print("GMC_SETUP: Arquivo .env carregado.")
    else:
        print("GMC_SETUP: AVISO - Arquivo .env não encontrado.")

    root_logger = logging.getLogger()
    if not root_logger.handlers:
        root_logger.setLevel(logging.INFO)
        # Handler para o console
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(module)s: %(message)s', datefmt='%H:%M:%S'))
        root_logger.addHandler(console_handler)
        # Handler para o arquivo
        try:
            log_path = "processamento_cobrancas.log"
            file_handler = logging.FileHandler(log_path, mode='a', encoding='utf-8')
            file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s'))
            root_logger.addHandler(file_handler)
            logging.info(f"Log de arquivo configurado em: '{os.path.abspath(log_path)}'")
        except IOError as e:
            logging.error(f"Erro crítico ao criar log de arquivo: {e}")

# MUDANÇA: Executa o setup imediatamente ao importar o módulo
setup_environment_and_logging()


# --- Importação de Módulos do Projeto ---
try:
    from modules import leitor_planilha, construtor_mensagem, arquivo_txt_sender
    from config import MAPEAMENTO_LOTEAMENTO, MAPEAMENTO_EMPRESA_POR_CODIGO, EMPRESA_PADRAO
except ImportError as e:
    logging.critical(f"Erro Crítico ao importar submódulos ou config: {e}", exc_info=True)
    raise

# --- Carregamento das Configurações do .env ---
# MUDANÇA: Agrupado em um dicionário para facilitar o acesso
CONFIG = {
    'col_nome': os.getenv('COLUNA_NOME'),
    'col_telefone': os.getenv('COLUNA_TELEFONE'),
    'col_lote': os.getenv('COLUNA_LOTE'),
    'col_valor': os.getenv('COLUNA_VALOR'),
    'col_vencimento': os.getenv('COLUNA_VENCIMENTO'),
    'col_loteamento': os.getenv('COLUNA_LOTEAMENTO'),
    'delay_segundos': int(os.getenv('DELAY_SEGUNDOS', '1')),
    'telefone_contato': os.getenv('TELEFONE_CONTATO', '[Seu Número de Telefone]')
}

# --- Funções Auxiliares ---

def _formatar_moeda(valor):
    """Tenta formatar um float como moeda BRL, com fallback robusto."""
    if not isinstance(valor, (int, float)):
        return "N/A"
    try:
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    except locale.Error:
        logging.warning("Locale 'pt_BR.UTF-8' não disponível. Usando formatação manual.")
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return locale.currency(valor, grouping=True, symbol='R$')

def _salvar_descartados(df_descartados, nome_arquivo_original):
    """
    # MUDANÇA: Função dedicada para salvar a planilha de contatos descartados.
    """
    if df_descartados.empty:
        return None, 0
    
    dir_descartados = "contatos_descartados"
    os.makedirs(dir_descartados, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    nome_base = os.path.splitext(os.path.basename(nome_arquivo_original))[0]
    caminho_arquivo = os.path.join(dir_descartados, f"{nome_base}_descartados_{timestamp}.xlsx")

    try:
        df_descartados.to_excel(caminho_arquivo, index=False)
        logging.info(f"{len(df_descartados)} contatos descartados salvos em: {caminho_arquivo}")
        return caminho_arquivo, len(df_descartados)
    except Exception as e:
        logging.error(f"Falha ao salvar arquivo de descartados: {e}")
        return None, 0

def _processar_grupo_cliente(telefone, grupo_df):
    """
    # MUDANÇA: Função dedicada para processar um grupo de parcelas de um mesmo telefone.
    Agrega dados, constrói e salva a mensagem.
    """
    nome_cliente = grupo_df[CONFIG['col_nome']].iloc[0]
    logging.info(f"Processando grupo para Cliente: {nome_cliente} | Telefone: {telefone} | Parcelas: {len(grupo_df)}")

    lista_parcelas_info = []
    valor_total = 0.0
    
    # Lógica para determinar a empresa (simplificada)
    codigos_loteamento = grupo_df[CONFIG['col_loteamento']].str.split('/').str[0].str.strip().unique()
    empresas_encontradas = {MAPEAMENTO_EMPRESA_POR_CODIGO.get(cod) for cod in codigos_loteamento if cod in MAPEAMENTO_EMPRESA_POR_CODIGO}
    
    if len(empresas_encontradas) == 1:
        nome_empresa = empresas_encontradas.pop()
    else:
        if len(empresas_encontradas) > 1:
            logging.warning(f"Cliente {nome_cliente} com múltiplas empresas. Usando empresa padrão.")
        nome_empresa = EMPRESA_PADRAO

    for _, parcela in grupo_df.iterrows():
        cod_loteamento = str(parcela.get(CONFIG['col_loteamento'], '')).split('/')[0].strip()
        valor_parcela = parcela.get(CONFIG['col_valor'], 0.0)
        
        info = {
            'loteamento_nome': MAPEAMENTO_LOTEAMENTO.get(cod_loteamento, f"[Cód:{cod_loteamento}]"),
            'lote': parcela.get(CONFIG['col_lote'], 'N/A'),
            'vencimento': parcela.get(CONFIG['col_vencimento']).strftime('%d/%m/%Y') if pd.notna(parcela.get(CONFIG['col_vencimento'])) else 'N/A',
            'valor': _formatar_moeda(valor_parcela)
        }
        lista_parcelas_info.append(info)
        if isinstance(valor_parcela, (int, float)):
            valor_total += valor_parcela
            
    # Cria e salva a mensagem
    mensagem = construtor_mensagem.criar_mensagem_consolidada(
        nome_cliente=nome_cliente,
        lista_parcelas_info=lista_parcelas_info,
        valor_total_fmt=_formatar_moeda(valor_total),
        nome_empresa=nome_empresa,
        telefone_contato=CONFIG['telefone_contato']
    )
    
    return arquivo_txt_sender.salvar_mensagem_em_txt(telefone, nome_cliente, mensagem)


# --- FUNÇÃO PRINCIPAL ---
def processar_cobrancas(arquivo_planilha_input):
    logging.info("--- [INÍCIO] Processamento de Cobranças ---")
    
    # Validação inicial de configuração
    if not CONFIG['col_nome'] or not CONFIG['col_telefone']:
        msg = "ERRO CRÍTICO: Nomes das colunas de NOME ou TELEFONE não definidos no .env."
        logging.error(msg)
        return {'status': 'falha', 'message': msg}

    # Carrega e pré-processa a planilha
    df = leitor_planilha.carregar_inadimplentes(arquivo_planilha_input, CONFIG)
    nome_arquivo = getattr(arquivo_planilha_input, 'name', 'arquivo_local.xlsx')
    
    resultados = {
        'status': 'falha', 'message': 'Processo não iniciado.',
        'planilha_processada': nome_arquivo, 'total_registros_carregados': 0,
        'output_dir': arquivo_txt_sender.DIRETORIO_SAIDA,
    }

    if df is None or df.empty:
        msg = f"Falha ao carregar dados de '{nome_arquivo}' ou planilha vazia/inválida."
        logging.error(msg)
        resultados['message'] = msg
        if df is not None: resultados['total_registros_carregados'] = len(df)
        return resultados

    resultados['total_registros_carregados'] = len(df)
    
    # Separa contatos válidos e descartados
    filtro_validos = df[CONFIG['col_telefone']].notna() & (df[CONFIG['col_telefone']] != '')
    df_validos = df[filtro_validos]
    df_descartados = df[~filtro_validos]
    
    # Salva descartados em um arquivo Excel separado
    caminho_descartados, num_descartados = _salvar_descartados(df_descartados, nome_arquivo)
    resultados.update({
        'caminho_arquivo_descartados': caminho_descartados,
        'num_descartados_arquivo': num_descartados,
        'total_com_telefone': len(df_validos)
    })

    if df_validos.empty:
        msg = "Nenhum registro com telefone válido foi encontrado para processar."
        logging.info(msg)
        resultados.update({'status': 'sucesso_sem_dados', 'message': msg})
        return resultados
        
    # Agrupa por telefone e processa cada grupo
    grupos = df_validos.groupby(CONFIG['col_telefone'], sort=False)
    resultados['total_telefones_unicos'] = len(grupos)
    
    sucessos = 0
    falhas = 0
    
    logging.info(f"Iniciando geração de mensagens para {len(grupos)} telefones únicos.")
    
    for i, (telefone, grupo_df) in enumerate(grupos):
        if _processar_grupo_cliente(telefone, grupo_df):
            sucessos += 1
        else:
            falhas += 1
        
        if i < len(grupos) - 1: # Evita delay após o último item
            time.sleep(CONFIG['delay_segundos'])
            
    # Finaliza e retorna os resultados consolidados
    if falhas == 0:
        message = "Processamento concluído com sucesso!"
        status = 'sucesso'
    else:
        message = f"Processo concluído com {falhas} falhas."
        status = 'sucesso_parcial'

    resultados.update({
        'status': status,
        'message': message,
        'mensagens_sucesso': sucessos,
        'falhas_envio': falhas
    })
    
    logging.info(f"--- [FIM] Processamento de Cobranças. Sucessos: {sucessos}, Falhas: {falhas} ---")
    return resultados

# --- Bloco para execução standalone ---
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog
    
    logging.info("Executando em modo standalone.")
    root = tk.Tk(); root.withdraw()
    caminho_arquivo_tk = filedialog.askopenfilename(title="Selecione a planilha Excel")
    if caminho_arquivo_tk:
        resultados_exec = processar_cobrancas(caminho_arquivo_tk)
        print("\n--- RESULTADO DA EXECUÇÃO ---")
        for k, v in resultados_exec.items():
            print(f"{k}: {v}")
    else:
        print("Nenhum arquivo selecionado. Encerrando.")