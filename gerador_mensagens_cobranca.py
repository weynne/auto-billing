#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# gerador_mensagens_cobranca.py

import time
import os
from dotenv import load_dotenv
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import logging
import sys
import locale # Para formatação de moeda

# --- Configuração do Logging ---
log_formatter_detalhado = logging.Formatter(
    '%(asctime)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
log_formatter_console = logging.Formatter('%(message)s')
logger_raiz = logging.getLogger()
logger_raiz.setLevel(logging.INFO)
if logger_raiz.hasHandlers(): logger_raiz.handlers.clear()
console_handler = logging.StreamHandler(sys.stdout) # Mudado para stdout
console_handler.setFormatter(log_formatter_console)
nome_arquivo_log = "processamento_cobrancas.log"
try:
    file_handler = logging.FileHandler(nome_arquivo_log, mode='a', encoding='utf-8')
    file_handler.setFormatter(log_formatter_detalhado)
    logger_raiz.addHandler(console_handler)
    logger_raiz.addHandler(file_handler)
    logging.info(f"--- Logging iniciado. Console: limpo | Arquivo: '{nome_arquivo_log}' (detalhado) ---")
except IOError as e:
    logger_raiz.addHandler(console_handler)
    logging.error(f"ERRO CRÍTICO ao criar/abrir log '{nome_arquivo_log}': {e}")
# ------------------------------

# --- Carregamento de Módulos e Env ---
try:
    from modules import leitor_planilha
    from modules import construtor_mensagem
    from modules import arquivo_txt_sender # Ou o futuro whatsapp_sender
    logging.info("Módulos locais importados.")
except ImportError as e:
    logging.exception("Erro Crítico: Falha ao importar módulos.")
    exit()

if load_dotenv(): logging.info("Arquivo .env carregado.")
else: logging.warning("AVISO: Arquivo .env não encontrado.")
# ------------------------------

# --- Configurações ---
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')
COLUNA_LOTE = os.getenv('LOTE') # Nome da coluna lote/ref
COLUNA_VALOR = os.getenv('COLUNA_VALOR') # Nome da coluna valor
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO') # Nome da coluna vencimento
try:
    valor_delay_str = os.getenv('DELAY_SEGUNDOS', '1')
    DELAY_ENTRE_MENSAGENS_SEGUNDOS = int(valor_delay_str)
    logging.info(f"Delay entre mensagens: {DELAY_ENTRE_MENSAGENS_SEGUNDOS}s.")
except ValueError:
    logging.warning(f"AVISO: Valor inválido ('{valor_delay_str}') para DELAY_SEGUNDOS. Usando padrão: 1s.")
    DELAY_ENTRE_MENSAGENS_SEGUNDOS = 1
# ------------------------------

# --- Configura Locale e Função Auxiliar de Moeda ---
use_locale = False
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    use_locale = True
    logging.info("Locale 'pt_BR.UTF-8' configurado para formatação de moeda.")
except locale.Error:
    # Este log agora funciona pois o logger está configurado
    logging.warning("Locale 'pt_BR.UTF-8' não encontrado. Usando formatação de moeda manual.")

def formatar_valor_moeda_local(valor_float):
    """Formata um float para moeda R$ (reutilizável)."""
    if not isinstance(valor_float, (int, float)) or pd.isna(valor_float):
         return "[Valor Inválido]"
    try:
        if use_locale:
            return locale.currency(valor_float, grouping=True, symbol='R$')
        else:
            return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception as e:
        logger.warning(f"Erro ao formatar valor {valor_float} para moeda: {e}")
        return "[Erro Formatação]"
# ------------------------------

# --- Função selecionar_arquivo_planilha (sem mudanças) ---
def selecionar_arquivo_planilha():
    """Abre janela gráfica para selecionar o arquivo da planilha."""
    logging.info("\nAbrindo janela para seleção da planilha...")
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    caminho_arquivo = filedialog.askopenfilename(title="Selecione a planilha Excel TRATADA", filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("Todos", "*.*")])
    root.destroy()
    if caminho_arquivo: logging.info(f"Arquivo selecionado: {os.path.basename(caminho_arquivo)}"); return caminho_arquivo
    else: logging.warning("Nenhum arquivo foi selecionado."); return None
# ------------------------------

# --- Função processar_cobrancas (LÓGICA DE CONSOLIDAÇÃO) ---
def processar_cobrancas():
    """ Orquestra leitura, AGRUPAMENTO, construção e salvamento/envio das mensagens consolidadas. """
    logging.info("\n--- Iniciando Processo de Geração CONSOLIDADA de Cobranças ---")

    # Verifica colunas essenciais para o processo
    if not COLUNA_NOME or not COLUNA_TELEFONE:
        logging.error("ERRO CRÍTICO: COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env.")
        return
    # Avisa sobre colunas opcionais que afetam a mensagem
    if not COLUNA_LOTE: logging.warning("AVISO: Variável 'LOTE' não definida no .env. Referência não será incluída.")
    if not COLUNA_VALOR: logging.warning("AVISO: Variável 'COLUNA_VALOR' não definida no .env. Valores não serão incluídos/somados.")
    if not COLUNA_VENCIMENTO: logging.warning("AVISO: Variável 'COLUNA_VENCIMENTO' não definida no .env. Datas de vencimento não serão incluídas.")

    caminho_planilha_selecionada = selecionar_arquivo_planilha()
    if not caminho_planilha_selecionada:
        logging.info("Processo cancelado. Encerrando.")
        return

    logging.info(f"\n--- Carregando e Processando Planilha: {os.path.basename(caminho_planilha_selecionada)} ---")
    dados_inadimplentes = leitor_planilha.carregar_inadimplentes(caminho_planilha_selecionada)

    if dados_inadimplentes is None or dados_inadimplentes.empty:
        logging.error("Falha ao carregar dados ou planilha vazia. Encerrando.")
        return

    # --- AGRUPAMENTO POR TELEFONE ---
    # Garante que a coluna de telefone não tenha NaNs/Inválidos antes de agrupar
    # A validação E.164 será feita no loop, aqui só removemos linhas sem telefone algum
    dados_validos = dados_inadimplentes.dropna(subset=[COLUNA_TELEFONE])
    # Valida o tipo da coluna telefone para o groupby (deve ser string)
    if not pd.api.types.is_string_dtype(dados_validos[COLUNA_TELEFONE]):
         logging.warning(f"Coluna de telefone '{COLUNA_TELEFONE}' não parece ser string após leitura. Tentando converter...")
         dados_validos[COLUNA_TELEFONE] = dados_validos[COLUNA_TELEFONE].astype(str)

    if dados_validos.empty:
         logging.info("Nenhum registro com número de telefone válido encontrado após limpeza inicial.")
         return

    logging.info("Agrupando registros por número de telefone...")
    try:
         colunas_ordenacao = [COLUNA_NOME]
         if COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in dados_validos.columns:
              # Tenta converter para data antes de ordenar (se ainda não for datetime)
              if not pd.api.types.is_datetime64_any_dtype(dados_validos[COLUNA_VENCIMENTO]):
                   dados_validos[COLUNA_VENCIMENTO] = pd.to_datetime(dados_validos[COLUNA_VENCIMENTO], errors='coerce', dayfirst=True)
              # Remove NaT antes de ordenar para evitar erros
              dados_validos_ord = dados_validos.dropna(subset=[COLUNA_VENCIMENTO])
              colunas_ordenacao.append(COLUNA_VENCIMENTO)
              dados_validos_ord = dados_validos_ord.sort_values(by=colunas_ordenacao)
              # Mantém linhas com data inválida sem ordenar por data
              dados_com_nat = dados_validos[dados_validos[COLUNA_VENCIMENTO].isnull()]
              dados_para_agrupar = pd.concat([dados_validos_ord, dados_com_nat])
         else:
              dados_para_agrupar = dados_validos.sort_values(by=colunas_ordenacao)


         grupos_por_telefone = dados_para_agrupar.groupby(COLUNA_TELEFONE, sort=False)
         total_telefones_unicos = len(grupos_por_telefone)
    except KeyError as e:
         logging.error(f"Erro ao agrupar/ordenar: Coluna '{e}' não encontrada.")
         return
    except Exception as e_group:
         logging.exception("Erro inesperado durante o agrupamento por telefone.")
         return
    # ---------------------------------

    enviados_sucesso = 0
    falhas_envio = 0
    telefones_processados = 0
    logging.info(f"\n--- Iniciando processamento CONSOLIDADO para {total_telefones_unicos} telefones únicos ---")

    # --- ITERAÇÃO POR GRUPO (TELEFONE) ---
    for telefone, grupo in grupos_por_telefone:
        telefones_processados += 1
        logging.info("-" * 20)
        logging.info(f"Processando Telefone {telefones_processados}/{total_telefones_unicos}: {telefone}")

        # Validar o telefone do grupo
        if not isinstance(telefone, str) or not telefone.startswith("+55") or not (len(telefone) == 13 or len(telefone) == 14):
            logging.warning(f"AVISO: Telefone do grupo inválido ou fora do padrão E.164 ('{telefone}'). Pulando este grupo.")
            falhas_envio += 1
            continue

        if grupo.empty:
            logging.warning(f"AVISO: Grupo para telefone {telefone} está vazio. Pulando.")
            continue

        # Pega o nome (já limpo) do primeiro registro do grupo
        nome_cliente = grupo[COLUNA_NOME].iloc[0] if COLUNA_NOME in grupo.columns and not grupo[COLUNA_NOME].empty else "Cliente"
        logging.info(f"  - Cliente: {nome_cliente}")
        logging.info(f"  - Nº de parcelas no grupo: {len(grupo)}")

        # --- Agregação dos dados das parcelas do grupo ---
        lista_parcelas_info = [] # Lista para guardar dicts com dados formatados
        valor_total_num = 0.0
        try:
            for _, parcela_row in grupo.iterrows():
                # Extrair dados já processados pelo leitor (data é obj, valor é numérico)
                lote_raw = parcela_row.get(COLUNA_LOTE) if COLUNA_LOTE else None
                venc_obj = parcela_row.get(COLUNA_VENCIMENTO) if COLUNA_VENCIMENTO else None
                valor_num = parcela_row.get(COLUNA_VALOR) if COLUNA_VALOR else None

                # Processar/Formatar Lote
                lote_info = "[Ref N/D]"
                if lote_raw and isinstance(lote_raw, str):
                    lote_texto = lote_raw.strip(); lote_info = lote_texto.split('/')[0].strip() if '/' in lote_texto else lote_texto
                elif lote_raw: lote_info = str(lote_raw).strip()

                # Formatar Vencimento
                venc_fmt = "[Data N/D]"
                if pd.notna(venc_obj) and hasattr(venc_obj, 'strftime'):
                    try: venc_fmt = venc_obj.strftime('%d/%m/%Y')
                    except ValueError: pass

                # Formatar Valor Individual e somar ao total
                valor_fmt = "[Valor N/D]"
                valor_parcela_num = 0.0
                if pd.notna(valor_num):
                     try:
                         valor_parcela_num = float(valor_num)
                         valor_fmt = formatar_valor_moeda_local(valor_parcela_num) # Usa helper local
                     except (ValueError, TypeError): pass
                else: valor_parcela_num = 0.0

                # Adiciona DADOS FORMATADOS à lista para o construtor
                lista_parcelas_info.append({
                    'lote': lote_info,
                    'vencimento': venc_fmt,
                    'valor': valor_fmt # String formatada R$ ...
                })
                valor_total_num += valor_parcela_num # Soma o valor numérico

        except Exception as e_agg:
             logging.error(f"ERRO ao agregar dados das parcelas para {nome_cliente} ({telefone}): {e_agg}")
             logging.exception("Detalhes da agregação:")
             falhas_envio += 1
             continue

        # Formata o valor total
        valor_total_fmt = formatar_valor_moeda_local(valor_total_num)

        # --- Construir Mensagem Consolidada ---
        try:
            # Chama a função no construtor que aceita a lista de parcelas FORMATADAS
            mensagem = construtor_mensagem.criar_mensagem_consolidada(
                nome_cliente,
                lista_parcelas_info,
                valor_total_fmt
            )
            logging.debug(f"Mensagem consolidada construída para {nome_cliente}")
        except Exception as e_msg:
            logging.error(f"ERRO: Falha ao construir mensagem CONSOLIDADA para '{nome_cliente}': {e_msg}.")
            logging.exception(f"Detalhes da construção:")
            falhas_envio += 1
            continue

        # --- Salvar/Enviar Mensagem Consolidada ---
        logging.info(f"  - Tentando salvar/enviar mensagem consolidada...")
        try:
            # Usa o sender atual (TXT) ou o futuro (WhatsApp)
            sucesso_envio = arquivo_txt_sender.salvar_mensagem_em_txt(telefone, nome_cliente, mensagem)
            # Ou: sucesso_envio = whatsapp_sender.enviar_mensagem(telefone, mensagem)
        except Exception as e_send:
            logging.error(f"ERRO: Falha inesperada ao tentar salvar/enviar msg consolidada para '{nome_cliente}': {e_send}")
            logging.exception(f"Detalhes do envio:")
            sucesso_envio = False

        if sucesso_envio:
            enviados_sucesso += 1
        else:
            falhas_envio += 1
            logging.info(f"  - Falha ao salvar/enviar a mensagem consolidada.")

        # --- Pausa entre TELEFONES ---
        if total_telefones_unicos > 1 and telefones_processados < total_telefones_unicos:
            logging.info(f"  - Aguardando {DELAY_ENTRE_MENSAGENS_SEGUNDOS} segundo(s)...")
            time.sleep(DELAY_ENTRE_MENSAGENS_SEGUNDOS)
    # --- Fim do loop por telefone ---

    # --- Resumo final ---
    logging.info("\n" + "="*40)
    logging.info("--- Processo CONSOLIDADO Concluído ---")
    logging.info(f"Planilha processada: {os.path.basename(caminho_planilha_selecionada)}")
    logging.info(f"Total de registros na planilha: {len(dados_inadimplentes)}")
    logging.info(f"Total de telefones únicos processados: {telefones_processados}/{total_telefones_unicos}")
    logging.info(f"Mensagens consolidadas salvas/enviadas com sucesso: {enviados_sucesso}")
    logging.info(f"Telefones com falha no processamento/envio: {falhas_envio}")
    if enviados_sucesso > 0 and hasattr(arquivo_txt_sender, 'DIRETORIO_SAIDA'):
        try: dir_saida = arquivo_txt_sender.DIRETORIO_SAIDA; logging.info(f"Verifique os arquivos TXT na pasta: '{dir_saida}'")
        except AttributeError: logging.warning("AVISO: Não foi possível determinar diretório de saída.")
    logging.info("="*40)
# ------------------------------

# --- Ponto de entrada ---
if __name__ == "__main__":
    processar_cobrancas()
# ------------------------------