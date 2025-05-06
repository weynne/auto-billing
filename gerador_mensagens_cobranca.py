#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# gerador_mensagens_cobranca.py

import time
import os
from dotenv import load_dotenv 
import pandas as pd
import logging
import sys
import locale

# --- Carregamento .env ---
if load_dotenv(): 
    logging.info("GMC_LOG: Arquivo .env carregado (por gerador_mensagens_cobranca).")
else: 
    logging.warning("GMC_LOG: AVISO - Arquivo .env não encontrado (por gerador_mensagens_cobranca).")

# --- Configuração do Logging ---
logger_raiz = logging.getLogger()
if logger_raiz.level == logging.NOTSET or logger_raiz.level > logging.INFO: # Garante que INFO passe
    logger_raiz.setLevel(logging.INFO)

# Adiciona handlers apenas se não existirem tipos similares para evitar duplicação
console_handler_exists = any(isinstance(h, logging.StreamHandler) and h.stream == sys.stdout for h in logger_raiz.handlers)
file_handler_log_path = os.path.abspath("processamento_cobrancas.log") 
file_handler_exists = any(isinstance(h, logging.FileHandler) and getattr(h, 'baseFilename', '') == file_handler_log_path for h in logger_raiz.handlers)

if not console_handler_exists:
    log_formatter_console = logging.Formatter('%(asctime)s [%(levelname)s] %(module)s L%(lineno)d: %(message)s', datefmt='%H:%M:%S')
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter_console)
    console_handler.setLevel(logging.INFO) 
    logger_raiz.addHandler(console_handler)

if not file_handler_exists:
    try:
        log_formatter_detalhado = logging.Formatter('%(asctime)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        file_handler = logging.FileHandler(file_handler_log_path, mode='a', encoding='utf-8')
        file_handler.setFormatter(log_formatter_detalhado)
        file_handler.setLevel(logging.INFO) 
        logger_raiz.addHandler(file_handler)
        # Não logar aqui para evitar recursão se o logging já estiver ativo por outro meio (Streamlit)
        # logging.info(f"GMC_LOG: File handler adicionado por gerador_mensagens_cobranca. Log: '{file_handler_log_path}'")
    except IOError as e_log:
        logging.error(f"GMC_LOG: ERRO CRÍTICO ao criar log de arquivo '{file_handler_log_path}': {e_log}")

# --- Carregamento de Módulos e Constantes do Projeto ---
try:
    from modules import leitor_planilha
    from modules import construtor_mensagem
    from modules import arquivo_txt_sender 
    from config import MAPEAMENTO_LOTEAMENTO, MAPEAMENTO_EMPRESA_POR_CODIGO, EMPRESA_PADRAO
except ImportError as e:
    logging.critical(f"GMC_LOG: Erro Crítico ao importar submódulos ou config: {e}", exc_info=True)
    if __name__ == "__main__": exit() # Se rodando standalone, pode sair
    else: raise # Se importado, levanta a exceção para o importador tratar

# --- Configurações do .env (lidas globalmente após load_dotenv) ---
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')
COLUNA_LOTE = os.getenv('COLUNA_LOTE') 
COLUNA_VALOR = os.getenv('COLUNA_VALOR')
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO')
COLUNA_LOTEAMENTO = os.getenv('COLUNA_LOTEAMENTO')
DELAY_ENTRE_MENSAGENS_SEGUNDOS_GLOBAL = 1 

# --- Configura Locale e Função Auxiliar de Moeda ---
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    try: 
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
        logging.warning("GMC_LOG: Locale 'pt_BR.UTF-8' não encontrado. Usando 'en_US.UTF-8'.")
    except locale.Error: 
        logging.warning("GMC_LOG: Nenhum locale (pt_BR, en_US) suportado encontrado. Formatação de moeda pode usar padrão do sistema.")

def formatar_valor_moeda_local(valor_float):
    if not isinstance(valor_float, (int, float)) or pd.isna(valor_float): return "[Valor Inválido]"
    try:
        # Tenta usar o locale para formatação, mas tem fallback para o método de substituição
        # se o locale não for explicitamente pt_BR.
        current_locale_numeric_tuple = locale.getlocale(locale.LC_NUMERIC) # (ex: ('pt_BR', 'UTF-8'))
        current_locale_numeric_str = current_locale_numeric_tuple[0] if current_locale_numeric_tuple else None
        
        if current_locale_numeric_str and 'pt_BR' in current_locale_numeric_str:
            return locale.currency(valor_float, grouping=True, symbol='R$')
        else: # Fallback para formatação manual se o locale não for pt_BR ou não estiver configurado
            valor_fmt = f'{valor_float:,.2f}' # Formato com vírgula como separador de milhar e ponto decimal
            if locale.localeconv()['decimal_point'] == '.': # Se o sistema usa ponto decimal
                 valor_fmt = valor_fmt.replace(',', '#TEMP#').replace('.', ',').replace('#TEMP#', '.')
            return f"R$ {valor_fmt}"
    except Exception as e:
        logging.warning(f"GMC_LOG: Erro ao formatar valor {valor_float} para moeda: {e}. Usando fallback.")
        # Fallback muito simples em caso de erro extremo
        return f"R$ {valor_float:.2f}".replace('.',',')


def selecionar_arquivo_planilha_standalone():
    import tkinter as tk
    from tkinter import filedialog
    logging.info("GMC_LOG (Standalone): Abrindo janela para seleção da planilha...")
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    caminho_arquivo = filedialog.askopenfilename(title="Selecione a planilha Excel TRATADA (.xlsx)", filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
    root.destroy()
    if caminho_arquivo: logging.info(f"GMC_LOG (Standalone): Arquivo selecionado: {os.path.basename(caminho_arquivo)}")
    else: logging.warning("GMC_LOG (Standalone): Nenhum arquivo foi selecionado.")
    return caminho_arquivo

def processar_cobrancas(arquivo_planilha_input=None, delay_override=None):
    logging.info("GMC_LOG: [INICIO] processar_cobrancas") 
    logging.debug(f"GMC_LOG: Handlers no logger raiz ao iniciar processar_cobrancas: {logging.getLogger().handlers}")

    current_delay_seconds = DELAY_ENTRE_MENSAGENS_SEGUNDOS_GLOBAL
    if delay_override is not None:
        try: current_delay_seconds = int(delay_override)
        except ValueError: logging.warning(f"GMC_LOG: Delay_override inválido '{delay_override}'. Usando padrão do .env ou código.")
    else:
        try: current_delay_seconds = int(os.getenv('DELAY_SEGUNDOS', str(DELAY_ENTRE_MENSAGENS_SEGUNDOS_GLOBAL)))
        except ValueError: logging.warning(f"GMC_LOG: DELAY_SEGUNDOS do .env inválido. Usando padrão do código: {current_delay_seconds}s.")
    logging.info(f"GMC_LOG: Delay entre mensagens: {current_delay_seconds}s.")

    resultados_processamento = {
        'status': 'falha', 'message': 'Processo não iniciado.',
        'planilha_processada': None, 'total_registros_carregados': 0,
        'total_descartados_sem_telefone': 0, 'total_com_telefone': 0,
        'total_telefones_unicos': 0, 'mensagens_sucesso': 0,
        'falhas_envio': 0, 'output_dir': arquivo_txt_sender.DIRETORIO_SAIDA,
        'caminho_arquivo_descartados': None, 'num_descartados_arquivo': 0
    }

    if not COLUNA_NOME or not COLUNA_TELEFONE:
        msg_erro = "GMC_LOG: ERRO CRÍTICO - COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env."
        logging.error(msg_erro)
        resultados_processamento['message'] = msg_erro; return resultados_processamento

    nome_arquivo_processado = "N/A"
    if arquivo_planilha_input:
        nome_arquivo_processado = getattr(arquivo_planilha_input, 'name', 'arquivo_da_interface.xlsx')
        logging.info(f"GMC_LOG: Usando planilha da interface: {nome_arquivo_processado}")
        dados_inadimplentes = leitor_planilha.carregar_inadimplentes(arquivo_planilha_input)
    else:
        caminho_arquivo_tk = selecionar_arquivo_planilha_standalone()
        if not caminho_arquivo_tk: resultados_processamento['message'] = "Processo cancelado (nenhum arquivo selecionado)."; return resultados_processamento 
        nome_arquivo_processado = os.path.basename(caminho_arquivo_tk)
        dados_inadimplentes = leitor_planilha.carregar_inadimplentes(caminho_arquivo_tk)
    
    resultados_processamento['planilha_processada'] = nome_arquivo_processado
    if dados_inadimplentes is None or dados_inadimplentes.empty:
        msg_erro = f"GMC_LOG: Falha ao carregar dados de '{nome_arquivo_processado}' ou planilha vazia."
        logging.error(msg_erro)
        resultados_processamento['message'] = msg_erro
        if dados_inadimplentes is not None: resultados_processamento['total_registros_carregados'] = len(dados_inadimplentes)
        return resultados_processamento
    resultados_processamento['total_registros_carregados'] = len(dados_inadimplentes)

    caminho_arquivo_descartados_gerado = None
    num_registros_no_arquivo_descartados = 0
    telefone_ausente_mask = pd.Series(False, index=dados_inadimplentes.index) 

    if COLUNA_TELEFONE in dados_inadimplentes.columns:
        telefone_ausente_mask = pd.isna(dados_inadimplentes[COLUNA_TELEFONE]) | (dados_inadimplentes[COLUNA_TELEFONE].astype(str).str.strip() == '')
        df_descartados_sem_telefone = dados_inadimplentes[telefone_ausente_mask].copy()
        if not df_descartados_sem_telefone.empty:
            num_registros_no_arquivo_descartados = len(df_descartados_sem_telefone)
            DIR_DESCARTADOS = "contatos_descartados" 
            os.makedirs(DIR_DESCARTADOS, exist_ok=True)
            timestamp_atual = time.strftime("%Y%m%d_%H%M%S")
            nome_base_planilha = os.path.splitext(os.path.basename(nome_arquivo_processado))[0]
            nome_arq_desc = f"{nome_base_planilha}_descartados_sem_telefone_{timestamp_atual}.xlsx"
            caminho_arquivo_descartados_gerado = os.path.join(DIR_DESCARTADOS, nome_arq_desc)
            try:
                df_descartados_sem_telefone.loc[:, 'MOTIVO_DESCARTE'] = 'Telefone ausente ou inválido na planilha original'
                df_descartados_sem_telefone.to_excel(caminho_arquivo_descartados_gerado, index=False)
                logging.info(f"GMC_LOG: Arquivo de {num_registros_no_arquivo_descartados} contatos descartados salvo em: {caminho_arquivo_descartados_gerado}")
            except Exception as e_excel:
                logging.error(f"GMC_LOG: Erro ao salvar arquivo de descartados '{caminho_arquivo_descartados_gerado}': {e_excel}")
                caminho_arquivo_descartados_gerado = None; num_registros_no_arquivo_descartados = 0 
        else: logging.info("GMC_LOG: Nenhum registro com telefone ausente para gerar arquivo de descartados.")
    else: logging.warning(f"GMC_LOG: Coluna '{COLUNA_TELEFONE}' não encontrada em dados_inadimplentes. Não foi possível gerar arquivo de descartados.")
    resultados_processamento['caminho_arquivo_descartados'] = caminho_arquivo_descartados_gerado
    resultados_processamento['num_descartados_arquivo'] = num_registros_no_arquivo_descartados

    dados_validos = dados_inadimplentes.dropna(subset=[COLUNA_TELEFONE])
    if not dados_validos.empty and COLUNA_TELEFONE in dados_validos.columns:
         dados_validos = dados_validos[dados_validos[COLUNA_TELEFONE].astype(str).str.strip() != '']
    total_apos_filtro_telefone = len(dados_validos)
    total_descartados_do_processamento = len(dados_inadimplentes) - total_apos_filtro_telefone
    resultados_processamento.update({
        'total_descartados_sem_telefone': total_descartados_do_processamento,
        'total_com_telefone': total_apos_filtro_telefone
    })
    if total_descartados_do_processamento > 0:
        logging.warning(f"GMC_LOG: --- {total_descartados_do_processamento} registros foram REMOVIDOS do processamento por telefone inválido/ausente. ---")
        if dados_inadimplentes[telefone_ausente_mask].index.any():
            # A condição de logar apenas se o arquivo de descarte não foi gerado foi removida
            # para sempre logar alguns detalhes, se houver descartes.
            logging.info("GMC_LOG: Detalhando alguns descartados (por falta de telefone) no log:")
            count_log_descartados = 0
            for idx_descartado in dados_inadimplentes[telefone_ausente_mask].index: 
                if count_log_descartados < 10: 
                    try:
                        nome_descartado_log = dados_inadimplentes.loc[idx_descartado, COLUNA_NOME] if COLUNA_NOME in dados_inadimplentes.columns else "[Nome N/D]"
                        # ***** ESTA É A LINHA CORRIGIDA *****
                        logging.warning(f"GMC_LOG:  - DESCARTADO (Ref. Linha Planilha: {idx_descartado + 1}): Cliente '{nome_descartado_log}' - Telefone ausente.")
                    except Exception as e_log_det:
                        # ***** E ESTA TAMBÉM *****
                        logging.warning(f"GMC_LOG:  - DESCARTADO (Ref. Linha Planilha: {idx_descartado + 1}): Telefone ausente (Erro ao buscar detalhes: {e_log_det})")
                    finally:
                        count_log_descartados += 1
                else:
                    if total_descartados_do_processamento > count_log_descartados:
                        logging.info(f"GMC_LOG:  ... e mais {total_descartados_do_processamento - count_log_descartados} registros descartados não detalhados no log.")
                    break
        elif total_descartados_do_processamento > 0 : 
             logging.warning(f"GMC_LOG: {total_descartados_do_processamento} registros descartados, mas não foi possível detalhá-los individualmente no log (verifique a máscara 'telefone_ausente_mask').")


    if dados_validos.empty:
        msg_info = "GMC_LOG: Nenhum registro com telefone válido para processar mensagens."
        logging.info(msg_info); resultados_processamento.update({'message': msg_info, 'status': 'sucesso_sem_dados_para_msg'})
        return resultados_processamento

    logging.info(f"GMC_LOG: Agrupando {len(dados_validos)} registros válidos por telefone...")
    try:
        colunas_ordenacao = [COLUNA_NOME] if COLUNA_NOME and COLUNA_NOME in dados_validos.columns else []
        dados_para_agrupar = dados_validos.copy() 
        if COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in dados_para_agrupar.columns and \
           pd.api.types.is_datetime64_any_dtype(dados_para_agrupar[COLUNA_VENCIMENTO]):
            if COLUNA_VENCIMENTO not in colunas_ordenacao: colunas_ordenacao.append(COLUNA_VENCIMENTO)
        elif COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in dados_para_agrupar.columns:
             logging.warning(f"GMC_LOG: Coluna '{COLUNA_VENCIMENTO}' não é data/datetime para ordenação.")
        if colunas_ordenacao:
            na_pos = 'last' if COLUNA_VENCIMENTO in colunas_ordenacao else 'first'
            dados_para_agrupar.sort_values(by=colunas_ordenacao, na_position=na_pos, inplace=True)
        else: logging.debug("GMC_LOG: Nenhuma coluna válida para ordenação pré-agrupamento.") # Mudado para debug
        
        grupos_por_telefone = dados_para_agrupar.groupby(COLUNA_TELEFONE, sort=False)
        total_telefones_unicos = len(grupos_por_telefone)
        resultados_processamento['total_telefones_unicos'] = total_telefones_unicos
    except Exception as e_group:
        msg_erro = f"GMC_LOG: Erro no agrupamento/ordenação: {e_group}"
        logging.exception(msg_erro); resultados_processamento.update({'message': msg_erro, 'status': 'falha_critica'})
        return resultados_processamento

    enviados_sucesso_count = 0; falhas_envio_count = 0; telefones_processados_count = 0
    logging.info(f"GMC_LOG: \n--- Iniciando processamento CONSOLIDADO para {total_telefones_unicos} telefones únicos ---")
    
    for telefone, grupo in grupos_por_telefone:
        telefones_processados_count += 1; telefone_str = str(telefone)
        logging.info(f"GMC_LOG: Processando Telefone {telefones_processados_count}/{total_telefones_unicos}: {telefone_str}")
        
        if not telefone_str.startswith("+55") or not (len(telefone_str) == 13 or len(telefone_str) == 14):
            logging.warning(f"GMC_LOG: AVISO - Telefone '{telefone_str}' inválido. Pulando grupo.")
            falhas_envio_count += 1; continue
        
        nome_cliente_valido = "[Cliente N/D]"
        if COLUNA_NOME in grupo.columns and not grupo[COLUNA_NOME].dropna().empty:
            nome_cliente_valido = grupo[COLUNA_NOME].dropna().iloc[0]
        elif COLUNA_NOME in grupo.columns and not grupo.empty:
             nome_cliente_valido = grupo[COLUNA_NOME].iloc[0] if pd.notna(grupo[COLUNA_NOME].iloc[0]) else "[Cliente N/D]"
        
        logging.info(f"GMC_LOG:   - Cliente: {nome_cliente_valido} | Parcelas: {len(grupo)}")

        lista_parcelas_info = []; valor_total_num = 0.0
        nome_empresa_grupo = EMPRESA_PADRAO; primeiro_nome_empresa_encontrado = None
        inconsistencia_empresa = False; primeira_parcela_valida_para_empresa = True
        try:
            for idx, parcela_row in grupo.iterrows(): # idx é o índice do DataFrame original (dados_validos)
                lote_code = None; nome_empresa_parcela = None
                duplicata_raw = parcela_row.get(COLUNA_LOTEAMENTO)
                
                if COLUNA_LOTEAMENTO and isinstance(duplicata_raw, str) and '/' in duplicata_raw:
                    parts = duplicata_raw.split('/'); lote_code = parts[0].strip()
                    nome_empresa_parcela = MAPEAMENTO_EMPRESA_POR_CODIGO.get(lote_code)
                
                if COLUNA_LOTEAMENTO and not lote_code and duplicata_raw and pd.notna(duplicata_raw): 
                    logging.warning(f"GMC_LOG: Parcela (Cli:{nome_cliente_valido}, Idx DF: {idx}): Cód. loteamento inválido '{duplicata_raw}'.")
                
                if nome_empresa_parcela:
                    if primeira_parcela_valida_para_empresa:
                        primeiro_nome_empresa_encontrado = nome_empresa_parcela; nome_empresa_grupo = nome_empresa_parcela
                        primeira_parcela_valida_para_empresa = False
                    elif not inconsistencia_empresa and primeiro_nome_empresa_encontrado != nome_empresa_parcela:
                        inconsistencia_empresa = True; nome_empresa_grupo = EMPRESA_PADRAO
                        logging.warning(f"GMC_LOG: Cliente {nome_cliente_valido} ({telefone_str}) com empresas diferentes. Usando padrão.")
                
                nome_loteamento_parcela = MAPEAMENTO_LOTEAMENTO.get(lote_code, f'[Cód:{lote_code}]') if lote_code else '[Loteam. N/D]'
                lote_raw = parcela_row.get(COLUNA_LOTE) if COLUNA_LOTE else None
                venc_obj = parcela_row.get(COLUNA_VENCIMENTO) if COLUNA_VENCIMENTO else None
                valor_num = parcela_row.get(COLUNA_VALOR) if COLUNA_VALOR else None

                if COLUNA_LOTE and pd.isna(lote_raw): logging.warning(f"GMC_LOG: Parcela (Cli:{nome_cliente_valido}, Idx DF: {idx}): Lote ('{COLUNA_LOTE}') ausente.")
                if COLUNA_VENCIMENTO and pd.isna(venc_obj): logging.warning(f"GMC_LOG: Parcela (Cli:{nome_cliente_valido}, Idx DF: {idx}): Venc. ('{COLUNA_VENCIMENTO}') ausente/inválido. Lido: '{parcela_row.get(COLUNA_VENCIMENTO)}'.")
                if COLUNA_VALOR and pd.isna(valor_num): logging.warning(f"GMC_LOG: Parcela (Cli:{nome_cliente_valido}, Idx DF: {idx}): Valor ('{COLUNA_VALOR}') ausente/inválido. Lido: '{parcela_row.get(COLUNA_VALOR)}'.")
                
                lote_info = "[Ref N/D]"; venc_fmt = "[Data N/D]"; valor_fmt = "[Valor N/D]"; valor_parcela_num = 0.0
                if COLUNA_LOTE and pd.notna(lote_raw): lote_str = str(lote_raw).strip(); lote_info = lote_str.split('/')[0].strip() if '/' in lote_str else lote_str
                if pd.notna(venc_obj) and hasattr(venc_obj, 'strftime'):
                    try: venc_fmt = venc_obj.strftime('%d/%m/%Y')
                    except ValueError: pass
                if pd.notna(valor_num):
                    try: valor_parcela_num = float(valor_num); valor_fmt = formatar_valor_moeda_local(valor_parcela_num)
                    except (ValueError, TypeError): valor_fmt = "[Erro Format. Valor]"
                
                lista_parcelas_info.append({'lote': lote_info, 'vencimento': venc_fmt, 'valor': valor_fmt, 'loteamento_nome': nome_loteamento_parcela})
                if pd.notna(valor_num): valor_total_num += valor_parcela_num
        except Exception as e_agg:
            logging.error(f"GMC_LOG: ERRO ao agregar dados para {nome_cliente_valido} ({telefone_str}): {e_agg}", exc_info=True)
            falhas_envio_count += 1; continue
        
        valor_total_fmt = formatar_valor_moeda_local(valor_total_num)
        try:
            mensagem = construtor_mensagem.criar_mensagem_consolidada(nome_cliente_valido, lista_parcelas_info, valor_total_fmt, nome_empresa_grupo)
        except Exception as e_msg:
            logging.error(f"GMC_LOG: ERRO ao construir mensagem para '{nome_cliente_valido}': {e_msg}.", exc_info=True)
            falhas_envio_count += 1; continue
        
        sucesso_envio_msg = False
        try:
            sucesso_envio_msg = arquivo_txt_sender.salvar_mensagem_em_txt(telefone_str, nome_cliente_valido, mensagem)
        except Exception as e_send:
            logging.error(f"GMC_LOG: ERRO ao salvar msg para '{nome_cliente_valido}': {e_send}", exc_info=True)
        
        if sucesso_envio_msg: enviados_sucesso_count += 1
        else: falhas_envio_count += 1; logging.warning(f"GMC_LOG:   - Falha ao salvar/enviar msg para {nome_cliente_valido} ({telefone_str}).")
        
        if total_telefones_unicos > 1 and telefones_processados_count < total_telefones_unicos and current_delay_seconds > 0:
            time.sleep(current_delay_seconds)

    resultados_processamento.update({
        'status': 'sucesso' if falhas_envio_count == 0 else 'sucesso_parcial', 
        'message': 'Processo concluído.' if falhas_envio_count == 0 else f'Processo concluído com {falhas_envio_count} falhas em grupos de telefone.', 
        'mensagens_sucesso': enviados_sucesso_count, 
        'falhas_envio': falhas_envio_count
    })
    logging.info(f"GMC_LOG: [FIM] processar_cobrancas. Sucesso: {enviados_sucesso_count}, Falhas: {falhas_envio_count}")
    return resultados_processamento

if __name__ == "__main__":
    logging.info("GMC_LOG: Executando 'gerador_mensagens_cobranca.py' em modo standalone.")
    resultados = processar_cobrancas() 
    if resultados:
        logging.info(f"GMC_LOG: Resultado final (standalone): {resultados.get('message')}")
        if resultados.get('caminho_arquivo_descartados'):
            logging.info(f"GMC_LOG: Arquivo de descartados: {resultados['caminho_arquivo_descartados']}")
    else:
        logging.error("GMC_LOG: Processamento standalone não retornou resultados.")