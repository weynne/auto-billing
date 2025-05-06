#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# gerador_mensagens_cobranca.py

import time
import os
from dotenv import load_dotenv
import tkinter as tk
from tkinter import filedialog
import pandas as pd # Import pandas
import logging
import sys
import locale

# --- Carregamento de Módulos e Env ---
try:
    from modules import leitor_planilha
    from modules import construtor_mensagem
    from modules import arquivo_txt_sender
    # Importa os mapeamentos/constantes de config.py
    from config import MAPEAMENTO_LOTEAMENTO, MAPEAMENTO_EMPRESA_POR_CODIGO, EMPRESA_PADRAO
    logging.info("Módulos locais e configuração importados.")
except ImportError as e:
    logging.exception("Erro Crítico: Falha ao importar módulos ou config.py.")
    exit()
except NameError as e:
    logging.exception("Erro Crítico: Falha ao importar mapeamentos/constantes de config.py.")
    exit()

# --- Configuração do Logging ---
log_formatter_detalhado = logging.Formatter(
    '%(asctime)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
log_formatter_console = logging.Formatter('%(message)s')
logger_raiz = logging.getLogger()
logger_raiz.setLevel(logging.INFO)
if logger_raiz.hasHandlers(): logger_raiz.handlers.clear()
console_handler = logging.StreamHandler(sys.stdout)
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

# --- Carregamento .env ---
if load_dotenv(): logging.info("Arquivo .env carregado.")
else: logging.warning("AVISO: Arquivo .env não encontrado.")
# ------------------------------

# --- Configurações do .env ---
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')
COLUNA_LOTE = os.getenv('COLUNA_LOTE')
COLUNA_VALOR = os.getenv('COLUNA_VALOR')
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO')
COLUNA_LOTEAMENTO = os.getenv('COLUNA_LOTEAMENTO')

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
    # Tenta configurar para Português do Brasil com UTF-8
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    use_locale = True
    logging.info("Locale 'pt_BR.UTF-8' configurado para formatação de moeda.")
except locale.Error:
    # Fallback para Inglês (ou padrão do sistema) se pt_BR não estiver disponível
    try:
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
        logging.warning("Locale 'pt_BR.UTF-8' não encontrado. Usando 'en_US.UTF-8' e formatação manual para R$.")
    except locale.Error:
        logging.warning("Locales 'pt_BR.UTF-8' e 'en_US.UTF-8' não encontrados. Usando formatação de moeda manual.")

def formatar_valor_moeda_local(valor_float):
    """Formata um float para moeda R$ (reutilizável)."""
    if not isinstance(valor_float, (int, float)) or pd.isna(valor_float):
        return "[Valor Inválido]"
    try:
        valor_fmt = f'{valor_float:,.2f}'
        # Garante o padrão brasileiro trocando os separadores se necessário
        if locale.localeconv()['decimal_point'] == '.':
            # Se o locale base usa ponto decimal, troca ponto por X, vírgula por ponto, X por vírgula
            valor_fmt = valor_fmt.replace(',', 'X').replace('.', ',').replace('X', '.')
        return f"R$ {valor_fmt}"
    except Exception as e:
        logging.warning(f"Erro ao formatar valor {valor_float} para moeda: {e}")
        return "[Erro Formatação]"
# ------------------------------


# --- Função selecionar_arquivo_planilha ---
def selecionar_arquivo_planilha():
    """Abre janela gráfica para selecionar o arquivo da planilha."""
    logging.info("\nAbrindo janela para seleção da planilha...")
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha Excel TRATADA (.xlsx)",
        filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
    root.destroy()
    if caminho_arquivo: logging.info(f"Arquivo selecionado: {os.path.basename(caminho_arquivo)}"); return caminho_arquivo
    else: logging.warning("Nenhum arquivo foi selecionado."); return None
# ------------------------------

# --- Função processar_cobrancas ---
def processar_cobrancas():
    logging.info("\n--- Iniciando Processo de Geração CONSOLIDADA de Cobranças ---")

    # --- Verificações Iniciais ---
    if not COLUNA_NOME or not COLUNA_TELEFONE:
        logging.error("ERRO CRÍTICO: COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env.")
        return
    # Avisos sobre colunas opcionais
    if not COLUNA_LOTE: logging.warning("AVISO: Variável 'COLUNA_LOTE' não definida no .env.")
    if not COLUNA_VALOR: logging.warning("AVISO: Variável 'COLUNA_VALOR' não definida no .env.")
    if not COLUNA_VENCIMENTO: logging.warning("AVISO: Variável 'COLUNA_VENCIMENTO' não definida no .env.")
    if not COLUNA_LOTEAMENTO: logging.warning("AVISO: Variável 'COLUNA_LOTEAMENTO' não definida no .env. Nomes de loteamento e empresa dinâmica não funcionarão.")

    # --- Seleção e Leitura da Planilha ---
    caminho_planilha_selecionada = selecionar_arquivo_planilha()
    if not caminho_planilha_selecionada:
        logging.info("Processo cancelado. Encerrando.")
        return
    logging.info(f"\n--- Carregando e Processando Planilha: {os.path.basename(caminho_planilha_selecionada)} ---")
    dados_inadimplentes = leitor_planilha.carregar_inadimplentes(caminho_planilha_selecionada)
    if dados_inadimplentes is None or dados_inadimplentes.empty:
        logging.error("Falha ao carregar dados ou planilha vazia. Encerrando.")
        return

    # --- Filtro e Log de Registros Sem Telefone ---
    total_antes_filtro_telefone = len(dados_inadimplentes)
    logging.info(f"Verificando coluna de telefone ('{COLUNA_TELEFONE}') para {total_antes_filtro_telefone} registros carregados...")

    # Guarda os índices originais antes de remover linhas sem telefone
    indices_originais = dados_inadimplentes.index

    # Filtra removendo linhas onde a coluna de telefone está vazia/nula
    dados_validos = dados_inadimplentes.dropna(subset=[COLUNA_TELEFONE])
    indices_validos = dados_validos.index
    total_apos_filtro_telefone = len(dados_validos)
    total_descartados_sem_telefone = total_antes_filtro_telefone - total_apos_filtro_telefone

    logging.info(f"Encontrados {total_apos_filtro_telefone} registros com telefone presente na coluna '{COLUNA_TELEFONE}'.")

    # LOG ADICIONAL: Informa quantos e quais registros foram descartados por falta de telefone
    if total_descartados_sem_telefone > 0:
        logging.warning(f"--- ATENÇÃO: {total_descartados_sem_telefone} registros foram DESCARTADOS por não possuírem telefone na coluna '{COLUNA_TELEFONE}'. ---")
        # Logar detalhes de TODOS os descartados
        indices_descartados = indices_originais.difference(indices_validos)
        for idx_descartado in indices_descartados:
            try:
                # Tenta pegar o nome do cliente descartado para dar contexto
                nome_descartado = dados_inadimplentes.loc[idx_descartado, COLUNA_NOME] if COLUNA_NOME in dados_inadimplentes.columns else "[Nome N/D]"
                logging.warning(f"  - DESCARTADO (Índice Original Aprox.: {idx_descartado}): Cliente '{nome_descartado}' - Telefone ausente.")
            except KeyError:
                 logging.warning(f"  - DESCARTADO (Índice Original Aprox.: {idx_descartado}): Telefone ausente (Erro ao buscar nome).")
            except Exception as e_log_desc:
                 logging.warning(f"  - DESCARTADO (Índice Original Aprox.: {idx_descartado}): Telefone ausente (Erro inesperado ao buscar detalhes: {e_log_desc}).")

    # Garante que a coluna de telefone seja string (ajustado para usar .loc)
    if not dados_validos.empty and not pd.api.types.is_string_dtype(dados_validos[COLUNA_TELEFONE]):
        logging.warning(f"Coluna de telefone '{COLUNA_TELEFONE}' não parece ser string. Tentando converter...")
        dados_validos.loc[:, COLUNA_TELEFONE] = dados_validos[COLUNA_TELEFONE].astype(str)

    # Se após remover linhas sem telefone, não sobrar nada válido...
    if dados_validos.empty:
        logging.info("Nenhum registro com número de telefone válido encontrado após filtragem. Encerrando.")
        return

    # --- Agrupamento por Telefone (dos dados válidos) ---
    logging.info("Agrupando registros válidos por número de telefone...")
    try:
        # Ordenação
        colunas_ordenacao = [COLUNA_NOME]
        dados_para_agrupar = dados_validos # << USA OS DADOS FILTRADOS
        if COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in dados_para_agrupar.columns:
            if pd.api.types.is_datetime64_any_dtype(dados_para_agrupar[COLUNA_VENCIMENTO]):
                colunas_ordenacao.append(COLUNA_VENCIMENTO)
            else:
                logging.warning(f"Coluna '{COLUNA_VENCIMENTO}' existe mas não é data/datetime após leitura. Não será usada para ordenação interna do grupo.")
        colunas_ordenacao = [col for col in colunas_ordenacao if col and col in dados_para_agrupar.columns]

        if len(colunas_ordenacao) > 0:
            logging.debug(f"Ordenando dados pré-agrupamento por: {colunas_ordenacao}")
            na_pos = 'last' if COLUNA_VENCIMENTO in colunas_ordenacao else 'first'
            dados_para_agrupar = dados_para_agrupar.sort_values(by=colunas_ordenacao, na_position=na_pos)
        else:
            logging.warning("Nenhuma coluna válida encontrada para ordenação pré-agrupamento.")

        # Agrupa os dados que RESTARAM (com telefone)
        grupos_por_telefone = dados_para_agrupar.groupby(COLUNA_TELEFONE, sort=False)
        total_telefones_unicos = len(grupos_por_telefone)

    except KeyError as e:
        logging.error(f"Erro ao agrupar/ordenar: Coluna '{e}' não encontrada nos dados filtrados.")
        return
    except Exception as e_group:
        logging.exception("Erro inesperado durante o agrupamento/ordenação por telefone.")
        return

    # --- Processamento dos Grupos ---
    enviados_sucesso = 0
    falhas_envio = 0
    telefones_processados = 0
    logging.info(f"\n--- Iniciando processamento CONSOLIDADO para {total_telefones_unicos} telefones únicos ---")

    for telefone, grupo in grupos_por_telefone:
        telefones_processados += 1
        logging.info("-" * 20)
        telefone_str = str(telefone)
        logging.info(f"Processando Telefone {telefones_processados}/{total_telefones_unicos}: {telefone_str}")

        if not telefone_str.startswith("+55") or not (len(telefone_str) == 13 or len(telefone_str) == 14):
            logging.warning(f"AVISO: Telefone '{telefone_str}' (Cliente: {grupo[COLUNA_NOME].iloc[0] if COLUNA_NOME in grupo else 'N/D'}) inválido/fora do padrão E.164. Pulando este grupo.")
            falhas_envio += 1 # Conta como falha para o grupo
            continue
        if grupo.empty:
            logging.warning(f"AVISO: Grupo para telefone {telefone_str} está vazio inesperadamente. Pulando.")
            continue

        nome_cliente = "[Cliente N/D]"
        if COLUNA_NOME in grupo.columns and not grupo[COLUNA_NOME].empty:
             primeiro_nome_valido = grupo[COLUNA_NOME].dropna().iloc[0] if not grupo[COLUNA_NOME].dropna().empty else None
             if primeiro_nome_valido:
                 nome_cliente = primeiro_nome_valido
             else:
                 logging.warning(f"AVISO: Grupo para telefone {telefone_str} não contém nome de cliente válido na coluna '{COLUNA_NOME}'. Usando '{nome_cliente}'.")
        else:
            logging.warning(f"AVISO: Coluna '{COLUNA_NOME}' não encontrada no grupo do telefone {telefone_str}. Usando '{nome_cliente}'.")

        logging.info(f"  - Cliente: {nome_cliente}")
        logging.info(f"  - Nº de parcelas no grupo: {len(grupo)}")

        lista_parcelas_info = []
        valor_total_num = 0.0
        nome_empresa_grupo = EMPRESA_PADRAO
        primeiro_nome_empresa_encontrado = None
        inconsistencia_empresa = False
        primeira_parcela_valida_para_empresa = True

        try:
            for idx, parcela_row in grupo.iterrows():
                lote_code = None
                nome_empresa_parcela = None
                duplicata_raw_original_para_log = parcela_row.get(COLUNA_LOTEAMENTO, "[NÃO FORNECIDO NA PLANILHA]")

                if COLUNA_LOTEAMENTO:
                    duplicata_raw_valor_atual = parcela_row.get(COLUNA_LOTEAMENTO)
                    if isinstance(duplicata_raw_valor_atual, str) and '/' in duplicata_raw_valor_atual:
                        parts = duplicata_raw_valor_atual.split('/')
                        if len(parts) >= 1:
                            lote_code = parts[0].strip()
                            nome_empresa_parcela = MAPEAMENTO_EMPRESA_POR_CODIGO.get(lote_code)

                if COLUNA_LOTEAMENTO and not lote_code:
                    logging.warning(f"Parcela (Índice Planilha: {idx}, Cliente: {nome_cliente}, Tel: {telefone_str}): Código de loteamento não pôde ser extraído da coluna '{COLUNA_LOTEAMENTO}'. Valor encontrado: '{duplicata_raw_original_para_log}'. Nomes de loteamento e empresa podem ser afetados.")

                if nome_empresa_parcela:
                    if primeira_parcela_valida_para_empresa:
                        primeiro_nome_empresa_encontrado = nome_empresa_parcela
                        nome_empresa_grupo = nome_empresa_parcela
                        primeira_parcela_valida_para_empresa = False
                    elif not inconsistencia_empresa and primeiro_nome_empresa_encontrado != nome_empresa_parcela:
                        inconsistencia_empresa = True
                        logging.warning(f"Cliente {nome_cliente} ({telefone_str}) possui parcelas de empresas diferentes ({primeiro_nome_empresa_encontrado} vs {nome_empresa_parcela}). Usando nome de empresa padrão: '{EMPRESA_PADRAO}' para este cliente.")
                        nome_empresa_grupo = EMPRESA_PADRAO

                nome_loteamento_parcela = MAPEAMENTO_LOTEAMENTO.get(lote_code, f'[Cód Loteam. {lote_code}?]') if lote_code else '[Loteamento N/D]'

                lote_raw_da_parcela = parcela_row.get(COLUNA_LOTE) if COLUNA_LOTE else None
                venc_obj_da_parcela = parcela_row.get(COLUNA_VENCIMENTO) if COLUNA_VENCIMENTO else None
                valor_num_da_parcela = parcela_row.get(COLUNA_VALOR) if COLUNA_VALOR else None

                if COLUNA_LOTE and pd.isna(lote_raw_da_parcela):
                    logging.warning(f"Parcela (Índice Planilha: {idx}, Cliente: {nome_cliente}, Tel: {telefone_str}): Informação de Lote (coluna '{COLUNA_LOTE}') ausente ou inválida.")

                if COLUNA_VENCIMENTO and pd.isna(venc_obj_da_parcela):
                    valor_original_venc_planilha = parcela_row.get(COLUNA_VENCIMENTO)
                    logging.warning(f"Parcela (Índice Planilha: {idx}, Cliente: {nome_cliente}, Tel: {telefone_str}): Informação de Vencimento (coluna '{COLUNA_VENCIMENTO}') ausente ou inválida (NaT). Valor lido: '{valor_original_venc_planilha}'.")

                if COLUNA_VALOR and pd.isna(valor_num_da_parcela):
                    valor_original_saldo_planilha = parcela_row.get(COLUNA_VALOR)
                    logging.warning(f"Parcela (Índice Planilha: {idx}, Cliente: {nome_cliente}, Tel: {telefone_str}): Informação de Saldo/Valor (coluna '{COLUNA_VALOR}') ausente ou inválida (NaN). Valor lido: '{valor_original_saldo_planilha}'.")

                lote_info = "[Ref N/D]"
                if COLUNA_LOTE and pd.notna(lote_raw_da_parcela):
                    lote_str_format = str(lote_raw_da_parcela).strip()
                    lote_info = lote_str_format.split('/')[0].strip() if '/' in lote_str_format else lote_str_format

                venc_fmt = "[Data N/D]"
                if pd.notna(venc_obj_da_parcela) and hasattr(venc_obj_da_parcela, 'strftime'):
                    try: venc_fmt = venc_obj_da_parcela.strftime('%d/%m/%Y')
                    except ValueError: pass

                valor_fmt = "[Valor N/D]"
                valor_parcela_num = 0.0
                if pd.notna(valor_num_da_parcela):
                    try:
                        valor_parcela_num = float(valor_num_da_parcela)
                        valor_fmt = formatar_valor_moeda_local(valor_parcela_num)
                    except (ValueError, TypeError):
                        valor_fmt = "[Erro Formatação Valor]"

                lista_parcelas_info.append({
                    'lote': lote_info,
                    'vencimento': venc_fmt,
                    'valor': valor_fmt,
                    'loteamento_nome': nome_loteamento_parcela,
                })
                if pd.notna(valor_num_da_parcela):
                    valor_total_num += valor_parcela_num
        except Exception as e_agg:
            logging.error(f"ERRO CRÍTICO ao processar parcelas para {nome_cliente} ({telefone_str}): {e_agg}")
            logging.exception("Detalhes da exceção ao processar parcelas:")
            falhas_envio += 1 # Conta como falha para o grupo
            continue

        logging.debug(f"Nome da empresa final definido para o grupo de {nome_cliente}: {nome_empresa_grupo}")

        valor_total_fmt = formatar_valor_moeda_local(valor_total_num)
        try:
            mensagem = construtor_mensagem.criar_mensagem_consolidada(
                nome_cliente,
                lista_parcelas_info,
                valor_total_fmt,
                nome_empresa_grupo
            )
            logging.debug(f"Mensagem consolidada construída para {nome_cliente}")
        except Exception as e_msg:
            logging.error(f"ERRO ao construir mensagem consolidada para '{nome_cliente}': {e_msg}.")
            logging.exception("Detalhes da exceção ao construir mensagem:")
            falhas_envio += 1
            continue

        logging.info(f"  - Tentando salvar/enviar mensagem consolidada...")
        sucesso_envio_msg = False
        try:
            sucesso_envio_msg = arquivo_txt_sender.salvar_mensagem_em_txt(telefone_str, nome_cliente, mensagem)
        except Exception as e_send:
            logging.error(f"ERRO inesperado ao tentar salvar/enviar msg para '{nome_cliente}': {e_send}")
            logging.exception("Detalhes da exceção ao salvar/enviar:")

        if sucesso_envio_msg:
             enviados_sucesso += 1
        else:
             falhas_envio += 1
             logging.warning(f"  - Falha ao salvar/enviar msg para {nome_cliente} ({telefone_str}).")

        if total_telefones_unicos > 1 and telefones_processados < total_telefones_unicos:
            if DELAY_ENTRE_MENSAGENS_SEGUNDOS > 0:
                logging.info(f"  - Aguardando {DELAY_ENTRE_MENSAGENS_SEGUNDOS} segundo(s)...")
                time.sleep(DELAY_ENTRE_MENSAGENS_SEGUNDOS)

    # --- Resumo final ---
    logging.info("\n" + "="*40)
    logging.info("--- Processo CONSOLIDADO Concluído ---")
    logging.info(f"Planilha processada: {os.path.basename(caminho_planilha_selecionada)}")
    logging.info(f"Total de registros carregados da planilha: {total_antes_filtro_telefone}")
    if total_descartados_sem_telefone > 0:
        logging.info(f"Total de registros DESCARTADOS por falta de telefone: {total_descartados_sem_telefone}")
    logging.info(f"Total de registros COM telefone (considerados para agrupamento): {total_apos_filtro_telefone}")
    logging.info(f"Total de telefones únicos processados (grupos): {total_telefones_unicos}")
    logging.info(f"Mensagens consolidadas salvas/enviadas com sucesso: {enviados_sucesso}")
    logging.info(f"Grupos de telefone com falha no processamento/envio: {falhas_envio}")
    if enviados_sucesso > 0 and hasattr(arquivo_txt_sender, 'DIRETORIO_SAIDA'):
        try:
            dir_saida = arquivo_txt_sender.DIRETORIO_SAIDA
            logging.info(f"Verifique os arquivos TXT na pasta: '{dir_saida}'")
        except AttributeError:
            logging.warning("AVISO: Não foi possível determinar diretório de saída a partir do módulo.")
    logging.info("="*40)
# ------------------------------

# --- Ponto de entrada ---
if __name__ == "__main__":
    processar_cobrancas()
# ------------------------------