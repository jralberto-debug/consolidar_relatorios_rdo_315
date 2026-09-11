import os
import pandas as pd
import openpyxl
from pathlib import Path
import logging
from datetime import datetime
import sys

def configurar_logging():
    """Configura o sistema de logging para acompanhar o processo"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('extrator_rdo.log', encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

def criar_pasta_destino(caminho_destino):
    """Cria a pasta de destino se ela não existir"""
    try:
        pasta = Path(caminho_destino)
        pasta.mkdir(parents=True, exist_ok=True)
        logging.info(f"Pasta de destino verificada/criada: {caminho_destino}")
        return True
    except Exception as e:
        logging.error(f"Erro ao criar pasta de destino: {str(e)}")
        return False

def encontrar_aba_efetivo(workbook):
    """Encontra a aba que contém 'EFETIVO' no nome"""
    for sheet_name in workbook.sheetnames:
        if 'EFETIVO' in sheet_name.upper():
            logging.info(f"Aba encontrada: {sheet_name}")
            return workbook[sheet_name]
    
    # Se não encontrar, tenta outras variações
    for sheet_name in workbook.sheetnames:
        name_upper = sheet_name.upper()
        if any(palavra in name_upper for palavra in ['EFETIV', 'PESSOAL', 'EQUIPE']):
            logging.info(f"Aba alternativa encontrada: {sheet_name}")
            return workbook[sheet_name]
    
    logging.warning("Nenhuma aba com 'EFETIVO' encontrada")
    return None

def encontrar_linha_referencia(worksheet):
    """Encontra a linha que contém o texto de referência"""
    textos_referencia = [
        "(CHUVAS / ALERTA VERMELHO / TREINAMENTOS DE SMS / INSPEÇÕES / AUDITORIAS, ETC..",
        "CHUVAS / ALERTA VERMELHO / TREINAMENTOS",
        "PARALISAÇÕES",
        "PARALIZAÇÃO",
        "INTERRUPÇÕES"
    ]
    
    max_row = min(worksheet.max_row, 200)  # Limita busca às primeiras 200 linhas
    
    for row_num in range(1, max_row + 1):
        for col_num in range(1, min(worksheet.max_column + 1, 20)):  # Limita às primeiras 20 colunas
            try:
                cell = worksheet.cell(row=row_num, column=col_num)
                if cell.value and isinstance(cell.value, str):
                    cell_text = cell.value.strip().upper()
                    for texto_ref in textos_referencia:
                        if texto_ref.upper() in cell_text:
                            logging.info(f"Linha de referência encontrada na linha {row_num}: '{texto_ref}'")
                            return row_num
            except Exception:
                continue
    
    logging.warning("Linha de referência não encontrada")
    return None

def encontrar_cabecalho(worksheet, linha_inicio):
    """Encontra a linha de cabeçalho com as colunas específicas"""
    colunas_esperadas = {
        'DATA': ['DATA', 'DT', 'DATE'],
        'DIA': ['DIA', 'DAY', 'SEMANA'],
        'OCORRÊNCIA': ['OCORRÊNCIA', 'OCORRENCIA', 'MOTIVO', 'DESCRIÇÃO', 'DESCRICAO', 'CAUSA', 'EVENTO', 'TIPO'],
        'INÍCIO': ['INÍCIO', 'INICIO', 'INI', 'HORA INICIO', 'H.INI', 'HORA INI', 'START', 'COMEÇO'],
        'TERMINO': ['TERMINO', 'TÉRMINO', 'FIM', 'HORA FIM', 'H.FIM', 'FINAL', 'END'],
        'DURAÇÃO': ['DURAÇÃO', 'DURACAO', 'TEMPO', 'HORAS', 'HRS', 'H', 'DURATION']
    }
    
    # Procura nas próximas 20 linhas após a linha de referência
    inicio_busca = linha_inicio + 1 if linha_inicio else 1
    fim_busca = min(inicio_busca + 20, worksheet.max_row)
    
    logging.info(f"Procurando cabeçalho entre as linhas {inicio_busca} e {fim_busca}")
    
    for row_num in range(inicio_busca, fim_busca + 1):
        posicoes_colunas = {}
        colunas_encontradas = 0
        linha_debug = []
        
        # Verifica todas as colunas da linha
        for col_num in range(1, min(worksheet.max_column + 1, 30)):
            try:
                cell = worksheet.cell(row=row_num, column=col_num)
                cell_value = cell.value
                
                if cell_value:
                    linha_debug.append(str(cell_value))
                    
                if cell_value and isinstance(cell_value, str):
                    cell_text = cell_value.strip().upper()
                    
                    # Verifica se contém alguma das palavras-chave
                    for col_padrao, variantes in colunas_esperadas.items():
                        if col_padrao not in posicoes_colunas:  # Evita duplicatas
                            for variante in variantes:
                                if variante in cell_text:
                                    posicoes_colunas[col_padrao] = col_num - 1
                                    colunas_encontradas += 1
                                    logging.info(f"  Coluna '{col_padrao}' encontrada na posição {col_num}: '{cell_text}'")
                                    break
            except Exception as e:
                continue
        
        # Debug: mostra o conteúdo da linha
        if linha_debug:
            logging.info(f"Linha {row_num}: {' | '.join(linha_debug[:10])}")  # Primeiras 10 colunas
        
        # Se encontrou pelo menos 2 colunas (mais flexível)
        if colunas_encontradas >= 2:
            logging.info(f"Cabeçalho encontrado na linha {row_num} com {colunas_encontradas} colunas")
            logging.info(f"Colunas mapeadas: {list(posicoes_colunas.keys())}")
            return row_num, posicoes_colunas
    
    logging.warning("Cabeçalho não encontrado - Tentando busca mais ampla...")
    
    # Segunda tentativa: busca mais ampla
    for row_num in range(1, min(worksheet.max_row + 1, 100)):
        linha_texto = []
        for col_num in range(1, min(worksheet.max_column + 1, 15)):
            try:
                cell = worksheet.cell(row=row_num, column=col_num)
                if cell.value:
                    linha_texto.append(str(cell.value).strip().upper())
            except:
                continue
        
        linha_completa = ' '.join(linha_texto)
        
        # Verifica se a linha contém pelo menos 2 palavras-chave
        palavras_encontradas = 0
        for col_list in colunas_esperadas.values():
            for palavra in col_list:
                if palavra in linha_completa:
                    palavras_encontradas += 1
                    break
        
        if palavras_encontradas >= 2:
            logging.info(f"Possível cabeçalho encontrado na linha {row_num} (busca ampla)")
            logging.info(f"Conteúdo: {linha_completa}")
            
            # Mapeia as posições novamente
            posicoes_colunas = {}
            for col_num in range(1, min(worksheet.max_column + 1, 15)):
                try:
                    cell = worksheet.cell(row=row_num, column=col_num)
                    if cell.value and isinstance(cell.value, str):
                        cell_text = cell.value.strip().upper()
                        for col_padrao, variantes in colunas_esperadas.items():
                            if col_padrao not in posicoes_colunas:
                                for variante in variantes:
                                    if variante in cell_text:
                                        posicoes_colunas[col_padrao] = col_num - 1
                                        break
                except:
                    continue
            
            if posicoes_colunas:
                return row_num, posicoes_colunas
    
    logging.warning("Cabeçalho não encontrado mesmo com busca ampla")
    return None, None

def extrair_dados_planilha(caminho_arquivo):
    """Extrai dados de uma planilha específica"""
    dados_extraidos = []
    
    try:
        logging.info(f"Processando arquivo: {os.path.basename(caminho_arquivo)}")
        
        # Abre o arquivo Excel
        workbook = openpyxl.load_workbook(caminho_arquivo, data_only=True)
        
        # Encontra a aba com "EFETIVO"
        worksheet = encontrar_aba_efetivo(workbook)
        if not worksheet:
            logging.warning(f"Aba 'EFETIVO' não encontrada em {os.path.basename(caminho_arquivo)}")
            return dados_extraidos
        
        # Encontra a linha de referência
        linha_referencia = encontrar_linha_referencia(worksheet)
        
        # Encontra o cabeçalho
        linha_cabecalho, posicoes_colunas = encontrar_cabecalho(worksheet, linha_referencia)
        if not linha_cabecalho or not posicoes_colunas:
            logging.warning(f"Cabeçalho não encontrado em {os.path.basename(caminho_arquivo)}")
            return dados_extraidos
        
        # Extrai os dados
        linha_atual = linha_cabecalho + 1
        linhas_vazias_consecutivas = 0
        max_linhas_vazias = 3
        
        while linha_atual <= worksheet.max_row and linhas_vazias_consecutivas < max_linhas_vazias:
            linha_tem_dados = False
            dados_linha = {'ARQUIVO_ORIGEM': os.path.basename(caminho_arquivo)}
            
            # Extrai dados de cada coluna mapeada
            for coluna_padrao, col_index in posicoes_colunas.items():
                try:
                    cell = worksheet.cell(row=linha_atual, column=col_index + 1)  # Ajusta para índice 1
                    valor = cell.value
                    
                    if valor is not None:
                        # Limpa e formata o valor
                        if isinstance(valor, str):
                            valor = valor.strip()
                            if valor:  # Se não é string vazia após strip
                                linha_tem_dados = True
                        else:
                            linha_tem_dados = True
                    
                    dados_linha[coluna_padrao] = valor
                    
                except Exception as e:
                    dados_linha[coluna_padrao] = None
            
            # Se a linha tem dados válidos, adiciona à lista
            if linha_tem_dados:
                dados_extraidos.append(dados_linha)
                linhas_vazias_consecutivas = 0
            else:
                linhas_vazias_consecutivas += 1
            
            linha_atual += 1
        
        logging.info(f"Extraídas {len(dados_extraidos)} linhas de {os.path.basename(caminho_arquivo)}")
        
    except Exception as e:
        logging.error(f"Erro ao processar {os.path.basename(caminho_arquivo)}: {str(e)}")
    
    return dados_extraidos

def processar_todos_arquivos():
    """Função principal que processa todos os arquivos RDO"""
    print("=" * 60)
    print("         EXTRATOR DE DADOS RDO - REPAR")
    print("=" * 60)
    
    # Configurações
    pasta_origem = r"C:\Users\Engevale\Documents\OneDrive\315.[TANQUES][ENGEVALE][REPAR][ONE]\315\RDO"
    pasta_destino = r"C:\Users\Engevale\Documents\OneDrive\315.[TANQUES][ENGEVALE][REPAR][ONE]\315\RDO\etapa4_paralizado"
    arquivo_destino = "paralisacoes_extraidas.xlsx"
    
    # Configura logging
    configurar_logging()
    logging.info("Iniciando extração de dados RDO")
    
    # Verifica se a pasta de origem existe
    if not os.path.exists(pasta_origem):
        print(f"❌ ERRO: Pasta de origem não encontrada: {pasta_origem}")
        logging.error(f"Pasta de origem não encontrada: {pasta_origem}")
        return False
    
    # Cria pasta de destino
    if not criar_pasta_destino(pasta_destino):
        print("❌ ERRO: Não foi possível criar a pasta de destino")
        return False
    
    # Lista todos os arquivos que começam com "RDO"
    pasta_origem_path = Path(pasta_origem)
    arquivos_rdo = list(pasta_origem_path.glob("RDO*.xlsx"))
    
    if not arquivos_rdo:
        print(f"❌ ERRO: Nenhum arquivo RDO encontrado em: {pasta_origem}")
        logging.error("Nenhum arquivo RDO encontrado na pasta de origem")
        return False
    
    print(f"📁 Encontrados {len(arquivos_rdo)} arquivos RDO")
    logging.info(f"Encontrados {len(arquivos_rdo)} arquivos RDO")
    
    # Processa cada arquivo
    todos_dados = []
    arquivos_processados = 0
    
    for i, arquivo in enumerate(arquivos_rdo, 1):
        print(f"📄 Processando ({i}/{len(arquivos_rdo)}): {arquivo.name}")
        dados = extrair_dados_planilha(str(arquivo))
        
        if dados:
            todos_dados.extend(dados)
            arquivos_processados += 1
            print(f"   ✅ {len(dados)} linhas extraídas")
        else:
            print("   ⚠️  Nenhum dado extraído")
    
    if not todos_dados:
        print("❌ ERRO: Nenhum dado foi extraído de nenhum arquivo")
        logging.error("Nenhum dado foi extraído dos arquivos")
        return False
    
    try:
        # Cria DataFrame
        df = pd.DataFrame(todos_dados)
        
        # Organiza as colunas na ordem desejada
        colunas_ordenadas = ['ARQUIVO_ORIGEM', 'DATA', 'DIA', 'OCORRÊNCIA', 'INÍCIO', 'TERMINO', 'DURAÇÃO']
        
        # Adiciona colunas faltantes se necessário
        for col in colunas_ordenadas:
            if col not in df.columns:
                df[col] = None
        
        # Reordena as colunas
        df = df[colunas_ordenadas]
        
        # Salva o arquivo final
        caminho_completo = os.path.join(pasta_destino, arquivo_destino)
        df.to_excel(caminho_completo, index=False, engine='openpyxl')
        
        # Relatório final
        print("\n" + "=" * 60)
        print("✅ EXTRAÇÃO CONCLUÍDA COM SUCESSO!")
        print("=" * 60)
        print(f"📁 Arquivos processados: {arquivos_processados}/{len(arquivos_rdo)}")
        print(f"📊 Total de linhas extraídas: {len(todos_dados)}")
        print(f"💾 Arquivo salvo em: {caminho_completo}")
        print(f"📋 Log detalhado salvo em: extrator_rdo.log")
        
        logging.info(f"Processo concluído com sucesso!")
        logging.info(f"Arquivos processados: {arquivos_processados}/{len(arquivos_rdo)}")
        logging.info(f"Total de linhas extraídas: {len(todos_dados)}")
        logging.info(f"Arquivo salvo em: {caminho_completo}")
        
        return True
        
    except Exception as e:
        print(f"❌ ERRO ao salvar arquivo final: {str(e)}")
        logging.error(f"Erro ao salvar arquivo final: {str(e)}")
        return False

if __name__ == "__main__":
    try:
        sucesso = processar_todos_arquivos()
        
        if not sucesso:
            print("\n⚠️  Processo finalizado com erros. Verifique o log para mais detalhes.")
        
        input("\n🔄 Pressione ENTER para fechar...")
        
    except KeyboardInterrupt:
        print("\n❌ Processo interrompido pelo usuário")
        logging.info("Processo interrompido pelo usuário")
        
    except Exception as e:
        print(f"\n❌ ERRO GERAL: {str(e)}")
        logging.error(f"Erro geral: {str(e)}")
        input("\n🔄 Pressione ENTER para fechar...")