"""
Dados.py - Consolidação de erros de CEP
"""
import pandas as pd
import os
import re

def consolidar_erros_cep():
    """Consolida erros de CEP de múltiplos arquivos"""
    
    pasta = r'i:\RECENTES\DIRETORIA DE PREVIDÊNCIA\COORDENAÇÃO DE CADASTRO E ATENDIMENTO\GERENCIA DE RECADASTRAMENTO E PROVA DE VIDA\CARTAS 2026\04- ABRIL'
    pasta_erro = pasta + '\\ERRO CEP'
    
    # Mapeamento de arquivos de erro para arquivos CSV originals
    mapeamento = {
        'Pasta1': 'Pasta1.csv',
        'Pasta2': 'Pasta2.csv',
        'Pasta3': 'Pasta3.csv',
        'Pasta4': 'Pasta4.csv',
        'Pasta5': 'Pasta5.csv'
    }
    
    todos_dados = []
    
    # Ler cada arquivo de erro
    txt_files = sorted([f for f in os.listdir(pasta_erro) if f.endswith('.txt')])
    
    for txt_file in txt_files:
        # Identificar qual pasta
        pasta_num = re.search(r'Pasta(\d+)', txt_file)
        if pasta_num:
            num_pasta = pasta_num.group(1)
            nome_csv = mapeamento.get(f'Pasta{num_pasta}')
            
            if nome_csv:
                caminho_csv = os.path.join(pasta, nome_csv)
                if os.path.exists(caminho_csv):
                    colunas = ['SR_A', 'NOME', 'c1', 'c2', 'c3', 'CPF', 'CEP', 'c4', 'ENDERECO', 'BAIRRO', 'MUNICIPIO', 'UF', 'N', 'COMPLEMENTO', 'c5']
                    df = pd.read_csv(caminho_csv, dtype=str, encoding='utf-8', on_bad_lines='skip', sep=';', header=None, names=colunas, keep_default_na=False)
                    
                    # Ler erros
                    caminho_erro = os.path.join(pasta_erro, txt_file)
                    with open(caminho_erro, 'r', encoding='utf-8', errors='replace') as f:
                        linhas = f.readlines()
                    
                    for linha in linhas[1:]:
                        if 'Linha:' in linha:
                            match_linha = re.search(r'Linha: (\d+)', linha)
                            erro = ''
                            linha_lower = linha.lower()
                            if 'não encontrado' in linha_lower:
                                erro = 'CEP não encontrado'
                            elif 'inválido' in linha_lower:
                                erro = 'CEP inválido'
                            elif 'obrigatório' in linha_lower:
                                erro = 'Campo obrigatório'
                            
                            if match_linha:
                                num_linha = int(match_linha.group(1))
                                idx = num_linha - 1  # Sem cabeçalho, 0-based
                                
                                if idx >= 0 and idx < len(df):
                                    row = df.iloc[idx].to_dict()
                                    row['arquivo_origem'] = nome_csv
                                    row['linha_erro'] = num_linha
                                    row['tipo_erro'] = erro
                                    row['arquivo_erro'] = txt_file
                                    todos_dados.append(row)
    
    df_resultado = pd.DataFrame(todos_dados)
    print(f'Total de registros com erro: {len(df_resultado)}')
    
    # Limpar caracteres inválidos
    for col in df_resultado.columns:
        df_resultado[col] = df_resultado[col].astype(str).str.replace(r'[\x00-\x1F\x7F-\x9F]', '', regex=True)
    
    # Formatar CPF e CEP com zeros à esquerda
    if 'CPF' in df_resultado.columns:
        df_resultado['CPF'] = df_resultado['CPF'].apply(lambda x: str(x).zfill(11) if str(x).strip() not in ['nan', ''] else x)
    if 'CEP' in df_resultado.columns:
        df_resultado['CEP'] = df_resultado['CEP'].apply(lambda x: str(x).zfill(8) if str(x).strip() not in ['nan', ''] else x)
    
    # Salvar
    df_resultado.to_excel('erros_ceps_consolidado.xlsx', index=False)
    print('Salvo: erros_ceps_consolidado.xlsx')
    
    return df_resultado

def buscar_cep_api(endereco, bairro, municipio, uf):
    """Busca CEP usando a API do CepRua"""
    import requests
    
    # Tratar valores NaN/None
    endereco = str(endereco) if pd.notna(endereco) else ''
    bairro = str(bairro) if pd.notna(bairro) else ''
    municipio = str(municipio) if pd.notna(municipio) else ''
    uf = str(uf) if pd.notna(uf) else ''
    
    if not uf or not municipio or uf == 'nan' or municipio == 'nan':
        return None
    
    uf = uf.strip().upper()
    municipio = municipio.strip().upper()
    endereco = endereco.strip().upper() if endereco else ''
    bairro = bairro.strip().upper() if bairro else ''
    
    buscas = [
        f"{endereco} {bairro} {municipio} {uf}",
        f"{endereco} {municipio} {uf}",
        f"{bairro} {municipio} {uf}",
        f"{municipio} {uf}"
    ]
    
    for query in buscas:
        if not query or len(query) < 5:
            continue
        try:
            url = f"https://ceprua.com.br/api/buscar?q={query}"
            res = requests.get(url, timeout=5)
            if res.status_code == 200:
                dados = res.json()
                if 'resultados' in dados and dados['resultados']:
                    for item in dados['resultados']:
                        cep = item.get('cep', '').replace('-', '')
                        item_uf = item.get('uf', '').upper().strip()
                        if len(cep) == 8 and item_uf == uf:
                            return cep
                if 'redirect' in dados:
                    cep = dados['redirect'].get('cep', '')
                    if cep:
                        return cep.replace('-', '')
        except:
            pass
    return None

def adicionar_cep_encontrado():
    """Adiciona coluna com CEP encontrado via API"""
    df = pd.read_excel('erros_ceps_consolidado.xlsx')
    
    print(f"Processando {len(df)} registros...")
    
    # Adicionar coluna vazia
    df['CEP_encontrado'] = ''
    
    for idx, row in df.iterrows():
        endereco = row.get('ENDERECO', '')
        bairro = row.get('BAIRRO', '')
        municipio = row.get('MUNICIPIO', '')
        uf = row.get('UF', '')
        
        # Buscar CEP via API
        novo_cep = buscar_cep_api(endereco, bairro, municipio, uf)
        if novo_cep:
            df.at[idx, 'CEP_encontrado'] = novo_cep
            nome = str(row.get('NOME', ''))[:30]
            print(f"{idx+2}: {nome}... -> {novo_cep}")
        else:
            df.at[idx, 'CEP_encontrado'] = ''
            endereco = row.get('ENDERECO', '')
            bairro = row.get('BAIRRO', '')
            municipio = row.get('MUNICIPIO', '')
            uf = row.get('UF', '')
            
            novo_cep = buscar_cep_api(endereco, bairro, municipio, uf)
            if novo_cep:
                df.at[idx, 'CEP_encontrado'] = novo_cep
                print(f"Linha {idx+2}: {row.get('NOME', '')[:30]}... -> {novo_cep}")
            else:
                df.at[idx, 'CEP_encontrado'] = ''
    
    df.to_excel('erros_ceps_consolidado.xlsx', index=False)
    print("Salvo com sucesso!")

if __name__ == '__main__':
    consolidar_erros_cep()
    adicionar_cep_encontrado()