import pandas as pd
import os
import re

pasta = r'i:\RECENTES\DIRETORIA DE PREVIDÊNCIA\COORDENAÇÃO DE CADASTRO E ATENDIMENTO\GERENCIA DE RECADASTRAMENTO E PROVA DE VIDA\CARTAS 2026\04- ABRIL'
pasta_erro = pasta + '\\ERRO CEP'

# Mapeamento de arquivos de erro para arquivos Excel
mapeamento = {
    'Pasta1': '2026 05 07 - prova de vida pendentes ABRIL- copia.xlsx',
    'Pasta2': '2026 05 07 - prova de vida pendentes ABRIL.xlsx',
    'Pasta3': '2026 05 06 - prova de vida pendentes ABRIL.xlsx',
    'Pasta4': '2026 04 07 - prova de vida pendentes ABRIL.xlsx',
    'Pasta5': '2026 04 07 - prova de vida pendentes ABRIL.xlsx'
}

todos_dados = []

# Ler cada arquivo de erro
txt_files = sorted([f for f in os.listdir(pasta_erro) if f.endswith('.txt')])

for txt_file in txt_files:
    # Identificar qual pasta
    pasta_num = re.search(r'Pasta(\d+)', txt_file)
    if pasta_num:
        num_pasta = pasta_num.group(1)
        nome_excel = mapeamento.get(f'Pasta{num_pasta}')
        
        if nome_excel:
            caminho_excel = os.path.join(pasta, nome_excel)
            if os.path.exists(caminho_excel):
                df = pd.read_excel(caminho_excel, dtype=str)
                
                # Ler erros
                caminho_erro = os.path.join(pasta_erro, txt_file)
                with open(caminho_erro, 'r', encoding='utf-8', errors='ignore') as f:
                    linhas = f.readlines()
                
                for linha in linhas[1:]:
                    if 'Linha:' in linha:
                        match_linha = re.search(r'Linha: (\d+)', linha)
                        erro = ''
                        if 'não encontrado' in linha.lower():
                            erro = 'CEP não encontrado'
                        elif 'inválido' in linha.lower():
                            erro = 'CEP inválido'
                        
                        if match_linha:
                            num_linha = int(match_linha.group(1))
                            idx = num_linha - 2  # Ajustar índice
                            
                            if idx >= 0 and idx < len(df):
                                row = df.iloc[idx].to_dict()
                                row['arquivo_origem'] = nome_excel
                                row['linha_erro'] = num_linha
                                row['tipo_erro'] = erro
                                row['arquivo_erro'] = txt_file
                                todos_dados.append(row)

df_resultado = pd.DataFrame(todos_dados)
print(f'Total de registros com erro: {len(df_resultado)}')
print(df_resultado[['NOME', 'CEP', 'tipo_erro', 'arquivo_origem']].head(10))

df_resultado.to_excel('erros_ceps_consolidado.xlsx', index=False)
print('\nSalvo: erros_ceps_consolidado.xlsx')