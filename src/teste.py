"""
Teste.py - Ler planilha, validar CEPs na API e retornar resultado

Este script:
1. Lê a planilha Excel
2. Para cada CEP, faz requisição na API
3. Adiciona coluna com resultado da API
"""
import pandas as pd
import requests
import time
import re

# ============================================
# CONFIGURAÇÃO
# ============================================

ARQUIVO_EXCEL = r'c:\Users\jose.junior\Desktop\Python\Projetos\CEP_Paperclipe\CEP_paperclipe\Exemplos de planilha\2026 04 07 - prova de vida pendentes MARÇO.xlsx'

# ============================================
# FUNÇÕES
# ============================================

def limpar_cep(cep):
    """Remove caracteres não numéricos do CEP"""
    cep_str = str(cep).strip()
    return re.sub(r'\D', '', cep_str)

def verificar_cep(cep):
    """Verifica se o CEP existe na API"""
    cep_limpo = limpar_cep(cep)
    
    if len(cep_limpo) != 8:
        return {"valido": False, "erro": "Formato inválido", "cep_api": ""}
    
    # Tentar AwesomeAPI
    try:
        url = f"https://cep.awesomeapi.com.br/json/{cep_limpo}"
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            dados = res.json()
            if 'cep' in dados:
                return {
                    "valido": True,
                    "erro": "",
                    "cep_api": dados.get('cep', ''),
                    "endereco_api": dados.get('address', ''),
                    "bairro_api": dados.get('district', ''),
                    "cidade_api": dados.get('city', ''),
                    "uf_api": dados.get('state', '')
                }
    except Exception as e:
        pass
    
    # Tentar ViaCEP
    try:
        url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            dados = res.json()
            if isinstance(dados, dict) and 'erro' not in dados:
                return {
                    "valido": True,
                    "erro": "",
                    "cep_api": dados.get('cep', '').replace('-', ''),
                    "endereco_api": dados.get('logradouro', ''),
                    "bairro_api": dados.get('bairro', ''),
                    "cidade_api": dados.get('localidade', ''),
                    "uf_api": dados.get('uf', '')
                }
    except Exception as e:
        pass
    
    return {"valido": False, "erro": "CEP não encontrado", "cep_api": ""}


# ============================================
# MAIN
# ============================================

print("=" * 60)
print("CEP Paperclipe - Teste de Validação")
print("=" * 60)

# 1. Ler planilha
print("\n1. Lendo planilha...")
df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
print(f"   Total de registros: {len(df)}")
print(f"   Colunas: {list(df.columns)}")

# 2. Encontrar coluna CEP
colunas_cep = [c for c in df.columns if 'CEP' in c.upper()]
print(f"\n2. Colunas com 'CEP': {colunas_cep}")

if not colunas_cep:
    print("   ERRO: Nao encontrou coluna de CEP!")
    exit()

col_cep = colunas_cep[0]
print(f"   OK - Usando coluna: {col_cep}")

# 3. Mostrar exemplos de CEP
print(f"\n3. Primeiros CEPs da planilha:")
for i, cep in enumerate(df[col_cep].head(5)):
    print(f"   {i+1}. {cep}")

# 4. Processar CEPs
print(f"\n4. Processando CEPs na API...")
print("   (Esto puede demorar varios minutos...)")

resultados = []
total = len(df)

for idx, row in df.iterrows():
    cep = str(row.get(col_cep, ''))
    print(cep)
    # Verificar CEP na API
    resultado = verificar_cep(cep)
    resultados.append(resultado)
    print(resultado)
    
    # Mostrar progresso a cada 50 registros
    if (idx + 1) % 50 == 0:
        print(f"   Processado: {idx + 1}/{total}")
    
    # Delay para não sobrecarregar a API
    time.sleep(0.1)

# 5. Adicionar colunas de resultado
print("\n5. Adicionando colunas de resultado...")
df['CEP_VALIDO'] = [r['valido'] for r in resultados]
df['CEP_ERRO'] = [r['erro'] for r in resultados]
df['CEP_API'] = [r['cep_api'] for r in resultados]
df['ENDERECO_API'] = [r.get('endereco_api', '') for r in resultados]
df['BAIRRO_API'] = [r.get('bairro_api', '') for r in resultados]
df['CIDADE_API'] = [r.get('cidade_api', '') for r in resultados]
df['UF_API'] = [r.get('uf_api', '') for r in resultados]

# 6. Salvar resultado
print("\n6. Salvando resultado...")
arquivo_saida = ARQUIVO_EXCEL.replace('.xlsx', '_validado.xlsx')
df.to_excel(arquivo_saida, index=False)
print(f"   Salvo em: {arquivo_saida}")

# 7. Estatísticas
validos = sum(1 for r in resultados if r['valido'])
invalidos = total - validos

print("\n" + "=" * 60)
print("RESULTADO")
print("=" * 60)
print(f"Total de registros: {total}")
print(f"CEPs válidos: {validos}")
print(f"CEPs inválidos: {invalidos}")

# Mostrar alguns exemplos de inválidos
print("\nExemplos de CEPs inválidos:")
invalidos_lista = [(i, r['erro']) for i, r in enumerate(resultados) if not r['valido']]
for i, erro in invalidos_lista[:10]:
    cep_original = df.iloc[i][col_cep]
    print(f"   Linha {i+2}: {cep_original} - {erro}")

print("\nPROCESSAMENTO CONCLUIDO!")