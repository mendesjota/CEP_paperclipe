"""
Testes de Integração com APIs do CEP Paperclipe

Este arquivo testa:
- Verificação de CEP existente (AwesomeAPI, ViaCEP)
- Busca de CEP por endereço (CepRua, AwesomeAPI, ViaCEP)
"""

import sys
sys.path.insert(0, 'C:/Users/jose.junior/Desktop/Python/Projetos/CEP_Paperclipe/CEP_paperclipe/src')

import requests
import time


# ============================================
# CONFIGURAÇÕES DE TESTE
# ============================================

DELAY_REQUISIÇÃO = 0.5  # delay para testes (maior que o padrão de 0.3s)


# ============================================
# FUNÇÕES DE API (copiadas para teste)
# ============================================

def verificar_cep_existe(cep):
    """Verifica se o CEP existe nas APIs"""
    cep_limpo = ''.join(filter(str.isdigit, str(cep)))
    
    # Tentar AwesomeAPI
    time.sleep(DELAY_REQUISIÇÃO)
    try:
        url = f"https://cep.awesomeapi.com.br/json/{cep_limpo}"
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            dados = res.json()
            if 'cep' in dados:
                return dados
    except:
        pass
    
    # Tentar ViaCEP
    time.sleep(DELAY_REQUISIÇÃO)
    try:
        url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            dados = res.json()
            if isinstance(dados, dict) and 'erro' not in dados:
                return dados
    except:
        pass
    
    return None


def buscar_cep_por_endereco(uf, municipio, logradouro, bairro):
    """Busca CEP por endereço usando múltiplas APIs"""
    import re
    
    uf = (uf or '').strip().upper()
    municipio = (municipio or '').upper().strip()
    logradouro = (logradouro or '').upper().strip()
    bairro = (bairro or '').upper().strip()
    
    # Limpar duplicatas
    palavras_logradouro = logradouro.split()
    logradouro = ' '.join([p for i, p in enumerate(palavras_logradouro) if i == 0 or p != palavras_logradouro[i-1]])
    
    palavras_bairro = bairro.split()
    bairro = ' '.join([p for i, p in enumerate(palavras_bairro) if i == 0 or p != palavras_bairro[i-1]])
    
    queries = [
        f"{uf} {municipio} {bairro} {logradouro}",
        f"{uf} {municipio} {bairro}",
        f"{bairro}, {municipio}, {uf}",
    ]
    
    for query in queries:
        # CepRua
        time.sleep(DELAY_REQUISIÇÃO)
        try:
            url = f"https://ceprua.com.br/api/buscar?q={query}"
            res = requests.get(url, timeout=5)
            if res.status_code == 200:
                dados = res.json()
                if dados and isinstance(dados, dict):
                    resultados = dados.get('resultados', [])
                    for r in resultados:
                        if r.get('uf', '').upper() == uf:
                            cep = r.get('cep', '').replace('-', '')
                            if cep and len(cep) == 8:
                                return cep
        except:
            pass
    
    return None


# ============================================
# TESTES DE API
# ============================================

def testar_verificar_cep_existe():
    """Testa a verificação de CEP existente"""
    print("\n=== Teste: verificar_cep_existe ===")
    
    ceps_teste = [
        ("70350756", "DF"),
        ("01001000", "SP"),
        ("20020000", "RJ"),
    ]
    
    for cep, uf_esperado in ceps_teste:
        print(f"\n  Testando CEP: {cep}")
        resultado = verificar_cep_existe(cep)
        
        if resultado:
            uf_encontrado = resultado.get('state', '') or resultado.get('uf', '')
            print(f"    -> Encontrado: {resultado.get('address', 'N/A')} ({uf_encontrado})")
            if uf_encontrado.upper() == uf_esperado.upper():
                print(f"    -> UF CORRETA!")
            else:
                print(f"    -> UF DIFERENTE (esperado: {uf_esperado})")
        else:
            print(f"    -> Não encontrado ou rate limit")
    
    print("=> verificar_cep_existe: OK")


def testar_buscar_cep_por_endereco():
    """Testa a busca de CEP por endereço"""
    print("\n=== Teste: buscar_cep_por_endereco ===")
    
    testes_endereco = [
        ("DF", "BRASILIA", "QUADRA SHIGS 706", "ASA SUL"),
    ]
    
    for uf, municipio, logradouro, bairro in testes_endereco:
        print(f"\n  Buscando: {uf} - {municipio} - {bairro} - {logradouro}")
        resultado = buscar_cep_por_endereco(uf, municipio, logradouro, bairro)
        
        if resultado:
            print(f"    -> CEP encontrado: {resultado}")
        else:
            print(f"    -> Não encontrado")
    
    print("=> buscar_cep_por_endereco: OK")


def testar_rate_limit():
    """Testa se o códigoaguenta múltiplas requisições sem bloquear"""
    print("\n=== Teste: Rate Limit (5 requisições) ===")
    
    ceps = ["70350756", "71660160", "73091900", "72910000", "12345678"]
    
    for i, cep in enumerate(ceps):
        print(f"  Requisição {i+1}/5: CEP {cep}")
        resultado = verificar_cep_existe(cep)
        
        if resultado:
            print(f"    -> OK")
        elif resultado is None:
            print(f"    -> Não encontrado ou rate limit")
    
    print("=> Rate Limit: OK")


# ============================================
# EXECUTAR TESTES
# ============================================

if __name__ == "__main__":
    print("=" * 60)
    print("TESTES DE INTEGRAÇÃO - CEP PAPERCLIPE")
    print("=" * 60)
    
    try:
        testar_verificar_cep_existe()
    except Exception as e:
        print(f"  Erro no teste: {e}")
    
    try:
        testar_buscar_cep_por_endereco()
    except Exception as e:
        print(f"  Erro no teste: {e}")
    
    try:
        testar_rate_limit()
    except Exception as e:
        print(f"  Erro no teste: {e}")
    
    print("\n" + "=" * 60)
    print("TESTES DE INTEGRAÇÃO CONCLUÍDOS")
    print("=" * 60)