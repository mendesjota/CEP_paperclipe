"""
CEP Paperclipe - Script de Validação e Correção de CEPs

Este script processa planilhas de prova de vida, valida CEPs via API,
busca correções para CEPs inválidos usando o endereço, e exporta
os resultados em 4 planilhas distintas.

Fluxo:
1. Upload da planilha Excel (.xlsx)
2. Validar CEPs (formato 8 dígitos + existência na API)
3. Buscar CEPs inválidos por endereço
4. Gerar 4 planilhas de saída

Saída:
- 1_ceps_validos.xlsx (CEPs originais corretos)
- 2_ceps_corrigidos.xlsx (CEPs encontrados via endereço)
- 3_ceps_nao_encontrados.xlsx (CEPs não encontrados)
- 4_base_consolidada.xlsx (base inteira com novo CEP e status)
"""

import io
import pandas as pd
import streamlit as st
import requests
import re
import time


# ============================================
# CONFIGURAÇÕES
# ============================================

# Colunas esperadas na planilha de entrada
CABEÇALHO_ESPERADO = [
    'CPF', 'MATRICULA', 'NOME', 'DATA NASCIMENTO',
    'ENDERECO', 'COMPLEMENTO DO ENDERECO', 'BAIRRO',
    'MUNICIPIO', 'UF', 'CEP'
]

# Limites técnicos
LOTE_CORREIOS = 300
DELAY_REQUISIÇÃO = 0.1  # segundos - reduzido para velocidade

# Abreviações normalizadas
ABREVIACOES = {
    r'\bR\b': 'RUA',
    r'\bAV\b': 'AVENIDA',
    r'\bAVENIDA\b': 'AVENIDA',
    r'\bQD\b': 'QUADRA',
    r'\bQN\b': 'QUADRA',
    r'\bCS\b': 'CASA',
    r'\bAP\b': 'APARTAMENTO',
    r'\bBL\b': 'BLOCO',
    r'\bLT\b': 'LOTE',
    r'\bCONJ\b': 'CONJUNTO',
    r'\bCJ\b': 'CONJUNTO',
}


# ============================================
# CONFIGURAÇÃO DA PÁGINA STREAMLIT
# ============================================

st.set_page_config(page_title="CEP Paperclipe", page_icon="📮")
st.title("📮 CEP Paperclipe")
st.markdown("### Validar, Corrigir CEPs e Gerar Lista para Envio")


# ============================================
# SESSION STATE
# ============================================

if 'processado' not in st.session_state:
    st.session_state.processado = False
if 'resultados' not in st.session_state:
    st.session_state.resultados = None
if 'df_original' not in st.session_state:
    st.session_state.df_original = None
if 'status_count' not in st.session_state:
    st.session_state.status_count = None


# ============================================
# FUNÇÕES AUXILIARES
# ============================================

def normalizar_texto(texto):
    """
    Normaliza texto: uppercase, remove múltiplos espaços,
    e expande abreviações comuns de endereços.
    """
    if not texto:
        return ""
    texto = str(texto).upper().strip()
    for padrao, substituicao in ABREVIACOES.items():
        texto = re.sub(padrao, substituicao, texto, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', texto).strip()


def limpar_cep(cep):
    """
    Remove todos os caracteres não numéricos do CEP.
    Trata CEP como string para preservar zeros à esquerda.
    """
    cep_str = str(cep).strip()
    return re.sub(r'\D', '', cep_str)


def validar_formato_cep(cep):
    """
    Verifica se o CEP tem exatamente 8 dígitos numéricos.
    """
    cep_limpo = limpar_cep(cep)
    return len(cep_limpo) == 8 and cep_limpo.isdigit()


def limpar_palavras_duplicadas(texto):
    """
    Remove palavras duplicadas em sequência.
    Ex: "QUADRA QUADRA" -> "QUADRA"
    """
    if not texto:
        return ""
    palavras = texto.upper().split()
    resultado = []
    for palavra in palavras:
        if not resultado or palavra != resultado[-1]:
            resultado.append(palavra)
    return ' '.join(resultado)


def verificar_cep_existe(cep):
    """
    Verifica se o CEP existe nas APIs (AwesomeAPI, ViaCEP).
    Retorna dados do CEP se existir, None caso contrário.
    """
    cep_limpo = limpar_cep(cep)
    
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


def buscar_cep_por_endereco(uf, municipio, logradouro, bairro, complemento):
    """
    Busca CEP por endereço usando múltiplas APIs em cascata.
    Conforme ideia.MD: concatenar ENDERECO + COMPLEMENTO + BAIRRO + MUNICIPIO + UF
    
    Retorna o CEP encontrado ou None se não encontrar.
    """
    # Normalizar entradas
    uf = (uf or '').strip().upper()
    municipio = normalizar_texto(municipio)
    logradouro = normalizar_texto(logradouro)
    bairro = normalizar_texto(bairro)
    complemento = normalizar_texto(complemento)
    
    # Limpar duplicatas
    logradouro = limpar_palavras_duplicadas(logradouro)
    bairro = limpar_palavras_duplicadas(bairro)
    municipio = limpar_palavras_duplicadas(municipio)
    
    # Criar queries de busca (mais específica para menos específica)
    queries = [
        f"{uf} {municipio} {bairro} {logradouro} {complemento}",
        f"{uf} {municipio} {bairro} {logradouro}",
        f"{uf} {municipio} {bairro}",
        f"{uf} {municipio} {logradouro}",
        f"{bairro}, {municipio}, {uf}",
    ]
    
    for query in queries:
        # Tentar CepRua (melhor para texto livre)
        time.sleep(DELAY_REQUISIÇÃO)
        try:
            url = f"https://ceprua.com.br/api/buscar?q={query}"
            res = requests.get(url, timeout=5)
            if res.status_code == 200:
                dados = res.json()
                if dados and isinstance(dados, dict):
                    resultados = dados.get('resultados', [])
                    # Filtrar por UF
                    for r in resultados:
                        if r.get('uf', '').upper() == uf:
                            cep = r.get('cep', '').replace('-', '')
                            if cep and len(cep) == 8:
                                return cep
        except:
            pass
        
        # Tentar AwesomeAPI
        time.sleep(DELAY_REQUISIÇÃO)
        try:
            url = f"https://cep.awesomeapi.com.br/search?q={query}"
            res = requests.get(url, timeout=5)
            if res.status_code == 200:
                dados = res.json()
                if dados.get('results'):
                    for r in dados['results']:
                        if r.get('state', '').upper() == uf:
                            return r.get('cep', '')
        except:
            pass
        
        # Tentar ViaCEP (parâmetros separados)
        time.sleep(DELAY_REQUISIÇÃO)
        try:
            base = logradouro.split(',')[0].strip() if logradouro else ''
            if base and len(base) >= 3:
                url = f"https://viacep.com.br/ws/{uf}/{municipio}/{base}/json/"
                res = requests.get(url, timeout=5)
                if res.status_code == 200:
                    dados = res.json()
                    if dados and isinstance(dados, list):
                        for r in dados:
                            if r.get('uf', '').upper() == uf:
                                return r.get('cep', '').replace('-', '')
        except:
            pass
    
    return None


def processar_linha(args):
    """
    Processa uma única linha da planilha.
    
    Fluxo:
    1. Verifica se o CEP tem formato válido (8 dígitos)
    2. Verifica se o CEP existe na API
    3. Se não existir, busca por endereço
    4. Retorna o resultado com status
    """
    i, row = args
    
    # Extrair dados da linha
    cep_bruto = str(row.get('CEP', ''))
    cep_original = limpar_cep(cep_bruto)
    uf = str(row.get('UF', '')).strip()
    municipio = str(row.get('MUNICIPIO', '')).strip()
    logradouro = str(row.get('ENDERECO', '')).strip()
    bairro = str(row.get('BAIRRO', '')).strip()
    complemento = str(row.get('COMPLEMENTO DO ENDERECO', '')).strip()
    nome = str(row.get('NOME', '')).strip()
    cpf = str(row.get('CPF', '')).strip()
    
    cep_corrigido = ""
    status = ""
    
    # ========================================
    # PASSO 1: Verificar formato e existência
    # ========================================
    
    if validar_formato_cep(cep_bruto):
        # CEP tem formato válido, verificar se existe na API
        dados_api = verificar_cep_existe(cep_original)
        
        if dados_api:
            # CEP existe! Verificar se UF bate
            uf_api = dados_api.get('state', '') or dados_api.get('uf', '')
            if uf.upper() == uf_api.upper():
                # UF bate - CEP válido
                cep_corrigido = cep_original
                status = "Válido"
            else:
                # UF diferente - buscar por endereço
                novo_cep = buscar_cep_por_endereco(uf, municipio, logradouro, bairro, complemento)
                if novo_cep:
                    cep_corrigido = novo_cep
                    status = "Corrigido"
                else:
                    # Não achou - manter original
                    cep_corrigido = cep_original
                    status = "Não Encontrado"
        else:
            # CEP não existe na API - buscar por endereço
            novo_cep = buscar_cep_por_endereco(uf, municipio, logradouro, bairro, complemento)
            if novo_cep:
                cep_corrigido = novo_cep
                status = "Corrigido"
            else:
                cep_corrigido = cep_original
                status = "Não Encontrado"
    else:
        # CEP com formato inválido - buscar por endereço
        novo_cep = buscar_cep_por_endereco(uf, municipio, logradouro, bairro, complemento)
        if novo_cep:
            cep_corrigido = novo_cep
            status = "Corrigido"
        else:
            cep_corrigido = cep_original
            status = "Não Encontrado"
    
    modificado = "Sim" if cep_corrigido != cep_original else "Não"
    
    return i, {
        'nome': nome,
        'cpf': cpf,
        'cep_original': cep_original,
        'cep_corrigido': cep_corrigido,
        'endereco': logradouro,
        'bairro': bairro,
        'municipio': municipio,
        'uf': uf,
        'complemento': complemento,
        'status': status,
        'modificado': modificado
    }


# ============================================
# INTERFACE STREAMLIT
# ============================================

st.markdown("---")

# Modelo de planilha
st.markdown("## 📋 Formato da Planilha")
st.info("A planilha deve conter o seguinte cabeçalho:")
st.code("CPF | MATRICULA | NOME | DATA NASCIMENTO | ENDERECO | COMPLEMENTO DO ENDERECO | BAIRRO | MUNICIPIO | UF | CEP")

# Download do modelo
col1, col2 = st.columns([3, 1])
with col1:
    st.markdown("### 📤 Upload da Planilha")
with col2:
    modelo_df = pd.DataFrame(columns=CABEÇALHO_ESPERADO)
    modelo_buffer = io.BytesIO()
    modelo_df.to_excel(modelo_buffer, index=False)
    st.download_button(
        label="📥 Baixar Modelo",
        data=modelo_buffer.getvalue(),
        file_name="modelo_planilha.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Clique para baixar o modelo de planilha"
    )

# Upload do arquivo
uploaded_file = st.file_uploader("Carregar arquivo Excel (.xlsx)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        # Ler planilha (tratar CEP como string para preservar zeros)
        df = pd.read_excel(uploaded_file, dtype=str)
        total = len(df)
        st.session_state.df_original = df.copy()
        st.markdown(f"**{total} registros carregados**")
        
        # Verificar colunas
        colunas_planilha = list(df.columns)
        colunas_faltando = [c for c in CABEÇALHO_ESPERADO if c not in colunas_planilha]
        
        if colunas_faltando:
            st.warning(f"⚠️ Colunas faltando: {colunas_faltando}")
        else:
            st.success("✅ Cabeçalho correto!")
        
        st.markdown("---")
        
        # Botão processar
        processar = st.button("✅ Processar Planilha", type="primary", use_container_width=True, key="btn_processar")
        
        if processar:
            progress_bar = st.progress(0)
            status_text = st.empty()
            status_text.text("Iniciando processamento...")
            
            resultados = [None] * len(df)
            status_count = {}
            
            # Processamento sequencial (mais confiável para Streamlit)
            progress_bar.progress(0)
            status_text.text("Processando registros...")
            
            resultados = []
            status_count = {}
            
            for idx, row in df.iterrows():
                resultado = processar_linha((idx, row))
                resultados.append(resultado[1])
                status_count[resultado[1]['status']] = status_count.get(resultado[1]['status'], 0) + 1
                
                if idx % 10 == 0:
                    progress = (idx + 1) / total
                    progress_bar.progress(progress)
                    status_text.text(f"Processando... {idx + 1}/{total}")
            
            status_text.text("✅ Processamento concluído!")
            
            # Salvar resultados na session state
            st.session_state.resultados = resultados
            st.session_state.status_count = status_count
            st.session_state.processado = True
        
        # Mostrar resultados se já processado
        if st.session_state.processado and st.session_state.resultados:
            resultados = st.session_state.resultados
            status_count = st.session_state.status_count
            df_original = st.session_state.df_original
            
            st.markdown("---")
            st.markdown("## 📊 Resultados")
            
            # Métricas
            c1, c2, c3, c4 = st.columns(4)
            validos = status_count.get('Válido', 0)
            corrigidos = status_count.get('Corrigido', 0)
            nao_encontrados = status_count.get('Não Encontrado', 0)
            
            c1.metric("Total", len(resultados))
            c2.metric("Válidos", validos, delta_color="normal")
            c3.metric("Corrigidos", corrigidos, delta_color="normal")
            c4.metric("Não Encontrados", nao_encontrados, delta_color="inverse")
            
            st.write("### Detalhes por status:")
            for status, count in sorted(status_count.items()):
                st.write(f"- {status}: {count}")
            
            # ========================================
            # GERAR AS 4 PLANILHAS DE SAÍDA
            # ========================================
            
            # Converter resultados para DataFrame
            df_resultados = pd.DataFrame(resultados)
            
            # 1. CEPs válidos (originais corretos)
            df_validos = df_resultados[df_resultados['status'] == 'Válido'].copy()
            
            # 2. CEPs corrigidos (encontrados via endereço)
            df_corrigidos = df_resultados[df_resultados['status'] == 'Corrigido'].copy()
            
            # 3. CEPs não encontrados
            df_nao_encontrados = df_resultados[df_resultados['status'] == 'Não Encontrado'].copy()
            
            # 4. Base consolidada (toda a base com colunas adicionais)
            df_consolidada = df_original.copy()
            df_consolidada['Novo_CEP'] = [r['cep_corrigido'] for r in resultados]
            df_consolidada['Status_Processamento'] = [r['status'] for r in resultados]
            
            # Criar buffers para download
            buffer_validos = io.BytesIO()
            df_validos.to_excel(buffer_validos, index=False)
            
            buffer_corrigidos = io.BytesIO()
            df_corrigidos.to_excel(buffer_corrigidos, index=False)
            
            buffer_nao_encontrados = io.BytesIO()
            df_nao_encontrados.to_excel(buffer_nao_encontrados, index=False)
            
            buffer_consolidada = io.BytesIO()
            df_consolidada.to_excel(buffer_consolidada, index=False)
            
            # ========================================
            # DOWNLOADS
            # ========================================
            
            st.markdown("---")
            st.markdown("## 📥 Downloads")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.download_button(
                    label=f"📄 1_ceps_validos ({len(df_validos)})",
                    data=buffer_validos.getvalue(),
                    file_name="1_ceps_validos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="CEPs originais que já estavam corretos"
                )
            
            with col2:
                st.download_button(
                    label=f"📄 2_ceps_corrigidos ({len(df_corrigidos)})",
                    data=buffer_corrigidos.getvalue(),
                    file_name="2_ceps_corrigidos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="CEPs que a API encontrou através do Endereço"
                )
            
            with col3:
                st.download_button(
                    label=f"📄 3_ceps_nao_encontrados ({len(df_nao_encontrados)})",
                    data=buffer_nao_encontrados.getvalue(),
                    file_name="3_ceps_nao_encontrados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="CEPs que não foram encontrados"
                )
            
            with col4:
                st.download_button(
                    label=f"📄 4_base_consolidada ({len(df_consolidada)})",
                    data=buffer_consolidada.getvalue(),
                    file_name="4_base_consolidada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Base inteira com Novo_CEP e Status_Processamento"
                )
            
            # Preview
            st.markdown("---")
            st.markdown("## 👀 Preview (primeiros 10)")
            
            for i, r in enumerate(resultados[:10]):
                status = r['status']
                modificado = r['modificado']
                emoji = "🟢" if status == "Válido" else "🟡" if modificado == "Sim" else "🔴"
                st.write(f"{emoji} {r['nome'][:30]}... | {r['cep_original']} → {r['cep_corrigido']} ({status})")
            
            # Botão limpar
            if st.button("🔄 Limpar e Recomeçar"):
                st.session_state.processado = False
                st.session_state.resultados = None
                st.session_state.df_original = None
                st.session_state.status_count = None
                st.rerun()
    
    except Exception as e:
        st.error(f"Erro: {e}")

st.markdown("---")
st.caption("CEP Paperclipe v3.0 - Seguindo ideia.MD")