"""
CEP Paperclipe - Sistema de Validação e Correção de CEPs
Seguindo rigorosamente o idea.MD

4 Planilhas de Saída:
1. 1_ceps_validos.xlsx - Dados originais no modelo Correios
2. 2_ceps_corrigidos.xlsx - Dados atualizados no modelo Correios  
3. 3_ceps_nao_encontrados.xlsx - Log de erros
4. 4_base_consolidada.xlsx - Dados originais + modificado
"""
import io
import pandas as pd
import streamlit as st
import requests
import re
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.styles import numbers
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="CEP Paperclipe", page_icon="📮")
st.title("📮 CEP Paperclipe")
st.markdown("### Validar, Corrigir CEPs e Gerar Lista para Envio")

# Colunas esperadas
CABEÇALHO_ESPERADO = ['CPF', 'MATRICULA', 'NOME', 'DATA NASCIMENTO', 'ENDERECO', 
                       'COMPLEMENTO DO ENDERECO', 'BAIRRO', 'MUNICIPIO', 'UF', 'CEP']

# Session state
if 'processado' not in st.session_state:
    st.session_state.processado = False
if 'resultados' not in st.session_state:
    st.session_state.resultados = None


# ============================================
# FUNÇÕES
# ============================================

def excel_to_bytes_text(df):
    """Gera Excel com todas as colunas em formato TEXTO (@)"""
    from openpyxl.styles import numbers
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"
    
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal='left')
            cell.number_format = numbers.FORMAT_TEXT
            if r_idx == 1:
                cell.font = Font(bold=True)
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def excel_correios_format(rows):
    """Gera Excel no modelo Correios com formato personalizado (zeros)"""
    from openpyxl.styles import numbers
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Correios"
    
    # Cabeçalho
    headers = ['PREFIXO', 'NOME', 'CPF', 'CEP', 'LOGRADOURO', 'ENDERECO', 'BAIRRO', 'MUNICIPIO', 'UF', 'FLAG', 'COMPLEMENTO', 'NUMERO']
    for c_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=header)
        cell.font = Font(bold=True)
    
    # Dados
    for r_idx, row in enumerate(rows, 2):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal='left')
            
            # Formato especial para CPF (11 zeros) e CEP (8 zeros)
            if c_idx == 3:  # CPF
                cell.number_format = '00000000000'
            elif c_idx == 4:  # CEP
                cell.number_format = '00000000'
    
    # Ajustar largura
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 5
    ws.column_dimensions['F'].width = 35
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 8
    ws.column_dimensions['J'].width = 5
    ws.column_dimensions['K'].width = 40
    ws.column_dimensions['L'].width = 8
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def limpar_cep(cep):
    """Remove caracteres não numéricos"""
    return re.sub(r'\D', '', str(cep).strip())

def normalizar_cep(cep):
    """Completa CEP com zeros à esquerda para ter 8 dígitos"""
    cep_limpo = limpar_cep(cep)
    return cep_limpo.zfill(8)

def normalizar_cpf(cpf):
    """Completa CPF com zeros à esquerda para ter 11 dígitos"""
    cpf_limpo = re.sub(r'\D', '', str(cpf).strip())
    return cpf_limpo.zfill(11)

def validar_formato_cep(cep):
    """Verifica se tem 8 dígitos"""
    cep_limpo = limpar_cep(cep)
    return len(cep_limpo) == 8 and cep_limpo.isdigit()

def verificar_cep(cep):
    """Verifica CEP na API CepRua"""
    cep_limpo = limpar_cep(cep)
    if not validar_formato_cep(cep):
        return {"valido": False, "cep": "", "status": "Formato inválido"}
    
    try:
        url = f"https://ceprua.com.br/api/buscar?q={cep_limpo}"
        res = requests.get(url, timeout=3)
        if res.status_code == 200:
            dados = res.json()
            if 'redirect' in dados:
                return {"valido": True, "cep": cep_limpo, "status": "Válido"}
    except:
        pass
    
    return {"valido": False, "cep": cep_limpo, "status": "Não encontrado"}

def buscar_cep_fallback(endereco, complemento, bairro, municipio, uf):
    """Tenta buscar CEP usando várias combinações de endereço - SÓ retorna se encontrar no mesmo UF"""
    uf_upper = uf.upper().strip() if uf else ''
    municipio_upper = municipio.upper().strip() if municipio else ''
    endereco_upper = endereco.upper() if endereco else ''
    partes = endereco_upper.split() if endereco else []
    
    buscas = []
    
    # 1. Bairro + cidade + UF
    buscas.append(f"{bairro} {municipio} {uf}".strip())
    
    # 2. Palavras do endereço + cidade + UF
    if len(partes) >= 2:
        for i in range(1, min(4, len(partes))):
            palavras = ' '.join(partes[:i])
            if len(palavras) > 5:
                buscas.append(f"{palavras} {municipio} {uf}".strip())
    
    # 3. Endereço completo
    buscas.append(f"{endereco} {bairro} {municipio} {uf}".strip())
    
    # 4. Padrões específicos do DF (SHIN, SHIS, SHIGS, SMPW, QI, QN, etc)
    padroes_df = ['SHIN', 'SHIS', 'SHIGS', 'SMPW', 'SQN', 'SQS', 'SRES', 'SQSW', 'SQNW', 'QI ', 'QN ', 'QS ']
    for padrao in padroes_df:
        if padrao in endereco_upper:
            idx = endereco_upper.find(padrao)
            parte = endereco[idx:idx+10].strip()
            if parte:
                buscas.append(f"{parte} {municipio} {uf}".strip())
                buscas.append(f"{parte} DF".strip())
                # Também tentar só a área (SHIN, SHIS, etc)
                if 'SHIN' in parte.upper():
                    buscas.append("SHIN DF")
                    buscas.append("ASA NORTE DF")
                elif 'SHIS' in parte.upper():
                    buscas.append("SHIS DF")
                    buscas.append("ASA SUL DF")
                elif 'QI ' in parte.upper() or 'QN ' in parte.upper():
                    buscas.append("ASA NORTE DF")
                    buscas.append("ASA SUL DF")
    
    # 5. Sem "QUADRA" ou "CONDOMINIO" + bairro + cidade + UF
    for palavra in ['QUADRA', 'CONDOMINIO', 'EDIFICIO', 'VILA', 'SETOR']:
        if palavra in endereco_upper:
            temp = endereco_upper.replace(palavra, '').strip()
            if len(temp) > 5:
                buscas.append(f"{temp} {bairro} {municipio} {uf}".strip())
    
    # 6. só a primeira palavra significativa do endereço + DF
    if partes:
        buscas.append(f"{partes[0]} DF".strip())
    
    # 7. Cidade + UF
    buscas.append(f"{municipio} {uf}".strip())
    
    for query in buscas:
        if not query or len(query) < 4:
            continue
        try:
            url = f"https://ceprua.com.br/api/buscar?q={query}"
            res = requests.get(url, timeout=3)
            if res.status_code == 200:
                dados = res.json()
                if 'resultados' in dados and dados['resultados']:
                    for item in dados['resultados']:
                        cep = item.get('cep', '').replace('-', '')
                        item_uf = item.get('uf', '').upper().strip()
                        if len(cep) == 8 and item_uf == uf_upper:
                            return cep
                if 'redirect' in dados:
                    cep = dados['redirect'].get('cep', '')
                    if cep:
                        return cep.replace('-', '')
        except:
            pass
    
    return ''

def gerar_linha_correios(nome, cpf, cep, endereco, bairro, municipio, uf, complemento):
    """Gera linha no formato Correios - TODOS os campos como TEXTO com aspas"""
    # Normalizar CPF (11 dígitos) e CEP (8 dígitos)
    cpf = normalizar_cpf(cpf)
    cep = normalizar_cep(cep)
    
    # Garantir que são strings
    nome = str(nome) if nome else ''
    endereco = str(endereco) if endereco else ''
    bairro = str(bairro) if bairro else ''
    municipio = str(municipio) if municipio else ''
    uf = str(uf) if uf else ''
    complemento = str(complemento) if complemento else ''
    
    # Limitar complemento a 40 caracteres
    if len(complemento) > 40:
        complemento = complemento[:40]
    
    # Formato com aspas para forçar texto
    return f'SR(A);"{nome.upper()}";"{cpf}";"{cep}";"";"{endereco.upper()}";"{bairro.upper()}";"{municipio.upper()}";"{uf.upper()}";"N";"{complemento.upper()}";"0"'


# ============================================
# INTERFACE
# ============================================

st.markdown("---")
st.markdown("## 📋 Formato da Planilha")
st.info("Colunas: CPF | MATRICULA | NOME | DATA NASCIMENTO | ENDERECO | COMPLEMENTO DO ENDERECO | BAIRRO | MUNICIPIO | UF | CEP")

# Download modelo
col1, col2 = st.columns([3, 1])
with col1:
    st.markdown("### 📤 Upload da Planilha")
with col2:
    modelo_df = pd.DataFrame(columns=CABEÇALHO_ESPERADO)
    modelo_buffer = excel_to_bytes_text(modelo_df)
    st.download_button("📥 Baixar Modelo", modelo_buffer.getvalue(), "modelo_planilha.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

uploaded_file = st.file_uploader("Carregar arquivo Excel (.xlsx)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=str)
        df.columns = df.columns.str.strip()
        for col in df.columns:
            df[col] = df[col].astype(str)
        
        colunas_faltando = [c for c in CABEÇALHO_ESPERADO if c not in df.columns]
        if colunas_faltando:
            st.warning(f"Colunas faltando: {colunas_faltando}")
        else:
            st.success(f"✅ {len(df)} registros carregados")
        
        st.markdown("---")
        
        if st.button("✅ Processar Planilha", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            status_text.text("Processando CEPs...")
            
            total = len(df)
            st.session_state.resultados = []
            status_count = {"Válido": 0, "Não encontrado": 0, "Formato inválido": 0}
            
            for idx, row in df.iterrows():
                cep_original = str(row.get('CEP', '')).strip()
                cep_limpo = normalizar_cep(cep_original)
                
                # Validar CEP na API
                resultado = verificar_cep(cep_original)
                
                cep_final = resultado["cep"] if resultado["valido"] else cep_limpo
                status = resultado["status"]
                modificado = "Sim" if cep_final != cep_limpo else "Não"
                
                status_count[status] = status_count.get(status, 0) + 1
                
                st.session_state.resultados.append({
                    "nome": str(row.get('NOME', '')),
                    "cpf": normalizar_cpf(str(row.get('CPF', ''))),
                    "cep_original": cep_limpo,
                    "cep_corrigido": cep_final,
                    "endereco": str(row.get('ENDERECO', '')),
                    "bairro": str(row.get('BAIRRO', '')),
                    "municipio": str(row.get('MUNICIPIO', '')),
                    "uf": str(row.get('UF', '')),
                    "complemento": str(row.get('COMPLEMENTO DO ENDERECO', '')),
                    "status": status,
                    "modificado": modificado,
                    "cep_sugerido": ""
                })
                
                if (idx + 1) % 100 == 0:
                    progress_bar.progress((idx + 1) / total)
                    status_text.text(f"Processando... {idx + 1}/{total}")
            
            # FALLBACK: Buscar CEPs não encontrados
            status_text.text("Buscando CEPs sugeridos...")
            
            nao_encontrados = [i for i, r in enumerate(st.session_state.resultados) if r['status'] == "Não encontrado"]
            
            for i, idx in enumerate(nao_encontrados):
                r = st.session_state.resultados[idx]
                cep_sugerido = buscar_cep_fallback(
                    r['endereco'], r['complemento'], r['bairro'], r['municipio'], r['uf']
                )
                st.session_state.resultados[idx]['cep_sugerido'] = normalizar_cep(cep_sugerido) if cep_sugerido else ''
                
                if (i + 1) % 50 == 0:
                    progress_bar.progress((i + 1) / len(nao_encontrados))
                    status_text.text(f"Buscando sugeridos... {i + 1}/{len(nao_encontrados)}")
            
            # ========================================
            # SEM FALLBACK - pulando busca de CEPs sugeridos para deixar mais rápido
            # ========================================
            status_text.text("✅ Concluído!")
            
            # ========================================
            # GERAR AS 4 PLANILHAS
            # ========================================
            
            # 1. CSV - Dados ORIGINAIS - dividido em blocos de 300
            linhas_original = []
            for r in st.session_state.resultados:
                linhas_original.append(
                    f'SR(A);{r["nome"].upper()};;;;{r["cpf"]};{r["cep_original"]};;'
                    f'{r["endereco"].upper()};{r["bairro"].upper()};{r["municipio"].upper()};{r["uf"].upper()};N;'
                    f'{r["complemento"].upper()[:40]};0'
                )
            
            # Criar blocos de 300
            bloco_size = 300
            st.session_state.blocos_original = []
            for i in range(0, len(linhas_original), bloco_size):
                bloco = '\ufeff' + '\n'.join(linhas_original[i:i+bloco_size])
                st.session_state.blocos_original.append(bloco)
            
            csv_original = '\ufeff' + '\n'.join(linhas_original)
            st.session_state.csv_original = csv_original
            
            # 2. CSV - Dados CORRIGIDOS - dividido em blocos de 300
            linhas_corrigidos = []
            for r in st.session_state.resultados:
                # Usar CEP corrigido ou sugerido se válido
                if r['status'] == 'Válido':
                    cep_usar = r['cep_corrigido']
                elif r['cep_sugerido']:
                    cep_usar = r['cep_sugerido']
                else:
                    cep_usar = r['cep_original']
                
                linhas_corrigidos.append(
                    f'SR(A);{r["nome"].upper()};;;;{r["cpf"]};{cep_usar};;'
                    f'{r["endereco"].upper()};{r["bairro"].upper()};{r["municipio"].upper()};{r["uf"].upper()};N;'
                    f'{r["complemento"].upper()[:40]};0'
                )
            
            # Criar blocos de 300
            st.session_state.blocos_corrigidos = []
            for i in range(0, len(linhas_corrigidos), bloco_size):
                bloco = '\ufeff' + '\n'.join(linhas_corrigidos[i:i+bloco_size])
                st.session_state.blocos_corrigidos.append(bloco)
            
            csv_corrigidos = '\ufeff' + '\n'.join(linhas_corrigidos)
            st.session_state.csv_corrigidos = csv_corrigidos
            st.session_state.count_corrigidos = len(linhas_corrigidos)
            
            # 3. Excel - Log de status (sem validação API agora)
            # Como não validamos mais, este arquivo mostra todos os registros processados
            df_erros = pd.DataFrame(st.session_state.resultados)
            df_erros = df_erros[['nome', 'cpf', 'cep_original', 'endereco', 
                                  'bairro', 'municipio', 'uf', 'complemento']]
            df_erros = df_erros.rename(columns={
                'cep_original': 'CEP'
            })
            for col in df_erros.columns:
                df_erros[col] = df_erros[col].astype(str)
            buffer_erros = excel_to_bytes_text(df_erros)
            st.session_state.buffer_erros = buffer_erros.getvalue()
            st.session_state.count_erros = len(df_erros)
            
            # 4. Excel - Base consolidada - formato TEXTO
            df_consolidada = df.copy()
            for col in df_consolidada.columns:
                df_consolidada[col] = df_consolidada[col].astype(str)
            buffer_consolidada = excel_to_bytes_text(df_consolidada)
            st.session_state.buffer_consolidada = buffer_consolidada.getvalue()
            
            st.session_state.status_count = status_count
            st.session_state.processado = True
    
    except Exception as e:
        st.error(f"Erro: {e}")

# ========================================
# MOSTRAR RESULTADOS (se já processado)
# ========================================
if st.session_state.get('processado', False):
    st.markdown("---")
    st.markdown("## 📊 Resultados")
    
    c1, c2 = st.columns(2)
    c1.metric("Total processado", len(st.session_state.resultados))
    c2.metric("Status", "Pendente validação")
    
    st.markdown("---")
    st.markdown("## 📥 Downloads")
    
    st.markdown("### 📥 Download Completo")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.download_button(f"📄 1_Original ({len(st.session_state.resultados)})", st.session_state.csv_original, "1_ceps_validos.csv", "text/csv; charset=utf-8-sig", use_container_width=True)
    
    with col2:
        st.download_button(f"📄 2_Corrigidos ({st.session_state.count_corrigidos})", st.session_state.csv_corrigidos, "2_ceps_corrigidos.csv", "text/csv; charset=utf-8-sig", use_container_width=True)
    
    with col3:
        st.download_button(f"📄 3_Erros ({st.session_state.count_erros})", st.session_state.buffer_erros, "3_ceps_nao_encontrados.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    
    with col4:
        st.download_button(f"📄 4_Consolidada ({len(st.session_state.resultados)})", st.session_state.buffer_consolidada, "4_base_consolidada.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    
    st.markdown("### 📥 Blocos de 300 (1_Original)")
    for i, bloco in enumerate(st.session_state.get('blocos_original', [])):
        st.download_button(f"Bloco {i+1} ({min(i*300+300, len(st.session_state.resultados))})", bloco, f"1_ceps_validos_bloco_{i+1}.csv", "text/csv; charset=utf-8-sig", use_container_width=True)
    
    st.markdown("### 📥 Blocos de 300 (2_Corrigidos)")
    for i, bloco in enumerate(st.session_state.get('blocos_corrigidos', [])):
        st.download_button(f"Bloco {i+1} ({min(i*300+300, len(st.session_state.resultados))})", bloco, f"2_ceps_corrigidos_bloco_{i+1}.csv", "text/csv; charset=utf-8-sig", use_container_width=True)
    
    st.success("✅ Processamento concluído!")

st.markdown("---")
st.caption("CEP Paperclipe v5.0 - Seguindo idea.MD")