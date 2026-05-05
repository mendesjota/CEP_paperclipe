import io

import pandas as pd
import streamlit as st
from cep.cep import get_address_from_cep, get_cep_from_address

st.set_page_config(
    page_title="Consulta CEP",
    page_icon="📮",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stApp {
        background-color: #ffffff;
    }
    .header-title {
        font-size: 2.5rem !important;
        font-weight: 700 !important;
        color: #1e3a5f !important;
        text-align: center;
        margin-bottom: 0.5rem !important;
    }
    .header-subtitle {
        font-size: 1.1rem !important;
        color: #6c757d !important;
        text-align: center;
        margin-bottom: 2rem !important;
    }
    .result-card {
        background-color: #f8f9fa;
        border-radius: 12px;
        padding: 1.5rem;
        border: 1px solid #e9ecef;
    }
    .stSuccess {
        background-color: #d4edda !important;
        border-radius: 8px !important;
    }
    .stError {
        background-color: #f8d7da !important;
        border-radius: 8px !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="header-title">📮 Consulta de CEP</p>', unsafe_allow_html=True)
st.markdown('<p class="header-subtitle">Encontre endereços brasileiros pelo CEP ou endereço</p>', unsafe_allow_html=True)

col_search, col_info = st.columns([3, 1])

with col_search:
    search_type = st.radio("", ["Buscar por CEP", "Buscar por Endereço"], horizontal=True)

if search_type == "Buscar por CEP":
    cep_input = st.text_input("Digite o CEP:", placeholder="12345-678", max_chars=9)
    col_btn, col_spacer = st.columns([1, 4])
    with col_btn:
        search_btn = st.button("Consultar", type="primary", use_container_width=True)
    
    if search_btn:
        if not cep_input:
            st.error("Por favor, digite um CEP válido.")
        else:
            with st.spinner("Consultando..."):
                result = get_address_from_cep(cep_input)
                
                if result:
                    st.success("✅ Endereço encontrado!")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("### 📍 Endereço Principal")
                        st.markdown(f"**Logradouro:** {result.get('logradouro', 'N/A')}")
                        st.markdown(f"**Bairro:** {result.get('bairro', 'N/A')}")
                        st.markdown(f"**Cidade/UF:** {result.get('localidade', 'N/A')}/{result.get('uf', 'N/A')}")
                        st.markdown(f"**CEP:** {result.get('cep', 'N/A')}")
                    with col2:
                        st.markdown("### ℹ️ Dados Complementares")
                        st.markdown(f"**Complemento:** {result.get('complemento', 'N/A')}")
                        st.markdown(f"**IBGE:** {result.get('ibge', 'N/A')}")
                        st.markdown(f"**GIA:** {result.get('gia', 'N/A')}")
                        st.markdown(f"**DDD:** {result.get('ddd', 'N/A')}")

                    df = pd.DataFrame([{
                        "CEP": result.get("cep", ""),
                        "Logradouro": result.get("logradouro", ""),
                        "Bairro": result.get("bairro", ""),
                        "Cidade": result.get("localidade", ""),
                        "UF (Estado)": result.get("uf", ""),
                        "Complemento": result.get("complemento", ""),
                        "IBGE": result.get("ibge", ""),
                        "GIA": result.get("gia", ""),
                        "DDD": result.get("ddd", "")
                    }])

                    buffer = io.BytesIO()
                    df.to_excel(buffer, index=False, engine="openpyxl")
                    buffer.seek(0)

                    cep_for_filename = result.get("cep", "unknown").replace("-", "")
                    st.download_button(
                        label="📥 Exportar para Excel",
                        data=buffer,
                        file_name=f"consulta_cep_{cep_for_filename}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("❌ CEP não encontrado ou inválido.")
else:
    address_input = st.text_input("Digite o endereço:", placeholder="Rua Example, São Paulo, SP")
    col_btn, col_spacer = st.columns([1, 4])
    with col_btn:
        search_btn = st.button("Buscar CEP", type="primary", use_container_width=True)
    
    if search_btn:
        if not address_input:
            st.error("Por favor, digite um endereço válido.")
        else:
            with st.spinner("Consultando..."):
                result = get_cep_from_address(address_input)
                
                if result:
                    st.success(f"✅ CEP encontrado: **{result}**")

                    full_address = get_address_from_cep(result)
                    if full_address:
                        df = pd.DataFrame([{
                            "CEP": full_address.get("cep", ""),
                            "Logradouro": full_address.get("logradouro", ""),
                            "Bairro": full_address.get("bairro", ""),
                            "Cidade": full_address.get("localidade", ""),
                            "UF (Estado)": full_address.get("uf", ""),
                            "Complemento": full_address.get("complemento", ""),
                            "IBGE": full_address.get("ibge", ""),
                            "GIA": full_address.get("gia", ""),
                            "DDD": full_address.get("ddd", "")
                        }])

                        buffer = io.BytesIO()
                        df.to_excel(buffer, index=False, engine="openpyxl")
                        buffer.seek(0)

                        cep_for_filename = result.replace("-", "")
                        st.download_button(
                            label="📥 Exportar para Excel",
                            data=buffer,
                            file_name=f"consulta_cep_{cep_for_filename}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                else:
                    st.error("❌ Endereço não encontrado ou inválido.")

with col_info:
    st.markdown("### ℹ️ Sobre")
    st.info("""
    **Fonte de Dados:** API ViaCEP
    
    Esta ferramenta consultation endereços postais brasileiros utilizando o banco de dados dos Correios.
    """)