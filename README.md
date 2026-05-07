# CEP Paperclipe - Sistema de Validação e Correção de CEPs

Sistema para processar planilhas de prova de vida, validar CEPs automaticamente, corrigir CEPs inválidos usando APIs públicas e gerar arquivos prontos para envio aos Correios.

---

## Problema

Muitos beneficiários informam CEPs incorretos ou CEPs gerais da localidades (ex: CEP da cidade ao invés do CEP específico do endereço), o que causa erros no envio de cartas pelos Correios.

Este sistema resolve esse problema validando cada CEP e, quando necessário, buscando o CEP correto usando o endereço completo.

---

## Formato das Planilhas

### Entrada: Prova de Vida (Excel)

| CPF | MATRICULA | NOME | DATA NASCIMENTO | ENDERECO | COMPLEMENTO DO ENDERECO | BAIRRO | MUNICIPIO | UF | CEP |
|-----|----------|------|----------------|----------|-------------------------|-------|-----------|-----|-----|
| 12345678901 | 001234 | JOAO SILVA | 15/03/1980 | QUADRA SHIGS 706 | CASA 15 | ASA SUL | BRASILIA | DF | 70350756 |

### Saída: 4 Planilhas Excel

| Arquivo | Descrição |
|---------|-----------|
| `1_ceps_validos.xlsx` | Apenas os CEPs originais que já estavam corretos |
| `2_ceps_corrigidos.xlsx` | CEPs que a API encontrou e corrigiu via endereço |
| `3_ceps_nao_encontrados.xlsx` | CEPs inválidos que a API não encontrou |
| `4_base_consolidada.xlsx` | Base inteira com colunas `Novo_CEP` e `Status_Processamento` |

---

## Status de Processamento

| Status | Significado |
|--------|-------------|
| **Válido** | CEP original está correto e existe na API |
| **Corrigido** | CEP estava errado/inválido, a API encontrou o correto via endereço |
| **Não Encontrado** | CEP inválido e a API não encontrou correspondência |

---

## Fluxo de Processamento

```
1. Upload Planilha Excel (.xlsx)
         ↓
2. Validar CEP (formato 8 dígitos + existência na API)
         ↓
3. Se inválido → Buscar CEP por endereço (ENDERECO + COMPLEMENTO + BAIRRO + MUNICIPIO + UF)
         ↓
4. Gerar 4 planilhas de saída
```

---

## Como o Código Funciona

### 1. Validação do CEP

```python
def validar_formato_cep(cep):
    cep_limpo = ''.join(filter(str.isdigit, str(cep)))
    return len(cep_limpo) == 8 and cep_limpo.isdigit()
```

- Se o CEP tiver menos de 8 dígitos → **INVÁLIDO**
- Se o CEP tiver letras → **INVÁLIDO**
- Se o CEP tiver 8 dígitos → **VÁLIDO (formato)**

### 2. Verificar se o CEP existe na API

```python
def verificar_cep_existe(cep):
    # Tenta AwesomeAPI
    url = f"https://cep.awesomeapi.com.br/json/{cep}"
    res = requests.get(url, timeout=5)
    if res.status_code == 200 and 'cep' in res.json():
        return dados  # CEP existe
    
    # Tenta ViaCEP como backup
    url = f"https://viacep.com.br/ws/{cep}/json/"
    res = requests.get(url, timeout=5)
    if res.status_code == 200 and 'erro' not in res.json():
        return dados
    
    return None  # Não encontrado
```

### 3. Busca de CEP por Endereço

Quando o CEP está inválido ou não existe, o código busca usando o endereço:

```
QUERY = UF + MUNICIPIO + BAIRRO + ENDERECO + COMPLEMENTO
```

**Exemplo:**
- UF: GO | MUNICIPIO: AGUAS LINDAS DE GOIAS | BAIRRO: PARQUE DA BARRAGEM V
- ENDERECO: QUADRA QUADRA 14 | COMPLEMENTO: CASA 71
- **Query:** `GO AGUAS LINDAS DE GOIAS PARQUE DA BARRAGEM V QUADRA QUADRA 14 CASA 71`

### 4. APIs de Busca (3 em cascata)

| # | API | Endpoint | Tipo |
|---|-----|----------|------|
| 1 | CepRua | `ceprua.com.br/api/buscar?q={query}` | Texto livre (prioritário) |
| 2 | AwesomeAPI | `cep.awesomeapi.com.br/search?q={query}` | Texto livre |
| 3 | ViaCEP | `viacep.com.br/ws/{UF}/{Cidade}/{Logradouro}/json/` | Parâmetros separados |

### 5. Normalização de Endereços

O código normaliza abreviações para melhorar a taxa de sucesso:

| Abreviação | Normalizado |
|------------|-------------|
| R | RUA |
| AV | AVENIDA |
| QD / QN | QUADRA |
| CS | CASA |
| AP | APARTAMENTO |
| BL | BLOCO |
| LT | LOTE |
| CONJ / CJ | CONJUNTO |

### 6. Limpeza de Palavras Duplicadas

Remove palavras duplicadas em sequência:
- `"QUADRA QUADRA"` → `"QUADRA"`
- `"RUA RUA RUA"` → `"RUA"`

---

## Estrutura de Arquivos

```
CEP_paperclipe/
├── src/
│   ├── app.py                  # Interface Streamlit
│   ├── ideia.MD                # Base de conhecimento
│   ├── tests/
│   │   ├── test_funcoes.py     # Testes unitários
│   │   └── test_api.py         # Testes de integração
│   └── requirements.txt         # Dependências
├── README.md
└── .gitignore
```

---

## Como Executar

### Via Streamlit (Navegador)

```bash
cd src
.venv\Scripts\streamlit.exe run app.py
```

Acesse: http://localhost:8501

---

## Executar Testes

```bash
cd src
.venv\Scripts\python.exe tests/test_funcoes.py
```

---

## Dependências

```
pandas
requests
openpyxl
streamlit
```

---

## Limites Técnicos

| Item | Limite |
|------|--------|
| Lote Correios | 300 objetos por envio |
| Delay entre requisições | 0.1 segundos |
| Timeout API | 5 segundos |

---

## Histórico de Versões

| Versão | Data | Descrição |
|--------|------|-----------|
| 3.0 | 2026-05-07 | Versão atual seguindo ideia.MD - 4 planilhas de saída + testes |

---

## Contato

Desenvolvedor responsável pelo projeto.