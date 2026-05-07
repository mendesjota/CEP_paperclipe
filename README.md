# CEP Paperclipe - Sistema de Validação e Correção de CEPs

Sistema para processar planilhas de prova de vida, validar CEPs automaticamente, corrigir CEPs inválidos usando APIs públicas e gerar arquivos prontos para envio aos Correios.

---

## Problema

Muitos beneficiários informam CEPs incorretos ou CEPs gerais da localidades (ex: CEP da cidade instead do CEP específico do endereço), o que causa erros no envio de cartas pelos Correios.

Este sistema resolve esse problema validando cada CEP e, quando necessário, buscando o CEP correto usando o endereço completo.

---

## Fluxo de Processamento

```
PLANILHA ENTRADA                    PROCESSAMENTO                     PLANILHA SAÍDA
(Prova de Vida)                     (Código Python)                   (Lista para Envio)
┌─────────────┐                    ┌──────────────────┐              ┌─────────────┐
│ NOME        │                    │ 1. Ler Excel     │              │ SR(A);...   │
│ CPF         │ ───────────────►   │ 2. Validar CEP   │ ──────────► │ CSV formato │
│ ENDEREÇO   │                    │ 3. Buscar API    │              │ Correios    │
│ COMPLEMENTO│                    │ 4. Corrigir      │              │             │
│ BAIRRO     │                    │ 5. Gerar CSV     │              │             │
│ MUNICIPIO  │                    └──────────────────┘              └─────────────┘
│ UF         │
│ CEP        │
└─────────────┘
```

---

## Formato das Planilhas

### Entrada: Prova de Vida (Excel)

| Campo | Descrição |
|-------|------------|
| CPF | Número do CPF |
| NOME | Nome do beneficiário |
| ENDEREÇO | Rua/Quadra/Avenida |
| COMPLEMENTO DO ENDEREÇO | Casa/Apartamento/Bloco |
| BAIRRO | Bairro |
| MUNICIPIO | Cidade |
| UF | Estado |
| CEP | CEP informado pelo beneficiário |

### Saída: Lista para Envio (CSV)

```
SR(A);NOME;;;;CPF;CEP;;ENDERECO;BAIRRO;MUNICIPIO;UF;N;COMPLEMENTO;0
```

| Campo | Posição | Exemplo |
|-------|---------|---------|
| Prefixo | 1 | SR(A) |
| Nome | 2 | WILSON CARLOS DE SOUZA |
| Campos vazios | 3-5 | ;;; |
| CPF | 6 | 13650106 |
| CEP | 7 | 70350756 |
| Campo vazio | 8 | ; |
| Endereço | 9 | QUADRA SHIGS 706 BLOCO F |
| Bairro | 10 | ASA SUL |
| Município | 11 | BRASILIA |
| UF | 12 | DF |
| Flag | 13 | N |
| Complemento | 14 | CASA NR.28 |
| Número | 15 | 0 |

---

## Como o Código Funciona

### 1. Validação do CEP

Para cada registro, o código executa uma sequência de validações:

#### Passo 1: Verificar formato
```python
def validar_cep(cep):
    c = limpar_cep(cep)  # Remove tudo que não é número
    return len(c) == 8 and c.isdigit()  # Tem exatamente 8 dígitos?
```

- Se o CEP tiver menos de 8 dígitos → **INVÁLIDO**
- Se o CEP tiver letras → **INVÁLIDO**
- Se o CEP tiver 8 dígitos → **VÁLIDO (formato)**

#### Passo 2: Verificar se existe na API
```python
def verificar_cep(cep):
    url = f"https://cep.awesomeapi.com.br/json/{cep}"
    res = requests.get(url, timeout=3)
    if res.status_code == 200 and 'cep' in res.json():
        return dados  # CEP existe
    return None  # CEP não existe
```

- Se a API não retornar dados → **NÃO ENCONTRADO**
- Se a API retornar dados → **Existe**

#### Passo 3: Verificar se o CEP bate com o endereço
```python
def verificar_endereco_match(endereco_original, endereco_api, bairro_original, bairro_api, uf_original, uf_api):
    # Normaliza os endereços
    n_end_orig = normalizar(endereco_original)
    n_end_api = normalizar(endereco_api)
    
    # Verifica se são compatíveis
    end_match = n_end_orig in n_end_api or n_end_api in n_end_orig
    bairro_match = n_bairro_orig in n_bairro_api or n_bairro_api in n_bairro_orig
    uf_match = uf_orig.upper() == uf_api.upper()
    
    return end_match and bairro_match and uf_match
```

- Se o CEP não bater com o endereço → **CEP ERRADO**
- Se bater → **OK**

---

### 2. Busca de CEP por Endereço

Quando o CEP está inválido ou errado, o código busca o CEP correto usando o endereço:

```
QUERY DE BUSCA = UF + MUNICIPIO + BAIRRO + ENDEREÇO + COMPLEMENTO
```

**Exemplo:**
- UF: GO
- MUNICIPIO: AGUAS LINDAS DE GOIAS
- BAIRRO: PARQUE DA BARRAGEM V
- ENDEREÇO: QUADRA QUADRA 14
- COMPLEMENTO: CASA 71
- **Query final:** `GO AGUAS LINDAS DE GOIAS PARQUE DA BARRAGEM V QUADRA QUADRA 14 CASA 71`

#### APIs de Busca (3 em cascata)

| # | API | Endpoint | Tipo |
|---|-----|----------|------|
| 1 | AwesomeAPI | `cep.awesomeapi.com.br/search?q={query}` | Texto livre (mais flexível) |
| 2 | CepRua | `ceprua.com.br/api/buscar?q={query}` | Texto livre (backup) |
| 3 | ViaCEP | `viacep.com.br/ws/{UF}/{Cidade}/{Logradouro}/json/` | Parâmetros separados (último recurso) |

O código tenta cada API em ordem até encontrar um CEP. Se nenhuma funcionar, marca como "NÃO ENCONTRADO".

---

### 3. Normalização de Endereços

Para melhorar a taxa de sucesso na busca, o código normaliza abreviações:

| Abreviação | Normalizado |
|------------|-------------|
| R | RUA |
| AV | AVENIDA |
| QD | QUADRA |
| QN | QUADRA |
| CS | CASA |
| AP | APARTAMENTO |
| BL | BLOCO |
| LT | LOTE |
| CONJ | CONJUNTO |
| CJ | CONJUNTO |

```python
def normalizar(texto):
    texto = str(texto).upper()
    substituicoes = {
        r'\bR\b': 'RUA',
        r'\bQD\b': 'QUADRA',
        # ... outros
    }
    for padrao, substituto in substituicoes.items():
        texto = re.sub(padrao, substituto, texto)
    return texto.strip()
```

---

### 4. Processamento Paralelo

Para速度speed up o processamento de milhares de registros, o código usa processamento paralelo:

```python
from concurrent.futures import ThreadPoolExecutor

with ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(processar_linha, (i, row)): i for i, row in df.iterrows()}
    for future in as_completed(futures):
        # Processa resultados
```

Isso permite fazer múltiplas requisições API simultaneamente, reduciendo o tempo total de processamento.

---

## Scripts do Projeto

### Arquivos Principais

| Arquivo | Função |
|---------|--------|
| `processador.py` | Módulo principal com funções de validação, normalização e busca |
| `buscar_corrigido.py` | Script que busca CEPs inválidos usando as APIs |
| `gerar_final.py` | Script que gera o CSV no formato Correios |
| `app.py` | Interface Streamlit para uso via navegador |

---

## Como Executar

### Via Terminal (Python)

```bash
cd src
.venv\Scripts\python.exe processar_validar.py
```

### Via Streamlit (Navegador)

```bash
cd src
.venv\Scripts\streamlit.exe run app.py
```

Acesse: http://localhost:8501

---

## Resultados do Processamento

Exemplo com a planilha "2026 04 07 - prova de vida pendentes MARÇO.xlsx":

| Status | Quantidade | Descrição |
|--------|------------|-----------|
| CEPs mantidos (originais válidos) | 1205 | CEPs que já estavam corretos |
| CEPs corrigidos (buscados) | 385 | CEPs encontrados via API |
| Não encontrados | 0 | Endereços que a API não retornou |
| **Total** | **1590** | Total de registros processados |

---

## Formato de Saída

O arquivo gerado `LISTA_PARA_ENVIO.csv` está pronto para upload no sistema dos Correios.

Cada linha segue o padrão:
```
SR(A);NOME;;;;CPF;CEP;;ENDERECO;BAIRRO;MUNICIPIO;UF;N;COMPLEMENTO;0
```

Exemplo completo:
```
SR(A);WILSON CARLOS DE SOUZA;;;;13650106;70350756;;QUADRA SHIGS 706 BLOCO F;ASA SUL;BRASILIA;DF;N;CASA NR.28;0
```

---

## Limites e Considerações

| Item | Limite |
|------|--------|
| Lote Correios | 500 objetos por envio |
| Rate limit API | ~10 requisições/segundo |
| Timeout API | 3-5 segundos por requisição |

### Observações

1. **APIs públicas**: As APIs usadas (AwesomeAPI, CepRua, ViaCEP) são gratuitas mas podem ter limites diários.

2. **Endereços incomuns**: Alguns endereços podem não ser encontrados pelas APIs, especialmente em áreas rurais ou novos loteamentos.

3. **Tempo de processamento**: O tempo total depende da quantidade de registros e da velocidade das APIs. Para 1590 registros, o processamento leva aproximadamente 5-10 minutos.

---

## Estrutura de Arquivos

```
CEP_paperclipe/
├── src/
│   ├── app.py                  # Interface Streamlit
│   ├── processador.py          # Módulo de processamento
│   ├── buscar_corrigido.py     # Script de busca
│   ├── gerar_final.py         # Script de geração CSV
│   ├── ideia.MD                # Base de conhecimento
│   └── requirements.txt        # Dependências
├── Exemplos de planilha/
│   ├── 2026 04 07 - prova de vida pendentes MARÇO.xlsx
│   └── LISTA_PARA_ENVIO.csv
├── README.md
└── .gitignore
```

---

## Dependências

```
pandas
requests
openpyxl
streamlit
```

Para instalar:
```bash
pip install pandas requests openpyxl streamlit
```

---

## Contato

Desenvolvedor responsável pelo projeto.

---

## Histórico de Versões

| Versão | Data | Descrição |
|--------|------|-----------|
| 1.0 | 2026-05-06 | Versão inicial com validação 3 camadas |
| 2.0 | 2026-05-06 | Adicionado busca por endereço completo (ENDERECO + COMPLEMENTO + BAIRRO + MUNICIPIO + UF) |
| 2.1 | 2026-05-06 | Correção na API AwesomeAPI (formato JSON com chave "results") |