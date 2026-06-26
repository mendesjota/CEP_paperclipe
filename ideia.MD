# CEP Paperclipe - Base de Conhecimento

## Visão Geral do Projeto

Sistema para processar planilhas de prova de vida, buscar CEPs automaticamente e gerar arquivos prontos para envio aos Correios.

---

## Contexto do Cliente

- **Órgão**: governo (contrato com Correios)
- **Volume**: 1.000 a 10.000 cartas por vez
- **Problema atual**: erros de CEP constantes, necessidade de dividir em lotes de 500

---

## Formato das Planilhas

### Entrada: Prova de Vida (Excel)
| CPF | MATRICULA | NOME | DATA NASCIMENTO | ENDERECO | COMPLEMENTO DO ENDERECO | BAIRRO | MUNICIPIO | UF | CEP |
|-----|----------|------|----------------|----------|-------------------------|-------|-----------|-----|-----|

### Saída: Lista para Envio (CSV) - Formato Correios
| Prefixo | Nome | ... | CPF | CEP | ... | Endereço | Bairro | Cidade | UF | Flag | Complemento | Número |
|--------|------|-----|-----|-----|-----|---------|--------|--------|---|---|------|------------|--------|

---

## 4 Planilhas de Saída

### 1_ceps_validos.xlsx - Dados ORIGINAIS (modelo Correios)
- Conteúdo: Todos os registros com CEP original da planilha
- Formato: Excel com colunas: PREFIXO, NOME, CPF, CEP, LOGRADOURO, ENDERECO, BAIRRO, MUNICIPIO, UF, FLAG, COMPLEMENTO, NUMERO
- Todas as células em formato TEXTO
- Uso: Lista de envio com dados originais

### 2_ceps_corrigidos.xlsx - Dados CORRIGIDOS (modelo Correios)
- Conteúdo: CEPs válidos + CEPs sugeridos (para não encontrados)
- Regra: Se CEP validado → usa CEP válido; Se não validado → usa CEP sugerido
- Formato: Excel com colunas: PREFIXO, NOME, CPF, CEP, LOGRADOURO, ENDERECO, BAIRRO, MUNICIPIO, UF, FLAG, COMPLEMENTO, NUMERO
- Todas as células em formato TEXTO
- Uso: Lista de envio com CEPs corrigidos/sugeridos

### 3_ceps_nao_encontrados.xlsx - Log de ERROS
- Conteúdo: Apenas registros onde CEP não foi validado
- Colunas:
  - Nome, CPF, Endereco, Complemento, Bairro, Municipio, UF (dados originais)
  - CEP_ORIGINAL (CEP da planilha)
  - CEP_INPUTADO (CEP verificado na API)
  - CEP_SUGERIDO (CEP encontrado pelo fallback)
- Uso: Lista para revisão manual

### 4_base_consolidada.xlsx - Dados ORIGINAIS + MODIFICADOS
- Conteúdo: Todos os registros da planilha original + colunas de resultado
- Colunas adicionadas:
  - CEP_CORRIGIDO (CEP válido ou sugerido)
  - STATUS (Válido / Não encontrado / Formato inválido)
  - MODIFICADO (Sim / Não)
  - CEP_SUGERIDO (para não encontrados)
- Uso: Análise completa de resultados

---

## APIs de Busca CEP (3 em cascata)

| # | API | Endpoint | Tipo |
|---|-----|----------|------|
| 1 | AwesomeAPI | `cep.awesomeapi.com.br/search?q={query}` | Texto livre |
| 2 | CepRua | `ceprua.com.br/api/buscar?q={query}` | Texto livre |
| 3 | ViaCEP | `viacep.com.br/ws/{UF}/{Cidade}/{Logradouro}/json/` | Parâmetros separados |

### Abreviações Normalizadas
```
R → RUA
AV → AVENIDA
QD/QN → QUADRA
CS → CASA
AP → APARTAMENTO
BL → BLOCO
LT → LOTE
```

---

## Limites Técnicos

| Item | Limite |
|------|-------|
| Lote Correios | 500 objetos |
| Delay entre requisições | 0.3s (padrão) |
| Rate limit | 10 req/segundo |

---

## Como Rodar

```bash
cd src
.venv\Scripts\streamlit.exe run app.py
```

Acesse: http://localhost:8501

---

## Fluxo de Uso

### Passo 1: Input
- Usuário acessa a página Streamlit
- Faz upload da planilha Excel (Entrada: Prova de Vida)

### Passo 2: Busca de CEP
- Código faz requisição para API com os dados de endereço
- API retorna o CEP correto
- Se CEP já existir na planilha, substitui pelo novo CEP da API

### Passo 3: Output
- Gera planilha CSV (Lista para Envio) - Formato Correios
- Se API retornar erro, mantém CEP original
- Se >500 registros, divide automaticamente em lotes

---

## Estrutura de Arquivos Atual

```
CEP_paperclipe/
├── src/
│   └── app.py        ← Único arquivo com tudo!
├── README.md
├── .gitignore
└── "Exemplos de planilha/"  ← Dados sensíveis (ignorado)
```

---

## Credenciais Correios (para uso futuro)

- **Ambiente**: homologação (testes) ou produção
- **Dados necessários**: usuário, senha componente, cartão postagem
- **Serviço**: API Pré-Postagem (código 86720)

---

## ⚠️ Regra Importante: Tratar como TEXTO

**Todos os dados devem ser tratados como string (texto)** para evitar perda de zeros à esquerda.

- **CEP** pode começar com 0 (ex: 01234567)
- **CPF** pode começar com 0 (ex: 01234567890)
- Ao ler Excel: usar `dtype=str`
- Ao gerar Excel: forçar tipo string com `.astype(str)`
- Ao salvar Excel: usar openpyxl com células forçadas como texto (não general/number)
- **Todas as colunas de TODAS as planilhas Excel devem ter formato TEXTO (@)** para evitar perda de zeros à esquerda em CEPs e CPFs
  - Usar `cell.number_format = numbers.FORMAT_TEXT` em cada célula
- **CSV (arquivos .csv)**: TODOS os campos devem ter aspas duplas para forçar interpretação como texto
  - Formato: `SR(A);"NOME";"CPF";"CEP";"";"ENDERECO";"BAIRRO";"MUNICIPIO";"UF";"N";"COMPLEMENTO";"0"`
  - Isso evita que CEP e CPF sejam interpretados como números pelo Excel
- **Normalização de CPF e CEP**:
  - CPF deve ter sempre 11 dígitos: usar `zfill(11)` para completar com zeros à esquerda
  - CEP deve ter sempre 8 dígitos: usar `zfill(8)` para completar com zeros à esquerda
  - Exemplo: CPF "175340110" → "00175340110" | CEP "71065907" → "071065907"
- **Botões de download**: usar session_state para persistir os dados processados - os botões devem continuar visíveis após o clique para permitir baixar todas as planilhas
- **Fallback para CEPs não encontrados**: para CEPs com status "Não encontrado", fazer busca adicional usando:
  1. ENDERECO + COMPLEMENTO DO ENDERECO + BAIRRO + MUNICIPIO + UF
  2. Se não encontrar: ENDERECO + BAIRRO + MUNICIPIO + UF
  3. Se ainda não encontrar: BAIRRO + MUNICIPIO + UF (CEP geral do bairro)
  - Adicionar coluna "CEP_SUGERIDO" na planilha consolidada

---

## Contato

Desenvolvedor responsável pelo projeto.