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

## Contato

Desenvolvedor responsável pelo projeto.