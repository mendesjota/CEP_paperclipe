# CEP Paperclipe - Base de Conhecimento

## Visão Geral do Projeto

Sistema para processar planilhas de prova de vida, buscar CEPs automaticamente e gerar arquivos prontos para envio aos Correios.

# Contexto e Objetivo
Atue como um Engenheiro de Dados Sênior especialista em Python e Pandas. Preciso de um script `app.py` que leia uma planilha, valide os CEPs informados, busque correções para os CEPs inválidos usando o endereço, e por fim, exporte os resultados em 4 planilhas distintas.

# Estrutura de Dados Esperada
- As colunas de referência serão `CEP` (que contém o cep atual) e `Endereco` (que contém Rua, Cidade, Estado).


# Requisitos Técnicos e Boas Práticas
- Utilize as bibliotecas `pandas` e `requests`.
- Adicione logs ou `prints` no terminal (ex: "Processando linha 15 de 1000...") para que eu saiba que o programa não travou.
- Entregue o código completo, documentado e pronto para rodar.

---

## Contexto do Cliente


- **Órgão**: governo (contrato com Correios)
- **Volume**: 1.000 a 10.000 cartas por vez
- **Problema atual**: erros de CEP constantes, necessidade de dividir em lotes de 300

---

## Teste

Escreva os testes, todo código antes de ser apresentado deve passar por testes

todos os aquivos deveram ter testes e passar por testes para não subir nada em produção defeituoso

Crie uma pasta com todos os testes, garanta que o código supra todos os requisitos.

---
## Formato das Planilhas

### Entrada: Prova de Vida (Excel)
| CPF | MATRICULA | NOME | DATA NASCIMENTO | ENDERECO | COMPLEMENTO DO ENDERECO | BAIRRO | MUNICIPIO | UF | CEP |
|-----|----------|------|----------------|----------|-------------------------|-------|-----------|-----|-----|

### Saída: Lista para Envio (CSV) - Formato Correios
| Prefixo (coluna A) | Nome (coluna B) | --- (coluna C,D, E) "Vazios" | | CPF (Coluna F) | CEP (coluna G) | ... (coluna H) "vazio"  | Endereço (coluna I) | Bairro (coluna J) | Cidade (coluna K) | UF (coluna L) | Flag (coluna M) | Complemento (coluna N) | Número (coluna O)|
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
| Lote Correios | 300 objetos |
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

1. Upload da planilha Excel (.xlsx)
2. Selecionar colunas (Nome, Endereço, Cidade, UF, CEP, Complemento)
3. Clicar "Processar Planilha"
4. Download do CSV (formato Correios)
5. Se >500 registros, download automático de lotes

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

## Contato

Desenvolvedor responsável pelo projeto.

---

## Fluxo

**Passo 1: Limpeza e Validação Inicial**
- Crie uma função que remova pontuações do CEP atual e verifique se ele possui exatamente 8 dígitos numéricos.
- Faça uma validação via API (ex: ViaCEP) para confirmar se esse CEP realmente existe.

**Passo 2: Busca e Correção (Fallback)**
- Caso o CEP seja inválido (formato errado) ou não exista na API, o código deve pegar a string da coluna `Endereco`.
- Crie uma função que faça uma busca pelo logradouro na API do ViaCEP (ou Nominatim) para tentar encontrar o CEP correto.
- **CRÍTICO:** Implemente um `time.sleep(1.5)` entre as chamadas de API para evitar bloqueios de IP (Rate Limit). Adicione blocos `try-except` para que o código não pare caso uma requisição falhe.

**Passo 3: Geração das 4 Planilhas de Saída**
Com base no processamento, divida o DataFrame original e exporte para Excel (.xlsx) os seguintes 4 arquivos:
1. `1_ceps_validos.xlsx`: Apenas as linhas onde o CEP original já estava correto.
2. `2_ceps_corrigidos.xlsx`: Apenas as linhas onde o CEP estava errado, mas a API conseguiu encontrar o correto através do Endereço.
3. `3_ceps_nao_encontrados.xlsx`: Apenas as linhas onde o CEP estava errado e a API NÃO encontrou uma correspondência pelo Endereço.
4. `4_base_consolidada.xlsx`: A base original inteira, adicionando duas novas colunas: `Novo_CEP` e `Status_Processamento` (Válido, Corrigido, Não Encontrado).

# Requisitos Técnicos e Boas Práticas
- Utilize as bibliotecas `pandas` e `requests`.
- Adicione logs ou `prints` no terminal (ex: "Processando linha 15 de 1000...") para que eu saiba que o programa não travou.
- Entregue o código completo, documentado e pronto para rodar.

---

## Comentado 

Sempre que possivel comente os códigos em portugues do Brasil de maneira didatica visando ajudar na manutenção.


---

## Endereço

Algumas pessoas escreve o endereço correto na coluna "COMPLEMENTO DO ENDERECO", na hora de fazer uma requisição por endereço para encontrar o CEP, deve ser feito uma contatenação de colunas ENDERECO + COMPLEMENTO DO ENDERECO + BAIRRO + MUNICIPIO + UF

---

## instrução

Preciso garantir que os CEPs sejam certos de funcionais 

verifique se os CEPs estão ativos e funcionando 

se não tiver ativo use as informações das colunas ENDERECO + COMPLEMENTO DO ENDERECO + BAIRRO + MUNICIPIO + UF para achar o cep, se preciso for tire palavras dupicadas, faça varias tentativas, procure o cep de maneira aproximada, usando apenas o bairro, porém o cep deve ser valido.

***priorize rapidez***

Acredito que apenas os códigos que não retornarem um endereço valido devem passa novamente pela API