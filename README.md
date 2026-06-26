# CEP Paperclipe — Validação e Correção de CEPs para envio aos Correios

Aplicação web (Streamlit) que recebe a planilha de prova de vida, **valida e corrige os CEPs na base oficial dos Correios** e gera os arquivos prontos para importação no sistema do Correios — separando automaticamente o que precisa de **revisão manual**.

---

## Como rodar (passo a passo no terminal)

Abra o **PowerShell** e execute:

```powershell
cd "c:\Users\jose.junior\Desktop\Python\Projetos\CEP_Paperclipe\CEP_paperclipe\src"
.\.venv\Scripts\streamlit.exe run app.py
```

O navegador abre sozinho em **http://localhost:8501**.
Para **parar** o programa: `Ctrl + C` no terminal.

> Se a porta 8501 já estiver em uso, force outra (ou a mesma após parar a anterior):
> ```powershell
> .\.venv\Scripts\streamlit.exe run app.py --server.port 8502
> ```

Primeira vez na máquina (instalar dependências):

```powershell
cd "c:\Users\jose.junior\Desktop\Python\Projetos\CEP_Paperclipe\CEP_paperclipe\src"
.\.venv\Scripts\pip.exe install -r requirements.txt
```

---

## Como usar (passo a passo na tela)

1. Faça **upload** da planilha Excel de prova de vida (`.xlsx`).
2. Clique em **Processar Planilha**.
   - O sistema consulta a base oficial (ViaCEP/DNE) para cada CEP novo, então um lote grande leva alguns segundos por CEP inédito. Há cache: CEPs repetidos são instantâneos.
3. Veja os números: **Total · Válidos · Corrigidos · Revisão manual**.
4. Baixe os arquivos (ver abaixo).

### Planilha de entrada (colunas esperadas)

`CPF | MATRICULA | NOME | DATA NASCIMENTO | ENDERECO | COMPLEMENTO DO ENDERECO | BAIRRO | MUNICIPIO | UF | CEP`

(Há um botão **Baixar Modelo** na própria tela.)

---

## O que o sistema faz com cada CEP

Para cada registro, o CEP é classificado em um de três status:

| Status | O que significa | Vai para o envio? |
|--------|------------------|-------------------|
| **Válido** | O CEP informado existe na base oficial e a UF confere. | Sim |
| **Corrigido** | O CEP estava errado/ausente/com UF divergente, e o sistema encontrou o **CEP correto pelo endereço** (busca por logradouro no ViaCEP), confirmando a UF. | Sim |
| **Revisão** | Não foi possível corrigir com segurança (CEP inexistente sem endereço resolvível) **ou o CPF é inválido**. | Não — vai para revisão manual |

**Regra de segurança:** nenhum CEP entra no arquivo de envio sem existir na base oficial e ter a UF confirmada. O sistema **não “inventa” CEP**.

---

## Arquivos gerados (downloads)

| Arquivo | Conteúdo | Uso |
|---------|----------|-----|
| **envio_correios.csv** | Apenas registros **Válidos + Corrigidos**, no formato de importação do Correios (`SR(A);...`). | É o arquivo que vai para o Correios. |
| **Blocos de 300** | O mesmo envio dividido em lotes de 300 linhas. | Quando o sistema do Correios exige lotes menores. |
| **revisao_manual.xlsx** | Registros que caíram em **Revisão**, com a coluna **MOTIVO** explicando o problema. | Lista para a equipe tratar manualmente. |
| **base_consolidada.xlsx** | **Todos** os registros com `CEP_INFORMADO`, `CEP_FINAL`, `STATUS` e `MOTIVO`. | Auditoria/conferência completa. |

### Como tratar a `revisao_manual.xlsx`

A coluna **MOTIVO** diz o que corrigir:

- **`CPF/CNPJ com N dígitos (esperado 11 ou 14)`** → o CPF está errado na origem. **Só o cadastro (SIAPE/sistema) corrige** — nenhum serviço de CEP resolve isso. Acerte o CPF na planilha e reprocesse.
- **`CEP não encontrado na base oficial`** → confira o endereço (logradouro/bairro). Com o endereço certo, ao reprocessar o sistema costuma achar o CEP.
- **`CEP pertence a XX, diferente da UF informada (YY)`** → UF ou CEP digitado errado; ajuste e reprocesse.
- **`CEP com formato inválido (≠ 8 dígitos)`** → CEP truncado/colado errado; corrija e reprocesse.

Fluxo recomendado: corrija os itens da revisão na planilha original → reprocesse → a revisão diminui a cada rodada.

---

## Como funciona por dentro (resumo técnico)

1. **Normalização**: CPF para 11 dígitos (sem truncar — se vier com mais, é marcado inválido); CEP para 8 dígitos (completa zero à esquerda; não trunca).
2. **Validação na base oficial**: consulta o CEP no **ViaCEP** (espelha o DNE dos Correios) — confere existência, UF e cidade.
3. **Correção por endereço** (quando o CEP falha): busca **endereço → CEP** por logradouro no ViaCEP (retorna o CEP exato), com desambiguação por bairro/bloco; reserva: busca textual no CepRua. Sempre confirmando a UF.
4. **Saída**: monta os arquivos e separa a revisão manual.

> A função de consulta é **desacoplada**: no futuro dá para trocar o ViaCEP pela **API CEP oficial dos Correios** (exige token do contrato) sem mexer no resto.

---

## Dependências

```
streamlit
requests
pandas
openpyxl
```

(instaladas no ambiente virtual `src/.venv`)

---

## Observações importantes

- **Tudo é tratado como TEXTO** para não perder zeros à esquerda de CPF e CEP.
- O arquivo de envio usa o formato `SR(A);NOME;;;;CPF;CEP;;ENDERECO;BAIRRO;MUNICIPIO;UF;N;COMPLEMENTO;0`.
- Processamento depende de internet (consulta o ViaCEP).
- Dados sensíveis (planilhas reais) **não** ficam no repositório (ver `.gitignore`).
