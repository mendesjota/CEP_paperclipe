# AGENTS.md

Toda comunicação, código, comentários e docs em **pt-BR**.

## Estrutura

O git repo está em `CEP_paperclipe/` (subpasta aninhada). O projeto executável está em `CEP_paperclipe/src/`. O venv compromissado está em `CEP_paperclipe/src/.venv`.

## Comandos (PowerShell, sempre via .venv)

```powershell
cd CEP_paperclipe/src
.venv\Scripts\streamlit.exe run app.py
.venv\Scripts\python.exe tests\test_funcoes.py   # unitários, sem rede
.venv\Scripts\python.exe tests\test_api.py       # integração, APIs reais
.venv\Scripts\pip.exe install -r requirements.txt
```

## Armadilhas que um agente perderia sem ajuda

- **Testes copiam funções inline.** `test_funcoes.py` e `test_api.py` têm cópias literais das funções de `app.py` (porque `app.py` executa Streamlit na importação). Alterou função no `app.py`? Atualize a cópia no teste ou ele testará lógica defasada em silêncio.
- **Testes NÃO são pytest.** São scripts com `if __name__ == "__main__"` que printam `[OK]/[FALHOU]`. O `.vscode/settings.json` habilita pytest, mas ignore — pytest reportará "passed" sem asserções reais.
- **READme.md e ideia.MD são parciais/desatualizados.** Descrevem design idealizado (3 APIs, 4 planilhas .xlsx) que `app.py` (v6.2) não implementa. Confie no `app.py` como fonte da verdade.
- **Tudo é texto.** CPF/CEP têm zeros à esquerda. Leia Excel com `dtype=str`, force `.astype(str)`, normalize com `zfill(11)`/`zfill(8)`. Exporte Excel com `number_format = numbers.FORMAT_TEXT`.
- **CSV de envio é formato Correios.** Delimitado por `;`, prefixado com BOM (`\ufeff`), linha fixa `SR(A);NOME;;;;CPF;CEP;;ENDERECO;BAIRRO;MUNICIPIO;UF;N;COMPLEMENTO;0`. Complemento truncado em 40 caracteres.
- **APIs reais usadas:** ViaCEP (primário, busca CEP + logradouro) + CepRua (fallback textual) + BrasilAPI (backup). Com rate limiting global (0.15s entre chamadas) e cooldown por provedor (60s após 3 falhas).
- **Heurísticas de endereço em `buscar_cep_fallback` e `corrigir_cep_por_endereco` são tuned para Brasília/DF** (SHIN/SHIS/SQN/SQS, quadras, etc.). Alterá-las afeta diretamente a taxa de acerto.
- **Scripts avulsos** (`teste.py`, `dados.py`, `consolidar_erros.py`) têm paths absolutos Windows hard-coded (`i:\...`, `Exemplos de planilha\...`). Não executam fora da máquina de origem sem edição.

## Pipeline do app (em ordem)

1. Upload .xlsx (lido com `dtype=str`, colunas em `CABEÇALHO_ESPERADO`)
2. `prevalidar_ceps()` consulta ViaCEP/BrasilAPI para CEPs únicos (cache em `_cache_dne`)
3. `avaliar_registro()` classifica cada linha: **Válido** (CEP+UF ok) | **Corrigido** (CEP via endereço) | **Revisão** (não corrigível ou CPF inválido)
4. Gera 2 CSVs (envio + blocos de 300) + 2 .xlsx (revisão_manual, base_consolidada) em `st.session_state`

## Saídas geradas

| Arquivo | Conteúdo |
|---------|----------|
| `envio_correios.csv` | Válidos + Corrigidos, formato Correios |
| Blocos de 300 linhas | Mesmo envio dividido |
| `revisao_manual.xlsx` | Registros em Revisão com coluna MOTIVO |
| `base_consolidada.xlsx` | Todos os registros (auditoria) |
