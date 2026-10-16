# Ambiente: escrever em modelos Excel neste computador

Leia quando for gravar o arquivo. São armadilhas já encontradas.

## Ferramentas
- Leitura: Python + openpyxl (`data_only=True` para valores em cache; `read_only=True` para planilhas grandes do RI, que travam em carga completa).
- Escrita: Excel via COM a partir do PowerShell (`New-Object -ComObject Excel.Application`). openpyxl não deve salvar o modelo: perde desenhos e deixa fórmulas sem valor em cache.
- PDF: extrair texto com `pypdf`; a leitura direta de PDF pode não estar disponível. Tabelas de release costumam ser imagem; o ITR tem o balanço em texto.

## Fluxo seguro
1. Copie o modelo e as fontes para a pasta temporária da sessão.
2. Gere um plano em JSON com Python (endereço → valor ou fórmula) e rode os testes de soma ali, antes de abrir o Excel.
3. Aplique o plano com um `.ps1` e salve.
4. Verifique com openpyxl comparando com o arquivo de partida.
5. Só então copie para a pasta do OneDrive, com nome novo. Abrir o Excel direto no OneDrive pode dar `RPC_E_CALL_REJECTED`.

## Armadilhas do Excel em pt-BR
- `Range.NumberFormat` é localizado: `'#.##0'`, `'0,00%'`, `'#.##0,0'`. Na dúvida, copie o `NumberFormat` de uma célula existente. `.Formula` continua em inglês.
- Para dar a uma coluna o formato de histórico: `Range("AE14:AE304").Copy()` e `PasteSpecial(-4122)` (só formatos) no destino.
- `$xl.CutCopyMode = 0` falha no PowerShell; simplesmente não use.

## Armadilhas do PowerShell
- Nomes de variáveis não diferenciam maiúsculas: `$r` sobrescreve `$R`.
- Scripts longos vão em arquivo `.ps1` executado com `powershell -NoProfile -ExecutionPolicy Bypass -File`. Um `.ps1` salvo sem BOM corrompe acentos: mantenha textos acentuados no JSON (lido com `-Encoding UTF8`) e caminhos com acento fora do `.ps1`.
- `ConvertFrom-Json` devolve inteiros como Int64: converta para `[double]` antes de escrever em célula.
- Use o caminho completo do usuário (com acento), não o caminho curto 8.3.

## Grades de colunas já vistas
- Modelo padrão de empresas: `Model` começa em D = 1T20, com coluna anual a cada cinco (H, M, R...). `Inputs (2)` começa em D = 1T21. Coluna do `Model` = coluna da `Inputs (2)` + 5.
- Base de bancos: `Inputs` bruta sem colunas anuais (E = 1T20); `Inputs (2)` e `Model` com colunas anuais (E = 1T20, I = FY20).
- Sempre confirme lendo os cabeçalhos; não presuma.
