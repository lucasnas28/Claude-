# Pipeline dos modelos 2T26

Gera os arquivos `Modelo_<Empresa>_2T26.xlsx` na **máscara vigente**, `Modelo_Base_credito_research_AJUSTADO_ratio.xlsx` (raiz do repositório). Regra: sempre a última versão da máscara que o usuário mandou.

O modelo é montado primeiro na máscara antiga, `Modelo_2.0_proprio_Padrao_AJUSTADO-2028_corrigido.xlsx`, cujo layout o `build_gen` conhece, e convertido no fim pelo `comum/to_new_mask.py`.

## Requisitos

- Python 3.10+ com `openpyxl` e `lxml`
- LibreOffice (`soffice`) com a ponte Python `uno`, usada em `comum/recalc.py` para recalcular as fórmulas e gravar o cache de valores
- `curl`
- Variável `SEC_UA` com nome e e-mail de contato. A SEC exige isso no cabeçalho User-Agent.

```bash
export SEC_UA="Seu Nome seu@email.com"
```

## Rodar uma companhia americana (dados da SEC)

```bash
bash pipeline/rodar_empresa.sh orcl 40 Oracle
```

- `orcl` é a pasta em `pipeline/empresas/`, que tem um `cfg.py` e um `spec.py`.
- `40` é o número de linhas reservadas no Model para o bloco de drivers. Se o `gen_drivers` acusar falta de linhas, aumente.
- `Oracle` vira o nome do arquivo: `Modelo_Oracle_2T26.xlsx`, gravado na raiz.

Os downloads e os arquivos intermediários ficam em `dados/<ticker>/`, que está no `.gitignore`. Para mudar o local, use `PIPE_DATA=/outro/lugar`.

O script termina mostrando a linha `CHECKS` (ativo, passivo + PL e DRE) e `hist errors`. Só entregue o modelo se os checks forem todos zero e não houver erro de fórmula no histórico.

## Etapas (o que o `rodar_empresa.sh` faz)

| # | Script | Faz |
|---|---|---|
| 1 | `sec/dl.sh CIK dir` | Baixa companyfacts, submissions e a instância XBRL de cada 10-Q/10-K desde 2019 |
| 2 | `sec/build_gen.py cfg.py MASCARA saida_a.xlsx` | Monta a aba Inputs a partir do XBRL: DRE e DFC trimestralizadas, balanço de face, códigos e sinais |
| 3 | `sec/post.sh prefixo` | Marca 1T26 e 2T26 como realizados, recalcula e envolve em IFERROR as linhas históricas com #DIV/0! |
| 4 | `comum/shift_rows.py in out Model 18 N` | Abre N linhas no Model, abaixo da receita líquida (linha 18), para os drivers |
| 5 | `sec/gen_drivers.py in out layout.json 18 N spec.py` | Escreve os drivers: códigos 45+ na Inputs e o bloco no Model |
| 6 | `comum/recalc.py` + `comum/inject.py` | Recalcula no LibreOffice e injeta os valores em cache no arquivo original, preservando a formatação |
| 7 | `comum/fix_panes.py` | Garante o painel congelado do Model na linha 11 |
| 8 | `comum/converter_mascara.sh MASCARA SAIDA arq:US$` | Passa para a máscara vigente (`to_new_mask.py`), recalcula e confere com `check_modelo.py` |

## Máscara vigente e conversão

`comum/to_new_mask.py ANTIGO.xlsx MASCARA.xlsx SAIDA.xlsx MOEDA ULTIMO_TRI` faz a conversão:

- **Inputs:** linha com um código e sinal +1 mantém o código. Linha com sinal −1 ou vários códigos vira uma linha bruta sem código e uma linha de lançamento por código (`=sinal × bruta`). Valores em mil.
- **Inputs (2):** flag de realizado na linha 3; código 44 = lucro reportado (base do check de DRE); códigos 45+ = drivers.
- **Model:** bloco de drivers copiado para a linha 18; linha 4 = lucro reportado.
- **Control:** companhia, ticker, moeda, unidade e último trimestre realizado.
- 2020 fica de fora: a grade da máscara começa no 1T21. Variações de saldo do 1T21 e de 2021 (capex, capital de giro, DFC indireta) ficam vazias por falta do saldo de 2020.

Correções que o conversor aplica sobre a máscara (registradas no `.log` de cada arquivo):

- dívida LP e capital + APIC de 1T26 e 2T26 sem ligação com a Inputs (2);
- linha 196 com união de células `(CFO,CFI)`, que dava #VALUE!;
- 4T25, 1T26, 2T26 e o anual de 2025 sem a fórmula histórica (capex, imobilizado, intangível, DSO/DIO/DPO etc.): repete o padrão do 3T25;
- linha 150 (intangível do roll-forward) apontava para o goodwill no histórico;
- linha 189 (dívida LP no cronograma) com zeros digitados nas projeções, o que zerava a dívida projetada;
- data de 31/12/2020 em H5, base do DSO do 1T21;
- IFERROR nas linhas 228 a 231 e nas fórmulas de crescimento da Valuation.

Para reconverter tudo quando chegar máscara nova: `bash pipeline/comum/converter_mascara.sh NOVA.xlsx saida/ Modelo_X_2T26.xlsx:US$ Modelo_Y_2T26.xlsx:R$ ...`. Os arquivos de entrada precisam estar na máscara antiga (com fórmulas). Para partir de arquivos já na máscara vigente, o conversor precisa ser adaptado.

## Nova companhia

1. Ache o CIK em sec.gov/edgar.
2. Crie `pipeline/empresas/<ticker>/cfg.py`:
   ```python
   NAME = 'Empresa X Inc.'; TICKER = 'XXX'; CIK = '0001234567'
   DL = __import__('os').environ.get('PIPE_DATA', 'dados') + '/<ticker>/dl'
   ```
   Opções do `cfg.py`, todas opcionais. Os exemplos estão nas outras empresas.
   - `REV`, `COGS`, `OPINC`, `PRETAX`, `TAX`, `NI`, `MI`, `INTEXP`, `DA`, `DIV`: lista de tags us-gaap em ordem de preferência.
   - `COGS_SUM` (Oracle), `DA_SUM` (Broadcom, Oracle) e `DA_FUNC`/`DA_DESC` (AMD): somas ou regras próprias.
   - `OP_FUNC`/`OP_DESC` (Bunge): lucro operacional derivado quando a companhia não divulga.
   - `DA_ADDBACK = 2`: devolve a D&A ao custo em vez da despesa (Freeport, Bunge).
   - `BSMAP`, `UNSKIP`: mapeamento extra de contas do balanço (preferenciais no PL, por exemplo).
   - `FACE_OVR`: balanço de um trimestre lido de outro filing, em geral a coluna comparativa (Constellation 4T21, CoreWeave 4T24).
   - `MONTHMAP`: mês de fechamento → trimestre-calendário (Costco).
   - `COGS_DIMS` (GE): custo só aberto por produto e serviço.
3. Liste os membros dimensionais para escolher os segmentos:
   ```bash
   python3 -I pipeline/sec/mem.py dados/<ticker>/dl
   ```
4. Crie o `spec.py` com `segspec.make(...)`. Os modelos mais simples são `lmt`, `uber` e `meli`. Para linhas extras, como RPO, veja `orcl` e `crwv`; para regras próprias, `fcx` e `bg`.
5. Rode `rodar_empresa.sh` e confira os segmentos contra o release. A soma deve fechar com a receita, ou a diferença deve ter explicação escrita na nota do bloco.

## Particularidades já resolvidas no código (não desfazer)

- **Trimestralização:** usa o 3M quando existe; senão, YTD menos o YTD anterior, sempre um par consistente (os dois reapresentados ou os dois originais). Ver `Co._pick` em `sec/sx.py`.
- **Contexto de 3M no 10-K com o valor do ano** (erro da Oracle): é ignorado, e o 4T é calculado como anual − 9M.
- **Ano fiscal fora do calendário:** cada fechamento vai para o trimestre-calendário mais próximo, ou segue o `MONTHMAP`.
- **Balanço de face:**
  - lido na ordem do documento até `LiabilitiesAndStockholdersEquity`;
  - subtotais detectados e excluídos;
  - resíduos por seção vão para outros ativos e outros passivos;
  - diferenças de arredondamento menores que US$ 50 mi vão para uma linha própria.
- **Painel congelado:** o `shift_rows.py` só desloca o congelamento quando ele fica abaixo da linha de inserção. O bug antigo levava o painel para a linha ~350.

## Companhias brasileiras e bancos

Os modelos brasileiros (CVM ou planilha de fundamentos da companhia) e os de bancos foram montados com scripts próprios, guardados em `pipeline/legado/`. Eles servem de referência, não de pipeline genérico:

- `legado/auren`: leitor de dados abertos da CVM (ITR/DFP).
- `legado/brasil`: construtores usados para Alupar, Fleury, bancos etc.

Também rodam com `PIPE_DATA` e esperam os arquivos de fonte que estão na raiz do repositório.
