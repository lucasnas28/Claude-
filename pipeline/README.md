# Pipeline dos modelos 2T26

Gera os arquivos `Modelo_<Empresa>_2T26.xlsx` a partir da máscara `Modelo_2.0_proprio_Padrao_AJUSTADO-2028_corrigido.xlsx` (raiz do repositório).

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
