# Passagem de trabalho: modelos de valuation 2T26

Situação em 16/10/2026. Use este documento para retomar o trabalho em outra ferramenta (Cursor, outra sessão do Claude).

## O que existe

- **Máscara vigente (regra: sempre a última versão):** `Modelo_Base_credito_research_AJUSTADO_ratio.xlsx`, com as abas Control → Inputs → Inputs (2) → Model → Valuation e grade de 1T21 a 4T26. Todos os modelos não bancários estão nela. As correções que o conversor aplica sobre a máscara estão em `pipeline/README.md`.
- **Máscara de montagem:** `Modelo_2.0_proprio_Padrao_AJUSTADO-2028_corrigido.xlsx`, usada só dentro do pipeline (correções em `prompts/Correcoes_mascara_padrao.md`).
- **Modelos prontos:** `Modelo_<Empresa>_2T26.xlsx` na raiz, com histórico de 1T21 a 2T26 e 1T26 e 2T26 marcados como realizados. A Auren foi montada direto na máscara vigente (a máscara saiu dela). Os bancos seguem na base própria.
- **Pipeline:** `pipeline/`, com instruções em `pipeline/README.md`.
- **Metodologia** (página para compartilhar): https://claude.ai/artifact/3NFU22unSonQBUYgKVjsEA
- **Habilidades** com as regras de preenchimento, atualização e projeção: `docs/skills/`.

## Status por companhia

### Feitos (SEC XBRL, EUA)

| Companhia | Ressalva |
|---|---|
| Alphabet, Microsoft, Meta, NVIDIA, Apple, AMD, Visa, Mastercard, Coca-Cola, Eli Lilly, Lockheed Martin, Caterpillar, Uber, UnitedHealth, MercadoLibre | Sem ressalva relevante. Na UnitedHealth, a soma dos segmentos inclui vendas intersegmento. |
| Tesla | O 10-Q do 1T23 trocou os segmentos Automotive e Energy; corrigido no spec |
| Micron, Broadcom, Deere, Walmart | Ano fiscal fora do calendário |
| Costco | Trimestres de 12 e 16 semanas. A partir do ano fiscal 2025, o lucro antes do IR é etiquetado como `IncomeLossAttributableToParent` (já no cfg) |
| Exxon, Chevron | Lucro operacional derivado (receita − custos totais + juros). Exxon: a receita inclui equivalência. Chevron: segmentos só até 3T24 |
| GE Aerospace | O perímetro muda com as cisões de HealthCare e Vernova |
| Netflix | A soma regional difere em até US$ 46 mi em alguns 4T; assinantes só até 4T24 |
| Blackstone | A receita pode ser negativa |
| NextEra | Segmentos medidos por lucro líquido |
| Constellation | A série começa em 2021; o balanço do 4T21 vem do comparativo do 10-Q de mar/22 |
| Freeport | Receita do segmento = terceiros + intersegmento (até 3T25 vinham separadas). D&A devolvida ao custo |
| Bunge | Bunge Limited (CIK 1144519) até 3T23 e Bunge Global (CIK 1996862) depois. Para rodar de novo, baixe os dois CIKs e junte `accns.txt` e `inst/` numa só pasta `dl`. A Viterra quebra a comparabilidade a partir do 3T25 |
| Oracle | Ano fiscal fecha em maio; o 10-K etiqueta o 4T fiscal com o valor do ano (tratado no `sx.py`). Custo = cloud e software + hardware + serviços. Preferenciais série D no PL |
| CoreWeave | IPO em mar/25: a série começa no 1T24 e o balanço no 4T24 |
| CHS Inc. | Cooperativa agrícola (preferenciais CHSCP na Nasdaq). Ano fiscal fecha em agosto. Balanço de 1T20 e 2T20 vazio (instâncias em ordem alfabética); balanço de agosto vem do comparativo do 10-Q de novembro. Segmentos: Energy, Ag, Nitrogen e Corporate até FY25; Energy, Grains, Agronomy e demais a partir de 4T25 |

### Feitos (Brasil)

- **Fonte CVM ou planilha de fundamentos:** Alupar, Auren (base CVM reportada; o pro forma fica como memória), Cury, Direcional, Fleury, Grupo Mateus, Movida, Orizon, Sanepar, SIMPAR, Track & Field, Vamos e Vivara.
- **Bancos** (base própria de banco): ABC Brasil, Agibank, BMG, BR Partners, BTG Pactual e PicPay.

### Bancos dos EUA

- **Base:** `MODELO_IF_BANCOS_EUA_v1.xlsx`, a adaptação da base brasileira de bancos para US GAAP, CECL e Basileia III dos EUA. O que mudou e como gerar está em `pipeline/bancos_eua/README.md`.
- **JPMorgan** (`Modelo_JPMorgan_2T26.xlsx`), via SEC XBRL, 1T20 a 2T26:
  - checks de balanço, DRE e PL zerados;
  - segmentos CCB, CIB, AWM e Corporate fechando com a receita líquida;
  - o CIB é calculado por diferença, porque em 2024 a JPM juntou CIB e Commercial Banking e reutilizou a etiqueta antiga;
  - NII ex-Markets, nonaccrual total, LCR e NSFR estão vazios: dependem do earnings supplement;
  - o EPS do modelo usa as ações no fim do período, não a média diluída do release.

### Pendentes

| Companhia | Motivo |
|---|---|
| SpaceX | IPO em jun/26, só há um 10-Q |
| Shell, TotalEnergies, BP, Rio Tinto, Glencore, BHP, Canadian Natural, TSMC, ASML, Spotify | Reportam por 20-F ou 40-F, sem trimestre na SEC. Precisam das planilhas de RI de cada companhia |

A lista completa do pedido está na aba "Planilha2" do arquivo `BALANCOS PARA FAZER - COMPLETO.xlsx`, que fica com o usuário.

## Máscara vigente: o que mudou nos números

- A grade começa no 1T21; 2020 saiu. A DFC indireta do 1T21 (e de 2021 no anual) fica vazia por falta do balanço de 2020.
- Auren: variação do direito de uso estava zerada na DFC (plug de caixa de até R$ 80 mi); restaurada a fórmula da máscara, plug zerado.
- Vamos: a fonte só tem DRE desde o 1T23, então 2021-22 ficam fora (o balanço desses anos existia no modelo antigo, sem DRE). A DFC começa no 2T23.
- Dívida LP e caixa disponível do cronograma em 1T26/2T26 usam o saldo real do balanço (o modelo antigo repetia o 4T25).
- Diferenças receita x soma dos segmentos que vêm da fonte (já existiam): AMD 4T21/4T22, GE Aerospace (cisões), Bunge até US$ 80 mi em 2024-25, MercadoLibre US$ 622 mi no 4T23, Netflix até US$ 46 mi, Oracle/Uber/Eli Lilly até US$ 8 mi (arredondamento).
- Conferência: `pipeline/comum/check_modelo.py` (checks e erros) e `pipeline/comum/dupla_checagem.py ORIGEM CONVERTIDO` (receita, caixa, checks e plug da DFC por trimestre).

## Bugs já corrigidos (não reintroduzir)

1. **Painel congelado do Model na linha ~350:** a planilha não rolava. Causa: o `shift_rows.py` deslocava o `topLeftCell`. Corrigido no script e nos 43 arquivos.
2. **Lucro antes do IR vazio** quando a companhia troca a tag (Costco, Oracle): a lista padrão de `PRETAX` foi ampliada. Sempre confira se a linha "Lucro antes do IR" da Inputs tem os 26 trimestres. Se faltar algum, o resultado cai em "resultado após IR não classificado" e a alíquota de IR sai errada, mesmo com os checks zerados.
3. **3M com o valor do ano no 10-K** (Oracle): tratado em `Co._pick`.

## Próximos passos sugeridos

- Outros bancos dos EUA (BAC, C, WFC, GS, MS) na base EUA, copiando o `build_jpm.py`; NII ex-Markets, nonaccrual e LCR a partir do earnings supplement.
- Estrangeiras a partir das planilhas de RI, uma de cada vez.
- Atualização para o 3T26 com a habilidade `docs/skills/atualizar-trimestre`.
