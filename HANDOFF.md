# Passagem de trabalho: modelos de valuation 2T26

Situação em 16/10/2026. Use este documento para retomar o trabalho em outra ferramenta (Cursor, outra sessão do Claude).

## O que existe

- **Máscara:** `Modelo_2.0_proprio_Padrao_AJUSTADO-2028_corrigido.xlsx`, com as abas Inputs → Inputs (2) → Model → Valuation. As correções feitas na máscara estão em `prompts/Correcoes_mascara_padrao.md`.
- **Modelos prontos:** `Modelo_<Empresa>_2T26.xlsx` na raiz, com histórico de 1T20 a 2T26 e 1T26 e 2T26 marcados como realizados.
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

### Feitos (Brasil)

- **Fonte CVM ou planilha de fundamentos:** Alupar, Auren (base CVM reportada; o pro forma fica como memória), Cury, Direcional, Fleury, Grupo Mateus, Movida, Orizon, Sanepar, SIMPAR, Track & Field, Vamos e Vivara.
- **Bancos** (base própria de banco): ABC Brasil, Agibank, BMG, BR Partners, BTG Pactual e PicPay.

### Pendentes

| Companhia | Motivo |
|---|---|
| JPMorgan | Banco: precisa da estrutura de modelo bancário |
| SpaceX | IPO em jun/26, só há um 10-Q |
| Shell, TotalEnergies, BP, Rio Tinto, Glencore, BHP, Canadian Natural, TSMC, ASML, Spotify | Reportam por 20-F ou 40-F, sem trimestre na SEC. Precisam das planilhas de RI de cada companhia |

A lista completa do pedido está na aba "Planilha2" do arquivo `BALANCOS PARA FAZER - COMPLETO.xlsx`, que fica com o usuário.

## Bugs já corrigidos (não reintroduzir)

1. **Painel congelado do Model na linha ~350:** a planilha não rolava. Causa: o `shift_rows.py` deslocava o `topLeftCell`. Os 43 arquivos já entregues foram corrigidos para `ySplit=11` / `topLeftCell=D12`. A máscara ainda herdava `topLeftCell=D336` com `ySplit` já em 11, e o `fix_panes.py` só atuava quando `ySplit > 20` — ao regenerar a Oracle o bug voltava. O script agora corrige também o `topLeftCell`, e a máscara foi alinhada.
2. **Lucro antes do IR vazio** quando a companhia troca a tag (Costco, Oracle): a lista padrão de `PRETAX` foi ampliada. Sempre confira se a linha "Lucro antes do IR" da Inputs tem os 26 trimestres. Se faltar algum, o resultado cai em "resultado após IR não classificado" e a alíquota de IR sai errada, mesmo com os checks zerados.
3. **3M com o valor do ano no 10-K** (Oracle): tratado em `Co._pick`.

## Próximos passos sugeridos

- JPMorgan: a base de bancos do repositório é a máscara COSIF/IF.data dos bancos brasileiros (BTG, ABC, Agibank etc.). Não há planilha de RI da JPM nem máscara bancária em US GAAP. Antes de montar, confirmar se copiamos essa máscara e mapeamos o 10-Q da SEC, ou se entra uma máscara bancária americana.
- Estrangeiras a partir das planilhas de RI, uma de cada vez.
- Atualização para o 3T26 com a habilidade `docs/skills/atualizar-trimestre`.
