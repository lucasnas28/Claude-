# Base de bancos dos EUA

`MODELO_IF_BANCOS_EUA_v1.xlsx` (raiz do repositório) é a base de bancos brasileira (`MODELO_IF_BANCOS_BASE_v7.xlsx`) adaptada para bancos dos EUA (US GAAP, CECL, Basileia III dos EUA). A mecânica é a mesma:
- Inputs → Inputs (2) → Model → Valuation;
- checks de balanço, DRE e roll-forward do PL;
- plug de funding;
- glossário;
- quatro métodos de valuation.

## O que mudou em relação à v7

| Bloco | v7 (Brasil) | EUA |
|---|---|---|
| Balanço | Cosif (aplicações interfinanceiras, TVM própria/vinculada, LF/LCI/LCA, repasses) | Fed funds/compromissadas, securities borrowed, trading, AFS, HTM, loans, ACL, repos, trading liabilities, VIEs, long-term debt |
| PL | Capital, reservas, ORA, tesouraria | + **ações preferenciais**; PL ordinário, MSR e TCE nas bases de cálculo |
| DRE | Intermediação financeira com TVM dentro | NII → provisão → receitas não-juros (tarifas abertas + **principal transactions**) → despesas não-juros |
| Lucro | Lucro dos controladores | + dividendo preferencial → **lucro ao acionista comum** (base de EPS, ROE ordinário, ROTCE e valuation) |
| Distribuição | Dividendos + JCP | Dividendos ordinários + **recompras** (roll-forward, payout e DDM) |
| Crédito | Estágios 1/2/3 (Res. 4.966), NPL 15-90 e 90 | CECL: nonaccrual, NPA, 30-89 e 90+ ainda acumulando, **net charge-offs** |
| Capital | PR, Nível I, Basileia | Total capital, Tier 1, **CET1** (Standardized e Advanced), **SLR**, Tier 1 leverage, LCR |
| Segmentos | Carteira por produto | 4 segmentos + reconciliação (receita e lucro), com check contra a receita líquida total |
| Macro | Selic, CDI, IPCA, TLP, câmbio | Fed funds (meta e efetivo), SOFR, Treasury 10 anos, CPI, PIB real (FRED) |
| Valuation | NTN-B, ERP Brasil, pares ITUB4… | Treasury 10 anos, ERP dos EUA, pares BAC/C/WFC/GS/MS/PNC/USB; **P/TBV, ROTCE e TBVPS** |

## Como gerar

```bash
# séries do FRED (macro)
for s in DFEDTARU DFF CPIAUCSL GDPC1 DGS10 SOFR; do curl -sS -o dados/fred/$s.csv "https://fred.stlouisfed.org/graph/fredgraph.csv?id=$s&cosd=2019-01-01"; done
python3 pipeline/bancos_eua/build_base_eua.py MODELO_IF_BANCOS_BASE_v7.xlsx dados/base_eua.xlsx dados/fred
# banco (ex.: JPMorgan): baixa o XBRL e preenche
bash pipeline/sec/dl.sh 0000019617 dados/jpm/dl
python3 pipeline/bancos_eua/build_jpm.py dados/base_eua.xlsx dados/jpm_a.xlsx dados/jpm/dl
python3 pipeline/comum/recalc.py dados/jpm_a.xlsx dados/jpm_a_recalc.xlsx
python3 pipeline/comum/inject.py dados/jpm_a.xlsx dados/jpm_a_recalc.xlsx Modelo_JPMorgan_2T26.xlsx
```

Para outro banco (BAC, C, WFC…), copie o `build_jpm.py` e troque:
- os membros dos segmentos;
- as tags específicas (por exemplo, nomes de membros `jpm:`).

O mapeamento de balanço e DRE usa tags us-gaap e serve para a maioria dos grandes bancos. Confira sempre os checks do Model, o CHECK da DRE na Inputs e o check dos segmentos.

## Limitações da fonte XBRL

Estas métricas só existem no earnings supplement e ficam vazias, com o motivo escrito:
- NII ex-Markets;
- nonaccrual total e NPA;
- vencidos fora do cartão;
- LCR e NSFR;
- saldos médios;
- headcount.
