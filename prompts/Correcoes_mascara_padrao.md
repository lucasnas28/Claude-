# Correções na máscara padrão (Modelo_2.0_proprio_Padrao_AJUSTADO-2028)

Arquivo corrigido: `Modelo_2.0_proprio_Padrao_AJUSTADO-2028_corrigido.xlsx`.
Nenhuma aba, nome definido ou formatação condicional foi removido. A Inputs e a Inputs (2) não mudaram, exceto o rótulo AG2.

## Corrigido

| # | Onde | Problema | Correção |
|---|---|---|---|
| 1 | Model I134, AF134 | capex sem fórmula no 1Q21 e no 4Q25 (check de caixa da linha 281 não fechava) | fórmula do trimestre vizinho |
| 2 | Model AF (19 linhas) | 4Q25 sem os índices (crescimento, margens, IR, prazos, variações) | fórmula do trimestre vizinho |
| 3 | Valuation D:R | colunas liam o Model dois anos antes (2023 lia 2021) | referências deslocadas para o ano certo |
| 4 | Valuation S:V e terminal | horizonte ia até 2041, mas o Model termina em 2037 | horizonte encurtado para 2037 (valor terminal, soma dos fluxos, data terminal D59, CAGR) |
| 5 | Valuation G24 | IR de 2026 apontava para a linha 99, que é vazia | linha 98 |
| 6 | Model linhas 3 e 4 | check de DRE comparava o lucro com a linha 4, que era vazia | linha 4 recebe o lucro reportado; o check fica em branco enquanto ela estiver vazia |
| 7 | Inputs (2) AG2 | rótulo 2025 repetido | 2026 |
| 8 | Model linha 300 | dívida/EBITDA dividia pela linha 16 (% de crescimento) | divide por EBITDA × 4 |
| 9 | Model linha 302 | dívida líquida/EBITDA usava linhas erradas a partir de 1Q24 (184 e 205; faltavam 178 e 201) | (178+179+180+183+201+204 − 110) / (EBITDA × 4) em todas as colunas |
| 10 | Model linha 303 | ex IFRS 16 usava linhas erradas a partir de 1Q25 (175, 176, 181, 202, 106, 28) | (178+179+183+204 − 110) / (EBITDA × 4) |
| 11 | Model linha 163 | total do ativo com #REF! em todo 1º trimestre projetado | mesma soma dos outros trimestres |
| 12 | Model linha 150 (histórico) | intangível espelhava o goodwill (145) | espelha o intangível (146) |
| 13 | Model linhas 137, 150, 189 | projeção partia de coluna anual vazia (zero) ou pulava um trimestre; AI189 era 0 digitado | cada trimestre parte do trimestre anterior; a dívida LP do 1Q26 parte do saldo real (linha 183) |
| 14 | Model linhas 225, 230, 231 | preço médio e ações pulavam o 2º trimestre e partiam de coluna anual vazia | trimestre anterior; o 1Q26 continua partindo de AG225/AG230 |
| 15 | Model linha 196 | `(X258,X265)` é união de células e dava #VALUE! (154 células) | CFO (259) + CFI (266) + fluxos de capital (268:274) |

Teste no modelo da Fleury preenchido: #VALUE! e #REF! zeraram, nenhum erro nas colunas históricas e os checks continuam zerados.

## Células de input destacadas em amarelo

- **AG197:** caixa mínimo como % da receita
- **AG225:** preço atual da ação (semente da linha 225)
- **AG230:** ações em circulação no fim do último trimestre (semente da linha 230)

## Não corrigido (depende de decisão de desenho)

- **Linhas 185, 187, 198 e 199:** a mecânica da parcela circulante e da linha de crédito rotativo referencia trimestres de forma inconsistente, e o 1º trimestre de cada ano zera a linha de crédito.
- **Linhas 148 e 149:** a compra de intangível e a amortização leem as linhas 49 e 50 (caixa e juros sobre caixa). Falta definir a divisão da D&A entre imobilizado e intangível (linha 152).
- **Linha 226:** o crescimento de LPA de consenso aponta para AL231 (ações médias de 2026).
- **Bloco macro (linhas 8 a 10):** IPCA e CDI digitados à mão, sem fonte.
- **#DIV/0! nas projeções:** existe porque as premissas estão vazias. Some quando os drivers forem preenchidos.
