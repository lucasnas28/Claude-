# PREENCHIMENTO DO MODELO BASE (EMPRESAS NÃO FINANCEIRAS) — FLEURY S.A. (FLRY3)

Preencher o histórico trimestral da Fleury S.A. na máscara
Modelo_Base_Empresas_Nao_Financeiras.xlsx, usando a planilha de fundamentos
"Dados Históricos - 2T26.xlsm" como fonte primária e o "Earnings Release -
2T26.pdf" como documento oficial de conferência.

Entregar como: Modelo_Fleury_2T26.xlsx

## 0. Escopo

Mexer apenas em:
- aba Inputs, células de valor, coluna A (conta original) e célula C3 (unidade
  original da fonte e fator aplicado)
- aba Inputs (2), linha 3 (Actual flag)
- aba Control, células B3, B4, B7 e B12

Não alterar nenhuma fórmula das abas Inputs (2), Model ou Valuation.
Não alterar premissas, projeções ou valuation.
Não usar Power Query, macro, VBA ou vínculo externo.
A fonte é .xlsm: abrir só para leitura, sem executar macro, lendo os valores
calculados (as colunas anuais das abas "Fleury" têm fórmula =SUM).
Trabalhar sobre cópia. Nunca sobrescrever o original.

Erro preexistente na máscara fora dessas células vai para lista separada no
relatório, sem correção automática.

### 0.1 Conferir a máscara antes de tudo

Confirmar que o arquivo recebido é de fato a máscara Base: aba Control
existe, Inputs tem os 41 códigos na coluna B e o texto padrão em C3,
Inputs (2) tem a linha 3 de Actual flag e fórmulas SUMIFS com /1000 nas
células trimestrais.

Se faltar qualquer um desses elementos (por exemplo, a versão antiga
"Modelo_2.0_proprio_Padrao_AJUSTADO-2028.xlsx", que tem a Inputs vazia, sem
Control e sem Actual flag), parar e avisar. Não adaptar a máscara.

## 1. Validação de entrada

Antes de escrever qualquer célula, reportar em texto os itens abaixo. Os
valores esperados vêm de uma leitura prévia da fonte e precisam ser
confirmados, não assumidos:

- empresa: capa "Fleury S.A. – Dados Históricos"
- moeda: BRL
- unidade: R$ mil em todas as abas usadas (DRE, BP e DFC).
  ATENÇÃO: o cabeçalho da aba "Combinada BP" diz "(R$ Milhões)", mas os
  valores estão em R$ mil (ex.: Caixa 1T22 = 264.905 = R$ 265 milhões).
  Confirmar contra o balanço do release antes de definir o fator.
- consolidado: sim (Grupo Fleury consolidado)
- cobertura: 1T21 a 2T26, sem buraco, costurando duas fontes (ver item 1.1)
- DRE e DFC vêm com trimestre isolado (os quatro trimestres somam o ano).
  Confirmar numericamente em pelo menos um ano de cada aba.
- 4T existe na fonte como coluna própria em DRE e DFC

### 1.1 Qual aba usar em cada período

A companhia combinou negócios com o Instituto Hermes Pardini. A transação
fechou em 28/04/2023 e a consolidação contábil começa em maio/2023.

| Períodos | Aba | Observação |
|---|---|---|
| 1T21 a 1T23 | "Fleury DRE", "Fleury BP", "Fleury DFC" | contábil, antes da combinação |
| 2T23 | "Combinada DRE"/"Combinada DFC": coluna "2T23 Contábil". "Combinada BP": coluna "2T23" | contábil, com Pardini a partir de maio |
| 3T23 a 2T26 | abas "Combinada" | contábil |

Proibido usar:
- as abas "Combinada" de 1T22 a 1T23: são pro forma (Fleury + Pardini como se
  já estivessem combinadas). Ex.: receita líquida pro forma 1T22 = 1.632.173
  contra 1.089.913 contábil.
- as colunas "2T23 " e "2023" da Combinada DRE/DFC: são pro forma. Usar
  "2T23 Contábil" e "2023 contábil".
- as abas "Pardini": não fazem parte do consolidado antes de maio/2023.

A quebra de perímetro no 2T23 é legítima (é o contábil). Registrar no
relatório e não ajustar.

Na "Fleury BP", as colunas trimestrais 4T21 e 4T22 estão vazias. O saldo de
31/12 está nas colunas anuais 2021 e 2022. Usar esse saldo como 4T.

Na "Combinada DFC", D&A e dividendos ficam zerados de 1T22 a 1T23 e na
coluna "2T23" pro forma. Isso confirma que esses períodos devem sair das abas
"Fleury" e da coluna "2T23 Contábil".

Se a fonte for de outra empresa, estiver corrompida ou protegida, parar e
avisar. Não reconstruir o modelo.

## 2. As 41 linhas, e só elas

A máscara deriva sozinha, a partir das variações de balanço:
capex, variação de capital de giro, fluxo de caixa operacional, de
investimento e de financiamento, dívida líquida, ROIC e todos os índices.

Portanto NÃO procurar na fonte: DFC completa, capex, variação de capital de
giro, dívida líquida, EBITDA divulgado, quantidade de ações.

Procurar exatamente os 41 códigos da coluna B da Inputs. Nada além disso.

Da DFC ou de nota explicativa vêm apenas duas coisas: depreciação e
amortização (código 4) e dividendos pagos (código 10).

## 3. Unidade

A Inputs é a camada bruta normalizada em R$ mil. É o que o título A1 e a
linha 3 declaram, e é o que a fórmula /1000 da Inputs (2) assume.

A fonte está em R$ mil (ver item 1), então o fator é 1: manter os valores.
Registrar em Inputs!C3: "Unidade original: R$ mil (cabeçalho da Combinada BP
diz R$ milhões, mas os valores estão em R$ mil, conferido contra o release).
Fator aplicado: 1".

Control!B6 já vem preenchido com "R$ milhões" (unidade final do modelo). Não
alterar essa célula: ela está fora do escopo da seção 0 e já está correta.

Gravar sempre número puro na célula. Nunca texto, nunca zero para tapar
buraco. Célula sem dado fica realmente vazia.

## 4. Convenção de sinal

Não é opcional e não é inferível. É o que as fórmulas do Model exigem:

Positivos:
  1  Net Revenues
  6  Nonoperating Income / (Loss), quando for receita
  8  Equity in Income of Affiliates, after tax
  11 a 22  todas as contas de ativo
  23 a 34  todas as contas de passivo
  35 Minority interest
  36 Convertible Preferred Stock
  37 Common Stock and APIC
  41 Retained earnings, se positivo

Negativos:
  2  COGS (cash)
  3  SG&A and Other (cash)
  4  D&A
  5  Interest Expense
  7  Taxes, quando houver despesa
  9  Minority Interest Expense, after tax
  10 Common Dividends
  38 Treasury stock (conta redutora)
  39 Comprehensive (accumulated) loss, se negativo

Na fonte Fleury, custos, despesas, IR, despesas financeiras e dividendos já
vêm negativos, e "Ações em Tesouraria" já vem negativa. Conferir linha a
linha e não inverter o que já está na convenção.

Teste do sinal: a linha 103 do Model (Net Income) tem que bater com o lucro
líquido divulgado. Ela é SUM das linhas 96, 97, 101 e 102.

## 5. COGS e SG&A são ex-depreciação

Os rótulos dizem "(cash)". O Model calcula:
  EBITDA = Receita + COGS + SG&A
  EBIT   = EBITDA + D&A

Se a depreciação embutida no custo e na despesa não for retirada, ela entra
duas vezes: o EBITDA sai subestimado e o EBIT sai errado.

Na Fleury a DRE da própria fonte já abre a D&A, então não estimar nada:
- dentro de "Custo do serviço prestado": "Depreciação e Amortização Custos"
- dentro de "(Despesas) Receitas Operacionais": "Depreciação e Amortização"
  e "Amortização de Ágio"

Código 2 = custo do serviço prestado sem a linha de D&A de custos.
Código 3 = despesas operacionais sem as linhas de D&A e amortização de ágio.
Código 4 = soma das três linhas de D&A (custos, despesas e ágio).

Conferir o código 4 contra "Depreciações e amortizações" da DFC no mesmo
trimestre. Diferença relevante vai para o relatório com a causa.

Atenção ao bloco "Abertura Antiga:" na DRE: é uma reabertura do mesmo custo
(Pessoal e Médicos, Materiais e Terceirizações, etc.), inclusive com outra
linha "Depreciação e Amortização". Não codificar nenhuma linha desse bloco,
senão o custo entra duas vezes.

## 6. Trimestralização

Só para contas de resultado (códigos 1 a 10), e só quando a fonte vier
acumulada no ano:
  1T isolado = 3M
  2T isolado = 6M menos 3M
  3T isolado = 9M menos 6M
  4T isolado = anual (DFP) menos 9M

A fonte Fleury já vem com trimestre isolado em DRE e DFC. Confirmado isso
(item 1), não trimestralizar nada.

Nunca trimestralizar contas de balanço (códigos 11 a 41). Para essas, usar o
saldo de encerramento de cada trimestre.

## 7. Mapeamento

Na coluna A de cada linha, escrever a conta original da fonte que alimentou
aquele código, indicando também a aba (ex.: "Combinada BP - Debêntures PC").
Isso é o que torna o preenchimento auditável.

Várias linhas da fonte podem apontar para o mesmo código. A fórmula da
Inputs (2) é SUMIFS por código, então isso é o uso pretendido, não uma
exceção. Para fazer isso, duplicar a linha na Inputs mantendo o mesmo número
na coluna B e descrevendo na coluna A qual conta é.

Proibido: dar código ao subtotal e às parcelas dele ao mesmo tempo. É a
única forma de gerar dupla contagem aqui. Na fonte Fleury, os subtotais a
deixar sem código são: Receita Bruta (usar Receita Líquida), Lucro Bruto,
Total do Circulante, Total do Realizável a Longo Prazo, Total do Não
Circulante, Total do Passivo, PL Atribuído aos Proprietários da Controladora,
Total do PL e todo o bloco "Reconciliação do EBITDA".

### Mapeamento sugerido para a Fleury

Adotar este mapeamento como ponto de partida. Mudar só com justificativa no
relatório.

Resultado:
- 1: Receita Líquida
- 2: Custo do serviço prestado, sem "Depreciação e Amortização Custos"
- 3: Gerais e Administrativas, Comercial, Outras receitas (despesas)
  operacionais líquidas, Reversão de (provisão) para riscos, todas sem as
  linhas de D&A e de amortização de ágio
- 4: Depreciação e Amortização Custos + Depreciação e Amortização (despesas)
  + Amortização de Ágio
- 5: Despesas financeiras
- 6: Receitas financeiras + Variação Cambial Líquida
- 7: Imposto de Renda e Contribuição Social (corrente + diferido)
- 8: Equivalência Patrimonial
- 9: Participação dos não-controladores
- 10: Dividendos e/ou juros sobre capital próprio pagos (DFC, atividades de
  financiamento)

Ativo:
- 11: Caixa e Equivalentes + Títulos e Valores Mobiliários AC
- 12: Contas a Receber AC
- 13: Estoques AC
- 14: IR e CS Diferidos AC + IR e CS Diferidos ANC
- 15: Despesas do Exercício Seguinte AC
- 16: Impostos a Recuperar AC, IRPJ e CSLL a compensar, Partes Relacionadas
  AC, Derivativos AC, Notas Promissórias AC, Outros AC, Ativos Mantidos para
  Venda
- 17: Investimentos
- 18: Imobilizado
- 19: Direito de Uso
- 20: vazio, a menos que a nota de intangível abra o ágio (ver código 21)
- 21: Intangível (inclui ágio)
- 22: todo o Realizável a Longo Prazo restante (Partes Relacionadas,
  Contas a Receber ANC, Derivativos ANC, Impostos a Recuperar ANC, IRPJ e
  CSLL a compensar ANC, Depósitos Judiciais, TVM ANC, Notas Promissórias ANC,
  Créditos a Receber ANC, Outros ANC)

Passivo:
- 23: Fornecedores
- 24: Obrigações Trabalhistas (na aba Fleury: Salários e Encargos a Recolher)
- 25: Obrigações Tributárias, Provisão para IR e CS, Impostos e Contribuições
  a Recolher PC
- 26: Derivativos PC, Dividendos e JCP PC, Contas a Pagar - Aquisição de
  Empresas PC, Partes Relacionadas PC, Outras Contas a Pagar PC
- 27: Empréstimos e Financiamentos PC + Debêntures PC
- 28: vazio (sem linha de crédito rotativo separada)
- 29: Arrendamento Mercantil Financeiro PC
- 30: Empréstimos e Financiamentos PNC + Debêntures PNC
- 31: Arrendamento Mercantil Financeiro PNC
- 32: vazio (sem conversíveis)
- 33: IR e CS Diferidos PC + IR e CS Diferidos PNC
- 34: Derivativos PNC, Provisões para riscos e contingências, Parcelamentos
  tributários, Provisão para perdas em investimentos, Impostos a Recolher PNC,
  Partes Relacionadas PNC, Contas a Pagar - Aquisição de Empresas PNC,
  Dividendos a Pagar PNC, Outros PNC

Patrimônio:
- 35: Minoritários
- 36: vazio
- 37: Capital Social + Reserva de Capital + Reserva de Capital - Opções
  Outorgadas
- 38: Ações em Tesouraria
- 39: Ajustes de Avaliação Patrimonial
- 40: Reserva de Reavaliação, Reserva Legal, Reserva para Investimentos,
  Dividendo Adicional Proposto, Juros sobre Capital Próprio
- 41: Lucros Acumulados + Lucro do Período

As contas "Contas a Pagar - Aquisição de Empresas" têm natureza de dívida
(earn-out e parcelas de M&A). Ficam em 26 e 34 para não distorcer a dívida
bancária, mas o saldo de cada trimestre precisa constar no relatório, porque
o analista pode querer somá-lo à dívida líquida.

### Regra do resíduo, que é o que faz o balanço fechar

A linha 163 do Model soma exatamente os códigos 11 a 22.
A linha 213 soma exatamente os códigos 23 a 34.
A linha 237 soma exatamente os códigos 35 a 41.

Logo, toda conta do balanço da companhia precisa cair em um desses 31
códigos. Nada pode ficar de fora.

Contas que não têm destino óbvio vão para:
  16  Other current assets
  22  Other non-current assets
  26  Other current liabilities (non-debt)
  34  Other non-current liabilities
  40  Other equity account

Mesma lógica no resultado. A linha 103 depende só dos códigos 1 a 9. Conta de
resultado sem destino óbvio vai para o código 3 (se operacional) ou o
código 6 (se não operacional).

### Cuidados específicos

Código 14 e código 33 têm o mesmo nome (Deferred income taxes). O 14 é ativo
(Model linha 119), o 33 é passivo (Model linha 205). Não trocar.

Código 5 é despesa financeira bruta, negativa. Não lançar resultado
financeiro líquido aqui. As linhas 298, 299 e 304 do Model calculam cobertura
de juros em cima dela.
Receita financeira vai no código 6.

Códigos 29 e 31 (arrendamentos IFRS 16) devem ficar separados da dívida
onerosa. A linha 303 do Model calcula dívida líquida ex IFRS 16 e depende
dessa separação.

Os nomes das contas mudam entre as abas "Fleury" e "Combinada" (ex.:
"Salários e Encargos a Recolher" contra "Obrigações Trabalhistas"; a aba
Fleury não tem linha de Debêntures separada). Mapear pelo conteúdo e
registrar a equivalência na coluna A.

## 8. Mapa de colunas

Inputs e Inputs (2) usam o mesmo layout. Deslocamento fixo de 5 colunas em
relação ao Model.

  D→I  1T21    E→J  2T21    F→K  3T21    G→L  4T21    H→M  2021
  I→N  1T22    J→O  2T22    K→P  3T22    L→Q  4T22    M→R  2022
  N→S  1T23    O→T  2T23    P→U  3T23    Q→V  4T23    R→W  2023
  S→X  1T24    T→Y  2T24    U→Z  3T24    V→AA 4T24    W→AB 2024
  X→AC 1T25    Y→AD 2T25    Z→AE 3T25    AA→AF 4T25   AB→AG 2025
  AC→AH 1T26   AD→AI 2T26

O último trimestre da Fleury é o 2T26 (coluna AD). As colunas AE, AF e AG
(3T26, 4T26 e 2026) ficam vazias: não há dado na fonte.

## 9. Colunas anuais

A coluna anual da Inputs (2) já vem cabeada e não deve ser tocada:
- para as contas de resultado (códigos 1 a 10), soma os quatro trimestres
  da própria Inputs (2), ex.: `=IF(COUNT(D4:G4)=4,SUM(D4:G4),"")`
- para as contas de balanço (códigos 11 a 41), repete o 4T da própria
  Inputs (2), ex.: `=IF(G15="","",G15)`

Essas fórmulas não leem a coluna anual da Inputs em nenhum momento. Portanto
a coluna anual da Inputs (H, M, R, W, AB) não precisa de fórmula nem de
preenchimento: ela não é usada pelo Model. Deixar vazia. Não criar SUM nem
digitar valor ali.

O único ponto anual que exige atenção é o Actual flag (ver item 10): a
coluna anual da Inputs (2), linha 3, é valor digitado manualmente, não
fórmula.

## 10. Actual flag

Linha 3 da Inputs (2), colunas D até AG.

  1  trimestre efetivamente divulgado e carregado
  0  trimestre sem dado ou futuro

Coluna anual recebe 1 somente quando os quatro trimestres daquele ano
estiverem com flag 1. Ano incompleto recebe 0.

Para a Fleury:
- 1T21 a 2T26 (D até AD, exceto colunas anuais): 1
- anuais 2021, 2022, 2023, 2024 e 2025 (H, M, R, W, AB): 1
- 1T26 e 2T26 com flag 1; 3T26 e 4T26 (AE, AF): 0
- anual 2026 (AG): 0

O flag precisa ser contínuo do 1T21 até o último trimestre divulgado. Buraco
no meio da série faz o bloco de fluxo de caixa do Model retornar #VALUE!,
porque a Inputs (2) devolve texto vazio quando o flag é 0 e as linhas 134,
246 a 255, 268 a 275 e 278 do Model usam subtração direta.

A costura Fleury (até 1T23) + Combinada (a partir do 2T23) não deixa buraco.
Se mesmo assim aparecer algum período sem dado no meio da série, parar e
avisar antes de gravar.

## 11. Control

Preencher apenas:
  B3  Fleury S.A.
  B4  FLRY3
  B7  30/06/2026, como DATA (não texto). Precisa bater exatamente com a data
      da linha 5 do Model, porque o Valuation faz MATCH contra ela.
  B12 data de valuation: [DEFINIR — se não for informada, usar a data de
      hoje e registrar no relatório]

Não preencher B5, B6, B8, B9, B10, B11 nem o bloco E4:E15 nesta etapa. B5 e
B6 já vêm preenchidos (moeda e unidade) e não precisam de ação.

## 12. Validações antes de entregar

Recalcular o arquivo e conferir:

  1. abre normalmente como .xlsx
  2. Model linha 2 (check de balanço) igual a zero, ou dentro de 0,1 em
     R$ milhões, em todos os trimestres com flag 1
  3. Model linha 103 (Net Income) bate com o lucro líquido divulgado em pelo
     menos três trimestres: 1T21, 2T23 (primeiro trimestre combinado) e 2T26
  4. Model linha 15 (Net Revenues) bate com a receita líquida divulgada nos
     mesmos três trimestres
  5. Model linha 35 (EBITDA) bate com EBIT mais D&A da própria CVM.
     Declarar no relatório se a companhia usa outra definição no release.
     EBITDA ajustado do release é referência, não gabarito.
  6. nenhum #REF!, #VALUE!, #NAME?, #DIV/0!, #NUM! ou #N/A nas colunas com
     flag 1
  7. margem bruta (linha 21) e margem EBITDA (linha 290) plausíveis e
     estáveis ao longo da série, sem salto que denuncie erro de sinal. Um
     degrau no 2T23 é esperado (entrada do Pardini) e não é erro.
  8. dívida total (linha 300) e caixa (linha 110) conferidos contra a
     divulgação
  9. nenhum valor digitado nas células trimestrais da Inputs (2)
 10. nenhum vínculo externo
 11. nenhum dado ou nome de outra empresa no arquivo

Âncoras do Earnings Release 2T26 (R$ milhões), para os itens 3 a 5:

| | 2T25 | 2T26 |
|---|---:|---:|
| Receita líquida | 2.024,5 | 2.316,7 |
| EBITDA (release) | 532,1 | 612,9 |
| Lucro líquido | 152,3 | 221,9 |

Para 1T21 e 2T23, usar o release ou o ITR do próprio trimestre. A conferência
dos itens 3, 4, 5 e 8 é contra release, ITR ou DFP. Não vale conferir a
planilha de fundamentos contra ela mesma.

Se o motor de recálculo puder degradar o arquivo, não recalcular. Deixar
recálculo automático ao abrir e avisar no relatório.

## 13. Eficiência

Primeiro montar o inventário das contas da fonte e o mapeamento completo
conta-fonte para código. Só depois escrever.

Escrever em bloco, nunca célula a célula.

Não consultar o RI enquanto a planilha de fundamentos resolver. Usar o RI só
para: derivação do 4T se faltar, divergência relevante e as validações 3 a 5
e 8. A D&A já está aberta na própria DRE da fonte.

Nunca substituir silenciosamente número da fonte primária por número do RI.
Divergência relevante vai para o relatório, com a causa investigada
(reapresentação, reclassificação, aquisição, operação descontinuada,
consolidado versus controladora, contábil versus pro forma).

Dúvida de mapeamento reversível: adotar o tratamento mais defensável,
registrar e explicar como reverter. Não interromper.
Interromper só por: máscara diferente da Base (item 0.1), moeda diferente,
buraco no meio da série, ou risco de danificar o arquivo.

## 14. Relatório final

Em texto, curto:
- período preenchido e períodos faltantes
- unidade da fonte e fator aplicado, inclusive a divergência do cabeçalho da
  Combinada BP
- abas usadas em cada período (Fleury até 1T23, Combinada contábil a partir
  do 2T23) e confirmação de que nenhum número pro forma entrou
- fonte primária e documentos oficiais consultados
- tratamento dado ao D&A embutido em custo e despesa
- contas mapeadas para linha residual, e quais
- saldo de "Contas a Pagar - Aquisição de Empresas" por trimestre
- códigos que ficaram vazios, e por quê
- linhas trimestralizadas (esperado: nenhuma)
- resultado das 11 validações, com os números das conferências 3 a 5 e 8
- divergências e causas
- erros preexistentes da máscara encontrados, sem correção
- confirmação de que Inputs (2) segue só com fórmula
- confirmação de que o original não foi sobrescrito
