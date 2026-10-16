---
name: "projecao-receita-volume-preco"
description: "Use antes de projetar receita de uma empresa (DCF ou multi-ano), para decompor cada segmento em quantidade x preço e raciocinar as premissas antes de montar fórmulas."
---

# Projeção de receita por volume x preço

Esta skill é um framework de raciocínio, não de modelagem. O objetivo é formar a tese de premissas ANTES de escrever qualquer fórmula no Excel. Só depois de passar por essas perguntas com o analista é que a projeção vira números.

## Passo 1: Decompor cada segmento em quantidade x preço

Para cada linha de receita relevante, identifique o par (quantidade, preço) que a multiplica. Exemplos:
- Locadora: diárias x diária média; carros vendidos x ticket médio
- Banco: carteira de crédito x spread/taxa
- Varejo: same-store sales (ou lojas x ticket médio)
- Telecom: assinantes x ARPU
- Seguradora: apólices x prêmio médio

Se a empresa já reporta esses drivers operacionais (releases trimestrais, apresentações), use como base histórica. Se não reporta, estime implicitamente (receita / preço observável em outra fonte).

## Passo 2: Para cada driver de QUANTIDADE, identificar o teto estrutural

Quantidade não cresce livre, ela é limitada por algo. Pergunte:
- Existe uma restrição de capital/capex para crescer (frota, lojas, capacidade)? A empresa está em ciclo de expansão ou de disciplina/desalavancagem?
- O crescimento depende de penetração de mercado que ainda tem espaço (ex: outsourcing de frota B2B ainda baixo no Brasil), ou já está maduro/saturado?
- Existe sazonalidade estrutural (ex: viagens no verão, datas comemorativas no varejo) que precisa entrar na trajetória trimestral, não só anual?
- Qual o custo de capital hoje (Selic/CDI, custo de dívida) e como isso afeta o apetite de expansão da empresa?

A resposta aqui não é um número, é uma direção qualitativa (crescimento moderado continua vs reaceleração vs desaceleração) que depois vira taxa de crescimento assumida.

## Passo 3: Para cada driver de PREÇO, identificar o mecanismo de repasse

Preço raramente é livre também. Pergunte:
- Existe repasse de inflação (IPCA) mais ou menos automático no contrato/precificação?
- Existe elasticidade ligada à oferta e demanda (ex: utilização de frota, ocupação, utilização de capacidade)? Se a oferta (quantidade) cresce mais rápido que a demanda, o preço tende a ceder, e vice-versa.
- Existe componente de câmbio ou insumo importado que afeta o custo repassado ao preço?
- A empresa tem poder de precificação (marca, posição competitiva) ou é tomadora de preço de mercado?

## Passo 4: Checar o feedback entre quantidade e preço

Quantidade e preço não são independentes. O caso mais comum: se a empresa expande capacidade/oferta mais rápido que a demanda cresce, a métrica de utilização/ocupação cai, e isso pressiona preço. Sempre pergunte se a trajetória de quantidade assumida é consistente com a trajetória de preço assumida, ou se elas se contradizem.

## Passo 5: Linhas que são consequência mecânica, não driver independente

Algumas linhas de receita não merecem uma tese própria de demanda, elas decorrem de uma política operacional já definida. Exemplo: receita de venda de ativos usados (seminovos) decorre da política de renovação de frota (idade média de venda), não de uma demanda própria a ser projetada. Nesses casos, ancore a premissa na política declarada da empresa (ex: meta de idade de frota) e no proxy de preço de mercado (ex: índice de veículos usados), em vez de tratar como incerteza aberta.

## Passo 6: Documentar a tese antes de quantificar

Antes de colocar número no Excel, resuma em uma frase por driver: direção (alta/estável/queda), racional (o porquê), e o que precisa ser verdade para essa tese se confirmar (ex: "frota cresce 8% a.a. se a empresa mantiver disciplina de capital atual e CDI não subir mais"). Isso vira a base das premissas trimestre a trimestre ou ano a ano do modelo.

Só depois desse racional estar claro e validado com o analista é que a etapa de modelagem (fórmulas, projeção no Excel) deve começar.