---
name: "projecao-tres-demonstracoes"
description: "Use ao projetar DRE, balanço e fluxo de caixa completos de uma empresa a partir do histórico (ex: horizonte de DCF), depois que a receita já foi projetada."
---

# Projeção completa de três demonstrações (DRE, Balanço, Fluxo de Caixa)

Esta skill continua de onde a skill projecao-receita-volume-preco termina. Ela pressupõe que a receita já foi raciocinada (quantidade x preço por segmento, com tese qualitativa por driver). A partir daí, o racional segue uma cadeia de dependência, não são premissas soltas: cada linha se apoia na anterior.

## Passo 1: Estudar o passado antes de projetar

Antes de qualquer premissa futura, olhe a série histórica (pelo menos 8-12 trimestres) de cada linha e pergunte: essa margem/relação é estável, cíclica ou está em tendência estrutural? Não projete a partir de um único trimestre recente, projete a partir do padrão.

## Passo 2: Custos e despesas (chegando no EBITDA)

Para cada linha de custo/despesa, identifique se ela é mais fixa ou mais variável em relação à receita/volume:
- Custos variáveis: projete como % da receita ou como custo unitário x quantidade (mesmo racional de volume x preço da receita, só que do lado de custo)
- Custos fixos: projete com repasse de inflação e ganho de escala (diluem como % de receita se a empresa cresce)
- Sempre confira se a margem EBITDA implícita da projeção é plausível frente ao histórico e frente a pares do setor

## Passo 3: D&A (depreciação e amortização)

D&A quase nunca é uma tese independente, ela decorre do estoque de ativos e da política de investimento:
- Empresas asset-heavy (frota, imóveis, equipamentos): D&A segue o saldo médio de ativo fixo e a vida útil/política de renovação já definida no racional de receita (ex: idade média de venda de frota)
- Empresas asset-light: D&A tende a ser mais estável como % de receita ou ligada a capex de TI/software

## Passo 4: CapEx

O capex de crescimento já foi implicitamente definido no racional de quantidade da receita (ex: crescimento de frota = capex líquido de expansão). O capex de manutenção segue o ciclo de reposição do ativo (ex: venda e recompra de frota). Junte os dois em uma política de capex total, e confira se ela é sustentável frente à geração de caixa e à política de alavancagem da empresa.

## Passo 5: Capital de giro

Estude a relação histórica entre receita/custos e as principais linhas de giro (prazo médio de recebimento, estoque, prazo médio de pagamento). Projete essas linhas como dias em relação à receita/CMV, não como valores absolutos, para manter a proporcionalidade quando a empresa cresce.

## Passo 6: Estrutura de capital e despesa financeira

Defina a política de alavancagem (a empresa está desalavancando, mantendo, ou alavancando para crescer). A despesa financeira decorre do saldo médio de dívida projetado x custo de dívida, que por sua vez segue a curva de juros esperada (CDI/Selic no Brasil) mais o spread de crédito da empresa. Não trate despesa financeira como premissa isolada, ela é resultado do resto do modelo (capex, geração de caixa, política de dividendos).

## Passo 7: Balanço como amarração

O balanço não é uma demonstração projetada de forma independente, ele é o resultado acumulado do DRE e do fluxo de caixa. Monte o balanço como rolagem: saldo inicial mais as movimentações que vêm do DRE (lucro retido) e do fluxo de caixa (caixa, dívida, capex). Isso é o que garante que o modelo feche.

## Passo 8: Fluxo de caixa e checagem

O fluxo de caixa deve ser derivado do DRE e da variação do balanço, não construído do zero. Ao final, sempre valide:
- O balanço fecha (ativo = passivo + patrimônio) em todos os períodos projetados
- A variação de caixa no fluxo bate com a variação de caixa no balanço
- As margens e retornos implícitos (ROIC, margem EBITDA, alavancagem) são plausíveis frente ao histórico e à tese de longo prazo

## Passo 9: Convergência de longo prazo (quando o horizonte é tipo DCF)

Quanto mais distante no horizonte, menos a projeção deve depender de ciclos de curto prazo e mais deve convergir para uma condição estrutural sustentável: crescimento perto do PIB ou do teto de mercado, margem estável, ROIC compatível com o custo de capital. Trate os primeiros 1-2 anos com mais detalhe operacional (trimestral, sazonalidade) e os anos seguintes com premissas mais simples e estáveis até o ano terminal.

## Ordem de execução

Sempre siga a cadeia: receita (via projecao-receita-volume-preco) -> custos/EBITDA -> D&A (via capex e política de ativo) -> resultado financeiro (via política de dívida) -> lucro líquido -> capital de giro e balanço -> fluxo de caixa -> checagem de consistência. Não pule etapas construindo uma linha antes de ter a tese da linha anterior.