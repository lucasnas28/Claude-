import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/bx/dl',
 'DRIVERS DE RECEITA (codigos 45+): taxas de gestão e consultoria x receita de performance (incentive fees + alocação de carry) - SEC XBRL (USD, fator 0,000001).',
 'Notas: a receita da Blackstone inclui a alocação de carried interest e ganhos de investimento (podem ser negativos, ex.: 1T20). A diferença para a linha 15 são receitas de investimento e outras.',
 [('adv', 'Taxas de gestão e consultoria (investment advice)', ['us-gaap:InvestmentAdviceMember'], False), ('perf', 'Receita de performance (incentive fees)', ['us-gaap:InvestmentPerformanceMember'], False)], SUMCHK=False))
