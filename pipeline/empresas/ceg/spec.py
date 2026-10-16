import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/ceg/dl',
 'DRIVERS DE RECEITA (codigos 45+): receita por região de geração/comercialização (Mid-Atlantic, Midwest, New York, ERCOT, outras regiões) - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. A soma das regiões não fecha com a receita consolidada: a diferença são outros negócios (gás natural, renováveis, resultado de derivativos economicamente hedgeados). A companhia existe como emissora separada desde a cisão da Exelon (fev/2022); os dados de 2021 vêm dos comparativos.',
 [('ma', 'Mid-Atlantic', ['ceg:ConstellationMidAtlanticMember'], False), ('mw', 'Midwest', ['ceg:ConstellationMidwestMember'], False), ('ny', 'New York', ['ceg:ConstellationNewYorkMember'], False),
  ('ercot', 'ERCOT', ['ceg:ConstellationERCOTMember'], False), ('oth', 'Outras regiões', ['ceg:ConstellationOtherRegionsMember'], False)], SUMCHK=False))
