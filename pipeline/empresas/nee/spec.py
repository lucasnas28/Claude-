import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/nee/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos FPL (Florida Power & Light), NEER (NextEra Energy Resources) e Gulf Power (até 2021) - receita e lucro líquido por segmento - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL. A NextEra mede os segmentos por lucro líquido (não por lucro operacional). Gulf Power foi incorporada à FPL em 2022. A diferença para a receita consolidada é Corporate & Eliminations.',
 [('fpl', 'FPL', ['nee:FPLSegmentMember', 'nee:FloridaPowerLightCompanyMember'], ['NetIncomeLoss']), ('neer', 'NEER', ['nee:NEERSegmentMember', 'nee:NexteraEnergyResourcesMember'], ['NetIncomeLoss']),
  ('gulf', 'Gulf Power', ['nee:GulfPowerSegmentMember'], ['NetIncomeLoss'])],
 EXTRA=[('corp', 'Corporate & Eliminations - receita', ['RegulatedAndUnregulatedOperatingRevenue'], ['nee:CorporateAndEliminationsMember'])],
 RV=('RegulatedAndUnregulatedOperatingRevenue', 'Revenues'), SUMCHK=False))
