import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/lmt/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Aeronautics, Missiles & Fire Control, Rotary & Mission Systems e Space (receita e lucro operacional), produtos x serviços - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores dimensionais das instâncias XBRL (receita líquida de vendas entre segmentos).',
 [('aero', 'Aeronautics', ['lmt:AeronauticsMember'], True), ('mfc', 'Missiles and Fire Control', ['lmt:MissilesAndFireControlMember'], True),
  ('rms', 'Rotary and Mission Systems', ['lmt:RotaryAndMissionSystemsMember'], True), ('space', 'Space', ['lmt:SpaceMember'], True)],
 EXTRA=[('prd', 'Receita de produtos', ['Revenues'], ['us-gaap:ProductMember']), ('srv', 'Receita de serviços', ['Revenues'], ['us-gaap:ServiceMember'])]))
