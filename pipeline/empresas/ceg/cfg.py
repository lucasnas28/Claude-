NAME = 'Constellation Energy Corporation'; TICKER = 'CEG'; CIK = '0001868275'
DL = __import__('os').environ.get('PIPE_DATA', 'dados')+'/ceg/dl'
FACE_OVR = {'4Q21': ('0001868275-22-000044', '2021-12-31')}   # o 10-K de 2021 etiquetou o balanço por entidade; usa a coluna comparativa do 10-Q de mar/22
