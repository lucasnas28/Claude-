import sys; sys.path.insert(0,__import__('os').environ.get('PIPE_DATA', 'dados')+'/work')
from sec import *
def ser(tags, kind, only=None):
    out={}; src={}
    for t in tags:
        for k,v in series(t,kind).items():
            if only and k not in only: continue
            if k not in out: out[k]=v[0]; src[k]=(t,)+v[1:]
    return out, src
FIS={'1':'1T','2':'2T','3':'3T','4':'4T'}
def fiscal(p):
    q=p[0]; y=2000+int(p[2:])+1
    return f'{FIS[q]} FY{y%100:02d}'
def enddate(p):
    for f in facts('Revenues'):
        if 'start' in f and label(f['end'])==p: return f['end']
def split_factor(filed):
    if filed<'2021-07-20': return 40
    if filed<'2024-06-10': return 10
    return 1
