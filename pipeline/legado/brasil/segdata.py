import json, sys, datetime as dt
sys.path.insert(0,__import__('os').environ.get('PIPE_DATA', 'dados')+'/work')
from sec import label, PERIODS
S=__import__('os').environ.get('PIPE_DATA', 'dados')+'/sec/inst'
d=json.load(open(S+'/rev.json'))
filed={l.split()[0]:l.split()[2] for l in open(S+'/accns.txt')}; filed['0001045810-21-000010']='2021-02-26'
def seg(member):
    """{period: (val, filed, accn)}: 3M for 1Q-3Q, FY 12M in the 4Q slot; latest filed wins."""
    out={}
    for a,facts in d.items():
        for tag,s,e,dims,v in facts:
            if len(dims)!=1 or tuple(dims[0])!=('srt:ProductOrServiceAxis','nvda:'+member+'Member'): continue
            days=(dt.date.fromisoformat(e)-dt.date.fromisoformat(s)).days
            p=label(e)
            if p not in PERIODS: continue
            q=int(p[0])
            if not ((q<4 and 80<=days<=100) or (q==4 and 350<=days<=380)): continue
            if p not in out or filed[a]>out[p][1]: out[p]=(v,filed[a],a)
    return out
