import json, datetime as dt
D=__import__('os').environ.get('PIPE_DATA', 'dados')+'/sec/facts.json'
F=json.load(open(D))['facts']
def d(s): return dt.date.fromisoformat(s)
def label(end):
    e=d(end); m=e.month; y=e.year
    if m in (4,5): return f'1Q{y%100:02d}'
    if m in (7,8): return f'2Q{y%100:02d}'
    if m in (10,11): return f'3Q{y%100:02d}'
    if m in (1,2): return f'4Q{(y-1)%100:02d}'
    return None
PERIODS=[f'{q}Q{y:02d}' for y in range(20,27) for q in range(1,5)][:26]  # 1Q20..2Q26
def facts(tag, ns='us-gaap', unit=None):
    t=F[ns].get(tag)
    if not t: return []
    u=unit or list(t['units'])[0]
    return t['units'][u]
def series(tag, kind, ns='us-gaap', unit=None):
    """kind: 'q' 3M discrete, 'ytd' cumulative fiscal YTD (3/6/9/12M), 'i' instant.
    Returns {label: (val, filed, form, accn)} choosing latest-filed fact."""
    out={}
    for f in facts(tag,ns,unit):
        if kind=='i':
            if 'start' in f: continue
            lab=label(f['end'])
        else:
            if 'start' not in f: continue
            days=(d(f['end'])-d(f['start'])).days
            lab=label(f['end'])
            if lab is None: continue
            q=int(lab[0])
            if kind=='q':
                if not 80<=days<=100: 
                    if not (q==4 and 350<=days<=380): continue
                    # FY fact stored in 4Q slot
                elif q==4: continue
            elif kind=='ytd':
                exp={1:91,2:182,3:273,4:364}[q]
                if abs(days-exp)>12: continue
        if lab not in PERIODS: continue
        prev=out.get(lab)
        if prev is None or f['filed']>prev[1]:
            out[lab]=(f['val'],f['filed'],f['form'],f['accn'])
    return out
