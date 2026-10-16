# saldo de fim de período (instantâneo, sem dimensão) por trimestre-calendário; usado para RPO/backlog
from sx import label, P
def inst(co, tags):
    out = {}
    for t in tags:
        got = {}
        for (n, ns, s, e, dd, v, u, a, fl, i) in co.facts:
            if n == t and s is None and not dd:
                lab = label(e)
                if lab in P and (lab not in got or fl > got[lab][1]): got[lab] = (v, fl)
        for k, (v, fl) in got.items(): out.setdefault(k, v)
    return out
