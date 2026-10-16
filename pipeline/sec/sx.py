"""Leitor genérico de instâncias XBRL da SEC (10-Q/10-K/20-F) para a máscara padrão.
Trimestres-calendário: cada data de fechamento vai para o fim de trimestre-calendário mais próximo."""
import os, glob, pickle, datetime as dt, collections
from lxml import etree

X = '{http://www.xbrl.org/2003/instance}'; XDI = '{http://xbrl.org/2006/xbrldi}'
P = [f'{q}Q{y:02d}' for y in range(20, 27) for q in range(1, 5)][:26]          # 1Q20..2Q26


def d(s): return dt.date.fromisoformat(s)


MONTHMAP = {}     # opcional por companhia: mês de fechamento -> trimestre-calendário (ex. Costco: fev -> 1Q, mai -> 2Q)


def label(end):
    e = d(end)
    if e.month in MONTHMAP: return f'{MONTHMAP[e.month]}Q{e.year % 100:02d}'

    best = None
    for y in (e.year - 1, e.year, e.year + 1):
        for q, (mm, dd) in enumerate([(3, 31), (6, 30), (9, 30), (12, 31)], 1):
            k = abs((e - dt.date(y, mm, dd)).days)
            if best is None or k < best[0]: best = (k, f'{q}Q{y % 100:02d}')
    return best[1]


def kind_of(start, end):
    n = (d(end) - d(start)).days
    if 80 <= n <= 115: return '3M'
    if 160 <= n <= 200: return '6M'
    if 245 <= n <= 285: return '9M'
    if 350 <= n <= 380: return '12M'


def prevq(p):
    q, y = int(p[0]), int(p[2:])
    return f'{q - 1}Q{y:02d}' if q > 1 else f'4Q{y - 1:02d}'


class Co:
    def __init__(self, base, monthmap=None):
        global MONTHMAP
        MONTHMAP.clear(); MONTHMAP.update(monthmap or {})
        self.B = base.rstrip('/') + '/'
        self.filings = {}
        for l in open(self.B + 'accns.txt'):
            t = l.split()
            if len(t) == 5: self.filings[t[0]] = dict(form=t[1], filed=t[2], rd=t[4])
        pk = self.B + 'facts.pkl'
        if os.path.exists(pk): self.facts = pickle.load(open(pk, 'rb'))
        else:
            self.facts = []
            for f in sorted(glob.glob(self.B + 'inst/*.xml')):
                a = os.path.basename(f)[:-4]
                if a not in self.filings or os.path.getsize(f) < 1000: continue
                try: t = etree.parse(f).getroot()
                except Exception: continue
                ctx, units = {}, {}
                for c in t.iter(X + 'context'):
                    p = c.find(X + 'period'); s = p.find(X + 'startDate'); e = p.find(X + 'endDate'); i = p.find(X + 'instant')
                    dd = tuple(sorted((m.get('dimension'), m.text) for m in c.iter(XDI + 'explicitMember')))
                    ctx[c.get('id')] = (s.text if s is not None else None, (e if e is not None else i).text, dd)
                for u in t.iter(X + 'unit'):
                    m = u.find('.//' + X + 'measure'); units[u.get('id')] = m.text if m is not None else ''
                for n, el in enumerate(t):
                    if not isinstance(el.tag, str) or el.get('contextRef') not in ctx or el.text is None: continue
                    try: v = float(el.text)
                    except ValueError: continue
                    q = etree.QName(el)
                    s, e, dd = ctx[el.get('contextRef')]
                    self.facts.append((q.localname, q.namespace, s, e, dd, v, units.get(el.get('unitRef'), ''), a, self.filings[a]['filed'], n))
            pickle.dump(self.facts, open(pk, 'wb'))
        # trimestre fiscal: data de fechamento do ano fiscal (10-K)
        ends = [f['rd'] for f in self.filings.values() if f['form'] in ('10-K', '20-F', '40-F')]
        self.fy_digit = collections.Counter(label(e)[0] for e in ends).most_common(1)[0][0]   # trimestre-calendário onde cai o FQ4

    def fq(self, p):
        """posição fiscal (1..4) do trimestre-calendário p."""
        return (int(p[0]) - int(self.fy_digit) - 1) % 4 + 1

    # ---------------- fluxos (sem dimensão)
    def flow(self, tags, kind, orig=False, dims=()):
        out = {}
        for tag in tags:
            got = {}
            for (n, ns, s, e, dd, v, u, a, fl, i) in self.facts:
                if n != tag or s is None or dd != dims or kind_of(s, e) != kind: continue
                lab = label(e)
                if lab not in got or (fl < got[lab][1] if orig else fl > got[lab][1]): got[lab] = (v, fl, a)
            for k, v in got.items(): out.setdefault(k, v)
        return out

    def ytd_kind(self, p):
        return {1: '3M', 2: '6M', 3: '9M', 4: '12M'}[self.fq(p)]

    def quarter(self, tags, orig=False, dims=()):
        """valor trimestral: 3M direto, senão YTD - YTD anterior do mesmo ano fiscal (par compatível, ver _pick).
        orig=True: tudo como divulgado originalmente."""
        def match(n, dd):
            return tags.index(n) if n in tags and dd == dims else None
        res = self._collect(match)
        if orig:
            for k in res:
                for lab in res[k]: res[k][lab]['new'] = res[k][lab]['old']
        out = {}
        for p in P + ['4Q19', '3Q19']:
            v = self._pick(res, p)
            if v is not None: out[p] = v
        return out

    # ---------------- balanço de face: instância do próprio 10-Q/10-K, ordem do documento até o total do passivo + PL
    def face(self, p, override=None):
        if override == 'skip': return None, None        # instância sem ordem de documento (ordem alfabética): balanço não confiável
        cands = [(a, f) for a, f in self.filings.items() if label(f['rd']) == p]
        if override:                                   # accn ou (accn, data): coluna comparativa de outro filing
            oa, od = override if isinstance(override, tuple) else (override, None)
            cands = [(oa, dict(self.filings[oa], rd=od or self.filings[oa]['rd']))]
        if not cands: return None, None
        a, f = sorted(cands, key=lambda x: x[1]['filed'])[0]
        rd = f['rd']
        rows = [(i, n, v) for (n, ns, s, e, dd, v, u, aa, fl, i) in self.facts if aa == a and s is None and e == rd and not dd and u.endswith('USD')]
        rows.sort()
        # o balanço começa na primeira linha típica de ativo circulante (descarta fatos de nota que vêm antes no documento)
        starts = [i for i, n, v in rows if n.startswith(('Cash', 'MarketableSecuritiesCurrent', 'ShortTermInvestments', 'AccountsReceivable', 'ReceivablesNet', 'Inventory', 'PropertyPlantAndEquipment', 'PublicUtilitiesPropertyPlantAndEquipment'))]
        if starts: rows = [x for x in rows if x[0] >= min(starts)]
        # linhas de face com um só membro (ex.: Tesla, veículos de leasing como PP&E por tipo): só as intercaladas entre as linhas de face
        idx = {n: i for i, n, v in rows}
        if 'AssetsCurrent' in idx and 'LiabilitiesAndStockholdersEquity' in idx:
            lo, hi = min(i for i, n, v in rows if i >= idx['AssetsCurrent'] - 400), idx['LiabilitiesAndStockholdersEquity']
            lo = rows[0][0] if rows else lo
            plain = sorted(i for i, n, v in rows)
            for (n, ns, s, e, dd, v, u, aa, fl, i) in self.facts:
                if aa == a and s is None and e == rd and len(dd) == 1 and u.endswith('USD') and lo < i < hi and 'VariableInterestEntity' not in dd[0][1] and not any(k in dd[0][0] for k in ('Segment', 'ConsolidationItems')) and n not in ('Assets', 'Liabilities', 'StockholdersEquity', 'LiabilitiesAndStockholdersEquity'):
                    # só se cair entre duas linhas de face consecutivas (distância curta)
                    if any(0 < i - j <= 8 for j in plain) and any(0 < j - i <= 8 for j in plain):
                        rows.append((i, f'{n}[{dd[0][1].split(":")[-1]}]', v))
            rows.sort()
        out, seen = [], set()
        for i, n, v in rows:
            if n in seen: continue
            seen.add(n); out.append((n, v))
            if n == 'LiabilitiesAndStockholdersEquity': break
        return out, a

    # ---------------- dimensional
    def _pick(self, res, p):
        """res[kind][lab] = {'new': (v, filed), 'old': (v, filed)}. 3M direto (mais recente); senão YTD - YTD anterior, usando o par
        mais recente só se o YTD maior não foi reapresentado depois do YTD menor; caso contrário o par original."""
        k = self.ytd_kind(p)
        if p in res['3M']:
            v3 = res['3M'][p]['new'][0]
            # contexto de 3M no 10-K com o valor do ano inteiro (erro de etiquetagem, ex.: Oracle): ignora e desacumula
            if not (k == '12M' and p in res[k] and abs(v3) > 0.5 and abs(res[k][p]['new'][0] - v3) < 0.5 and prevq(p) in res['9M']): return v3
        if k == '3M' or p not in res[k]: return None
        pk = {'6M': '3M', '9M': '6M', '12M': '9M'}[k]
        y0 = res[pk].get(prevq(p))
        if y0 is None: return None
        y = res[k][p]
        recast = abs(y['new'][0] - y['old'][0]) > 0.5 and y['new'][1] > y0['new'][1]
        if not recast: return y['new'][0] - y0['new'][0]
        return y['old'][0] - y0['old'][0]

    def _collect(self, match):
        res = {k: {} for k in ('3M', '6M', '9M', '12M')}
        for (n, ns, s, e, dd, v, u, a, fl, i) in self.facts:
            if s is None: continue
            kind = kind_of(s, e)
            if kind not in res: continue
            pri = match(n, dd)
            if pri is None: continue
            lab = label(e); cur = res[kind].setdefault(lab, {})
            if 'new' not in cur or (pri, -int(fl.replace('-', ''))) < cur['nk']: cur['new'] = (v, fl); cur['nk'] = (pri, -int(fl.replace('-', '')))
            if 'old' not in cur or (pri, int(fl.replace('-', ''))) < cur['ok']: cur['old'] = (v, fl); cur['ok'] = (pri, int(fl.replace('-', '')))
        return res

    def dq(self, tag, member, axis=None, orig=False):
        """série trimestral de um membro. member: str, tupla (combinação exata de membros fora de ConsolidationItems/OperatingSegments)
        ou lista de alternativas (preferência na ordem). Alternativas unidas antes de desacumular (a companhia troca o nome do membro
        entre o 10-Q e o 10-K do mesmo ano)."""
        alts = member if isinstance(member, list) else [member]
        wants = [tuple(sorted(mm)) if isinstance(mm, tuple) else (mm,) for mm in alts]
        def match(n, dd):
            if n != tag or not dd: return None
            seg = any(ax == 'srt:ConsolidationItemsAxis' and m == 'us-gaap:OperatingSegmentsMember' for ax, m in dd)
            mem = tuple(sorted(m for ax, m in dd if not (ax == 'srt:ConsolidationItemsAxis' and m == 'us-gaap:OperatingSegmentsMember')))
            # mesmo membro com e sem "OperatingSegments": o fato do segmento operacional tem prioridade
            return wants.index(mem) * 2 + (0 if seg else 1) if mem in wants else None
        res = self._collect(match)
        if orig:
            for k in res:
                for lab in res[k]: res[k][lab]['new'] = res[k][lab]['old']
        out = {}
        for p in P:
            v = self._pick(res, p)
            if v is not None: out[p] = v
        return out

    def members(self, tags=('Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'OperatingIncomeLoss')):
        c = collections.Counter()
        for (n, ns, s, e, dd, v, u, a, fl, i) in self.facts:
            if n in tags and s is not None and dd and kind_of(s, e) == '3M':
                mem = [(ax, m) for ax, m in dd if not (ax == 'srt:ConsolidationItemsAxis' and m == 'us-gaap:OperatingSegmentsMember')]
                if len(mem) == 1: c[(n, mem[0][0], mem[0][1])] += 1
        return c
