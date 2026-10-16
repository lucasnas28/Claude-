#!/bin/bash
# uso: dl.sh CIK(10 dígitos) DIR  -> DIR/facts.json, subs*.json, accns.txt, inst/*.xml (10-Q/10-K/20-F/40-F/6-K? só 10-Q/10-K/20-F)
CIK=$1; D=$2; UA="${SEC_UA:?defina SEC_UA, ex.: export SEC_UA=\"Seu Nome seu@email\"}"; mkdir -p $D/inst
curl -sS -m 180 -H "User-Agent: $UA" -o $D/facts.json https://data.sec.gov/api/xbrl/companyfacts/CIK$CIK.json -w "facts %{http_code} %{size_download}\n"
curl -sS -m 120 -H "User-Agent: $UA" -o $D/subs.json https://data.sec.gov/submissions/CIK$CIK.json -w "subs %{http_code}\n"
python3 -I - "$D" <<'PY'
import json, sys, subprocess
D = sys.argv[1]; UA = __import__('os').environ['SEC_UA']
s = json.load(open(D + '/subs.json'))
pages = [s['filings']['recent']]
for f in s['filings'].get('files', []):
    if f['filingTo'] >= '2019-06-01':
        subprocess.run(['curl', '-sS', '-m', '120', '-H', 'User-Agent: ' + UA, '-o', D + '/' + f['name'], 'https://data.sec.gov/submissions/' + f['name']])
        pages.append(json.load(open(D + '/' + f['name'])))
rows = []
for r in pages:
    for a, f, d, p, pd in zip(r['accessionNumber'], r['form'], r['filingDate'], r['primaryDocument'], r['reportDate']):
        if f in ('10-Q', '10-K', '20-F', '40-F', '10-KT') and d >= '2019-06-01' and pd: rows.append((a, f, d, p or 'x', pd))
rows = sorted(set(rows), key=lambda x: x[2], reverse=True)
open(D + '/accns.txt', 'w').write('\n'.join(' '.join(r) for r in rows) + '\n')
print(len(rows), 'filings')
PY
CIKN=$((10#$CIK))
while read a f d doc rd; do
  [ -z "$a" ] && continue; [ -s $D/inst/$a.xml ] && continue
  n=${a//-/}
  x=$(curl -sS -m 60 -H "User-Agent: $UA" https://www.sec.gov/Archives/edgar/data/$CIKN/$n/index.json | python3 -I -c "
import json,sys
L=[i['name'] for i in json.load(sys.stdin)['directory']['item']]
c=[n for n in L if n.endswith('_htm.xml')] or [n for n in L if n.endswith('.xml') and not any(k in n for k in ('_cal','_def','_lab','_pre','FilingSummary','MetaLinks'))]
print(c[0] if c else '')")
  [ -n "$x" ] && curl -sS -m 180 -H "User-Agent: $UA" -o $D/inst/$a.xml https://www.sec.gov/Archives/edgar/data/$CIKN/$n/$x
  echo "$a $f $rd $x $(wc -c < $D/inst/$a.xml 2>/dev/null)"; sleep 0.25
done < $D/accns.txt
