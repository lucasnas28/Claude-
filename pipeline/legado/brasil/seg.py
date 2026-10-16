import glob, os, json, datetime as dt
from lxml import etree
S=__import__('os').environ.get('PIPE_DATA', 'dados')+'/sec/inst'
X='{http://www.xbrl.org/2003/instance}'; XDI='{http://xbrl.org/2006/xbrldi}'
filed={}
for l in open(S+'/accns.txt'):
    a,fm,fd=l.split(); filed[a]=fd
filed['0001045810-21-000010']='2021-02-26'
REVTAGS={'Revenues','RevenueFromContractWithCustomerExcludingAssessedTax'}
def parse(path):
    t=etree.parse(path).getroot()
    ctx={}
    for c in t.iter(X+'context'):
        p=c.find(X+'period'); s=p.find(X+'startDate'); e=p.find(X+'endDate')
        if s is None: continue
        dims=tuple(sorted((m.get('dimension'),m.text) for m in c.iter(XDI+'explicitMember')))
        ctx[c.get('id')]=(s.text,e.text,dims)
    out=[]
    for el in t:
        tag=etree.QName(el).localname
        if tag in REVTAGS and el.get('contextRef') in ctx and el.text:
            s,e,dims=ctx[el.get('contextRef')]
            out.append((tag,s,e,dims,float(el.text)))
    return out
if __name__=='__main__':
    allf={}
    for f in sorted(glob.glob(S+'/0*.xml')):
        a=os.path.basename(f)[:-4]
        allf[a]=parse(f)
    json.dump(allf,open(S+'/rev.json','w'))
    # show dimension sets seen in latest and earliest
    for a in [sorted(allf)[0], '0001045810-24-000124', sorted(allf)[-1]]:
        ds=sorted({d for _,_,_,d,_ in allf[a]})
        print(a); [print('  ',d) for d in ds]
