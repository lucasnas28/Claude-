import sys, collections; sys.path.insert(0,__import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from sx import *
c=Co(sys.argv[1]); tags=sys.argv[2].split(',') if len(sys.argv)>2 else ['Revenues','RevenueFromContractWithCustomerExcludingAssessedTax','OperatingIncomeLoss']
cc=collections.defaultdict(set)
for (n,ns,s,e,dd,v,u,a,fl,i) in c.facts:
    if n in tags and s and dd:
        mem=tuple(sorted(m for ax,m in dd if not (ax=='srt:ConsolidationItemsAxis' and m=='us-gaap:OperatingSegmentsMember')))
        if len(mem)<=2: cc[(n[:12],mem)].add(label(e))
for k,v in sorted(cc.items(), key=lambda x:-len(x[1])):
    if len(v)>=4: print(len(v), min(v,key=lambda l:(l[2:],l[0])), max(v,key=lambda l:(l[2:],l[0])), k)
