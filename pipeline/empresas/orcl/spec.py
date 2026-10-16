import sys; sys.path.insert(0, __import__('os').environ.get('PIPE_SEC', 'pipeline/sec')); from segspec import make; from rpo import inst
RV = ('RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues')
globals().update(make(__import__('os').environ.get('PIPE_DATA', 'dados')+'/orcl/dl',
 'DRIVERS DE RECEITA (codigos 45+): segmentos Cloud & Software (antes Cloud & License), Hardware e Serviços (receita e margem do segmento); abertura de cloud (OCI x aplicações) e software; regiões; RPO (carteira contratada) - SEC XBRL (USD, fator 0,000001).',
 'Notas: valores das instâncias XBRL. Ano fiscal termina em 31/mai: ago -> 3T, nov -> 4T, fev -> 1T, mai -> 2T do calendário. O trimestre encerrado em ago/26 (1T do ano fiscal 2027) fica fora da grade, que termina no 2T26. A partir do ano fiscal 2026 a Oracle passou a abrir Cloud (infraestrutura + aplicações) separado de Software (licença + suporte), com comparativos desde 2T24; antes, a abertura era Cloud services & license support (incluía suporte de licenças) e Cloud license & on-premise license. As duas aberturas ficam em linhas separadas porque o escopo é diferente. Lucro do segmento = margem de contribuição divulgada pela Oracle (antes de P&D, G&A, amortização e reestruturação).',
 [('cs', 'Cloud & Software (antes Cloud & License)', ['orcl:CloudAndSoftwareBusinessMember', 'orcl:CloudAndLicenseBusinessMember'], True),
  ('hw', 'Hardware', ['orcl:HardwareBusinessMember'], True), ('sv', 'Serviços', ['orcl:ServicesBusinessMember'], True)],
 EXTRA=[('cl', 'Nova abertura - Cloud total', RV, ['orcl:CloudRevenuesMember']), ('oci', 'Nova abertura - Cloud Infrastructure (OCI)', RV, ['orcl:CloudInfrastructureMember']),
        ('capp', 'Nova abertura - Cloud Applications', RV, ['orcl:CloudApplicationsMember']), ('sws', 'Nova abertura - Software support', RV, ['orcl:SoftwareSupportMember']),
        ('swl', 'Nova abertura - Software license', RV, ['orcl:SoftwareLicenseMember']),
        ('csls', 'Abertura antiga - Cloud services & license support', RV, ['orcl:CloudServicesAndLicenseSupportMember']),
        ('icsls', 'Abertura antiga - Infrastructure cloud services & license support', RV, ['orcl:InfrastructureCloudServicesAndLicenseSupportMember']),
        ('acsls', 'Abertura antiga - Applications cloud services & license support', RV, ['orcl:ApplicationsCloudServicesAndLicenseSupportMember']),
        ('am', 'Américas', RV, ['srt:AmericasMember']), ('emea', 'EMEA', RV, ['us-gaap:EMEAMember']), ('apac', 'Ásia-Pacífico', RV, ['srt:AsiaPacificMember'])],
 RV=RV))
_items = items
def items(co):
    return _items(co) + [('rpo', 'RPO - obrigações de desempenho remanescentes, saldo no fim do período (USD)', 'moeda', inst(co, ['RevenueRemainingPerformanceObligation']), '#,##0', 'last')]
BLOCK += '''
rr += 1
hdr_row(rr, 'Carteira contratada (USD milhões)'); rr += 1
line('rpo', '   RPO - carteira contratada (saldo)', annual='last')
line('rpo_y', '   % YoY RPO', '0.0%', 'calc', yoy(R['rpo']))
'''
