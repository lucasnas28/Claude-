import sys, shutil, datetime
sys.path.insert(0, __import__('os').environ.get('PIPE_DATA', 'dados')+'/bank')
from ifd import name as ifname
import openpyxl
from openpyxl.utils import get_column_letter as CL
BASE, OUT = sys.argv[1:3]
SH = openpyxl.load_workbook(__import__('os').environ.get('PIPE_DATA', 'dados')+'/bmgb/dl/sh.xlsx', data_only=True)
P = [f'{q}Q{y:02d}' for y in range(20, 27) for q in range(1, 5)][:26]
def ser(sheet, *rows, k=1.0):
    ws = SH[sheet]; out = {}
    hc = {str(ws.cell(1, c).value).strip(): c for c in range(2, ws.max_column + 1) if ws.cell(1, c).value is not None}
    for p in P:
        c = hc.get(f'{p[0]}T{p[2:]}')
        if c is None: continue
        vs = [ws.cell(r, c).value for r in rows]
        if all(not isinstance(v, (int, float)) for v in vs): continue
        out[p] = k * sum(v for v in vs if isinstance(v, (int, float)))
    return out
shutil.copyfile(BASE, OUT)
wb = openpyxl.load_workbook(OUT); inp = wb['Inputs']
COL = {p: 5 + i for i, p in enumerate(P)}
ROW = {inp.cell(r, 2).value: r for r in range(5, 149) if inp.cell(r, 2).value}
written = 0
def put(code, vals, note):
    global written
    r = ROW[code]; inp.cell(r, 4).value = note
    for p, v in vals.items():
        if p in COL and v is not None: inp.cell(r, COL[p]).value = round(float(v), 6); written += 1
def empty(code, why): inp.cell(ROW[code], 4).value = f'[vazio] {why}'
S = 'Bmg_Planilhas_2T26 (RI BMG)'
def g(code, sheet, rows, desc, k=1.0): put(code, ser(sheet, *rows, k=k), f'{S} aba "{sheet}" l.{"+".join(map(str, rows))} · {desc} · fonte · moeda')
# BALANÇO (consolidado; desde 1T21)
g('disp', 'Ativo', [3], 'Disponibilidades')
g('aplic_di', 'Ativo', [5], 'Aplicações em depósitos interfinanceiros')
g('tvm_prop', 'Ativo', [6], 'TVM e instrumentos financeiros derivativos (TOTAL; fonte não separa)')
g('cred_op', 'Ativo', [7], 'Operações com características de concessão de crédito')
g('cred_pdd', 'Ativo', [8], 'Provisão para perdas associadas ao risco de crédito')
g('aplic_rg', 'Ativo', [9], 'Operações de seguros (ativo)')
g('rel_a', 'Ativo', [10], 'Relações interfinanceiras')
g('out_a', 'Ativo', [11, 12], 'Outros créditos + outros valores e bens (inclui ativos fiscais, não abertos)')
g('inv_col', 'Ativo', [15], 'Investimentos'); g('imob', 'Ativo', [16], 'Imobilizado de uso'); g('intang', 'Ativo', [17], 'Intangível')
g('dep_vista', 'Captação', [7], 'Depósito à vista')
g('dep_inter', 'Captação', [5], 'CDI (depósitos interfinanceiros)')
g('dep_prazo', 'Captação', [4, 6], 'CDB + DPGE (depósitos a prazo)')
g('mo_prop', 'Passivo', [5], 'Captações no mercado aberto - carteira própria')
g('div_emis', 'Passivo', [6], 'Recursos de aceites e emissão de títulos (LF sênior, subordinada e perpétua, LCA/LCI: abertura nas memórias)')
g('emp_empr', 'Passivo', [7], 'Obrigações por empréstimos e repasses')
g('deriv_p', 'Passivo', [8], 'Instrumentos financeiros derivativos')
g('prov_tec', 'Passivo', [9], 'Operações de seguros (passivo)')
g('rel_p', 'Passivo', [10], 'Relações interfinanceiras')
g('prov_dem', 'Passivo', [11], 'Provisões')
g('fisc_p', 'Passivo', [12], 'Obrigações fiscais diferidas')
g('out_p', 'Passivo', [13], 'Outras obrigações (inclui cessões de crédito e obrigações fiscais correntes)')
g('res_luc', 'Passivo', [19], 'Patrimônio líquido dos controladores TOTAL (a planilha não abre capital/reservas). Na planilha os rótulos das linhas 17 e 19 estão trocados: l.17 = PL total')
g('minor', 'Passivo', [18], 'Participação de acionistas não controladores')
for c_, why in [('aplic_mo', 'não aberto (dentro de instrumentos financeiros)'), ('aplic_pdd', 'não divulgada'), ('tvm_vinc', 'dentro de tvm_prop'), ('tvm_rg', 'dentro de aplic_rg'), ('tvm_pdd', 'não divulgada'),
                ('deriv_a', 'dentro de tvm_prop'), ('cred_tvm', 'dentro de cred_op'), ('fisc_a', 'dentro de out_a'), ('inv_out', 'dentro de inv_col'), ('dep_poup', 'BMG não capta poupança'),
                ('dep_out', 'não divulgado'), ('mo_terc', 'dentro de mo_prop'), ('mo_livre', 'idem'), ('div_ext', 'dentro de div_emis'), ('div_coe', 'idem'), ('div_sub', 'LF subordinada/perpétua dentro de div_emis (memórias)'),
                ('emp_rep', 'dentro de emp_empr'), ('prov_gar', 'dentro de prov_dem'), ('cap_social', 'PL total em res_luc'), ('res_cap', 'idem'), ('ora', 'idem'), ('tesour', 'idem')]:
    empty(c_, why)
# DRE GERENCIAL (recorrente) + ajuste para o contábil
D = 'DRE Gerencial'
g('rec_cred', D, [4], 'Operações de crédito'); g('rec_tvm', D, [5, 9], 'Resultado com TVM + resultado com derivativos')
g('serv', D, [6], 'Receitas de prestação de serviços'); g('desp_dep', D, [8], 'Captações no mercado'); g('desp_emp', D, [10], 'Empréstimos e repasses')
g('seguros', D, [11], 'Resultado de seguros'); g('pdd_desp', D, [13], 'Provisão para créditos de liquidação duvidosa'); g('pdd_recup', D, [14], 'Recuperação de créditos baixados para prejuízo')
g('out_desp', D, [15], 'Despesas de comissões de agentes (originação de consignado; a Cia as põe no custo do crédito)')
g('pessoal', D, [18, 27], 'Despesas de pessoal + participação nos lucros (PLR incluída aqui para o lucro fechar; também em pes_plr)')
g('admin', D, [19], 'Outras despesas administrativas'); g('desp_trib', D, [20], 'Despesas tributárias'); g('res_part', D, [21], 'Resultado de participações em coligadas e controladas')
g('out_rec', D, [22], 'Outras despesas/receitas operacionais (LÍQUIDO)'); g('ir_corr', D, [26], 'IR e CSLL TOTAL (fonte não separa)')
g('ll_minor', D, [28], 'Participação de não controladores'); g('pes_plr', D, [27], 'Participação nos lucros')
# memórias e ajuste não recorrente
r = 151; inp.cell(153, 3).value = None
inp.cell(r, 3).value = 'MEMÓRIAS DO RI (sem código) e AJUSTE DO GERENCIAL PARA O CONTÁBIL'; r += 1
MEM = [('Indicadores', 2, 'Lucro líquido contábil (aba Indicadores l.2)', 'll_c'), ('Indicadores', 3, 'Lucro líquido recorrente (aba Indicadores l.3)', 'll_r'),
       (D, 29, 'Lucro líquido da DRE gerencial (l.29) - conferência', None), ('Indicadores', 12, 'Inadimplência over 90 (% da carteira; aba Indicadores l.12)', 'npl_pct'),
       ('Margem Financeira', 3, 'Margem financeira gerencial (aba Margem Financeira l.3)', None), ('Margem Financeira', 15, 'Margem financeira % a.a. (l.15)', None),
       ('Carteira', 16, 'Carteira de crédito Brasil (aba Carteira l.16)', None), ('Carteira', 17, 'Empréstimo consignado nos EUA (aba Carteira l.17)', None),
       ('Captação', 9, 'LF sênior (aba Captação l.9)', None), ('Captação', 10, 'LF subordinada (l.10)', None), ('Captação', 11, 'LF perpétua (l.11)', None),
       ('Captação', 13, 'Cessões de crédito e securitização (l.13)', None), ('Captação', 17, 'Captação total (l.17)', None)]
MR = {}
for sh, rr, desc, key in MEM:
    inp.cell(r, 3).value = desc; inp.cell(r, 4).value = f'{S} · memoria'
    for p, v in ser(sh, rr).items():
        if p in COL and isinstance(v, (int, float)): inp.cell(r, COL[p]).value = v
    if key: MR[key] = r
    r += 1
DREC = ['rec_cred', 'rec_tvm', 'serv', 'desp_dep', 'desp_emp', 'seguros', 'pdd_desp', 'pdd_recup', 'out_desp', 'pessoal', 'admin', 'desp_trib', 'res_part', 'out_rec', 'ir_corr', 'll_minor']
inp.cell(ROW['nao_op'], 4).value = (f'{S}: RESÍDUO = lucro contábil (l.{MR["ll_c"]}) menos as demais linhas da DRE gerencial codificadas. Contém o resultado não operacional gerencial (l.24), '
                                    f'os efeitos não recorrentes (contábil - recorrente, líquidos de IR) e o arredondamento da planilha (inteiros) · derivada · moeda')
nop = ser(D, 24)
for p in P:
    if p in nop:
        c = CL(COL[p])
        inp.cell(ROW['nao_op'], COL[p]).value = f'=N({c}{MR["ll_c"]})-(' + '+'.join(f'N({c}{ROW[k]})' for k in DREC) + ')'
inp.cell(ROW['npl90'], 4).value = f'{S}: inadimplência over 90 (%) x carteira de crédito total (l.{MR["npl_pct"]} x cart_ampl) · derivada · moeda'
for p in P:
    c = CL(COL[p])
    if p in ser('Indicadores', 12) and isinstance(ser('Indicadores', 12)[p], (int, float)):
        inp.cell(ROW['npl90'], COL[p]).value = f'=IFERROR({c}{MR["npl_pct"]}*{c}{ROW["cart_ampl"]},"")'
inp.cell(r, 3).value = f'Espaço livre até a linha 190 (linhas {r} a 190) para novas linhas de origem.'
# COMPLEMENTARES
K = 'Carteira'
g('cart_ampl', K, [18], 'Carteira de crédito total (Brasil + consignado EUA)')
g('cart_pf', K, [3, 7, 17], 'PF: consignados + varejo PF + consignado nos EUA')
g('cart_pj', K, [13], 'Atacado (PJ)')
g('cart_consig', K, [3], 'Consignados (cartão consignado + cartão benefício + empréstimo consignado)')
g('cart_cartao', K, [11], 'Cartão de crédito (varejo)')
put('cart_pf_out', {p: ser(K, 7)[p] - (ser(K, 11).get(p) or 0) + (ser(K, 17).get(p) or 0) for p in ser(K, 7)}, f'{S} aba Carteira: varejo PF (l.7) menos cartão (l.11) mais consignado EUA (l.17) · derivada · moeda')
g('cart_mpe', K, [14], 'Empresas'); g('cart_grandes', K, [15], 'Operações estruturadas')
g('at_rent', 'Margem Financeira', [18], 'Ativos rentáveis MÉDIOS (a fonte publica a média)')
g('ib_rep', 'Indicadores', [6], 'Índice de Basileia'); g('in1_rep', 'Indicadores', [7], 'Tier I')
PRU = 'C0080178'
for code, pat in [('rwa', r'^Ativos Ponderados pelo Risco'), ('rwa_cred', r'^RWA para Risco de Crédito'), ('rwa_mkt', r'^RWA para Risco de Mercado'), ('rwa_op', r'^RWA para Risco Operacional'),
                  ('pr', r'^Patrimônio de Referência para Comparação'), ('nivel1', r'^Patrimônio de Referência Nível I'), ('cet1', r'^Capital Principal para Comparação')]:
    put(code, {p: ifname(p, 5, PRU, pat) for p in P if p >= '1Q21'[:0] and int(p[2:]) >= 21}, f'IF.data/BCB rel. 5, conglomerado prudencial {PRU}: "{pat[1:]}" (a planilha só traz os índices) · fonte · moeda')
put('icp_rep', {p: (ifname(p, 5, PRU, r'^Índice de Capital Principal') or 0) * 1e6 for p in P if int(p[2:]) >= 21}, f'IF.data/BCB rel. 5 ({PRU}): índice de capital principal · fonte · pct')
put('alav', {p: (ifname(p, 5, PRU, r'^Razão de Alavancagem') or 0) * 1e6 for p in P if int(p[2:]) >= 21}, f'IF.data/BCB rel. 5 ({PRU}): razão de alavancagem · fonte · pct')
for c_, why in [('cart_veic', 'não opera'), ('cart_imob', 'não opera'), ('cart_rural', 'não opera'), ('cart_pj_out', 'atacado aberto em empresas + estruturadas'),
                ('est1', 'não está na planilha'), ('est2', 'idem'), ('est3', 'idem'), ('prov_e1', 'idem'), ('prov_e2', 'idem'), ('prov_e3', 'idem'),
                ('prov_tot', 'provisão do balanço em cred_pdd'), ('npl15', 'não divulgado'), ('wo', 'não divulgado'), ('reneg', 'não divulgado'),
                ('mf_cli', 'só margem total (memória)'), ('mf_mkt', 'idem'), ('pas_oner', 'não divulgado'), ('lcr', 'não está na planilha'), ('nsfr', 'idem'),
                ('acoes', 'não está na planilha (preencher)'), ('div_decl', 'idem'), ('jcp_decl', 'idem'), ('preco', 'idem'), ('funcion', 'idem'), ('agencias', 'idem'), ('clientes', 'idem'),
                ('rec_aplic', 'dentro de TVM'), ('rec_rg', 'dentro de seguros'), ('desp_div', 'dentro de captações'), ('desp_prov_tec', 'dentro de seguros'),
                ('prov_civ', 'dentro de outras receitas/despesas'), ('prov_trab', 'idem'), ('prov_fisc', 'idem'), ('ir_dif', 'IR total em ir_corr')]:
    empty(c_, why)
for c_ in ['serv_tarifas', 'serv_cartoes', 'serv_adm_rec', 'serv_seg_corr', 'serv_merc_cap', 'serv_out', 'pes_proventos', 'pes_benef', 'pes_out', 'adm_proc', 'adm_alug', 'adm_mkt', 'adm_terc', 'adm_deprec', 'adm_out']:
    empty(c_, 'a planilha não abre')
inp['E197'] = 'Banco BMG S.A.'; inp['E198'] = 'BMGB4'
inp['E199'] = 'Banco de varejo (consignado) / BR GAAP (COSIF); DRE gerencial ajustada ao lucro contábil'
inp['E203'] = datetime.datetime(2026, 6, 30); inp['E204'] = datetime.datetime(2021, 3, 31)
inp['E203'].number_format = inp['E204'].number_format = 'dd/mm/yyyy'
inp['F1'] = ('BMG (BMGB4). Fonte: Bmg_Planilhas_2T26 do RI (bancobmg.mzweb.com.br): balanço consolidado, DRE gerencial (recorrente) + ajuste não recorrente para o lucro contábil, '
             'carteira, captação, indicadores; capital (RWA/PR) do IF.data/BCB. Série desde 1T21 (a planilha começa aí; 2020 vazio). R$ milhões inteiros (arredondamento de ±1).')
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('written', written, 'memo end', r)
