"""Atualização mensal do Tabelão: PDFs do mês + VIMOB -> dados.json.

1. Ajuste o bloco CFG e o PAY (rótulos do fluxo, copiados do cabeçalho de cada PDF do mês).
2. python3 atualizar_mes.py      -> reescreve dados.json (tabelas, widget, datas) e imprime o log de divergências.
3. python3 montar.py             -> gera ../index.html
Natury: mantida como está (valores do VIMOB) até existir tabela oficial; para atualizar, ver MANUAL_Tabelao_Yuny.md.
Leia sempre os Excel com engine='calamine'.
"""
import os, sys
CFG = {
    'pasta_pdf': 'pdf',                                           # pasta com os PDFs do mês
    'mes_pdf': 'NOVEMBRO - 2026',                                 # sufixo do nome dos PDFs: "<EMPREENDIMENTO> - <mes_pdf>.pdf"
    'vimob': 'Yuny_Estoque_Vendas_Unificado_2026-11.xlsx',        # aba Base Unificada
    'tabela_vendas': 'Tabela de vendas_3009.xlsx',                # coluna "Status - VIMOB" (sub-status das reservas)
    'mes': 'Novembro 2026', 'atualizado': '11/11/2026', 'status': '10/11/2026',
}
STD = ['1x Ato 11/2026', '1x Financiamento 02/2027', 'Total']
import re, json, unicodedata, subprocess
import pandas as pd
def strip(s): return unicodedata.normalize('NFD', s).encode('ascii','ignore').decode().upper()
def nu(u):
    d=re.sub(r'\D','',str(u)); return d.lstrip('0') or '0'
def obj_json(h, chave, abre, fecha):
    m = h.index(chave) + len(chave) - 1; d = 0; i = m
    while True:
        c = h[i]
        if c == abre: d += 1
        elif c == fecha:
            d -= 1
            if d == 0: break
        i += 1
    return m, i + 1
IDMAP={'camino':'CAMINO ALTO DA BOA VISTA','uape':'UAPÉ POR YUNY','jer155':'155 JERÔNIMO','jer177':'177 JERÔNIMO','zip':'ZIP BROOKLIN','organy':'ORGANY','esquina':'ESQUINA JARDYN','quaddra':'QUADDRA LORENA','uwin':'UWIN BROOKLIN','apeVilaEma':'APÊ VILA EMA','apeJardimSul':'APÊ JARDIM SUL'}
ORDEM=list(IDMAP)
def load_tv():
    df=pd.read_excel(pd.ExcelFile(CFG["tabela_vendas"],engine='calamine'),sheet_name='Tabela de vendas',header=0)
    return df[df['EST_CH_TIPOESTRUTURA']=='U']
def load_un():
    return pd.read_excel(pd.ExcelFile(CFG["vimob"],engine='calamine'),sheet_name='Base Unificada')
SMAP = {'LIVRE':'l','VENDIDA':'v','PERMUTADA':'p','INDISPONIVEL':'i','EM NEGOCIACAO':'n',
        'ENVIADA PARA APROVACAO':'a','ENVIADA P/ APROVACAO':'a','CONTRATO RODADO':'c',
        'EM TRATATIVA':'t','RESERVADA':'o','OUTRO':'o'}
SPAT = '|'.join(sorted(map(re.escape, SMAP), key=len, reverse=True))
ROW = re.compile(r'^\s*(\S+)\s+(' + SPAT + r')\s+([\d.,]+)\s*m\s+(.*)$', re.I)
def num(s): return float(s.replace('.', '').replace(',', '.'))
def parse_pdf(path, nome, npay):
    txt = subprocess.run(['pdftotext', '-layout', path, '-'], capture_output=True, text=True).stdout
    grupos, cur, ruins = [], None, []
    for ln in txt.splitlines():
        m = ROW.match(strip(ln))
        if m:
            un = ln.split()[0]
            vals = [num(x) for x in re.findall(r'\d[\d.]*(?:,\d+)?', m.group(4))]
            if len(vals) != npay + 1:
                ruins.append((un, ln.strip())); continue
            if cur is None: cur = {'cat': None, 'rows': []}; grupos.append(cur)
            cur['rows'].append([un, SMAP[m.group(2).upper()], num(m.group(3)), vals[0], vals[1:]])
        else:
            s = ln.strip()
            if (s and re.fullmatch(r"[A-ZÀ-Ü0-9 ()\-/.]+", s) and len(s) > 3
                    and s != nome and 'FORMA' not in s
                    and 'PRIVATIVA' not in strip(s) and 'UNIDADE' not in strip(s)
                    and not re.fullmatch(r'[\d.,% ]+', s)
                    and not re.search(r'\d{2}/\d{2}/\d{4}', s)):
                cur = {'cat': s, 'rows': []}; grupos.append(cur)
    return [g for g in grupos if g['rows']], ruins
PDFS={'camino':'CAMINO','uape':'UAPÉ','jer155':'155_JERÔNIMO','jer177':'177_JERÔNIMO','zip':'ZIP BROOKLIN','organy':'ORGANY','esquina':'ESQUINA','quaddra':'QUADDRA LORENA','uwin':'UWIN','apeVilaEma':'APÊ_VILA EMA','apeJardimSul':'APÊ_JARDIM SUL'}
PDFNOME={'camino':'CAMINO ALTO DA BOA VISTA','uape':'UAPÉ POR YUNY','jer155':'155 JERÔNIMO','jer177':'177 JERÔNIMO','zip':'ZIP BROOKLIN','organy':'ORGANY','esquina':'ESQUINA JARDYN','quaddra':'QUADDRA LORENA','uwin':'UWIN BROOKLIN','apeVilaEma':'APÊ VILA EMA','apeJardimSul':'APÊ JARDIM SUL'}
def parse_all():
    out={}
    for pid,f in PDFS.items():
        out[pid]=parse_pdf(os.path.join(CFG['pasta_pdf'],f'{f} - {CFG["mes_pdf"]}.pdf'),PDFNOME[pid],len(PAY[pid]))
    return out
RAW={'DISPONIVEL':'l','VENDIDA':'v','PERMUTADA':'p','INDISPONIVEL':'i','RESERVADA':'R'}
SUB={'EM TRATATIVA':'t','CONTRATO RODADO':'c','RESERVADA':'o'}
def status_maps():
    tv=load_tv(); un=load_un(); un=un[un['Tipo']=='Unidade']
    un=un[~un['Empreendimento'].str.contains('ESPELHO')]
    un['emp']=un['Empreendimento'].replace({'UAPÉ':'UAPÉ POR YUNY','QUADDRA JESUÍNO':'QUADDRA JESUINO'})
    tvsub={(e,b,nu(u)):s for e,b,u,s in zip(tv.Empreendimento,tv.Bloco,tv.Unidade,tv['Status - VIMOB'])}
    st={}
    for e,b,u,s in zip(un.emp,un.Bloco,un.Unidade,un.Status):
        st[(e,b,nu(u))]=RAW[strip(s)]
    return st,tvsub,un,tv
def build_T():
    P=parse_all(); st,tvsub,un,tv=status_maps(); T={}; log=[]
    for pid in ORDEM:
        g,_=P[pid]; emp=IDMAP[pid]
        for gg in g:
            for x in gg['rows']:
                b = 'APÊ JARDIM SUL' if gg['cat'] is None else gg['cat']
                k=(emp,b,nu(x[0])); pdfst=x[1]
                s=st.get(k)
                if s is None: log.append(('sem_planilha',pid,gg['cat'],x[0],pdfst)); continue
                if s=='R':
                    sub=SUB.get(tvsub.get(k)); s = sub or (pdfst if pdfst in 'cto' else 'o')
                if s!=pdfst: log.append(('pdf_vs_plan',pid,gg['cat'],x[0],pdfst,s))
                x[1]=s
        T[pid]={'payLabels':PAY[pid],'grupos':g}
        pdfkeys={( 'APÊ JARDIM SUL' if gg['cat'] is None else gg['cat'], nu(x[0])) for gg in g for x in gg['rows']}
        for (e,b,u),s in st.items():
            if e==emp and s in ('l','R') and (b,u) not in pdfkeys and b!='ESCRITÓRIOS':
                log.append(('disp_sem_pdf',pid,b,u,s))
    return T,log
HID = {'v','p','i'}
WLAB={'l':'LIVRE','t':'EM TRATATIVA','o':'RESERVADA','c':'CONTRATO RODADO','a':'ENVIADA P/ APROVAÇÃO','n':'EM NEGOCIAÇÃO'}
EXCL=('NATURY','JESU','1125')
def build_W(T, W_old):
    st,tvsub,un,tv=status_maps()
    card={}
    for p in T:
        for g in T[p]['grupos']:
            for x in g['rows']: card[(IDMAP[p],g['cat'] or 'APÊ JARDIM SUL',nu(x[0]))]=x
    out=[]
    for e in W_old:
        emp=e['emp']
        if any(z in strip(emp) for z in ('NATURY','JESUINO','1125')): continue
        if emp=='QUADDRA LORENA':
            torres=[]
            for g in T['quaddra']['grupos']:
                us=[{'u':('LOJA '+x[0].zfill(2)) if 'LOJA' in g['cat'] else nu(x[0]),'st':WLAB[x[1]],'metr':x[2],'m2':x[3],'total':x[4][-1]} for x in g['rows'] if x[1] not in HID]
                if us: torres.append({'torre':g['cat'],'unidades':us})
        else:
            sub=un[un.emp==emp]; tb={}
            for b,u,a,v,s in zip(sub.Bloco,sub.Unidade,sub['Área m²'],sub['Valor Atual'],sub.Status):
                k=(emp,b,nu(u)); c=RAW[strip(s)]
                if c in HID: continue
                if k in card and card[k][1] not in HID: c=card[k][1]
                elif c=='R': c=SUB.get(tvsub.get(k),'o')
                elif k in card: c=card[k][1] if card[k][1] not in HID else c
                tb.setdefault(b,[]).append({'u':nu(u),'st':WLAB[c],'metr':round(float(a),2),'m2':round(float(v)/float(a),2),'total':round(float(v),2)})
            torres=[{'torre':b,'unidades':tb[b]} for b in sorted(tb)]
        if torres: out.append({'emp':emp,'torres':torres})
    return out

PAY = {k: STD for k in PDFS}
PAY['jer177'] = ['1x Ato 11/2026', '2x Sinal/Entrada 12/2026', '8x Mensal 02/2027', '1x Única 11/2027', '1x Financiamento 01/2028', 'Total']
PAY['quaddra'] = ['1x Ato nov/26', '2x Sinal/Entrada dez/26', '19x Mensal fev/27', '1x Única mar/27', '1x Única mar/28', '1x Única out/28', '1x Financiamento dez/28', 'Total']

if __name__ == '__main__':
    import json
    AQUI = os.path.dirname(os.path.abspath(__file__))
    arq = os.path.join(AQUI, 'dados.json')
    d = json.load(open(arq, encoding='utf-8'))
    T, log = build_T()
    for p in d['projetos']:
        if p['id'] in T: p['tabela'] = T[p['id']]
    W = build_W(T, d['widget'])
    mantidos = [e for e in d['widget'] if 'NATURY' in strip(e['emp'])]   # Natury fica como está
    d['widget'] = W + mantidos
    d['mes'], d['atualizado'], d['status'] = CFG['mes'], CFG['atualizado'], CFG['status']
    saida = sys.argv[1] if len(sys.argv) > 1 else arq
    json.dump(d, open(saida, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    for x in log: print(*x)
    # conferência: área × R$/m² tem que fechar com o total (erro de PDF já aconteceu em Camino 116 e Jardim Sul 813)
    for pid, t in T.items():
        for g in t['grupos']:
            for u, st_, a, m2, v in g['rows']:
                if st_ not in HID and a and m2 and abs(a * m2 - v[-1]) / v[-1] > 0.01:
                    print('CONFERIR', pid, g['cat'], u, f'área {a} × R$/m² {m2} = {a*m2:,.0f} ≠ total {v[-1]:,.0f}  (vale o total do VIMOB; corrigir o campo errado)')
    print('dados gravados em', saida)
