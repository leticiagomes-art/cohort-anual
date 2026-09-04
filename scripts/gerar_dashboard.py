#!/usr/bin/env python3
"""
gerar_dashboard.py
------------------
Lê os CSVs acumulados (data/{pedidos,reembolsos,chargebacks}/AAAA-MM.csv.gz),
monta o payload do cohort semanal e injeta no template do dashboard.

Este repositório é dedicado só à análise de cohort — o dashboard mostra
uma única visão (Cohort · Refund rate / Chargeback rate por semana de
venda, geral e por produto). Não computa mais Gross/reembolso mensal,
afiliados, decisor, MaxWeb, reconciliação etc.: nada disso alimenta o
template atual, e manter esse cálculo só infla o payload e o tempo de
execução à toa.

Saída: index.html (raiz do repo, servido pelo GitHub Pages)

Uso:
    python scripts/gerar_dashboard.py

Dependências: pandas
"""
import sys, json, re, unicodedata
from pathlib import Path
from datetime import datetime
import pandas as pd

ROOT       = Path(__file__).resolve().parent.parent
DATA_DIR   = ROOT / "data"
TEMPLATE   = ROOT / "scripts" / "dashboard_template.html"
OUT        = ROOT / "index.html"

def norm(s):
    t = unicodedata.normalize('NFKD', str(s).lower()).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]', '', t)).strip()

def sf(v):
    """safe float — NaN/None → None"""
    try:
        f = float(v)
        return None if (f != f) else round(f, 6)
    except (TypeError, ValueError):
        return None

# Data em que cada produto virou funil 100% White (ou 50/50 no caso do
# BreathEaseX — ver README/CLAUDE.md). Usada só pra marcar cada semana do
# cohort como Black/White na coluna BW da matriz.
BW = {
    'BreathEaseX': '2026-08-12', 'NervoLyn': '2026-07-24',
    'Prostafense':  '2026-07-27', 'AudiLeaf': '2026-07-28',
    'VisiumPro':    '2026-07-28', 'MaroBrain': '2026-07-28',
    'GlucoRecover': '2026-07-31', 'FloraNew':  '2026-07-31',
    'AlphaErec':    '2026-08-04', 'BoosterXT': '2026-08-07',
    'NailsCleanPro':'2026-08-12', 'LipoBliss': '2026-08-13',
    'LipoPeak':     '2026-08-14', 'MounjaMelt':'2026-08-14',
    'ReduTide':     '2026-07-07', 'RedurBurn': '2026-07-07',
    'LipoVive':     '2026-08-14', 'VigorLong': '2026-08-07',
    'Lipotrine':    '2026-08-14',
}

# ── 1. carregar CSVs ──────────────────────────────────────────────────────────
print("→ Carregando CSVs particionados...")

def carregar_particoes(subpasta):
    """Lê e concatena todas as partições mensais de uma base (data/{subpasta}/AAAA-MM.csv.gz)."""
    pasta = DATA_DIR / subpasta
    if pasta.is_dir():
        arquivos = sorted(pasta.glob("*.csv.gz"))
        if arquivos:
            partes = [pd.read_csv(a, low_memory=False) for a in arquivos]
            df = pd.concat(partes, ignore_index=True)
            print(f"   {subpasta}: {len(arquivos)} mês(es), {len(df):,} linhas")
            return df
    legado = DATA_DIR / f"base_{subpasta}.csv.gz"
    if legado.exists():
        df = pd.read_csv(legado, low_memory=False)
        print(f"   {subpasta}: arquivo único (legado), {len(df):,} linhas")
        return df
    print(f"   {subpasta}: NENHUM DADO ENCONTRADO")
    return pd.DataFrame()

ped = carregar_particoes("pedidos")
ref = carregar_particoes("reembolsos")
cb  = carregar_particoes("chargebacks")

if not len(ped):
    print("ERRO: base de pedidos vazia — nada a gerar")
    sys.exit(1)

for d in (ped, ref, cb):
    # format='mixed' e obrigatorio aqui: o CSV acumulado mistura datas
    # "AAAA-MM-DD" (sem hora) com "AAAA-MM-DD HH:MM:SS" (com hora) na mesma
    # coluna, dependendo de qual arquivo/merge escreveu cada linha. Sem
    # format='mixed', pd.to_datetime trava no formato do PRIMEIRO valor nao
    # nulo e vira NaT em todo o resto.
    d['data_pedido'] = pd.to_datetime(d['data_pedido'], errors='coerce', format='mixed')
    if 'data_evento' in d.columns:
        d['data_evento'] = pd.to_datetime(d['data_evento'], errors='coerce', format='mixed')

# data_evento do reembolso e, por definicao, o refund_date. Linhas antigas do
# historico podem ter data_evento nula mesmo com refund_date preenchido (o
# merge_csvs.py so concatena, nunca recalcula coluna derivada de linha
# antiga) — recalcular sempre a partir do refund_date evita o cohort inteiro
# dar ~0 quando isso acontece.
if 'refund_date' in ref.columns:
    ref['data_evento'] = pd.to_datetime(ref['refund_date'], errors='coerce', format='mixed')

# âncora: última data com dado nos arquivos (não data de hoje)
DATA_REF = ped['data_pedido'].max().normalize()
ANO      = DATA_REF.year
print(f"   data de referência: {DATA_REF.date()} | ano: {ANO}")

ped_a = ped[ped['data_pedido'].dt.year == ANO].copy()
ref_a = ref[ref['data_pedido'].dt.year == ANO].copy()
cb_a  = cb[cb['data_pedido'].dt.year  == ANO].copy()

# semana_ini/semana_pedido_ini já estão nos CSVs — garantir dtype
for d in (ped_a, ref_a, cb_a):
    d['semana_ini'] = pd.to_datetime(d['semana_ini'], errors='coerce')
    if 'semana_pedido_ini' in d.columns:
        d['semana_pedido_ini'] = pd.to_datetime(d['semana_pedido_ini'], errors='coerce')

# ── 2. cohort ─────────────────────────────────────────────────────────────────
print("→ Cohort...")
prods = list(ped_a.groupby('produto')['amount'].sum().sort_values(ascending=False).index)

def mk_cohort(ped_b, ev_b, max_w=12):
    rows = []
    prod_sem = ped_b.groupby('semana_ini').apply(
        lambda g: g['produto'].mode().iloc[0] if len(g) else ''
    ).to_dict()
    for sem in sorted(ped_b['semana_ini'].dropna().unique()):
        pp = ped_b[ped_b['semana_ini'] == sem]
        n  = int(pp['order_id'].nunique())
        g  = float(pp['amount'].sum())
        if g < 1: continue
        d_ = int((DATA_REF - sem).days // 7)
        # semana_pedido_ini (semana do PEDIDO), nao semana_ini (semana do
        # EVENTO) — sao datas diferentes pra reembolso/chargeback, e usar a
        # errada aqui fazia o cohort inteiro dar ~0: o join so "acertava"
        # quando o reembolso, por coincidencia, acontecia na mesma semana
        # do calendario que a safra sendo iterada, em vez de olhar reembolsos
        # de QUALQUER semana que pertencam a essa safra de pedidos.
        col_sem_evento = 'semana_pedido_ini' if 'semana_pedido_ini' in ev_b.columns else 'semana_ini'
        ev_sem = ev_b[ev_b[col_sem_evento] == sem]
        ws = {}
        for w in range(1, max_w + 1):
            wf  = sem + pd.Timedelta(weeks=w)
            cnt = float(ev_sem.loc[ev_sem['data_evento'] < wf, 'amount'].sum()) \
                  if 'data_evento' in ev_sem.columns else 0
            ws[f'W{w}'] = round(cnt / g, 4) if g > 0 and w <= d_ + 1 else None
        prod = prod_sem.get(sem, '')
        virada = BW.get(prod)
        bw = 'white' if (virada and sem >= pd.Timestamp(virada)) else 'black'
        rows.append({
            's':  str(sem)[:10],
            'sf': pd.Timestamp(sem).strftime('%b %d, %Y'),
            'n': n, 'g': round(g, 2), 'd': d_, 'bw': bw, **ws
        })
    return rows

coh_rf = mk_cohort(ped_a, ref_a)
coh_cb = mk_cohort(ped_a, cb_a)

prf, pcb = {}, {}
for p in prods:
    prf[p] = mk_cohort(ped_a[ped_a['produto'] == p], ref_a[ref_a['produto'] == p])
    pcb[p] = mk_cohort(ped_a[ped_a['produto'] == p], cb_a[cb_a['produto']  == p])
    for row in prf[p] + pcb[p]:
        virada = BW.get(p)
        row['bw'] = 'white' if (virada and pd.Timestamp(row['s']) >= pd.Timestamp(virada)) else 'black'

# curvas de velocidade (Jan-Mai vs Jun-15Jul, pra comparar a aceleração)
def curva(rows, m0, m1):
    sel = [r for r in rows if m0 <= r['s'] <= m1 and r['g'] > 0]
    if not sel: return [0] * 13
    gt = sum(r['g'] for r in sel)
    return [0] + [round(sum((r.get(f'W{i+1}') or 0) * r['g'] for r in sel) / gt, 4)
                  for i in range(12)]

yr = str(ANO)
curva_a = curva(coh_rf, f'{yr}-01-01', f'{yr}-05-31')
curva_b = curva(coh_rf, f'{yr}-06-01', f'{yr}-07-15')

# cohortResumo (W+0/W+4/W+8 ponderado, semanas maduras ≥ 8)
cohort_resumo = {}
for p in prods:
    rows_p = prf[p]
    n = len(rows_p)
    mad = [r for i, r in enumerate(rows_p) if (n - 1 - i) >= 8 and (r['g'] or 0) > 0]
    if not mad: continue
    tot = sum(r['g'] for r in mad)
    cohort_resumo[p] = {
        'w0':     round(sum((r.get('W1') or 0) * r['g'] for r in mad) / tot, 5),
        'w4':     round(sum((r.get('W4')  or 0) * r['g'] for r in mad) / tot, 5),
        'w8':     round(sum((r.get('W8')  or 0) * r['g'] for r in mad) / tot, 5),
        'gross':  round(tot, 2),
        'semanas': len(mad),
    }

# ── 3. montar payload final ────────────────────────────────────────────────────
print("→ Montando payload...")
P = {
    'cohortResumo': cohort_resumo,
    'cohortEx': {
        'rf':    coh_rf, 'cb': coh_cb,
        'prods': prods,  'prf': prf, 'pcb': pcb,
        'bw':    BW,
        'curvaA': curva_a, 'curvaB': curva_b,
    },
    'curvas': {
        'colsW': [f'W+{i}' for i in range(13)],
        'A': curva_a, 'B': curva_b,
    },
    'gerado_em': str(DATA_REF.date()),
}

# ── 4. injetar no template ─────────────────────────────────────────────────────
print("→ Gerando HTML...")
if not TEMPLATE.exists():
    print(f"ERRO: template não encontrado em {TEMPLATE}")
    sys.exit(1)

template_html = TEMPLATE.read_text(encoding='utf-8')
payload_json  = json.dumps(P, ensure_ascii=False, separators=(',', ':'), default=str)
html          = template_html.replace('__PAYLOAD__', payload_json)

OUT.write_text(html, encoding='utf-8')
kb = OUT.stat().st_size // 1024
print(f"✓ {OUT}  ({kb} KB)  |  data ref: {DATA_REF.date()}")
