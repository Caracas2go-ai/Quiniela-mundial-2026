#!/usr/bin/env python3
"""
export_data.py — Exporta data COMPACTA de la Quiniela Mundial 2026 para la web app.
Optimizado para caber en GitHub vía MCP: combos como datos estructurados (no HTML),
picks como arrays por índice de partido, sin CSS ni logo embebidos.

Uso:
  python3 export_data.py --data-dir "<QUINIELA MUNDIAL 2026>" \
      --ranking-json "_ranking.json" --out "quiniela-web/data.json"
"""
import argparse, json, os, sys, glob
from datetime import datetime, timedelta, timezone
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generar_reporte_quiniela as g  # noqa: E402


def norm(s):
    return ''.join(ch for ch in s.lower().strip() if ch.isalnum())


def kickoff_iso(d, ts):
    frac = float(ts) % 1
    total = round(frac * 1440)
    h, m = (total // 60) % 24, total % 60
    return datetime(d.year, d.month, d.day, h, m, tzinfo=timezone(timedelta(hours=-4))).isoformat()


def mid(m):
    return f"{m['group']}_{m['round']}_{g.FLAG_MAP.get(m['home'],'un')}_{g.FLAG_MAP.get(m['away'],'un')}"


def _mult(factor, n):
    v = round(factor ** max(n, 1))
    return f'{v:,}x'.replace(',', '.') if v >= 1000 else f'{max(v,2)}x'


def _read_results(dd):
    """Resultados reales por nº de partido desde la hoja Fixture (A=nº, D=GL, E=GV)."""
    p = os.path.join(dd, 'ADMIN.xlsx')
    try:
        rows = g.read_xlsx_sheet(p, 'Fixture')
    except Exception:
        return {}
    out = {}
    for r_idx, cells in rows:
        a = (cells.get('A', '') or '').strip()
        try:
            num = int(float(a))
        except (ValueError, TypeError):
            continue
        d = (cells.get('D', '') or '').strip(); e = (cells.get('E', '') or '').strip()
        if d in ('', '-') or e in ('', '-'):
            continue
        try:
            out[num] = (int(float(d)), int(float(e)))
        except (ValueError, TypeError):
            pass
    return out


def _read_teams(dd):
    """48 equipos por grupo desde la hoja Equipos (A=Num, B=Nombre, C=Grupo, D=Rank)."""
    p = os.path.join(dd, 'ADMIN.xlsx')
    try:
        rows = g.read_xlsx_sheet(p, 'Equipos')
    except Exception:
        return []
    out = []
    for r_idx, cells in rows:
        a = (cells.get('A', '') or '').strip()
        nm = (cells.get('B', '') or '').strip()
        gr = (cells.get('C', '') or '').strip()
        rk = (cells.get('D', '') or '').strip()
        try:
            int(float(a))
        except (ValueError, TypeError):
            continue
        if not nm or nm == 'NombreEquipo':
            continue
        try:
            rk = int(float(rk))
        except (ValueError, TypeError):
            rk = None
        out.append({'name': nm, 'group': gr, 'rank': rk, 'code': g.FLAG_MAP.get(nm, 'un')})
    return out[:48]


# ── COMBOS ESTRUCTURADOS (misma lógica que el reporte, salida compacta) ─────────

def combos_1x2(ids, M):
    N = len(ids)
    def oc(i): return M[i]['oc']
    def top(i):
        o = oc(i); return max(o, key=o.get) if sum(o.values()) else '1'
    def least(i):
        o = oc(i); return min(o, key=o.get) if sum(o.values()) else '2'
    def strg(i):
        o = oc(i); t = sum(o.values()); return (max(o.values()) / t) if t else 0
    def inv(x): return '2' if x == '1' else ('1' if x == '2' else 'X')
    def lbl(i, out):
        m = M[i]; return f"Gana {m['home']}" if out == '1' else (f"Gana {m['away']}" if out == '2' else 'Empate')
    cons = [top(i) for i in ids]
    div = sorted(range(N), key=lambda k: strg(ids[k]))
    s1 = list(cons)
    s2 = list(cons);  s2[div[0]] = 'X' if div else None
    s3 = list(cons)
    for k in div[:2]: s3[k] = 'X'
    s4 = list(cons)
    if div: s4[div[0]] = inv(cons[div[0]])
    s5 = [inv(o) for o in cons]
    s6 = [least(i) for i in ids]
    def legs(outs): return [[ids[k], lbl(ids[k], outs[k])] for k in range(N)]
    D = [('g', '🛡️ La sólida', 1.5, 'rk-bajo', 'Riesgo bajo', s1),
         ('b', '⚖️ Favoritos + 1 empate', 1.9, 'rk-medio', 'Riesgo medio', s2),
         ('t', '🎯 Doble empate', 2.2, 'rk-alto', 'Riesgo medio-alto', s3),
         ('a', '🔥 Con sorpresa', 2.5, 'rk-alto', 'Riesgo alto', s4),
         ('p', '🔄 Contraataque', 3.2, 'rk-extremo', 'Riesgo extremo', s5),
         ('r', '💣 Todo o nada', 4.2, 'rk-extremo', 'Riesgo extremo', s6)]
    return [{'cls': c, 'n': n, 'm': _mult(f, N), 'rc': rc, 'rl': rl, 'legs': legs(o)} for c, n, f, rc, rl, o in D]


def combos_exacto(ids, M):
    N = len(ids)
    def srank(i, r):
        top = M[i].get('scoreTop', [])
        return top[r][0] if r < len(top) else (top[-1][0] if top else '1-1')
    def low(i):
        top = M[i].get('scoreTop', [])
        if not top: return '0-0'
        def gg(s):
            try: h, a = s.split('-'); return int(h) + int(a)
            except Exception: return 9
        return sorted(top, key=lambda t: gg(t[0]))[0][0]
    def legs(fn): return [[ids[k], fn(ids[k], k)] for k in range(N)]
    D = [('g', '🎯 Marcador favorito', 3.0, 'rk-alto', 'Riesgo alto', lambda i, k: srank(i, 0)),
         ('b', '🥈 Segundo marcador', 3.6, 'rk-alto', 'Riesgo alto', lambda i, k: srank(i, 1)),
         ('t', '🥉 Tercer marcador', 4.2, 'rk-extremo', 'Riesgo extremo', lambda i, k: srank(i, 2)),
         ('a', '🧱 Pocos goles', 3.8, 'rk-extremo', 'Riesgo extremo', lambda i, k: low(i)),
         ('p', '🎲 Mixta (1º y 2º)', 4.0, 'rk-extremo', 'Riesgo extremo', lambda i, k: srank(i, 0 if k % 2 == 0 else 1)),
         ('r', '💣 Los raros', 5.0, 'rk-extremo', 'Riesgo extremo', lambda i, k: srank(i, 3))]
    return [{'cls': c, 'n': n, 'm': _mult(f, N), 'rc': rc, 'rl': rl, 'legs': legs(fn)} for c, n, f, rc, rl, fn in D]


def combos_goles(ids, M):
    N = len(ids)
    def top(i):
        o = M[i]['oc']; return max(o, key=o.get) if sum(o.values()) else '1'
    def fav(i):
        m = M[i]; t = m['home'] if top(i) == '1' else m['away']; return t.split(' ')[0][:11]
    def lbl(i, key):
        return {'btts': 'Ambos marcan', 'o15': '+1.5 goles', 'o25': '+2.5 goles', 'o35': '+3.5 goles',
                'u25': '−2.5 goles', 'win2': f'Gana {fav(i)} +2', 'cs': f'{fav(i)} sin recibir'}.get(key, key)
    def legs(keys): return [[ids[k], lbl(ids[k], keys[k % len(keys)])] for k in range(N)]
    D = [('g', '⚽ Ambos marcan', 1.8, 'rk-medio', 'Riesgo medio', ['btts']),
         ('b', '🔀 Mixta de goles', 2.1, 'rk-medio', 'Riesgo medio', ['o25', 'btts', 'o15', 'o35']),
         ('t', '📈 Overs variados', 2.3, 'rk-alto', 'Riesgo medio-alto', ['o25', 'o15', 'o35', 'o25']),
         ('a', '🎯 Favoritos + goles', 2.6, 'rk-alto', 'Riesgo alto', ['win2', 'btts', 'cs', 'o25']),
         ('p', '🧮 Equilibrada', 2.4, 'rk-alto', 'Riesgo alto', ['btts', 'u25', 'o25', 'cs']),
         ('r', '🎆 Festival mixto', 3.0, 'rk-extremo', 'Riesgo extremo', ['o35', 'win2', 'o25', 'btts'])]
    return [{'cls': c, 'n': n, 'm': _mult(f, N), 'rc': rc, 'rl': rl, 'legs': legs(k)} for c, n, f, rc, rl, k in D]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data-dir', required=True)
    ap.add_argument('--ranking-json', required=True)
    ap.add_argument('--out', required=True)
    a = ap.parse_args()

    allm = [m for m in g.build_dated_matches(a.data_dir)
            if m['home'] in g.FLAG_MAP and m['away'] in g.FLAG_MAP]
    allm.sort(key=lambda m: ((m['date'].isoformat() if m.get('date') else '9999'), m.get('time_serial', 0)))
    dated = [m for m in allm if m.get('date')]
    vote = g.collect_votes(a.data_dir, allm)
    results = _read_results(a.data_dir)
    teams = _read_teams(a.data_dir)

    rk = json.load(open(a.ranking_json, encoding='utf-8'))
    allp = rk.get('all', [])
    name_to_stand = {norm(p['name']): p for p in allp}

    order = [mid(m) for m in allm]
    M = {}
    for m in allm:
        v = vote.get((m['group'], m['round'], m['home'], m['away']), {})
        oc = v.get('outcome_counts', Counter()); scc = v.get('score_counts', Counter()); tot = v.get('total', 0)
        cons = max(oc, key=oc.get) if oc else '1'
        num = m.get('match_num')
        try:
            rr = results.get(int(num)) if num not in (None, '') else None
        except (ValueError, TypeError):
            rr = None
        hd = bool(m.get('date'))
        M[mid(m)] = {
            'g': m['group'], 'r': m['round'], 'home': m['home'], 'away': m['away'],
            'hc': g.FLAG_MAP.get(m['home'], 'un'), 'ac': g.FLAG_MAP.get(m['away'], 'un'),
            'kick': kickoff_iso(m['date'], m.get('time_serial', 0)) if hd else '',
            'date': m['date'].isoformat() if hd else '',
            'cons': cons, 'consPct': g.pct(oc.get(cons, 0), tot),
            'consLbl': g.outcome_label(cons, m['home'], m['away']),
            'gs': (scc.most_common(1)[0][0] if scc else ''),
            'scoreTop': scc.most_common(5),
            'oc': {'1': oc.get('1', 0), 'X': oc.get('X', 0), '2': oc.get('2', 0)}, 'total': tot,
            'res': (f"{rr[0]}-{rr[1]}" if rr else ''),
        }

    bydate = defaultdict(list)
    for m in dated:
        bydate[m['date']].append(m)
    matchdays = []
    for d in sorted(bydate):
        ids = [mid(m) for m in sorted(bydate[d], key=lambda x: x.get('time_serial', 0))]
        matchdays.append({'date': d.isoformat(), 'label': g.date_es(d), 'matchIds': ids,
                          'combos': {'1x2': combos_1x2(ids, M), 'exacto': combos_exacto(ids, M), 'goles': combos_goles(ids, M)}})

    idx = {mi: k for k, mi in enumerate(order)}
    players = {}
    matched = 0
    for f in sorted(glob.glob(os.path.join(a.data_dir, 'Quinielas', '*.xlsx'))):
        disp = os.path.splitext(os.path.basename(f))[0].replace('_', ' ')
        picks = g.parse_pool(f, dated)
        arr = [''] * len(order)
        for m in dated:
            pk = g.get_pick_for_match(picks, m)
            if pk and pk.get('outcome') in ('1', 'X', '2'):
                arr[idx[mid(m)]] = f"{pk['outcome']}|{pk.get('score','')}"
        st = name_to_stand.get(norm(disp))
        if st: matched += 1
        pr = g.parse_predictions(f)
        players[disp] = {'p': st['pos'] if st else None, 't': st['pts'] if st else None,
                         'k': arr, 'pred': pr or None}

    out = {
        'gen': datetime.now(timezone(timedelta(hours=-4))).isoformat(),
        'total': rk.get('total', len(allp)),
        'leaderName': rk.get('leader_name'), 'leaderPts': rk.get('leader_pts'),
        'order': order, 'matches': M, 'matchdays': matchdays,
        'players': players,
        'standings': [{'pos': p['pos'], 'name': p['name'], 'pts': p['pts']} for p in allp],
        'dist': rk.get('dist'),
        'teams': teams,
    }
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    json.dump(out, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print(f"[export] players={len(players)} (rank:{matched}) matches={len(M)} days={len(matchdays)} "
          f"size={os.path.getsize(a.out)/1024:.0f}KB -> {a.out}")


if __name__ == '__main__':
    main()
