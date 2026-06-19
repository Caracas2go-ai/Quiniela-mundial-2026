#!/usr/bin/env python3
"""
build_ranking_json.py — construye _ranking.json para el reporte de la Quiniela
Lee el CSV completo exportado del Sheet "Standings" (los 240 participantes) y
arma el JSON que consume generar_reporte_quiniela.py, incluyendo el ranking
completo (all) y el top 3, además de neighbors y dist.

Uso:
  python3 build_ranking_json.py <standings.csv> <salida_ranking.json>

CSV: col0=índice, col1=Pos, col2=Jugador, col3=Puntos Totales.
Se ignoran las filas plantilla cuyo nombre empieza con "Pegar Valores".
"""
import csv, json, sys
from collections import Counter


def main():
    csv_path = sys.argv[1] if len(sys.argv) > 1 else '_standings.csv'
    out_path = sys.argv[2] if len(sys.argv) > 2 else '_ranking.json'

    rows = list(csv.reader(open(csv_path, encoding='utf-8')))
    allp = []
    for r in rows:
        if len(r) < 4:
            continue
        try:
            int(r[0]); pos = int(r[1]); pts = int(r[3])
        except (ValueError, IndexError):
            continue
        name = r[2].strip()
        if not name or name.startswith('Pegar Valores') or name == 'Jugador':
            continue
        allp.append({'pos': pos, 'name': name, 'pts': pts})

    if not allp:
        raise SystemExit(f'[build_ranking_json] No se parsearon participantes de {csv_path}')

    total = len(allp)
    saul_i = next((i for i, p in enumerate(allp)
                   if 'saul' in p['name'].lower() or 'bretto' in p['name'].lower()), None)
    if saul_i is None:
        raise SystemExit('[build_ranking_json] No se encontró a Saul Bretto en el ranking')

    saul = allp[saul_i]
    leader = allp[0]
    cc = Counter(p['pts'] for p in allp)
    dist = [[pts, cc[pts]] for pts in sorted(cc, reverse=True)[:8]]

    data = {
        'saul_pos': saul['pos'], 'saul_pts': saul['pts'], 'total': total,
        'leader_name': leader['name'], 'leader_pts': leader['pts'],
        'top3': allp[:3],
        'neighbors': allp[max(0, saul_i - 2):saul_i + 3],
        'dist': dist,
        'all': allp,
    }
    json.dump(data, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'[build_ranking_json] total={total} saul_pos={saul["pos"]} '
          f'saul_pts={saul["pts"]} leader={leader["name"]} all={len(allp)}')


if __name__ == '__main__':
    main()
