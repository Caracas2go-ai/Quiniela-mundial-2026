#!/usr/bin/env python3
"""
generar_reporte_quiniela.py
Quiniela Mundial 2026 — HTML report generator
Stdlib only: zipfile, xml.etree, datetime, argparse, glob, json, collections, shutil

Usage:
  python3 generar_reporte_quiniela.py \
    --data-dir "/path/to/QUINIELA MUNDIAL 2026" \
    --out "/path/to/quiniela-reporte-2026.html" \
    --fecha 2026-06-19 \
    [--ranking-json /path/to/ranking.json]
"""

import argparse, zipfile, xml.etree.ElementTree as ET, json, glob, os, shutil, base64
from datetime import date, timedelta
from collections import Counter, defaultdict

# ── CONFIG ────────────────────────────────────────────────────────────────────
EPOCH = date(1899, 12, 30)

FLAG_MAP = {
    "México": "mx", "Mexico": "mx",
    "Corea del Sur": "kr", "Corea": "kr",
    "Escocia": "gb-sct",
    "Marruecos": "ma",
    "Estados Unidos": "us", "EE.UU.": "us", "USA": "us",
    "Australia": "au",
    "Turquía": "tr", "Turquia": "tr",
    "Paraguay": "py",
    "Brasil": "br",
    "Haití": "ht", "Haiti": "ht",
    "Sudáfrica": "za", "Sudafrica": "za",
    "República Checa": "cz", "Republica Checa": "cz",
    "Bélgica": "be", "Belgica": "be",
    "Egipto": "eg",
    "Irán": "ir", "Iran": "ir",
    "Nueva Zelanda": "nz",
    "Canadá": "ca", "Canada": "ca",
    "Bosnia y Herzegovina": "ba",
    "Catar": "qa",
    "Suiza": "ch",
    "Alemania": "de",
    "Curazao": "cw",
    "Costa de Marfil": "ci",
    "Ecuador": "ec",
    "Países Bajos": "nl", "Paises Bajos": "nl",
    "Japón": "jp", "Japon": "jp",
    "Suecia": "se",
    "Túnez": "tn", "Tunez": "tn",
    "España": "es", "Espana": "es",
    "Cabo Verde": "cv",
    "Arabia Saudita": "sa",
    "Uruguay": "uy",
    "Francia": "fr",
    "Senegal": "sn",
    "Irak": "iq",
    "Noruega": "no",
    "Argentina": "ar",
    "Argelia": "dz",
    "Austria": "at",
    "Jordania": "jo",
    "Portugal": "pt",
    "RD Congo": "cd",
    "Uzbekistán": "uz", "Uzbekistan": "uz",
    "Colombia": "co",
    "Inglaterra": "gb-eng",
    "Croacia": "hr",
    "Ghana": "gh",
    "Panamá": "pa", "Panama": "pa",
}

ABBREV_MAP = {
    "Méx": "México", "Mex": "México",
    "Sud": "Sudáfrica",
    "Cor": "Corea del Sur",
    "Rep": "República Checa",
    "Can": "Canadá",
    "Bos": "Bosnia y Herzegovina",
    "Cat": "Catar",
    "Sui": "Suiza",
    "Bra": "Brasil",
    "Mar": "Marruecos",
    "Hai": "Haití",
    "Esc": "Escocia",
    "Est": "Estados Unidos",
    "Par": "Paraguay",
    "Aus": "Australia",
    "Tur": "Turquía",
    "Ale": "Alemania",
    "Cur": "Curazao",
    "Cos": "Costa de Marfil",
    "Ecu": "Ecuador",
    "Paí": "Países Bajos",
    "Jap": "Japón",
    "Sue": "Suecia",
    "Tún": "Túnez",
    "Bél": "Bélgica",
    "Egi": "Egipto",
    "Irá": "Irán",
    "Nue": "Nueva Zelanda",
    "Esp": "España",
    "Cab": "Cabo Verde",
    "Ara": "Arabia Saudita",
    "Uru": "Uruguay",
    "Fra": "Francia",
    "Sen": "Senegal",
    "Ira": "Irak",
    "Nor": "Noruega",
    "Arg": "Argentina",
    "Alg": "Argelia",
    "Aut": "Austria",
    "Jor": "Jordania",
    "Por": "Portugal",
    "RD ": "RD Congo",
    "Uzb": "Uzbekistán",
    "Col": "Colombia",
    "Ing": "Inglaterra",
    "Cro": "Croacia",
    "Gha": "Ghana",
    "Pan": "Panamá",
}

# ── XLSX HELPERS ──────────────────────────────────────────────────────────────

def read_xlsx_sheet(path, sheet_name):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        ss = []
        if 'xl/sharedStrings.xml' in names:
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            root = tree.getroot()
            ns = root.tag.split('}')[0] + '}' if '}' in root.tag else ''
            for si in root.iter(f'{ns}si'):
                text = ''.join(t.text or '' for t in si.iter(f'{ns}t'))
                ss.append(text)

        wb = ET.parse(z.open('xl/workbook.xml')).getroot()
        wb_ns = wb.tag.split('}')[0] + '}' if '}' in wb.tag else ''
        sheet_ids = {}
        for sh in wb.iter(f'{wb_ns}sheet'):
            sheet_ids[sh.get('name')] = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

        rels = ET.parse(z.open('xl/_rels/workbook.xml.rels')).getroot()
        rels_ns = rels.tag.split('}')[0] + '}' if '}' in rels.tag else ''
        rmap = {r.get('Id'): r.get('Target') for r in rels.iter(f'{rels_ns}Relationship')}

        if sheet_name not in sheet_ids:
            return []
        _t = rmap[sheet_ids[sheet_name]]
        if _t.startswith('/'):
            target = _t.lstrip('/')
        elif _t.startswith('xl/'):
            target = _t
        else:
            target = 'xl/' + _t

        ws = ET.parse(z.open(target)).getroot()
        ws_ns = ws.tag.split('}')[0] + '}' if '}' in ws.tag else ''

        rows = []
        for row in ws.iter(f'{ws_ns}row'):
            r_idx = int(row.get('r', 0))
            cells = {}
            for c in row:
                ref = c.get('r', '')
                t = c.get('t', '')
                v_el = c.find(f'{ws_ns}v')
                val = ''
                if v_el is not None and v_el.text:
                    val = ss[int(v_el.text)] if t == 's' else v_el.text
                col_str = ''.join(ch for ch in ref if ch.isalpha())
                if col_str:
                    cells[col_str] = val
            rows.append((r_idx, cells))
        return rows


def serial_to_date(serial_str):
    try:
        s = float(serial_str)
        if s < 1:
            return None
        return EPOCH + timedelta(days=int(s))
    except (ValueError, TypeError):
        return None


# ── PARSE HORARIOS ────────────────────────────────────────────────────────────

def parse_horarios(data_dir):
    admin = os.path.join(data_dir, 'ADMIN.xlsx')
    rows = read_xlsx_sheet(admin, 'Horarios')

    group_labels = {f'G{chr(ord("A")+i)}': chr(ord('A')+i) for i in range(12)}
    knockout_labels = {'DIECISEISAVOS', 'OCTAVOS', 'CUARTOS', 'SF', '3-4', 'F', 'CUARTOS DE FINAL', 'SEMIFINAL', 'FINAL'}

    schedule = []
    current_group = None
    group_row_count = 0

    for r_idx, cells in rows:
        a_val = cells.get('A', '').strip()

        if a_val in group_labels:
            current_group = group_labels[a_val]
            group_row_count = 0
            continue

        if a_val in knockout_labels:
            current_group = None
            group_row_count = 0
            continue

        if current_group and a_val:
            dt = serial_to_date(a_val)
            if dt:
                rnd_idx = group_row_count // 2
                rounds = ['J1', 'J2', 'J3']
                rnd = rounds[rnd_idx] if rnd_idx < 3 else None
                if rnd:
                    schedule.append({
                        'group': current_group,
                        'round': rnd,
                        'match_in_round': group_row_count % 2,
                        'date': dt,
                        'time_serial': float(a_val),
                        'row_idx': r_idx,
                    })
                group_row_count += 1

    return schedule


# ── PARSE FIXTURE ─────────────────────────────────────────────────────────────

def _strip_accents(s):
    import unicodedata
    return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn')


def load_group_teams(data_dir):
    """{group_letter: [team names]} desde la hoja Equipos (B=Nombre, C=Grupo)."""
    admin = os.path.join(data_dir, 'ADMIN.xlsx')
    try:
        rows = read_xlsx_sheet(admin, 'Equipos')
    except Exception:
        return {}
    out = {}
    for _r, cells in rows:
        nm = (cells.get('B', '') or '').strip()
        gr = (cells.get('C', '') or '').strip()
        if nm and nm != 'NombreEquipo' and gr:
            out.setdefault(gr.upper(), []).append(nm)
    return out


def resolve_team(abbr, group, gteams):
    """Resuelve una abreviatura de 3 letras al equipo correcto DENTRO de su grupo.
    Evita colisiones globales (p.ej. 'Aus' = Australia en D y Austria en J)."""
    abbr = (abbr or '').strip()
    if not abbr:
        return abbr
    teams = gteams.get((group or '').upper(), [])
    na = _strip_accents(abbr).lower()
    cands = [t for t in teams if _strip_accents(t).lower().startswith(na)]
    if len(cands) == 1:
        return cands[0]
    mapped = ABBREV_MAP.get(abbr)
    if mapped and (not teams or mapped in teams):
        return mapped
    if cands:
        return cands[0]
    return mapped or abbr


# Round-robin estándar de un grupo de 4 (por nº de siembra 1-4): oponente por jornada.
# J1: 1-2, 3-4 · J2: 1-3, 2-4 · J3: 1-4, 2-3
_RR_OPP = {
    'J1': {1: 2, 2: 1, 3: 4, 4: 3},
    'J2': {1: 3, 3: 1, 2: 4, 4: 2},
    'J3': {1: 4, 4: 1, 2: 3, 3: 2},
}


def _candidates(abbr, teams):
    na = _strip_accents(abbr).lower()
    cands = [t for t in teams if _strip_accents(t).lower().startswith(na)]
    if not cands:
        m = ABBREV_MAP.get(abbr)
        if m and m in teams:
            cands = [m]
    return cands


def resolve_pair(home_abbr, away_abbr, group, rnd, gteams):
    """Resuelve local y visitante a la vez. Maneja abreviaturas ambiguas dentro
    del grupo (p.ej. 'Arg' = Argentina o Argelia en J) usando el round-robin."""
    home_abbr = (home_abbr or '').strip()
    away_abbr = (away_abbr or '').strip()
    teams = gteams.get((group or '').upper(), [])
    if len(teams) < 4:
        return (resolve_team(home_abbr, group, gteams),
                resolve_team(away_abbr, group, gteams))
    seed = {i + 1: teams[i] for i in range(4)}
    pos = {teams[i]: i + 1 for i in range(4)}
    hc = _candidates(home_abbr, teams)
    ac = _candidates(away_abbr, teams)
    opp = _RR_OPP.get(rnd, {})

    def pick(cands, other):
        if len(cands) == 1:
            return cands[0]
        if len(other) == 1 and opp:
            o = opp.get(pos[other[0]])
            if o and seed[o] in cands:
                return seed[o]
        return cands[0] if cands else None

    home = pick(hc, ac)
    away = pick(ac, hc)
    if home and away and home == away:
        amb = hc if len(hc) >= 2 else ac
        if len(amb) >= 2:
            s = sorted(amb, key=lambda t: pos[t])
            home, away = s[0], s[1]
    if not home:
        home = resolve_team(home_abbr, group, gteams)
    if not away:
        away = resolve_team(away_abbr, group, gteams)
    return home, away


# Label digit → (round, match_in_round)
_FIX_LABEL_MAP = {'1': ('J1', 1), '2': ('J2', 0), '3': ('J2', 1), '4': ('J3', 0)}


def _parse_fixture_block(rows, cols, gteams):
    """Parsea un bloque de columnas del Fixture. cols = (num, group, home, away, label).
    Cada grupo tiene 6 partidos: header=J1 m0, '1X'=J1 m1, '2X'=J2 m0, '3X'=J2 m1,
    '4X'=J3 m0, última fila sin label=J3 m1."""
    num_c, grp_c, home_c, away_c, lbl_c = cols
    matches = []
    current_group = None
    group_match_count = 0

    for _r, cells in rows:
        a_val = (cells.get(num_c, '') or '').strip()
        b_val = (cells.get(grp_c, '') or '').strip()

        # New group header row
        if len(b_val) == 1 and b_val.upper() in 'ABCDEFGHIJKL':
            current_group = b_val.upper()
            group_match_count = 0
            home_abbr = (cells.get(home_c, '') or '').strip()
            away_abbr = (cells.get(away_c, '') or '').strip()
            if home_abbr and away_abbr:
                try:
                    match_num = int(float(a_val))
                except (ValueError, TypeError):
                    match_num = 0
                home, away = resolve_pair(home_abbr, away_abbr, current_group, 'J1', gteams)
                matches.append({
                    'match_num': match_num, 'group': current_group,
                    'round': 'J1', 'match_in_round': 0,
                    'home': home, 'away': away,
                })
                group_match_count = 1
            continue

        if not current_group or not a_val:
            continue

        try:
            match_num = int(float(a_val))
        except (ValueError, TypeError):
            continue

        home_abbr = (cells.get(home_c, '') or '').strip()
        away_abbr = (cells.get(away_c, '') or '').strip()
        match_label = (cells.get(lbl_c, '') or '').strip()

        if not home_abbr or not away_abbr:
            continue

        if match_label and len(match_label) >= 2:
            lbl_digit = match_label[0]
            grp_char = match_label[1].upper()
            if grp_char in 'ABCDEFGHIJKL':
                current_group = grp_char
            if lbl_digit in _FIX_LABEL_MAP:
                rnd, pos = _FIX_LABEL_MAP[lbl_digit]
                home, away = resolve_pair(home_abbr, away_abbr, current_group, rnd, gteams)
                matches.append({
                    'match_num': match_num, 'group': current_group,
                    'round': rnd, 'match_in_round': pos,
                    'home': home, 'away': away,
                })
                group_match_count += 1
                continue

        # No label — última fila sin label por grupo = J3 match 1
        if group_match_count >= 4:
            home, away = resolve_pair(home_abbr, away_abbr, current_group, 'J3', gteams)
            matches.append({
                'match_num': match_num, 'group': current_group,
                'round': 'J3', 'match_in_round': 1,
                'home': home, 'away': away,
            })
        group_match_count += 1

    return matches


def parse_fixture(data_dir):
    """Lee la hoja Fixture COMPLETA: dos bloques de columnas.
    Bloque izquierdo (Grupos A-F): num=A, grupo=B, local=C, visita=F, label=L.
    Bloque derecho  (Grupos G-L): num=M, grupo=N, local=O, visita=R, label=X.
    """
    admin = os.path.join(data_dir, 'ADMIN.xlsx')
    rows = read_xlsx_sheet(admin, 'Fixture')
    gteams = load_group_teams(data_dir)
    left = _parse_fixture_block(rows, ('A', 'B', 'C', 'F', 'L'), gteams)
    right = _parse_fixture_block(rows, ('M', 'N', 'O', 'R', 'X'), gteams)
    return left + right


# ── BUILD DATED MATCHES ───────────────────────────────────────────────────────

def build_dated_matches(data_dir):
    fixture = parse_fixture(data_dir)
    schedule = parse_horarios(data_dir)

    sched_idx = {}
    for s in schedule:
        key = (s['group'], s['round'], s['match_in_round'])
        sched_idx[key] = s

    dated = []
    for m in fixture:
        # Use match_in_round from fixture if available, else fallback
        pos = m.get('match_in_round', 0)
        sk = (m['group'], m['round'], pos)
        sched_entry = sched_idx.get(sk)
        m = dict(m)
        if sched_entry:
            m['date'] = sched_entry['date']
            m['time_serial'] = sched_entry['time_serial']
        else:
            m['date'] = None
            m['time_serial'] = 0
        dated.append(m)

    dated.sort(key=lambda x: x['time_serial'])
    return dated


# ── PARSE CLAS ────────────────────────────────────────────────────────────────

def parse_clas(data_dir, ranking_json=None):
    if ranking_json and os.path.isfile(ranking_json):
        with open(ranking_json) as f:
            data = json.load(f)
        total = int(data.get('total', 240))
        s_pos = int(data['saul_pos']); s_pts = int(data['saul_pts'])
        leader = {'rank_idx': 1, 'pos': 1, 'name': data.get('leader_name', 'Líder'),
                  'pts': int(data.get('leader_pts', s_pts))}
        ranking = [leader]
        saul_entry = None
        nb_list = []
        for nb in data.get('neighbors', []):
            e = {'rank_idx': 0, 'pos': int(nb['pos']), 'name': nb['name'], 'pts': int(nb['pts'])}
            ranking.append(e); nb_list.append(e)
            if 'saul' in nb['name'].lower() or 'bretto' in nb['name'].lower():
                saul_entry = e
        if saul_entry is None:
            saul_entry = {'rank_idx': 0, 'pos': s_pos, 'name': 'Saul Bretto', 'pts': s_pts}
            ranking.append(saul_entry); nb_list.append(saul_entry)
        saul_entry['_total'] = total
        saul_entry['_leader'] = leader
        saul_entry['_neighbors'] = nb_list
        saul_entry['_dist'] = data.get('dist')
        saul_entry['_top3'] = data.get('top3')
        saul_entry['_all'] = data.get('all')
        return ranking, saul_entry

    admin = os.path.join(data_dir, 'ADMIN.xlsx')
    rows = read_xlsx_sheet(admin, 'CLAS')

    ranking = []
    saul_entry = None
    for r_idx, cells in rows:
        b_val = cells.get('B', '').strip()
        c_val = cells.get('C', '').strip()
        d_val = cells.get('D', '').strip()
        a_val = cells.get('A', '').strip()

        try:
            pos = int(float(b_val))
        except (ValueError, TypeError):
            continue

        try:
            pts = int(float(d_val))
        except (ValueError, TypeError):
            pts = 0

        try:
            rank_idx = int(float(a_val))
        except (ValueError, TypeError):
            rank_idx = 0

        entry = {'rank_idx': rank_idx, 'pos': pos, 'name': c_val, 'pts': pts}
        ranking.append(entry)

        if 'Saul' in c_val or 'Bretto' in c_val or 'saul' in c_val.lower():
            saul_entry = entry

    return ranking, saul_entry


# ── PARSE POOL SHEET ──────────────────────────────────────────────────────────

def parse_pool(path, target_matches):
    try:
        rows = read_xlsx_sheet(path, 'Pool')
    except Exception:
        return {}

    pool_picks = {}
    for r_idx, cells in rows:
        a_val = cells.get('A', '').strip()
        b_val = cells.get('B', '').strip()
        c_val = cells.get('C', '').strip()

        if not a_val or not c_val or len(a_val) < 2:
            continue

        if len(a_val) == 2 and a_val[0].isalpha() and a_val[1].isdigit():
            grp_char = a_val[0].upper()
            rnd_char = a_val[1]
            if grp_char in 'ABCDEFGHIJKL' and rnd_char in '123':
                key = (grp_char, f'J{rnd_char}', b_val.strip())
                if '|' in c_val:
                    parts = c_val.split('|', 1)
                    outcome = parts[0].strip()
                    score = parts[1].strip() if len(parts) > 1 else ''
                    pool_picks[key] = {'outcome': outcome, 'score': score}

    if not pool_picks and target_matches:
        pool_picks = _recover_pool_from_worldcup(path, target_matches)
    return pool_picks


def _recover_pool_from_worldcup(path, target_matches):
    """Fallback para quinielas guardadas sin recalcular (hoja Pool vacía).
    La hoja WORLDCUP conserva el nº de partido (col AH) y los goles (AC=local,
    AD=visitante), así que reconstruimos los picks de fase de grupos desde ahí."""
    try:
        rows = read_xlsx_sheet(path, 'WORLDCUP')
    except Exception:
        return {}
    by_num = {}
    for m in target_matches:
        n = m.get('match_num')
        if n is not None:
            try:
                by_num[int(n)] = m
            except (ValueError, TypeError):
                pass

    def _goals(v):
        v = (v or '').strip()
        return int(float(v)) if v not in ('', '-') else 0

    out = {}
    for r_idx, cells in rows:
        ah = cells.get('AH', '').strip()
        try:
            num = int(float(ah))
        except (ValueError, TypeError):
            continue
        m = by_num.get(num)
        if not m:
            continue
        try:
            hg = _goals(cells.get('AC')); ag = _goals(cells.get('AD'))
        except (ValueError, TypeError):
            continue
        outcome = '1' if hg > ag else ('2' if hg < ag else 'X')
        out[(m['group'], m['round'], f"{m['home']}-{m['away']}")] = {
            'outcome': outcome, 'score': f"{hg}-{ag}"}
    return out


def get_pick_for_match(pool_picks, match):
    grp = match['group']
    rnd = match['round']
    home = match['home']
    away = match['away']

    # Try to find pick by group+round+matchup text similarity
    candidates = [(k, v) for k, v in pool_picks.items() if k[0] == grp and k[1] == rnd]

    if not candidates:
        return None

    # Score each candidate by how well the matchup string matches
    def score_match(matchup_str, home, away):
        s = 0
        # Use first 3-4 chars for fuzzy match
        for team in [home, away]:
            for n in [3, 4, 5, 6]:
                if team[:n] in matchup_str:
                    s += n
                    break
        return s

    best = None
    best_score = -1
    for k, v in candidates:
        ms = k[2]
        sc = score_match(ms, home, away)
        if sc > best_score:
            best_score = sc
            best = v

    return best


# ── COLLECT VOTES ─────────────────────────────────────────────────────────────

def collect_votes(data_dir, target_matches):
    quinielas_dir = os.path.join(data_dir, 'Quinielas')
    files = glob.glob(os.path.join(quinielas_dir, '*.xlsx'))

    vote_data = {}
    for m in target_matches:
        key = (m['group'], m['round'], m['home'], m['away'])
        vote_data[key] = {'outcome_counts': Counter(), 'score_counts': Counter(), 'total': 0}

    for fpath in files:
        pool_picks = parse_pool(fpath, target_matches)
        for m in target_matches:
            mkey = (m['group'], m['round'], m['home'], m['away'])
            pick = get_pick_for_match(pool_picks, m)
            if pick:
                outcome = pick.get('outcome', '')
                score = pick.get('score', '')
                if outcome in ('1', 'X', '2'):
                    vote_data[mkey]['outcome_counts'][outcome] += 1
                    vote_data[mkey]['total'] += 1
                    if score:
                        vote_data[mkey]['score_counts'][score] += 1

    return vote_data


def collect_saul_picks(data_dir, target_matches):
    saul_file = os.path.join(data_dir, 'Quinielas', 'Saul_Bretto.xlsx')
    if not os.path.isfile(saul_file):
        for f in glob.glob(os.path.join(data_dir, 'Quinielas', '*.xlsx')):
            if 'saul' in os.path.basename(f).lower():
                saul_file = f
                break

    return parse_pool(saul_file, target_matches)


def parse_predictions(path):
    """Hoja WORLDCUP → Cuadro de honor: campeón, subcampeón, 3º, botas y balones.
    Las etiquetas viven en la columna W y los valores en la columna AA."""
    try:
        rows = read_xlsx_sheet(path, 'WORLDCUP')
    except Exception:
        return {}
    label_map = [
        ('runnerup', 'subcampeón'), ('champion', 'campeón'), ('third', '3º puesto'),
        ('boot_gold', 'bota de oro'), ('boot_silver', 'bota de plata'), ('boot_bronze', 'bota de bronce'),
        ('ball_gold', 'balón de oro'), ('ball_silver', 'balón de plata'), ('ball_bronze', 'balón de bronce'),
    ]
    res = {}
    for r_idx, cells in rows:
        w = cells.get('W', '').strip().lower()
        aa = cells.get('AA', '').strip()
        if not w or not aa:
            continue
        for key, lbl in label_map:
            if lbl in w and key not in res:
                res[key] = aa
                break
    return res


# ── HELPERS ───────────────────────────────────────────────────────────────────

def flag_url(team_name, size='w80'):
    code = FLAG_MAP.get(team_name, 'un')
    return f"https://flagcdn.com/{size}/{code}.png"


def outcome_label(outcome, home, away):
    if outcome == '1':
        return f'Gana {home}'
    elif outcome == '2':
        return f'Gana {away}'
    elif outcome == 'X':
        return 'Empata'
    return outcome


def outcome_chip_class(outcome):
    if outcome == '1': return 'ch-grn'
    elif outcome == '2': return 'ch-red'
    return 'ch-amb'


def serial_to_time_str(serial):
    try:
        frac = float(serial) % 1
        total_minutes = round(frac * 24 * 60)
        hours = (total_minutes // 60) % 24
        minutes = total_minutes % 60
        ampm = 'AM' if hours < 12 else 'PM'
        h12 = hours % 12 or 12
        if minutes:
            return f"{h12}:{minutes:02d} {ampm} EST"
        return f"{h12}:00 {ampm} EST"
    except Exception:
        return ''


def round_label(rnd):
    if rnd == 'J1': return 'Jornada 1'
    if rnd == 'J2': return 'Jornada 2'
    if rnd == 'J3': return 'Jornada 3'
    return rnd


def date_es(d):
    days = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']
    months = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
              'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    return f"{days[d.weekday()]} {d.day} {months[d.month]}"


def pct(count, total):
    if not total: return 0
    return round(count / total * 100)


# ── CSS ───────────────────────────────────────────────────────────────────────

CSS = """    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    :root {
      --ink:#0B0B0C; --paper:#FFFFFF; --paper2:#ECECE7; --line:#E2E2DC; --mut:#86867E;
      --red:#E2231A; --grn:#1FA85A; --amb:#F5B419; --blu:#2C6FE0; --pur:#6C4FB8; --tea:#13B5A6; --pnk:#E43E8C; --coral:#FF6A4D; --yel:#EFE53B;
    }
    html { scroll-behavior: smooth; }
    body { background: var(--paper2); color: var(--ink); font-family: 'Noto Sans', system-ui, sans-serif; -webkit-font-smoothing: antialiased; }
    img { display: block; }
    .topbar { background: var(--ink); display: flex; align-items: center; justify-content: center; gap: 16px; padding: 14px 20px; }
    .tb-brand { display: flex; align-items: center; gap: 11px; }
    .tb-trophy { font-size: 26px; line-height: 1; }
    .tb-wm { font-family: 'Anton', sans-serif; color: #fff; font-size: 18px; letter-spacing: 1.5px; line-height: 1; text-align: center; }
    .tb-wm small { display: block; font-family: 'Noto Sans', sans-serif; font-size: 9px; font-weight: 800; letter-spacing: 2px; color: var(--yel); margin-top: 3px; }
    .tb-tag { font-family: 'Anton', sans-serif; font-size: 14px; color: var(--ink); background: var(--yel); padding: 5px 12px; border-radius: 6px; letter-spacing: 0.5px; }
    .strip { display: flex; height: 6px; position: relative; overflow: hidden; }
    .strip i { flex: 1; }
    .strip::after { content: ""; position: absolute; top: 0; bottom: 0; left: 0; width: 28%; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.6), transparent); animation: sheen 5s ease-in-out infinite; }
    .topbanner { position: relative; overflow: hidden; height: 36px; background: var(--ink); display: flex; align-items: center; justify-content: center; }
    .topbanner::before { content: ""; position: absolute; inset: 0; background: linear-gradient(90deg, var(--red), var(--amb), var(--grn), var(--tea), var(--blu), var(--pur), var(--pnk), var(--red)); background-size: 200% 100%; animation: flow 3s linear infinite; opacity: 0.92; }
    .topbanner span { position: relative; z-index: 1; font-family: 'Anton', sans-serif; font-size: 12px; letter-spacing: 3px; color: #fff; text-transform: uppercase; text-shadow: 0 1px 4px rgba(0,0,0,0.55); }
    .nav { position: sticky; top: 0; z-index: 20; background: rgba(11,11,12,0.97); display: flex; justify-content: center; gap: 4px; padding: 9px 14px; overflow-x: auto; -webkit-overflow-scrolling: touch; border-bottom: 1px solid rgba(255,255,255,0.08); }
    .nav::-webkit-scrollbar { display: none; }
    .nav a { flex-shrink: 0; font-size: 11px; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; color: #C9C9C2; text-decoration: none; padding: 7px 13px; border-radius: 20px; transition: background 0.15s, color 0.15s; }
    .nav a:hover, .nav a:focus, .nav a.on { background: var(--yel); color: var(--ink); }
    .wrap { max-width: 1120px; margin: 0 auto; padding: 16px; }
    .grid { display: grid; gap: 14px; grid-template-columns: 1fr; }
    section { scroll-margin-top: 56px; }
    .card { background: var(--paper); border: 1px solid var(--line); border-radius: 24px; padding: 20px; position: relative; }
    .card:not(.hero)::before { content: none; }
    @media (min-width: 780px) { .grid { grid-template-columns: 1fr 1fr; align-items: start; } .span2 { grid-column: 1 / -1; } }
    .hero { position: relative; overflow: hidden; margin-bottom: 14px; padding: 30px 22px; background: radial-gradient(130% 130% at 50% 22%, rgba(11,11,12,0.72), rgba(11,11,12,0.95)), repeating-linear-gradient(115deg, var(--red) 0 30px, var(--pur) 30px 60px, var(--blu) 60px 90px, var(--grn) 90px 120px, var(--tea) 120px 150px, var(--amb) 150px 180px); }
    .hero-26 { position: absolute; right: -14px; top: -6px; font-family: 'Anton', sans-serif; font-size: 168px; line-height: 1; color: var(--paper2); z-index: 0; letter-spacing: -8px; background: linear-gradient(135deg, var(--blu), var(--grn) 38%, var(--amb) 64%, var(--red)); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; color: transparent; opacity: 0.55; }
    .hero > * { position: relative; z-index: 1; }
    .eyebrow { font-size: 10px; font-weight: 800; letter-spacing: 2px; text-transform: uppercase; color: #C9C9C2; }
    .rank { font-family: 'Anton', sans-serif; font-size: 92px; line-height: 0.9; color: #fff; }
    .rank span { color: var(--yel); }
    .hsub { font-size: 13px; font-weight: 700; letter-spacing: 0.4px; text-transform: uppercase; color: #C9C9C2; margin-top: 4px; }
    .name { font-family: 'Anton', sans-serif; font-size: 24px; letter-spacing: 0.5px; margin-top: 16px; color: #fff; }
    .pts-pill { display: inline-block; margin-top: 11px; background: var(--grn); color: #fff; font-family: 'Anton', sans-serif; font-size: 15px; letter-spacing: 1px; padding: 7px 16px; border-radius: 7px; }
    .st { font-family: 'Anton', sans-serif; font-size: 18px; letter-spacing: 0.8px; display: flex; align-items: center; gap: 10px; margin-bottom: 16px; text-transform: uppercase; }
    .st .cbar { width: 6px; height: 22px; border-radius: 2px; flex-shrink: 0; }
    .st .smeta { margin-left: auto; font-family: 'Noto Sans', sans-serif; font-size: 10px; font-weight: 800; letter-spacing: 1px; color: var(--mut); white-space: nowrap; }
    .matches { display: grid; gap: 11px; grid-template-columns: 1fr; }
    @media (min-width: 620px) { .matches { grid-template-columns: 1fr 1fr; } }
    .match { padding: 14px; border: 1.5px solid var(--line); border-radius: 18px; }
    .fl { height: 14px; width: auto; border-radius: 2px; box-shadow: 0 0 0 0.5px rgba(0,0,0,0.18); }
    .sb { position: relative; overflow: hidden; background: var(--ink); border-radius: 12px; padding: 18px 12px; margin-bottom: 13px; display: grid; grid-template-columns: 1fr auto 1fr; align-items: center; gap: 8px; }
    .sb::before { content: ""; position: absolute; inset: 0; background: repeating-linear-gradient(115deg, var(--red) 0 26px, var(--pur) 26px 52px, var(--blu) 52px 78px, var(--grn) 78px 104px, var(--tea) 104px 130px, var(--amb) 130px 156px); opacity: 0.9; }
    .sb::after { content: ""; position: absolute; inset: 0; background: radial-gradient(120% 120% at 50% 50%, rgba(11,11,12,0.70), rgba(11,11,12,0.93)); }
    .sb > * { position: relative; z-index: 1; }
    .sb-team { display: flex; flex-direction: column; align-items: center; gap: 7px; }
    .cflag { width: 40px; height: 40px; border-radius: 50%; object-fit: cover; box-shadow: 0 0 0 2.5px #fff; }
    .sb-name { font-family: 'Anton', sans-serif; font-size: 13px; color: #fff; letter-spacing: 0.3px; text-transform: uppercase; text-align: center; line-height: 1.05; }
    .sb-vs { font-family: 'Anton', sans-serif; font-size: 15px; color: var(--yel); letter-spacing: 1px; }
    .m-meta { display: flex; align-items: center; justify-content: center; gap: 8px; flex-wrap: wrap; margin: 0 0 13px; }
    .m-grp { font-size: 10px; font-weight: 800; letter-spacing: 0.6px; text-transform: uppercase; color: var(--mut); }
    .m-when { font-size: 11px; font-weight: 800; letter-spacing: 0.3px; text-transform: uppercase; color: #fff; background: var(--ink); padding: 4px 11px; border-radius: 20px; }
    .cmp { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .cell { background: var(--paper2); border: 1px solid var(--line); border-radius: 14px; padding: 12px 9px; text-align: center; }
    .tag { display: block; font-size: 11px; font-weight: 900; letter-spacing: 0.5px; text-transform: uppercase; color: var(--ink); margin-bottom: 10px; }
    .tag-grp { color: var(--blu); }
    .chip { display: block; font-family: 'Anton', sans-serif; font-size: 17px; letter-spacing: 0.3px; padding: 11px 10px; border-radius: 9px; text-transform: uppercase; line-height: 1.12; min-height: 40px; background: var(--paper); border: 1px solid var(--line); border-left: 5px solid var(--mut); color: var(--ink); }
    .ch-grn { border-left-color: var(--grn); } .ch-amb { border-left-color: var(--ink); } .ch-red { border-left-color: var(--red); }
    .sub { display: block; font-family: 'Anton', sans-serif; font-size: 16px; letter-spacing: 0.5px; color: var(--ink); margin-top: 8px; }
    .m-agree { display: flex; justify-content: center; margin-top: 12px; }
    .m-agree span { display: inline-flex; align-items: center; gap: 5px; font-size: 10px; font-weight: 800; letter-spacing: 0.4px; text-transform: uppercase; padding: 5px 12px; border-radius: 8px; background: rgba(31,168,90,0.14); color: var(--grn); }
    .m-agree.no span { background: rgba(245,180,25,0.20); color: #8A6200; }
    .vote { margin-bottom: 16px; }
    .vote:last-child { margin-bottom: 0; }
    .v-head { display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 8px; flex-wrap: wrap; }
    .v-team { font-family: 'Anton', sans-serif; font-size: 13px; letter-spacing: 0.3px; display: flex; align-items: center; gap: 6px; text-transform: uppercase; }
    .v-pick { display: inline-flex; align-items: center; gap: 5px; font-size: 9px; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; color: #fff; background: var(--ink); padding: 4px 9px; border-radius: 5px; white-space: nowrap; }
    .v-bar { display: flex; height: 11px; border-radius: 6px; overflow: hidden; gap: 2px; margin-bottom: 7px; }
    .v-bar i { height: 100%; }
    .v-nums { display: flex; justify-content: space-between; font-size: 10px; font-weight: 700; color: var(--mut); }
    .v-nums .on { color: var(--ink); font-weight: 800; }
    .fav-grid { display: grid; gap: 14px; grid-template-columns: 1fr; }
    @media (min-width: 620px) { .fav-grid { grid-template-columns: 1fr 1fr; } }
    .fav-match { border: 1.5px solid var(--line); border-radius: 14px; padding: 14px; }
    .fav-head { display: flex; align-items: center; gap: 6px; font-family: 'Anton', sans-serif; font-size: 13px; letter-spacing: 0.3px; text-transform: uppercase; margin-bottom: 12px; flex-wrap: wrap; }
    .fav-head .fav-n { margin-left: auto; font-family: 'Noto Sans', sans-serif; font-size: 9px; font-weight: 800; letter-spacing: 0.5px; color: var(--mut); }
    .fav-row { display: flex; align-items: center; gap: 9px; margin: 7px 0; }
    .fav-sc { font-family: 'Anton', sans-serif; font-size: 15px; width: 40px; flex-shrink: 0; }
    .fav-track { flex: 1; height: 13px; background: var(--paper2); border: 1px solid var(--line); border-radius: 5px; overflow: hidden; }
    @keyframes flow { 0% { background-position: 0% 50%; } 100% { background-position: 200% 50%; } }
    @keyframes sheen { 0% { transform: translateX(-130%); } 55%, 100% { transform: translateX(360%); } }
    .fav-bar { height: 100%; border-radius: 5px; background: var(--ink); }
    .fav-meta { width: 66px; text-align: right; flex-shrink: 0; }
    .fav-pct { font-family: 'Anton', sans-serif; font-size: 14px; display: block; line-height: 1; }
    .fav-cnt { font-size: 9px; font-weight: 700; color: var(--mut); }
    .combos { display: grid; gap: 12px; grid-template-columns: 1fr; }
    @media (min-width: 560px) { .combos { grid-template-columns: 1fr; } }
    @media (min-width: 980px) { .combos { grid-template-columns: 1fr; } }
    .combo { background: var(--paper2); border: 1px solid var(--line); border-left: 5px solid var(--line); border-radius: 16px; padding: 13px 15px; display: flex; flex-direction: column; }
    .combo-h { display: flex; align-items: center; justify-content: space-between; gap: 8px; margin-bottom: 12px; }
    .combo-name { font-family: 'Anton', sans-serif; font-size: 16px; letter-spacing: 0.3px; }
    .combo-x { font-family: 'Anton', sans-serif; font-size: 23px; color: var(--grn); line-height: 1; }
    .combo-x small { display: block; font-family: 'Noto Sans', sans-serif; font-size: 8px; font-weight: 800; color: var(--mut); letter-spacing: 0.5px; text-align: right; }
    .pk { display: flex; align-items: center; gap: 8px; padding: 8px 0; border-top: 1px solid var(--line); }
    .pk-fl { display: flex; gap: 2px; flex-shrink: 0; }
    .pk-match { display: flex; align-items: center; gap: 7px; flex: 1; min-width: 0; }
    .pk-team { font-size: 12px; font-weight: 600; min-width: 0; line-height: 1.25; }
    .pk-r { font-family: 'Anton', sans-serif; font-size: 13.5px; letter-spacing: 0.4px; padding: 6px 2px; text-transform: uppercase; white-space: nowrap; flex-shrink: 0; color: var(--ink); text-align: right; }
    .r-g, .r-a, .r-b, .r-t, .r-r { background: transparent; color: var(--ink); }
    .pk-out { font-family: 'Noto Sans', sans-serif; font-size: 12px; font-weight: 500; color: var(--ink); text-align: right; flex-shrink: 0; white-space: nowrap; line-height: 1.2; padding-left: 10px; }
    .combo--g { border-left-color: var(--grn); } .combo--b { border-left-color: var(--blu); } .combo--a { border-left-color: var(--amb); } .combo--p { border-left-color: var(--pnk); } .combo--r { border-left-color: var(--red); } .combo--t { border-left-color: var(--tea); }
    .combo-risk { margin-top: auto; padding-top: 11px; }
    .combo-risk span { display: inline-block; font-size: 9px; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; padding: 3px 10px; border-radius: 10px; }
    .rk-bajo { background: rgba(31,168,90,0.15); color: var(--grn); } .rk-medio { background: rgba(245,180,25,0.20); color: #8A6200; }
    .rk-alto { background: rgba(226,35,26,0.14); color: var(--red); } .rk-extremo { background: var(--ink); color: var(--yel); }
    .stats { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .stat { background: var(--paper2); border-radius: 12px; padding: 14px; }
    .stat-num { font-family: 'Anton', sans-serif; font-size: 34px; line-height: 1; }
    .stat-lbl { font-size: 10px; font-weight: 800; letter-spacing: 0.8px; text-transform: uppercase; margin-top: 6px; }
    .stat-sub { font-size: 11px; font-weight: 600; color: var(--mut); margin-top: 3px; }
    .c-grn { color: var(--grn); } .c-red { color: var(--red); } .c-blu { color: var(--blu); } .c-amb { color: var(--amb); }
    .leader { display: flex; align-items: center; gap: 14px; background: var(--ink); border-radius: 12px; padding: 14px 16px; }
    .leader-badge { width: 40px; height: 40px; background: var(--yel); color: var(--ink); border-radius: 9px; display: flex; align-items: center; justify-content: center; font-family: 'Anton', sans-serif; font-size: 20px; flex-shrink: 0; }
    .leader-info { flex: 1; }
    .leader-name { font-family: 'Anton', sans-serif; font-size: 16px; color: #fff; letter-spacing: 0.3px; }
    .leader-cap { font-size: 10px; font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase; color: var(--mut); margin-top: 2px; }
    .leader-pts { font-family: 'Anton', sans-serif; font-size: 26px; color: var(--yel); text-align: right; }
    .leader-gap { font-size: 10px; font-weight: 700; color: var(--red); text-transform: uppercase; }
    .hero-emblem { position: absolute; right: 14px; top: 50%; transform: translateY(-50%); height: 168px; width: auto; z-index: 0; opacity: 0.92; filter: drop-shadow(0 6px 18px rgba(0,0,0,0.45)); }
    @media (max-width: 560px) { .hero-emblem { height: 112px; right: 6px; opacity: 0.8; } }
    .top3 { display: grid; gap: 9px; }
    .t3 { display: flex; align-items: center; gap: 13px; background: var(--ink); border-radius: 12px; padding: 13px 16px; }
    .t3-badge { width: 38px; height: 38px; border-radius: 9px; display: flex; align-items: center; justify-content: center; font-family: 'Anton', sans-serif; font-size: 19px; flex-shrink: 0; color: var(--ink); }
    .t3-1 .t3-badge { background: var(--yel); } .t3-2 .t3-badge { background: #C9C9C2; } .t3-3 .t3-badge { background: #D9A066; }
    .t3-name { flex: 1; font-family: 'Anton', sans-serif; font-size: 16px; color: #fff; letter-spacing: 0.3px; }
    .t3-pos { font-size: 9px; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; color: var(--mut); }
    .t3-pts { font-family: 'Anton', sans-serif; font-size: 23px; color: var(--yel); flex-shrink: 0; }
    .fullrk { border: 1.5px solid var(--line); border-radius: 12px; max-height: 540px; overflow-y: auto; }
    .fr-row { display: flex; align-items: center; gap: 10px; padding: 9px 14px; border-bottom: 1px solid var(--line); }
    .fr-row:last-child { border-bottom: none; }
    .fr-row.me { background: var(--yel); }
    .fr-pos { font-family: 'Anton', sans-serif; font-size: 13px; color: var(--mut); width: 38px; text-align: right; flex-shrink: 0; }
    .fr-row.me .fr-pos { color: var(--ink); }
    .fr-name { font-size: 12.5px; font-weight: 600; flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .fr-row.me .fr-name { font-weight: 800; }
    .fr-diff { font-size: 9px; font-weight: 800; letter-spacing: 0.3px; text-transform: uppercase; padding: 3px 8px; border-radius: 10px; flex-shrink: 0; }
    .fr-pts { font-family: 'Anton', sans-serif; font-size: 15px; width: 30px; text-align: right; flex-shrink: 0; }
    .stages { display: flex; gap: 7px; flex-wrap: wrap; }
    .stage { font-size: 11px; font-weight: 800; letter-spacing: 0.4px; text-transform: uppercase; padding: 6px 13px; border-radius: 20px; border: 1.5px solid var(--line); color: var(--mut); }
    .stage.on { background: var(--grn); border-color: var(--grn); color: #fff; }
    .nb { border: 1.5px solid var(--line); border-radius: 12px; overflow: hidden; }
    .nb-row { display: flex; align-items: center; gap: 11px; padding: 11px 14px; border-bottom: 1px solid var(--line); }
    .nb-row:last-child { border-bottom: none; }
    .nb-row.me { background: var(--yel); }
    .nb-pos { font-family: 'Anton', sans-serif; font-size: 14px; color: var(--mut); width: 34px; text-align: right; flex-shrink: 0; }
    .nb-row.me .nb-pos { color: var(--ink); }
    .nb-name { font-size: 13px; font-weight: 600; flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .nb-row.me .nb-name { font-weight: 800; }
    .nb-badge { font-size: 9px; font-weight: 800; letter-spacing: 0.4px; text-transform: uppercase; padding: 3px 8px; border-radius: 10px; flex-shrink: 0; }
    .bg-you { background: var(--ink); color: var(--yel); } .bg-up { background: rgba(31,168,90,0.16); color: var(--grn); } .bg-dn { background: rgba(226,35,26,0.14); color: var(--red); }
    .nb-pts { font-family: 'Anton', sans-serif; font-size: 17px; width: 32px; text-align: right; flex-shrink: 0; }
    .pt-up { color: var(--grn); } .pt-me { color: var(--ink); } .pt-dn { color: var(--red); }
    .dist-row { display: flex; align-items: center; gap: 11px; margin: 7px 0; }
    .dist-pts { font-family: 'Anton', sans-serif; font-size: 15px; width: 38px; text-align: right; flex-shrink: 0; color: var(--ink); }
    .dist-pts.me { color: var(--grn); }
    .dist-wrap { flex: 1; height: 22px; background: var(--paper2); border: 1px solid var(--line); border-radius: 6px; overflow: hidden; }
    .dist-bar { height: 100%; background: var(--ink); display: flex; align-items: center; padding-left: 8px; min-width: 22px; }
    .dist-bar.me { background: var(--grn); }
    .dist-c { font-size: 11px; font-weight: 800; color: #fff; }
    .dist-meta { font-size: 10px; font-weight: 700; letter-spacing: 0.2px; color: var(--mut); width: 106px; flex-shrink: 0; }
    .dist-meta.me { color: var(--grn); font-weight: 800; }
    .insight { background: var(--ink); border-radius: 14px; padding: 18px; }
    .insight-t { font-family: 'Anton', sans-serif; font-size: 13px; letter-spacing: 1px; color: var(--yel); text-transform: uppercase; margin-bottom: 10px; }
    .insight-x { font-size: 13px; line-height: 1.65; color: #D6D6D0; }
    .insight-x strong { color: #fff; font-weight: 800; }
    .note { font-size: 11px; font-weight: 600; color: var(--mut); margin-top: 13px; line-height: 1.55; }
    .note strong { color: var(--ink); font-weight: 800; }
    .ft { text-align: center; padding: 28px 20px; }
    .ft-26 { font-family: 'Anton', sans-serif; font-size: 46px; letter-spacing: 1px; background: linear-gradient(135deg, var(--blu), var(--grn) 35%, var(--amb) 60%, var(--red), var(--pnk)); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; color: transparent; }
    .ft-emblem { height: 74px; width: auto; margin: 0 auto; opacity: 0.95; filter: drop-shadow(0 3px 10px rgba(0,0,0,0.22)); }
    .ft-strip { display: flex; height: 5px; max-width: 120px; margin: 10px auto; border-radius: 3px; overflow: hidden; }
    .ft-strip i { flex: 1; }
    .ft-txt { font-size: 10px; font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase; color: var(--mut); margin-top: 8px; }
    @media (prefers-reduced-motion: reduce) { .topbanner::before, .fav-bar, .strip::after { animation: none !important; } }
    @media (min-width: 780px) { .card { transition: transform .18s ease, box-shadow .18s ease; } .card:hover { transform: translateY(-3px); box-shadow: 0 12px 32px rgba(11,11,12,0.10); } }"""


# ── HTML BUILDERS ─────────────────────────────────────────────────────────────

def build_match_card(m, vote, saul_pick):
    home, away = m['home'], m['away']
    group, rnd = m['group'], m['round']
    grp_label = f"Grupo {group}"
    t_str = serial_to_time_str(m.get('time_serial', 0))
    when_label = date_es(m['date']) + (f" · {t_str}" if t_str else "")

    home_code = FLAG_MAP.get(home, 'un')
    away_code = FLAG_MAP.get(away, 'un')

    total = vote.get('total', 0) if vote else 0
    oc = vote.get('outcome_counts', Counter()) if vote else Counter()

    if total and oc:
        cons_outcome = max(oc, key=oc.get)
        cons_pct = pct(oc.get(cons_outcome, 0), total)
        cons_label = outcome_label(cons_outcome, home, away)
        cons_chip = outcome_chip_class(cons_outcome)
        sc = vote.get('score_counts', Counter()) if vote else Counter()
        top_sc = sc.most_common(1)[0][0] if sc else ''
        cons_sub = f"{cons_pct}% del grupo · {top_sc}" if top_sc else f"{cons_pct}% del grupo"
    else:
        cons_outcome, cons_label, cons_chip, cons_sub = '1', f'Gana {home}', 'ch-grn', 'Sin datos'

    pick = saul_pick
    if pick:
        saul_outcome = pick.get('outcome', '')
        saul_score = pick.get('score', '')
        saul_chip = outcome_chip_class(saul_outcome)
        saul_label = outcome_label(saul_outcome, home, away)
        agree = (saul_outcome == cons_outcome)
        agree_class = '' if agree else ' no'
        agree_text = '✓ Tú y el grupo coinciden' if agree else '✕ No coincides con el grupo'
    else:
        saul_outcome, saul_score, saul_chip, saul_label = '?', '?', 'ch-amb', 'Sin pick'
        agree_class, agree_text = ' no', 'Sin pick registrado'

    return f"""          <div class="match">
            <div class="sb">
              <div class="sb-team"><img class="cflag" src="https://flagcdn.com/w80/{home_code}.png" alt="{home}"><span class="sb-name">{home}</span></div>
              <div class="sb-vs">VS</div>
              <div class="sb-team"><img class="cflag" src="https://flagcdn.com/w80/{away_code}.png" alt="{away}"><span class="sb-name">{away}</span></div>
            </div>
            <div class="m-meta"><span class="m-grp">{grp_label}</span><span class="m-when">{when_label}</span></div>
            <div class="cmp">
              <div class="cell"><span class="tag">👤 Tu pick</span><span class="chip {saul_chip}">{saul_label}</span><span class="sub">{saul_score}</span></div>
              <div class="cell"><span class="tag tag-grp">👥 El grupo</span><span class="chip {cons_chip}">{cons_label}</span><span class="sub">{cons_sub}</span></div>
            </div>
            <div class="m-agree{agree_class}"><span>{agree_text}</span></div>
          </div>"""


def build_vote_row(m, vote, saul_pick):
    home, away = m['home'], m['away']
    home_code = FLAG_MAP.get(home, 'un')
    away_code = FLAG_MAP.get(away, 'un')
    total = vote.get('total', 0) if vote else 0
    oc = vote.get('outcome_counts', Counter()) if vote else Counter()

    home_pct = pct(oc.get('1', 0), total)
    draw_pct = pct(oc.get('X', 0), total)
    away_pct = pct(oc.get('2', 0), total)

    saul_out = saul_pick.get('outcome', '') if saul_pick else ''
    saul_text_map = {'1': home[:12], '2': away[:12], 'X': 'Empate'}
    saul_text = saul_text_map.get(saul_out, 'Sin pick')

    def on_cls(out):
        return 'on' if saul_out == out else ''

    return f"""        <div class="vote">
          <div class="v-head"><span class="v-team"><img class="fl" src="https://flagcdn.com/h24/{home_code}.png"> {home[:10]} vs {away[:10]} <img class="fl" src="https://flagcdn.com/h24/{away_code}.png"></span><span class="v-pick">Tu pick: {saul_text}</span></div>
          <div class="v-bar"><i style="width:{home_pct}%;background:var(--grn)"></i><i style="width:{draw_pct}%;background:var(--amb)"></i><i style="width:{away_pct}%;background:var(--red)"></i></div>
          <div class="v-nums"><span class="{on_cls('1')}">{home[:10]} {home_pct}%</span><span class="{on_cls('X')}">Empate {draw_pct}%</span><span class="{on_cls('2')}">{away[:10]} {away_pct}%</span></div>
        </div>"""


def build_fav_match(m, vote, saul_pick):
    home, away = m['home'], m['away']
    home_code = FLAG_MAP.get(home, 'un')
    away_code = FLAG_MAP.get(away, 'un')
    total = vote.get('total', 0) if vote else 0
    sc = vote.get('score_counts', Counter()) if vote else Counter()
    saul_score = saul_pick.get('score', '') if saul_pick else ''

    top5 = sc.most_common(5)
    if not top5:
        return f"""          <div class="fav-match">
            <div class="fav-head"><img class="fl" src="https://flagcdn.com/h24/{home_code}.png"> {home} vs {away} <img class="fl" src="https://flagcdn.com/h24/{away_code}.png"><span class="fav-n">Sin datos</span></div>
          </div>"""

    max_c = top5[0][1]
    rows = []
    for score, cnt in top5:
        w = round(cnt / max_c * 100)
        p = pct(cnt, total)
        rows.append(f'            <div class="fav-row"><span class="fav-sc">{score}</span><div class="fav-track"><div class="fav-bar" style="width:{w}%"></div></div><div class="fav-meta"><span class="fav-pct">{p}%</span><span class="fav-cnt">{cnt} veces</span></div></div>')

    saul_note = f' · Tu pick: <strong>{saul_score}</strong>' if saul_score else ''
    return f"""          <div class="fav-match">
            <div class="fav-head"><img class="fl" src="https://flagcdn.com/h24/{home_code}.png"> {home} vs {away} <img class="fl" src="https://flagcdn.com/h24/{away_code}.png"><span class="fav-n">{total} votos{saul_note}</span></div>
{chr(10).join(rows)}
          </div>"""


def _mult_str(factor, n):
    v = round(factor ** max(n, 1))
    if v >= 1000:
        return f'{v:,}x'.replace(',', '.')
    return f'{max(v, 2)}x'


def build_combos_1x2(matches, vote_data):
    if not matches:
        return '<p style="color:var(--mut)">No hay partidos para combinar.</p>'
    ms = list(matches)
    N = len(ms)

    def vd(m):
        return vote_data.get((m['group'], m['round'], m['home'], m['away']), {})

    def top_out(m):
        oc = vd(m).get('outcome_counts', Counter())
        return max(oc, key=oc.get) if oc else '1'

    def least_out(m):
        oc = vd(m).get('outcome_counts', Counter())
        return min(oc, key=oc.get) if oc else '2'

    def strength(m):
        oc = vd(m).get('outcome_counts', Counter()); tot = vd(m).get('total', 0)
        return (max(oc.values()) / tot) if (tot and oc) else 0

    def invert(out):
        return '2' if out == '1' else ('1' if out == '2' else 'X')

    divided = sorted(range(N), key=lambda i: strength(ms[i]))

    def pk(m, out):
        hc = FLAG_MAP.get(m['home'], 'un'); ac = FLAG_MAP.get(m['away'], 'un')
        if out == '1':
            lbl = f'Gana {m["home"]}'
        elif out == '2':
            lbl = f'Gana {m["away"]}'
        else:
            lbl = 'Empate'
        team = f'{m["home"]} vs {m["away"]}'
        return f'            <div class="pk"><span class="pk-match"><img class="fl" src="https://flagcdn.com/h24/{hc}.png"><span class="pk-team">{team}</span><img class="fl" src="https://flagcdn.com/h24/{ac}.png"></span><span class="pk-out">{lbl}</span></div>'

    def rows(outs):
        return '\n'.join(pk(ms[i], outs[i]) for i in range(N))

    cons = [top_out(m) for m in ms]
    s1 = list(cons)
    s2 = list(cons)
    if divided:
        s2[divided[0]] = 'X'
    s3 = list(cons)
    for i in divided[:2]:
        s3[i] = 'X'
    s4 = list(cons)
    if divided:
        s4[divided[0]] = invert(cons[divided[0]])
    s5 = [invert(o) for o in cons]
    s6 = [least_out(m) for m in ms]

    def combo(cls, name, factor, risk_cls, risk_lbl, outs):
        return f"""          <div class="combo combo--{cls}">
            <div class="combo-h"><span class="combo-name">{name}</span><span class="combo-x">≈{_mult_str(factor, N)}<small>aprox</small></span></div>
{rows(outs)}
            <div class="combo-risk"><span class="{risk_cls}">{risk_lbl}</span></div>
          </div>"""

    return '\n'.join([
        combo('g', '🛡️ La sólida', 1.5, 'rk-bajo', 'Riesgo bajo', s1),
        combo('b', '⚖️ Favoritos + 1 empate', 1.9, 'rk-medio', 'Riesgo medio', s2),
        combo('t', '🎯 Doble empate', 2.2, 'rk-alto', 'Riesgo medio-alto', s3),
        combo('a', '🔥 Con sorpresa', 2.5, 'rk-alto', 'Riesgo alto', s4),
        combo('p', '🔄 Contraataque', 3.2, 'rk-extremo', 'Riesgo extremo', s5),
        combo('r', '💣 Todo o nada', 4.2, 'rk-extremo', 'Riesgo extremo', s6),
    ])


def build_combos_exacto(matches, vote_data):
    if not matches:
        return ''
    ms = list(matches)
    N = len(ms)

    def vd(m):
        return vote_data.get((m['group'], m['round'], m['home'], m['away']), {})

    def score_rank(m, rank=0):
        top = vd(m).get('score_counts', Counter()).most_common(6)
        if not top:
            return '1-1'
        return top[rank][0] if rank < len(top) else top[-1][0]

    def low_score(m):
        top = vd(m).get('score_counts', Counter()).most_common(6)
        if not top:
            return '0-0'
        def goals(s):
            try:
                h, a = s.split('-'); return int(h) + int(a)
            except Exception:
                return 9
        return sorted(top, key=lambda t: goals(t[0]))[0][0]

    def pk_ex(m, score):
        hc = FLAG_MAP.get(m['home'], 'un'); ac = FLAG_MAP.get(m['away'], 'un')
        try:
            h, a = score.split('-')
            css = 'r-g' if int(h) > int(a) else ('r-r' if int(h) < int(a) else 'r-a')
        except Exception:
            css = 'r-b'
        team = f'{m["home"]} vs {m["away"]}'
        return f'            <div class="pk"><span class="pk-match"><img class="fl" src="https://flagcdn.com/h24/{hc}.png"><span class="pk-team">{team}</span><img class="fl" src="https://flagcdn.com/h24/{ac}.png"></span><span class="pk-r {css}">{score}</span></div>'

    def rows(fn):
        return '\n'.join(pk_ex(ms[i], fn(ms[i], i)) for i in range(N))

    def combo(cls, name, factor, risk_cls, risk_lbl, fn):
        return f"""          <div class="combo combo--{cls}">
            <div class="combo-h"><span class="combo-name">{name}</span><span class="combo-x">≈{_mult_str(factor, N)}<small>aprox</small></span></div>
{rows(fn)}
            <div class="combo-risk"><span class="{risk_cls}">{risk_lbl}</span></div>
          </div>"""

    return '\n'.join([
        combo('g', '🎯 Marcador favorito', 3.0, 'rk-alto', 'Riesgo alto', lambda m, i: score_rank(m, 0)),
        combo('b', '🥈 Segundo marcador', 3.6, 'rk-alto', 'Riesgo alto', lambda m, i: score_rank(m, 1)),
        combo('t', '🥉 Tercer marcador', 4.2, 'rk-extremo', 'Riesgo extremo', lambda m, i: score_rank(m, 2)),
        combo('a', '🧱 Pocos goles', 3.8, 'rk-extremo', 'Riesgo extremo', lambda m, i: low_score(m)),
        combo('p', '🎲 Mixta (1º y 2º)', 4.0, 'rk-extremo', 'Riesgo extremo', lambda m, i: score_rank(m, 0 if i % 2 == 0 else 1)),
        combo('r', '💣 Los raros', 5.0, 'rk-extremo', 'Riesgo extremo', lambda m, i: score_rank(m, 3)),
    ])


def build_combos_goles(matches, vote_data):
    if not matches:
        return ''

    def vd(m):
        return vote_data.get((m['group'], m['round'], m['home'], m['away']), {})

    def top_out(m):
        oc = vd(m).get('outcome_counts', Counter())
        return max(oc, key=oc.get) if oc else '1'

    def fav_short(m):
        t = m['home'] if top_out(m) == '1' else m['away']
        return t.split(' ')[0][:11]

    def market_label(m, key):
        return {
            'btts': 'Ambos marcan',
            'o15': '+1.5 goles',
            'o25': '+2.5 goles',
            'o35': '+3.5 goles',
            'u25': '−2.5 goles',
            'ht': 'Gol 1er tiempo',
            'win2': f'Gana {fav_short(m)} +2',
            'cs': f'{fav_short(m)} sin recibir',
        }.get(key, key)

    def pk(m, key):
        hc = FLAG_MAP.get(m['home'], 'un'); ac = FLAG_MAP.get(m['away'], 'un')
        team = f'{m["home"]} vs {m["away"]}'
        return f'            <div class="pk"><span class="pk-match"><img class="fl" src="https://flagcdn.com/h24/{hc}.png"><span class="pk-team">{team}</span><img class="fl" src="https://flagcdn.com/h24/{ac}.png"></span><span class="pk-out">{market_label(m, key)}</span></div>'

    ms = list(matches)
    N = len(ms)

    def rows(keys):
        return '\n'.join(pk(ms[i], keys[i % len(keys)]) for i in range(N))

    def combo(cls, name, factor, risk_cls, risk_lbl, keys):
        return f"""          <div class="combo combo--{cls}">
            <div class="combo-h"><span class="combo-name">{name}</span><span class="combo-x">≈{_mult_str(factor, N)}<small>aprox</small></span></div>
{rows(keys)}
            <div class="combo-risk"><span class="{risk_cls}">{risk_lbl}</span></div>
          </div>"""

    return '\n'.join([
        combo('g', '⚽ Ambos marcan', 1.8, 'rk-medio', 'Riesgo medio', ['btts']),
        combo('b', '🔀 Mixta de goles', 2.1, 'rk-medio', 'Riesgo medio', ['o25', 'btts', 'o15', 'o35']),
        combo('t', '📈 Overs variados', 2.3, 'rk-alto', 'Riesgo medio-alto', ['o25', 'o15', 'o35', 'o25']),
        combo('a', '🎯 Favoritos + goles', 2.6, 'rk-alto', 'Riesgo alto', ['win2', 'btts', 'cs', 'o25']),
        combo('p', '🧮 Equilibrada', 2.4, 'rk-alto', 'Riesgo alto', ['btts', 'u25', 'o25', 'cs']),
        combo('r', '🎆 Festival mixto', 3.0, 'rk-extremo', 'Riesgo extremo', ['o35', 'win2', 'o25', 'btts']),
    ])


def build_neighbors(ranking, saul_entry):
    if not saul_entry:
        return '<p>Saul no encontrado en el ranking.</p>'
    saul_pts = saul_entry['pts']
    saul_name = saul_entry['name']
    saul_idx = next((i for i, r in enumerate(ranking) if r['name'] == saul_name), -1)
    nbs = saul_entry.get('_neighbors')
    if nbs:
        neighbors = nbs
    else:
        start = max(0, saul_idx - 3)
        end = min(len(ranking), saul_idx + 4)
        neighbors = ranking[start:end]
    rows = []
    for nb in neighbors:
        is_me = nb['name'] == saul_name
        me_cls = ' me' if is_me else ''
        if is_me:
            badge = '<span class="nb-badge bg-you">★ Tú</span>'
            pts_cls = 'pt-me'
        elif nb['pts'] > saul_pts:
            badge = f'<span class="nb-badge bg-up">▲ +{nb["pts"] - saul_pts}</span>'
            pts_cls = 'pt-up'
        elif nb['pts'] == saul_pts:
            badge = '<span class="nb-badge bg-dn">= empate</span>'
            pts_cls = 'pt-dn'
        else:
            badge = f'<span class="nb-badge bg-dn">▼ -{saul_pts - nb["pts"]}</span>'
            pts_cls = 'pt-dn'
        rows.append(f'          <div class="nb-row{me_cls}"><span class="nb-pos">{nb["pos"]}°</span><span class="nb-name">{nb["name"]}</span>{badge}<span class="nb-pts {pts_cls}">{nb["pts"]}</span></div>')
    return '\n'.join(rows)


def build_distribution(ranking, saul_entry):
    dist = saul_entry.get('_dist') if saul_entry else None
    if not ranking and not dist:
        return ''
    saul_pts = saul_entry['pts'] if saul_entry else 0
    total = (saul_entry.get('_total') if saul_entry else None) or len(ranking) or 1
    if dist:
        pts_counter = {int(p): int(c) for p, c in dist}
    else:
        pts_counter = Counter(r['pts'] for r in ranking)
    unique_pts = sorted(pts_counter.keys(), reverse=True)
    max_c = max(pts_counter.values()) if pts_counter else 1
    rows = []
    for pts in unique_pts[:8]:
        cnt = pts_counter[pts]
        w = round(cnt / max_c * 100)
        pctt = round(cnt / total * 100)
        is_me = (pts == saul_pts)
        me_cls = ' me' if is_me else ''
        if is_me:
            meta = f'{cnt} contigo · {pctt}%'
        else:
            meta = f'{cnt} jugador{"" if cnt == 1 else "es"} · {pctt}%'
        rows.append(f'        <div class="dist-row"><span class="dist-pts{me_cls}">{pts}</span><div class="dist-wrap"><div class="dist-bar{me_cls}" style="width:{max(8,w)}%"><span class="dist-c">{cnt}</span></div></div><span class="dist-meta{me_cls}">{meta}</span></div>')
    return '\n'.join(rows)


def _rank_rows(entries, saul_pts):
    rows = []
    for e in entries:
        is_me = ('saul' in e['name'].lower() or 'bretto' in e['name'].lower())
        me_cls = ' me' if is_me else ''
        d = e['pts'] - saul_pts
        if is_me:
            badge = '<span class="fr-diff bg-you">★ Tú</span>'; pts_cls = 'pt-me'
        elif d > 0:
            badge = f'<span class="fr-diff bg-up">▲ +{d}</span>'; pts_cls = 'pt-up'
        elif d == 0:
            badge = '<span class="fr-diff bg-dn">= 0</span>'; pts_cls = 'pt-me'
        else:
            badge = f'<span class="fr-diff bg-dn">▼ {d}</span>'; pts_cls = 'pt-dn'
        rows.append(f'          <div class="fr-row{me_cls}"><span class="fr-pos">{e["pos"]}°</span><span class="fr-name">{e["name"]}</span>{badge}<span class="fr-pts {pts_cls}">{e["pts"]}</span></div>')
    return '\n'.join(rows)


def build_top3(saul_entry):
    top3 = (saul_entry.get('_top3') if saul_entry else None) or []
    if not top3:
        return ''
    rows = []
    for i, e in enumerate(top3[:3]):
        rows.append(f'          <div class="t3 t3-{i+1}"><div class="t3-badge">{i+1}</div><div class="t3-name">{e["name"]}</div><div class="t3-pts">{e["pts"]}</div></div>')
    return '\n'.join(rows)


def build_top_ranking(saul_entry, n=10):
    allp = (saul_entry.get('_all') if saul_entry else None) or []
    if not allp:
        return '<p style="color:var(--mut)">Ranking no disponible.</p>'
    return _rank_rows(allp[:n], saul_entry['pts'])


def build_full_ranking(saul_entry):
    allp = (saul_entry.get('_all') if saul_entry else None) or []
    if not allp:
        return '<p style="color:var(--mut)">Ranking completo no disponible.</p>'
    return _rank_rows(allp, saul_entry['pts'])


def build_insight(saul_entry, ranking, target_matches):
    if not saul_entry: return 'Datos de ranking no disponibles.'
    saul_pts = saul_entry['pts']
    saul_pos = saul_entry['pos']
    total = (saul_entry.get('_total') if saul_entry else None) or len(ranking)
    top_pct = round(saul_pos / total * 100)
    leader = (saul_entry.get('_leader') if saul_entry else None) or (ranking[0] if ranking else None)
    gap = (leader['pts'] - saul_pts) if leader else 0
    dist = saul_entry.get('_dist') if saul_entry else None
    if dist:
        above_pts = [int(p) for p, _ in dist if int(p) > saul_pts]
        next_gap = (min(above_pts) - saul_pts) if above_pts else 1
    else:
        above = [r for r in ranking if r['pts'] > saul_pts]
        next_gap = (min(r['pts'] for r in above) - saul_pts) if above else 1
    return (f'Con <strong>{saul_pts} pts</strong>, estás en el top {top_pct}% de {total} participantes. '
            f'Solo <strong>{next_gap} punto{"" if next_gap == 1 else "s"} más</strong> para escalar. '
            f'El diferenciador clave llega en <strong>octavos</strong>, donde los picks de equipos generan los saltos grandes.')


# ── GENERATE HTML ─────────────────────────────────────────────────────────────

def generate_html(target_date, dated_matches, vote_data, saul_picks, ranking, saul_entry, logo_b64=''):
    fecha_str = target_date.strftime('%Y-%m-%d')
    date_label = date_es(target_date)
    today_matches = sorted([m for m in dated_matches if m.get('date') == target_date], key=lambda m: m.get('time_serial', 0))

    saul_pos = saul_entry['pos'] if saul_entry else '?'
    saul_pts = saul_entry['pts'] if saul_entry else 0
    total_part = (saul_entry.get('_total') if saul_entry else None) or len(set(r['name'] for r in ranking)) or len(ranking)
    top_pct = round(saul_pos / total_part * 100) if saul_entry and total_part else '?'
    leader = (saul_entry.get('_leader') if saul_entry else None) or (ranking[0] if ranking else None)
    leader_name = leader['name'].upper() if leader else 'N/A'
    leader_pts = leader['pts'] if leader else 0
    gap_leader = leader_pts - saul_pts

    rounds_present = sorted(set(round_label(m['round']) for m in today_matches))
    rnd_label = ' · '.join(rounds_present) if rounds_present else '–'
    smeta = f"{rnd_label} · {date_label.upper()}"

    match_cards = []
    vote_rows = []
    fav_matches = []
    for m in today_matches:
        mkey = (m['group'], m['round'], m['home'], m['away'])
        v = vote_data.get(mkey, {})
        sp = get_pick_for_match(saul_picks, m)
        match_cards.append(build_match_card(m, v, sp))
        vote_rows.append(build_vote_row(m, v, sp))
        fav_matches.append(build_fav_match(m, v, sp))

    matches_html = '\n'.join(match_cards) if match_cards else '<p style="color:var(--mut)">No hay partidos para esta fecha.</p>'
    voto_html = '\n'.join(vote_rows) if vote_rows else '<p style="color:var(--mut)">Sin datos de voto.</p>'
    fav_html = '\n'.join(fav_matches) if fav_matches else '<p style="color:var(--mut)">Sin datos de marcadores.</p>'

    sample_total = max((vote_data.get((m['group'], m['round'], m['home'], m['away']), {}).get('total', 0) for m in today_matches), default=0) if today_matches else 0
    fav_smeta = f"{rnd_label.upper()} · {sample_total} VOTOS" if sample_total else rnd_label.upper()

    combos_1x2 = build_combos_1x2(today_matches, vote_data)
    combos_exacto = build_combos_exacto(today_matches, vote_data)
    combos_goles = build_combos_goles(today_matches, vote_data)

    if saul_entry and saul_entry.get('_total'):
        nbs = saul_entry.get('_neighbors', [])
        _dm = {int(p): int(c) for p, c in (saul_entry.get('_dist') or [])}
        above_count = max(0, saul_pos - 1)
        tied_count = max(0, _dm.get(saul_pts, sum(1 for r in nbs if r['pts'] == saul_pts)) - 1)
        below_count = max(0, total_part - saul_pos - tied_count)
        pts_above_list = [r['pts'] for r in nbs if r['pts'] > saul_pts]
        next_pts = min(pts_above_list) if pts_above_list else saul_pts + 1
        pts_to_next = next_pts - saul_pts
        next_pos = saul_pos
    else:
        above_count = sum(1 for r in ranking if r['pts'] > saul_pts)
        tied_count = max(0, sum(1 for r in ranking if r['pts'] == saul_pts) - 1)
        below_count = total_part - above_count - tied_count - 1
        pts_above_list = [r['pts'] for r in ranking if r['pts'] > saul_pts]
        next_pts = min(pts_above_list) if pts_above_list else saul_pts + 1
        pts_to_next = next_pts - saul_pts
        next_pos = min((r['pos'] for r in ranking if r['pts'] > saul_pts), default=saul_pos)

    if saul_entry is not None:
        if not saul_entry.get('_all'):
            saul_entry['_all'] = [{'pos': r.get('pos'), 'name': r.get('name'), 'pts': r.get('pts')} for r in ranking]
        if not saul_entry.get('_top3'):
            saul_entry['_top3'] = saul_entry['_all'][:3]
    distribution_html = build_distribution(ranking, saul_entry)
    insight_text = build_insight(saul_entry, ranking, today_matches)
    top3_html = build_top3(saul_entry)
    top_ranking_html = build_top_ranking(saul_entry, 10)
    full_ranking_html = build_full_ranking(saul_entry)
    dist_note = (f'Cuántos participantes hay en cada nivel de puntos (los 8 niveles más altos). '
                 f'Tú estás en <strong>{saul_pts} pts</strong> junto a otros {tied_count}.')
    hero_bg = (f'<img class="hero-emblem" src="data:image/png;base64,{logo_b64}" alt="FIFA World Cup 26">'
               if logo_b64 else '<div class="hero-26">26</div>')
    ft_bg = (f'<img class="ft-emblem" src="data:image/png;base64,{logo_b64}" alt="FIFA World Cup 26">'
             if logo_b64 else '<div class="ft-26">26</div>')

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Quiniela Mundial 2026 — Saúl Bretto</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Anton&family=Noto+Sans:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
  <style>
{CSS}
  </style>
</head>
<body>

  <div class="topbanner"><span>★ FIFA World Cup 26 · Quiniela Saúl Bretto · Caracas2go ★</span></div>

  <div class="topbar">
    <div class="tb-brand">
      <span class="tb-trophy">🏆</span>
      <span class="tb-wm">FIFA WORLD CUP 26<small>QUINIELA · SAÚL BRETTO</small></span>
    </div>
    <span class="tb-tag">26</span>
  </div>
  <div class="strip"><i style="background:var(--red)"></i><i style="background:var(--pur)"></i><i style="background:var(--blu)"></i><i style="background:var(--grn)"></i><i style="background:var(--tea)"></i><i style="background:var(--yel)"></i><i style="background:var(--pnk)"></i></div>

  <nav class="nav">
    <a href="#posicion">Mi posición</a>
    <a href="#proximos">Próximos</a>
    <a href="#voto">Cómo votó</a>
    <a href="#top5">Top 5</a>
    <a href="#stake">Combinadas</a>
    <a href="#lideres">Líderes</a>
    <a href="#stats">Números</a>
    <a href="#ranking">Grupo</a>
    <a href="#ranking-full">Todos</a>
  </nav>

  <div class="wrap">

    <section id="posicion" class="card hero">
      {hero_bg}
      <div class="eyebrow">Tu posición en la quiniela</div>
      <div class="rank">{saul_pos}<span>°</span></div>
      <div class="hsub">de {total_part} participantes · top {top_pct}%</div>
      <div class="name">SAÚL BRETTO</div>
      <span class="pts-pill">{saul_pts} PUNTOS</span>
    </section>

    <div class="grid">

      <section id="proximos" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--blu)"></span>Próximos Partidos<span class="smeta">{smeta}</span></div>
        <div class="matches">
{matches_html}
        </div>
      </section>

      <section id="voto" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--pnk)"></span>Cómo votó el grupo<span class="smeta">{total_part} QUINIELAS</span></div>
{voto_html}
      </section>

      <section id="top5" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--amb)"></span>Marcadores favoritos por partido<span class="smeta">{fav_smeta}</span></div>
        <div class="fav-grid">
{fav_html}
        </div>
        <div class="note">El marcador exacto que más puso el grupo en cada partido. La barra más larga = el marcador que más se repitió.</div>
      </section>

      <section id="stake" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--pur)"></span>Combinadas Stake · 1X2<span class="smeta">{rnd_label.upper()}</span></div>
        <div class="combos">
{combos_1x2}
        </div>
      </section>

      <section id="stake-exacto" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--amb)"></span>Combinadas Stake · Resultado exacto<span class="smeta">{rnd_label.upper()}</span></div>
        <div class="combos">
{combos_exacto}
        </div>
      </section>

      <section id="stake-goles" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--blu)"></span>Combinadas Stake · Goles<span class="smeta">{rnd_label.upper()}</span></div>
        <div class="combos">
{combos_goles}
        </div>
        <div class="note">Multiplicadores <strong>aproximados</strong> — confírmalos en Stake antes de apostar. Apuesta con responsabilidad.</div>
      </section>

      <section id="lideres" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--yel)"></span>Líderes del torneo<span class="smeta">TOP 3</span></div>
        <div class="top3">
{top3_html}
        </div>
        <div class="st" style="margin:18px 0 12px;"><span class="cbar" style="background:var(--tea)"></span>Etapas del torneo</div>
        <div class="stages">
          <span class="stage on">⚡ Grupos</span><span class="stage">Octavos</span><span class="stage">Cuartos</span><span class="stage">Semis</span><span class="stage">Final</span>
        </div>
      </section>

      <section id="stats" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--grn)"></span>Tus números</div>
        <div class="stats">
          <div class="stat"><div class="stat-num c-grn">{saul_pts}</div><div class="stat-lbl">Puntos</div><div class="stat-sub">Fase de grupos</div></div>
          <div class="stat"><div class="stat-num c-red">-{gap_leader}</div><div class="stat-lbl">Gap al líder</div><div class="stat-sub">{leader_name}: {leader_pts}</div></div>
          <div class="stat"><div class="stat-num c-blu">{below_count}</div><div class="stat-lbl">Por debajo tuyo</div><div class="stat-sub">{above_count} arriba · {tied_count} empatados</div></div>
          <div class="stat"><div class="stat-num c-amb">+{pts_to_next}</div><div class="stat-lbl">Para subir</div><div class="stat-sub">→ Puesto {next_pos} con {next_pts} pts</div></div>
        </div>
      </section>

      <section id="dist" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--red)"></span>Distribución de puntos<span class="smeta">TOP 8 NIVELES</span></div>
{distribution_html}
        <div class="note">{dist_note}</div>
      </section>

      <section id="ranking" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--pur)"></span>Ranking Top 10</div>
        <div class="nb">
{top_ranking_html}
        </div>
      </section>

      <section class="card span2">
        <div class="insight">
          <div class="insight-t">💡 Insight del día · {date_label}</div>
          <div class="insight-x">{insight_text}</div>
        </div>
      </section>

      <section id="ranking-full" class="card span2">
        <div class="st"><span class="cbar" style="background:var(--blu)"></span>Ranking completo<span class="smeta">{total_part} PARTICIPANTES</span></div>
        <div class="fullrk">
{full_ranking_html}
        </div>
        <div class="note">Diferencia respecto a ti: <strong>▲</strong> por encima · <strong>▼</strong> por debajo. Estás resaltado en amarillo.</div>
      </section>

    </div>

    <div class="ft">
      {ft_bg}
      <div class="ft-strip"><i style="background:var(--red)"></i><i style="background:var(--blu)"></i><i style="background:var(--grn)"></i><i style="background:var(--yel)"></i><i style="background:var(--pnk)"></i></div>
      <div class="ft-txt">Generado por el AI OS de Saúl Bretto · Caracas2go · {fecha_str}</div>
    </div>

  </div>
  <script>
    (function () {{
      var links = [].slice.call(document.querySelectorAll('.nav a'));
      var map = links.map(function (a) {{ return {{ a: a, el: document.getElementById(a.getAttribute('href').slice(1)) }}; }}).filter(function (x) {{ return x.el; }});
      function onScroll() {{
        var y = window.scrollY + 90, cur = map[0];
        map.forEach(function (x) {{ if (x.el.offsetTop <= y) cur = x; }});
        links.forEach(function (a) {{ a.classList.remove('on'); }});
        if (cur) cur.a.classList.add('on');
      }}
      window.addEventListener('scroll', onScroll, {{ passive: true }});
      onScroll();
    }})();
  </script>
</body>
</html>"""


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Quiniela Mundial 2026 — Generador de reporte HTML')
    parser.add_argument('--data-dir', required=True, help='Path to "QUINIELA MUNDIAL 2026" folder')
    parser.add_argument('--out', required=True, help='Output HTML file path')
    parser.add_argument('--fecha', required=True, help='Target date YYYY-MM-DD')
    parser.add_argument('--ranking-json', default=None, help='Optional JSON ranking override [{pos,name,pts}]')
    args = parser.parse_args()

    y, mo, d = args.fecha.split('-')
    target_date = date(int(y), int(mo), int(d))
    print(f"[INFO] Target date: {target_date}")

    out_path = args.out
    if os.path.isfile(out_path):
        backup_path = out_path.replace('.html', '.BACKUP.html')
        shutil.copy2(out_path, backup_path)
        print(f"[INFO] Backed up existing HTML to {backup_path}")

    print("[INFO] Parsing fixture and schedule...")
    dated_matches = build_dated_matches(args.data_dir)
    print(f"[INFO] Total dated matches: {len(dated_matches)}")

    today_matches = [m for m in dated_matches if m.get('date') == target_date]
    print(f"[INFO] Matches on {target_date}: {len(today_matches)}")
    for m in today_matches:
        print(f"  Group {m['group']} {m['round']}: {m['home']} vs {m['away']}")

    print("[INFO] Parsing CLAS ranking...")
    ranking, saul_entry = parse_clas(args.data_dir, args.ranking_json)
    print(f"[INFO] Ranking entries: {len(ranking)}")
    if saul_entry:
        print(f"[INFO] Saul: pos={saul_entry['pos']}, pts={saul_entry['pts']}")

    print("[INFO] Collecting votes from all quinielas...")
    vote_data = collect_votes(args.data_dir, today_matches)
    for m in today_matches:
        mkey = (m['group'], m['round'], m['home'], m['away'])
        v = vote_data.get(mkey, {})
        print(f"  {m['home']} vs {m['away']}: {v.get('total',0)} votes, {dict(v.get('outcome_counts',{}))}")

    print("[INFO] Loading Saul's picks...")
    saul_picks = collect_saul_picks(args.data_dir, today_matches)

    logo_b64 = ''
    logo_path = os.path.join(args.data_dir, '2026_FIFA_World_Cup_emblem.svg.png')
    if os.path.isfile(logo_path):
        with open(logo_path, 'rb') as lf:
            logo_b64 = base64.b64encode(lf.read()).decode('ascii')
        print(f"[INFO] Embedded FIFA emblem ({len(logo_b64)} b64 chars)")
    else:
        print("[WARN] FIFA emblem not found, using '26' fallback")

    print("[INFO] Generating HTML...")
    html = generate_html(target_date, dated_matches, vote_data, saul_picks, ranking, saul_entry, logo_b64)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[INFO] Report written to: {out_path}")
    print(f"[INFO] HTML size: {len(html):,} bytes")


if __name__ == '__main__':
    main()
