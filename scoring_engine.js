/**
 * scoring_engine.js — Motor de puntaje Quiniela Mundial 2026 (validado 100% vs ADMIN de Alejandro)
 * Pensado para correr tal cual en un nodo Code de n8n.
 *
 * Entrada:
 *   data = { matches:[{code,home,away,group,result}], players:[name...], picks:{ name:[pickStr por índice de match] } }
 *          (pickStr formato "o|h-a", ej "1|2-0"; result formato "h-a")
 *   resultsMap (opcional) = resultados EN VIVO que pisan data.matches[i].result.
 *          claves admitidas: code (ej "A1") o "home|away".
 *
 * Reglas (de la hoja ADMIN):
 *   Fase de grupos: Signo 1X2 = +1 · Diferencia (si signo OK) = +1 · Exacto = +2  (máx 4)
 *   Eliminatorias (signo/dif/exacto por fase): 16avos 2/2/2 · 8vos 3/3/3 · 4tos 4/4/4 ·
 *                  Semis 5/5/5 · 3º-4º 6/6/6 · Final 8/8/8
 *   (Campeón 15, Subcampeón 10, Bota de Oro 8 y posición exacta de grupo +2 se suman aparte
 *    cuando se resuelvan; no afectan el cálculo por partido.)
 */
function stagePoints(code, group) {
  // group-stage codes: letra + dígito (A1, B2, L3...). Eliminatorias usan otros códigos.
  if (group && /^[A-L]$/.test(group)) return { s: 1, d: 1, e: 2 };
  const c = String(code || '').toUpperCase();
  if (/^[A-L]\d/.test(c)) return { s: 1, d: 1, e: 2 };          // grupos
  if (/16|R32|DIEC/.test(c)) return { s: 2, d: 2, e: 2 };       // 16avos
  if (/OCT|R16|8VO/.test(c)) return { s: 3, d: 3, e: 3 };       // octavos
  if (/CUART|QF|R8|4TO/.test(c)) return { s: 4, d: 4, e: 4 };   // cuartos
  if (/SEMI|SF/.test(c)) return { s: 5, d: 5, e: 5 };           // semis
  if (/3|TER|BRONCE/.test(c)) return { s: 6, d: 6, e: 6 };      // 3º y 4º
  if (/FINAL|^F$/.test(c)) return { s: 8, d: 8, e: 8 };         // final
  return { s: 1, d: 1, e: 2 };                                  // fallback grupos
}

function parseScore(s) {
  const m = /^\s*(\d+)\s*-\s*(\d+)\s*$/.exec(s || '');
  return m ? [parseInt(m[1], 10), parseInt(m[2], 10)] : null;
}
function outcome(h, a) { return h > a ? '1' : (h < a ? '2' : 'X'); }

function pointsFor(pick, score, st) {
  if (!pick || pick.indexOf('|') < 0) return 0;
  const parts = pick.split('|');
  const po = parts[0].trim();
  const pr = parseScore(parts[1]);
  const rr = parseScore(score);
  if (!rr) return 0;
  const ro = outcome(rr[0], rr[1]);
  let p = 0;
  if (po === ro) p += st.s;
  if (pr && po === ro && (pr[0] - pr[1]) === (rr[0] - rr[1])) p += st.d;
  if (pr && pr[0] === rr[0] && pr[1] === rr[1]) p += st.e;
  return p;
}

function computeStandings(data, resultsMap) {
  resultsMap = resultsMap || {};
  const totals = {};
  data.players.forEach(function (n) { totals[n] = 0; });
  const perMatch = [];
  data.matches.forEach(function (m, i) {
    const live = resultsMap[m.code] || resultsMap[m.home + '|' + m.away];
    const score = live || m.result || '';
    perMatch.push({ code: m.code, home: m.home, away: m.away, result: score });
    if (!score) return;
    const st = stagePoints(m.code, m.group);
    data.players.forEach(function (n) {
      const arr = data.picks[n];
      if (arr) totals[n] += pointsFor(arr[i], score, st);
    });
  });
  const standings = Object.keys(totals).map(function (name) {
    return { name: name, pts: totals[name] };
  }).sort(function (a, b) { return b.pts - a.pts || a.name.localeCompare(b.name, 'es'); });
  let pos = 0, last = null, lastPos = 0;
  standings.forEach(function (e, i) {
    pos = i + 1;
    if (e.pts === last) e.pos = lastPos; else { e.pos = pos; last = e.pts; lastPos = pos; }
  });
  return { standings: standings, results: perMatch };
}

// Export para Node (test local) y uso directo en n8n.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { computeStandings: computeStandings, pointsFor: pointsFor, stagePoints: stagePoints };
}
