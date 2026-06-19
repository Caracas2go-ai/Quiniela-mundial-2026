/**
 * live_update.js — Núcleo del nodo Code de n8n para la Quiniela 2026 EN VIVO.
 * Toma resultados reales de ESPN (gratis, sin key) + picks congelados → ranking en vivo.
 * Independiente de Alejandro (sirve para comparar). FotMob queda solo para links.
 *
 * Fuente: https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard?dates=YYYYMMDD
 */
var engine = (typeof require !== 'undefined') ? require('./scoring_engine.js') : null;

// FIFA-3 (ESPN abbreviation) -> ISO-2 (nuestros códigos de bandera)
var FIFA3 = {MEX:'mx',RSA:'za',KOR:'kr',CZE:'cz',SUI:'ch',BIH:'ba',CAN:'ca',QAT:'qa',USA:'us',PAR:'py',
BRA:'br',MAR:'ma',HAI:'ht',SCO:'gb-sct',AUS:'au',TUR:'tr',GER:'de',CUW:'cw',NED:'nl',JPN:'jp',CIV:'ci',
ECU:'ec',SWE:'se',TUN:'tn',ESP:'es',CPV:'cv',BEL:'be',EGY:'eg',KSA:'sa',URU:'uy',IRN:'ir',NZL:'nz',
FRA:'fr',SEN:'sn',IRQ:'iq',NOR:'no',ARG:'ar',ALG:'dz',AUT:'at',JOR:'jo',POR:'pt',COD:'cd',ENG:'gb-eng',
CRO:'hr',GHA:'gh',PAN:'pa',UZB:'uz',COL:'co'};

// Extrae partidos de la respuesta de ESPN (uno o varios días concatenados).
function parseEspn(jsonOrArray) {
  var jsons = Array.isArray(jsonOrArray) ? jsonOrArray : [jsonOrArray];
  var out = [];
  jsons.forEach(function (j) {
    (j && j.events || []).forEach(function (e) {
      var c = e.competitions && e.competitions[0]; if (!c) return;
      var cs = c.competitors || [];
      var h = cs.filter(function (x) { return x.homeAway === 'home'; })[0];
      var a = cs.filter(function (x) { return x.homeAway === 'away'; })[0];
      if (!h || !a) return;
      var st = e.status && e.status.type || {};
      out.push({
        hc: String(h.team.abbreviation || '').toUpperCase(),
        ac: String(a.team.abbreviation || '').toUpperCase(),
        hs: parseInt(h.score, 10), as: parseInt(a.score, 10),
        completed: !!st.completed, state: st.name || ''
      });
    });
  });
  return out;
}

// Construye el mapa de resultados {codigoPartido: "h-a"} en NUESTRA orientación.
// Mapea por código ISO (m.hc/m.ac vienen en los picks), no por nombre.
function buildResultsMap(espnEvents, picksData) {
  var byPair = {};
  picksData.matches.forEach(function (m) {
    if (!/^[A-L]\d/.test(String(m.code || ''))) return;          // solo fase de grupos (códigos A1..L6)
    if (m.hc && m.ac && m.hc !== 'un' && m.ac !== 'un') byPair[[m.hc, m.ac].sort().join('|')] = m;
  });
  var resultsMap = {}, live = {};
  espnEvents.forEach(function (ev) {
    var hi = FIFA3[ev.hc], ai = FIFA3[ev.ac];
    if (!hi || !ai || isNaN(ev.hs) || isNaN(ev.as)) return;
    var m = byPair[[hi, ai].sort().join('|')];
    if (!m) return;
    var hs = ev.hs, as = ev.as;
    if (m.hc === ai) { hs = ev.as; as = ev.hs; } // nuestra local = visitante de ESPN -> voltear
    var score = hs + '-' + as;
    var mid = m.group + '_J' + String(m.code).slice(1) + '_' + m.hc + '_' + m.ac; // = mid del sitio
    if (ev.completed) resultsMap[m.home + '|' + m.away] = score; // clave única para el motor
    live[mid] = { score: score, completed: ev.completed, state: ev.state, home: m.home, away: m.away };
  });
  return { resultsMap: resultsMap, live: live };
}

// Payload final que n8n publica (lo lee el sitio).
function buildLive(espnJson, picksData, teams, prevStandings) {
  var ev = parseEspn(espnJson);
  var rb = buildResultsMap(ev, picksData);
  var res = engine.computeStandings(picksData, rb.resultsMap);
  // resumen por jugador: cuánto sumó respecto al ranking previo (para "subiste X pts")
  var prev = {}; (prevStandings || []).forEach(function (e) { prev[e.name] = e.pts; });
  var standings = res.standings.map(function (e) {
    return { pos: e.pos, name: e.name, pts: e.pts, delta: (prev[e.name] != null ? e.pts - prev[e.name] : 0) };
  });
  return {
    generatedAt: new Date().toISOString(),
    finished: Object.keys(rb.resultsMap).length,
    results: rb.live,            // {code: {score, completed, state}}
    standings: standings
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parseEspn: parseEspn, buildResultsMap: buildResultsMap, buildLive: buildLive, FIFA3: FIFA3 };
}
