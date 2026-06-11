#!/usr/bin/env python3
"""
Generador del dashboard de la Porra Mundial 2026 (Reveni).

Lee el Excel de administración (Porra_Admin_v4_EN.xlsx) y produce un único
fichero HTML autocontenido con analíticas estilo "quién es el más rebelde",
"quién piensa como quién", estilo de apuesta, favoritos, etc.

Reejecutable: cuando se rellenen resultados/eliminatorias en el Excel, basta
con volver a correr el script y el HTML se regenera (incluyendo aciertos en
directo si ya hay resultados).

Uso:
    python3 generate_dashboard.py [ruta_xlsx] [ruta_salida_html]
"""

from __future__ import annotations

import json
import sys
import statistics
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

# --------------------------------------------------------------------------
# Configuración / constantes
# --------------------------------------------------------------------------

DEFAULT_XLSX = "Porra_Admin_v4_EN.xlsx"
DEFAULT_OUT = "index.html"  # raíz del sitio de GitHub Pages
LOGO_SVG_PATH = "reveni-logo.svg"

N_PARTICIPANTS_MAX = 30          # huecos P1..P30 en el Excel
RAW_FIRST_HOME_COL = 3           # columna C = primer "home" de P1
NAME_ROW = 6                     # fila con los nombres reales
GROUP_MATCH_ROWS = range(7, 79)  # 72 partidos de fase de grupos
QUALIFIER_ROWS = range(80, 116)  # 1st/2nd/3rd de cada grupo (A..L)
THIRDS_ROWS = range(117, 125)    # 8 mejores terceros

# Traducción EN -> ES (reconocible) + bandera
TEAMS = {
    "Mexico": ("México", "🇲🇽"), "South Africa": ("Sudáfrica", "🇿🇦"),
    "South Korea": ("Corea del Sur", "🇰🇷"), "Czech Rep.": ("Rep. Checa", "🇨🇿"),
    "Canada": ("Canadá", "🇨🇦"), "Bosnia-Herz.": ("Bosnia", "🇧🇦"),
    "Qatar": ("Catar", "🇶🇦"), "Switzerland": ("Suiza", "🇨🇭"),
    "Brazil": ("Brasil", "🇧🇷"), "Morocco": ("Marruecos", "🇲🇦"),
    "Haiti": ("Haití", "🇭🇹"), "Scotland": ("Escocia", "🏴󠁧󠁢󠁳󠁣󠁴󠁿"),
    "USA": ("EE. UU.", "🇺🇸"), "Paraguay": ("Paraguay", "🇵🇾"),
    "Australia": ("Australia", "🇦🇺"), "Turkey": ("Turquía", "🇹🇷"),
    "Germany": ("Alemania", "🇩🇪"), "Curacao": ("Curazao", "🇨🇼"),
    "Ivory Coast": ("Costa de Marfil", "🇨🇮"), "Ecuador": ("Ecuador", "🇪🇨"),
    "Netherlands": ("Países Bajos", "🇳🇱"), "Japan": ("Japón", "🇯🇵"),
    "Sweden": ("Suecia", "🇸🇪"), "Tunisia": ("Túnez", "🇹🇳"),
    "Belgium": ("Bélgica", "🇧🇪"), "Egypt": ("Egipto", "🇪🇬"),
    "Iran": ("Irán", "🇮🇷"), "New Zealand": ("Nueva Zelanda", "🇳🇿"),
    "Spain": ("España", "🇪🇸"), "Cape Verde": ("Cabo Verde", "🇨🇻"),
    "Uruguay": ("Uruguay", "🇺🇾"), "Saudi Arabia": ("Arabia Saudí", "🇸🇦"),
    "France": ("Francia", "🇫🇷"), "Senegal": ("Senegal", "🇸🇳"),
    "Iraq": ("Irak", "🇮🇶"), "Norway": ("Noruega", "🇳🇴"),
    "Argentina": ("Argentina", "🇦🇷"), "Algeria": ("Argelia", "🇩🇿"),
    "Austria": ("Austria", "🇦🇹"), "Jordan": ("Jordania", "🇯🇴"),
    "Portugal": ("Portugal", "🇵🇹"), "DR Congo": ("RD Congo", "🇨🇩"),
    "Uzbekistan": ("Uzbekistán", "🇺🇿"), "Colombia": ("Colombia", "🇨🇴"),
    "England": ("Inglaterra", "🏴󠁧󠁢󠁥󠁮󠁧󠁿"), "Croatia": ("Croacia", "🇭🇷"),
    "Ghana": ("Ghana", "🇬🇭"), "Panama": ("Panamá", "🇵🇦"),
}


ES2EN = {es: en for en, (es, _fl) in TEAMS.items()}


def team_es(name):
    if name is None:
        return None
    name = str(name).strip()
    return TEAMS.get(name, (name, "🏳️"))[0]


def team_flag(name):
    if name is None:
        return "🏳️"
    return TEAMS.get(str(name).strip(), (name, "🏳️"))[1]


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------

def parse_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    raw = wb["Raw data"]

    # Localiza columnas de participantes con nombre en NAME_ROW
    participants = []  # [{name, home_col, away_col, idx}]
    for slot in range(N_PARTICIPANTS_MAX):
        home_col = RAW_FIRST_HOME_COL + 2 * slot
        name = raw.cell(NAME_ROW, home_col).value
        if name is not None and str(name).strip():
            participants.append({
                "name": str(name).strip(),
                "home_col": home_col,
                "away_col": home_col + 1,
                "idx": len(participants),
            })
    names = [p["name"] for p in participants]
    n = len(participants)

    # Partidos de fase de grupos
    matches = []
    for r in GROUP_MATCH_ROWS:
        concept = raw.cell(r, 2).value
        if not concept or "vs" not in str(concept):
            continue
        concept = str(concept).strip()
        code, _, teams = concept.partition(":")
        code = code.strip()
        home_en, _, away_en = teams.strip().partition(" vs ")
        home_en, away_en = home_en.strip(), away_en.strip()
        group = code.split("-")[0].replace("G", "")  # "GA-M1" -> "A"
        picks = []
        for p in participants:
            hg = raw.cell(r, p["home_col"]).value
            ag = raw.cell(r, p["away_col"]).value
            picks.append((_num(hg), _num(ag)))
        matches.append({
            "code": code, "group": group,
            "home_en": home_en, "away_en": away_en,
            "home": team_es(home_en), "away": team_es(away_en),
            "home_flag": team_flag(home_en), "away_flag": team_flag(away_en),
            "picks": picks,
        })

    # Clasificados por grupo (1/2/3)
    qualifiers = defaultdict(dict)  # group -> {pos -> [team per participant]}
    for r in QUALIFIER_ROWS:
        label = raw.cell(r, 2).value
        if not label:
            continue
        label = str(label).strip().lower()
        pos = None
        if label.startswith("1st"):
            pos = 1
        elif label.startswith("2nd"):
            pos = 2
        elif label.startswith("3rd"):
            pos = 3
        if pos is None:
            continue
        group = label.split("group")[-1].strip().upper()
        picks = [raw.cell(r, p["home_col"]).value for p in participants]
        qualifiers[group][pos] = picks

    # 8 mejores terceros
    thirds = []  # list per row -> [team per participant]
    for r in THIRDS_ROWS:
        label = raw.cell(r, 2).value
        if not label:
            continue
        picks = [raw.cell(r, p["home_col"]).value for p in participants]
        thirds.append(picks)

    # ¿Hay resultados reales?
    results = parse_results(wb)

    return {
        "participants": participants,
        "names": names,
        "n": n,
        "matches": matches,
        "qualifiers": qualifiers,
        "thirds": thirds,
        "results": results,
    }


def parse_results(wb):
    """Lee resultados reales de fase de grupos si existen. Devuelve {code:(h,a)}."""
    rr = wb["Real results"]
    out = {}
    for r in range(1, rr.max_row + 1):
        concept = rr.cell(r, 2).value
        if not concept or "vs" not in str(concept):
            continue
        code = str(concept).split(":")[0].strip()
        h = _num(rr.cell(r, 3).value)
        a = _num(rr.cell(r, 4).value)
        if h is not None and a is not None:
            out[code] = (h, a)
    return out


def _num(v):
    if v is None or v == "":
        return None
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------
# Utilidades de analítica
# --------------------------------------------------------------------------

def outcome(hg, ag):
    if hg is None or ag is None:
        return None
    if hg > ag:
        return "1"
    if hg < ag:
        return "2"
    return "X"


def minmax_scale(values):
    lo, hi = min(values), max(values)
    if hi - lo < 1e-9:
        return [50.0 for _ in values]
    return [100.0 * (v - lo) / (hi - lo) for v in values]


def zscore(values):
    mu = statistics.mean(values)
    sd = statistics.pstdev(values) or 1.0
    return [(v - mu) / sd for v in values]


# --------------------------------------------------------------------------
# Motor de analíticas
# --------------------------------------------------------------------------

def compute(data):
    names = data["names"]
    n = data["n"]
    matches = data["matches"]
    nm = len(matches)

    # ---- Índice de Rebeldía (signo minoría + distancia marcador) ----
    raw_sign = [0.0] * n
    raw_score = [0.0] * n
    match_meta = []  # por partido: agreement, mediana, etc.

    for m in matches:
        picks = m["picks"]
        valid = [(i, h, a) for i, (h, a) in enumerate(picks) if h is not None and a is not None]
        if not valid:
            match_meta.append(None)
            continue
        outcomes = [outcome(h, a) for _, h, a in valid]
        oc = Counter(outcomes)
        total = len(valid)
        homes = [h for _, h, a in valid]
        aways = [a for _, h, a in valid]
        med_h = statistics.median(homes)
        med_a = statistics.median(aways)
        # distancia de marcador por persona
        sdist = {}
        for i, h, a in valid:
            sdist[i] = abs(h - med_h) + abs(a - med_a)
        max_sdist = max(sdist.values()) or 1.0
        for i, h, a in valid:
            o = outcome(h, a)
            share = oc[o] / total
            raw_sign[i] += (1.0 - share)
            raw_score[i] += sdist[i] / max_sdist
        # scoreline modal
        scl = Counter((h, a) for _, h, a in valid)
        modal_scl, modal_cnt = scl.most_common(1)[0]
        top_outcome, top_cnt = oc.most_common(1)[0]
        match_meta.append({
            "code": m["code"], "group": m["group"],
            "home": m["home"], "away": m["away"],
            "home_en": m["home_en"], "away_en": m["away_en"],
            "home_flag": m["home_flag"], "away_flag": m["away_flag"],
            "outcome_dist": dict(oc),
            "outcome_agreement": top_cnt / total,
            "top_outcome": top_outcome,
            "modal_scoreline": f"{modal_scl[0]}-{modal_scl[1]}",
            "modal_scoreline_share": modal_cnt / total,
            "median": f"{int(med_h)}-{int(med_a)}",
            "n": total,
        })

    raw_sign = [s / nm for s in raw_sign]
    raw_score = [s / nm for s in raw_score]
    z = [0.5 * a + 0.5 * b for a, b in zip(zscore(raw_sign), zscore(raw_score))]
    rebel_index = [round(v, 1) for v in minmax_scale(z)]

    rebeldia = sorted(
        [{"name": names[i], "index": rebel_index[i],
          "sign": round(raw_sign[i], 3), "score": round(raw_score[i], 3)}
         for i in range(n)],
        key=lambda x: -x["index"],
    )

    # ---- Similitud: % marcadores idénticos + distancia media de goles ----
    exact = [[0] * n for _ in range(n)]
    gdist = [[0.0] * n for _ in range(n)]
    comparable = 0
    for m in matches:
        picks = m["picks"]
        if all(h is not None for h, a in picks):
            comparable += 1
        for i in range(n):
            hi, ai = picks[i]
            if hi is None:
                continue
            for j in range(i + 1, n):
                hj, aj = picks[j]
                if hj is None:
                    continue
                if hi == hj and ai == aj:
                    exact[i][j] += 1
                    exact[j][i] += 1
                d = abs(hi - hj) + abs(ai - aj)
                gdist[i][j] += d
                gdist[j][i] += d

    sim_pct = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            if i != j and comparable:
                sim_pct[i][j] = round(100.0 * exact[i][j] / comparable, 1)

    # almas gemelas (máx % idéntico) y polos opuestos (máx distancia)
    best_pair = (None, None, -1)
    worst_pair = (None, None, -1)
    for i in range(n):
        for j in range(i + 1, n):
            if sim_pct[i][j] > best_pair[2]:
                best_pair = (i, j, sim_pct[i][j])
            avg_d = gdist[i][j] / (comparable or 1)
            if avg_d > worst_pair[2]:
                worst_pair = (i, j, round(avg_d, 2))

    # gemelo de cada uno
    twins = []
    for i in range(n):
        best_j, best_v = None, -1
        for j in range(n):
            if i != j and sim_pct[i][j] > best_v:
                best_v, best_j = sim_pct[i][j], j
        twins.append({"name": names[i], "twin": names[best_j], "pct": best_v})

    # ---- Estilo de apuesta ----
    style = []
    all_scorelines = Counter()
    for i in range(n):
        goals, draws, blowouts, biggest = [], 0, 0, (None, -1)
        for m in matches:
            h, a = m["picks"][i]
            if h is None:
                continue
            goals.append(h + a)
            if h == a:
                draws += 1
            if abs(h - a) >= 3:
                blowouts += 1
            if h + a > biggest[1]:
                biggest = ((m, h, a), h + a)
            all_scorelines[(h, a)] += 1
        cnt = len(goals) or 1
        big = None
        if biggest[0]:
            mm, h, a = biggest[0]
            big = {"home": mm["home"], "away": mm["away"],
                   "home_en": mm["home_en"], "away_en": mm["away_en"],
                   "score": f"{h}-{a}",
                   "flags": f"{mm['home_flag']}{mm['away_flag']}", "goals": h + a}
        style.append({
            "name": names[i],
            "avg_goals": round(sum(goals) / cnt, 2),
            "pct_draws": round(100 * draws / cnt, 1),
            "pct_blowouts": round(100 * blowouts / cnt, 1),
            "biggest": big,
        })

    common_scorelines = [
        {"score": f"{h}-{a}", "count": c,
         "pct": round(100 * c / (nm * n), 1)}
        for (h, a), c in all_scorelines.most_common(8)
    ]

    # ---- Lobo solitario: marcadores únicos (solo 1 persona) ----
    unique_count = [0] * n
    unique_examples = []  # picks únicos llamativos
    for m in matches:
        scl = Counter()
        for h, a in m["picks"]:
            if h is not None:
                scl[(h, a)] += 1
        for i, (h, a) in enumerate(m["picks"]):
            if h is not None and scl[(h, a)] == 1:
                unique_count[i] += 1
                diff = abs(h - a) + (h + a)
                unique_examples.append({
                    "name": names[i], "home": m["home"], "away": m["away"],
                    "home_en": m["home_en"], "away_en": m["away_en"],
                    "flags": f"{m['home_flag']}{m['away_flag']}",
                    "score": f"{h}-{a}", "spice": diff,
                })
    lobo = sorted(
        [{"name": names[i], "count": unique_count[i]} for i in range(n)],
        key=lambda x: -x["count"],
    )
    unique_examples.sort(key=lambda x: -x["spice"])

    # ---- Favoritos del torneo ----
    group_consensus = []
    respect = Counter()        # prestigio ponderado: 1º=3, 2º=2, 3º=1
    pos_weight = {1: 3, 2: 2, 3: 1}
    for group in sorted(data["qualifiers"].keys()):
        pos = data["qualifiers"][group]
        firsts = [team_es(t) for t in pos.get(1, []) if t]
        c = Counter(firsts)
        for p, slot_picks in pos.items():
            w = pos_weight.get(p, 1)
            for t in slot_picks:
                if t:
                    respect[team_es(t)] += w
        if c:
            top, cnt = c.most_common(1)[0]
            group_consensus.append({
                "group": group, "favorite": top,
                "flag": _flag_es(top),
                "agreement": round(100 * cnt / sum(c.values()), 1),
                "dist": dict(c.most_common()),
            })
    most_respected = [{"team": t, "flag": _flag_es(t), "count": c}
                      for t, c in respect.most_common(12)]
    least_respected = [{"team": t, "flag": _flag_es(t), "count": c}
                       for t, c in sorted(respect.items(), key=lambda x: x[1])[:8]]
    polarizing = sorted(group_consensus, key=lambda x: x["agreement"])[:5]

    # terceros más elegidos
    thirds_counter = Counter()
    for row in data["thirds"]:
        for t in row:
            if t:
                thirds_counter[team_es(t)] += 1
    top_thirds = [{"team": t, "flag": _flag_es(t), "count": c}
                  for t, c in thirds_counter.most_common(10)]

    # ---- Partidos divisivos / unánimes ----
    valid_meta = [m for m in match_meta if m]
    divisive = sorted(valid_meta, key=lambda m: (m["outcome_agreement"], m["modal_scoreline_share"]))[:6]
    unanimous = sorted(valid_meta, key=lambda m: (-m["modal_scoreline_share"], -m["outcome_agreement"]))[:6]

    # ---- Fichas por persona ----
    rank_of = {r["name"]: i + 1 for i, r in enumerate(rebeldia)}
    idx_of = {r["name"]: r["index"] for r in rebeldia}
    style_of = {s["name"]: s for s in style}
    twin_of = {t["name"]: t for t in twins}
    lobo_of = {l["name"]: l["count"] for l in lobo}
    cards = []
    for i, name in enumerate(names):
        st = style_of[name]
        cards.append({
            "name": name,
            "rebel_rank": rank_of[name],
            "rebel_index": idx_of[name],
            "avg_goals": st["avg_goals"],
            "pct_draws": st["pct_draws"],
            "biggest": st["biggest"],
            "twin": twin_of[name]["twin"],
            "twin_pct": twin_of[name]["pct"],
            "lobo": lobo_of[name],
            "label": style_label(st, idx_of[name]),
        })

    # ---- Premios ----
    goleador = max(style, key=lambda s: s["avg_goals"])
    cerrojo = min(style, key=lambda s: s["avg_goals"])
    empate = max(style, key=lambda s: s["pct_draws"])
    awards = {
        "rebelde": rebeldia[0],
        "borrego": rebeldia[-1],
        "gemelas": {"a": names[best_pair[0]], "b": names[best_pair[1]], "pct": best_pair[2]},
        "opuestos": {"a": names[worst_pair[0]], "b": names[worst_pair[1]], "dist": worst_pair[2]},
        "goleador": {"name": goleador["name"], "avg": goleador["avg_goals"]},
        "cerrojo": {"name": cerrojo["name"], "avg": cerrojo["avg_goals"]},
        "empate": {"name": empate["name"], "pct": empate["pct_draws"]},
        "lobo": lobo[0],
    }

    # ---- Stats del hero ----
    total_goals = sum(
        h + a for m in matches for (h, a) in m["picks"] if h is not None
    )
    hero = {
        "participants": n,
        "matches": nm,
        "groups": len(data["qualifiers"]),
        "total_goals": total_goals,
        "avg_goals_match": round(total_goals / (nm * n), 2) if nm and n else 0,
        "top_scoreline": common_scorelines[0] if common_scorelines else None,
        "has_results": bool(data["results"]),
    }

    # ---- Aciertos en directo (solo si hay resultados) ----
    live = None
    if data["results"]:
        live = compute_live(data, matches)

    return {
        "es2en": ES2EN,
        "hero": hero,
        "rebeldia": rebeldia,
        "matrix": {"names": names, "sim": sim_pct},
        "twins": twins,
        "best_pair": {"a": names[best_pair[0]], "b": names[best_pair[1]], "pct": best_pair[2]},
        "worst_pair": {"a": names[worst_pair[0]], "b": names[worst_pair[1]], "dist": worst_pair[2]},
        "style": style,
        "common_scorelines": common_scorelines,
        "lobo": lobo,
        "unique_examples": unique_examples[:12],
        "group_consensus": group_consensus,
        "most_respected": most_respected,
        "least_respected": least_respected,
        "polarizing": polarizing,
        "top_thirds": top_thirds,
        "divisive": divisive,
        "unanimous": unanimous,
        "cards": cards,
        "awards": awards,
        "live": live,
    }


def _flag_es(es_name):
    for en, (es, fl) in TEAMS.items():
        if es == es_name:
            return fl
    return "🏳️"


def style_label(st, rebel_index):
    """Devuelve una clave; el texto/emoji se traduce en el cliente (JS)."""
    if st["avg_goals"] >= 3.0:
        return "goleador"
    if st["avg_goals"] <= 2.0:
        return "cerrojo"
    if st["pct_draws"] >= 35:
        return "empate"
    if rebel_index >= 70:
        return "rebelde"
    if rebel_index <= 25:
        return "borrego"
    return "equilibrado"


def compute_live(data, matches):
    """Ranking simplificado de aciertos cuando hay resultados (signo=2, pleno=4)."""
    results = data["results"]
    names = data["names"]
    n = data["n"]
    pts = [0] * n
    exact = [0] * n
    sign = [0] * n
    played = 0
    for m in matches:
        if m["code"] not in results:
            continue
        rh, ra = results[m["code"]]
        ro = outcome(rh, ra)
        played += 1
        for i, (h, a) in enumerate(m["picks"]):
            if h is None:
                continue
            if h == rh and a == ra:
                pts[i] += 4
                exact[i] += 1
            elif outcome(h, a) == ro:
                pts[i] += 2
                sign[i] += 1
    table = sorted(
        [{"name": names[i], "pts": pts[i], "exact": exact[i], "sign": sign[i]}
         for i in range(n)],
        key=lambda x: -x["pts"],
    )
    return {"played": played, "table": table}


# --------------------------------------------------------------------------
# main (parte de verificación; el render se añade después)
# --------------------------------------------------------------------------

CSS = r"""
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#022f36;--bg2:#001a1f;--teal:#003C46;--surface:#063a44;--surface2:#0a4a56;
  --mint:#7afcd0;--mint2:#49d6bb;--mint-soft:#defff0;--line:rgba(122,252,208,.15);
  --text:#e2f6ef;--muted:#86b1aa;--gold:#ffd27a;--red:#ff7a7a;--maxw:1060px;
}
html{scroll-behavior:smooth}
body{background:radial-gradient(1200px 800px at 75% -10%,#0a4a56 0%,var(--bg) 45%,var(--bg2) 100%);
  color:var(--text);font-family:'Inter',system-ui,-apple-system,Segoe UI,Roboto,sans-serif;
  line-height:1.5;-webkit-font-smoothing:antialiased;min-height:100vh;overflow-x:hidden}
h1,h2,h3,.disp{font-family:'Space Grotesk','Inter',system-ui,sans-serif;font-weight:700;letter-spacing:-.01em;line-height:1.05}
a{color:inherit}
.mint{color:var(--mint)}.muted{color:var(--muted)}
/* NAV */
nav.rail{position:fixed;top:0;left:0;height:100vh;width:188px;padding:26px 18px;display:flex;flex-direction:column;gap:4px;
  border-right:1px solid var(--line);background:linear-gradient(180deg,rgba(0,26,31,.6),transparent);backdrop-filter:blur(6px);z-index:40}
nav.rail .brand{color:var(--mint);height:26px;margin-bottom:18px;display:block}
nav.rail .brand svg{height:26px;width:auto}
.langtoggle{display:flex;gap:4px;margin-bottom:16px;background:rgba(0,0,0,.2);border:1px solid var(--line);border-radius:10px;padding:3px}
.langtoggle button{flex:1;border:0;background:transparent;color:var(--muted);font:600 .78rem 'Space Grotesk',sans-serif;
  padding:6px 0;border-radius:7px;cursor:pointer;letter-spacing:.05em;transition:.18s}
.langtoggle button.on{background:var(--mint);color:#012}
nav.rail a{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:10px;color:var(--muted);
  text-decoration:none;font-size:.82rem;font-weight:600;transition:.2s}
nav.rail a .dot{width:7px;height:7px;border-radius:50%;background:currentColor;opacity:.5;transition:.2s}
nav.rail a:hover{color:var(--text);background:rgba(122,252,208,.05)}
nav.rail a.active{color:var(--mint);background:rgba(122,252,208,.1)}
nav.rail a.active .dot{opacity:1;box-shadow:0 0 10px var(--mint)}
.wrap{max-width:var(--maxw);margin:0 auto;padding:0 28px 0 calc(188px + 40px)}
/* HERO */
header.hero{min-height:100vh;display:flex;flex-direction:column;justify-content:center;position:relative;padding-top:40px}
.hero .eyebrow{color:var(--mint);font-weight:700;letter-spacing:.22em;text-transform:uppercase;font-size:.78rem;margin-bottom:18px}
.hero h1{font-size:clamp(2.6rem,8vw,5.6rem);background:linear-gradient(120deg,#fff 10%,var(--mint) 90%);
  -webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;margin-bottom:10px}
.hero .lead{font-size:clamp(1rem,2vw,1.3rem);color:var(--muted);max-width:620px;margin-bottom:42px}
.chips{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;max-width:920px}
.chip{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:18px 22px;min-width:0}
.chip .big{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.chip .big{font-family:'Space Grotesk';font-size:2rem;font-weight:700;color:var(--mint)}
.chip .lab{font-size:.78rem;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-top:4px}
.scrollcue{position:absolute;bottom:26px;left:0;color:var(--muted);font-size:.8rem;display:flex;align-items:center;gap:8px;animation:bob 1.8s ease-in-out infinite}
@keyframes bob{0%,100%{transform:translateY(0)}50%{transform:translateY(6px)}}
/* SECTIONS */
section.sec{padding:74px 0;border-top:1px solid var(--line)}
.sec-head{margin-bottom:34px}
.kicker{font-family:'Space Grotesk';color:var(--mint);font-weight:700;font-size:.9rem;letter-spacing:.1em;margin-bottom:8px}
.sec h2{font-size:clamp(1.8rem,4vw,2.9rem)}
.sec .sub{color:var(--muted);max-width:640px;margin-top:10px;font-size:1.02rem}
.grid{display:grid;gap:16px}
.g2{grid-template-columns:repeat(2,1fr)}.g3{grid-template-columns:repeat(3,1fr)}.g4{grid-template-columns:repeat(4,1fr)}
.card{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:22px}
.card.glow{box-shadow:0 0 0 1px rgba(122,252,208,.2),0 20px 60px rgba(122,252,208,.08)}
.card h3{font-size:1.05rem;margin-bottom:4px}
.card .k{font-size:.74rem;color:var(--muted);text-transform:uppercase;letter-spacing:.1em}
.big-num{font-family:'Space Grotesk';font-size:2.4rem;font-weight:700;color:var(--mint);line-height:1}
/* BARS */
.bar-row{display:grid;grid-template-columns:30px minmax(84px,170px) 1fr 54px;align-items:center;gap:14px;padding:7px 0}
.bar-rank{font-family:'Space Grotesk';color:var(--muted);font-size:.95rem;text-align:right;font-weight:700}
.bar-name{font-weight:600;font-size:.95rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar-co{display:flex;flex-direction:column;gap:3px;min-width:0}
.bar-track{height:13px;background:rgba(255,255,255,.06);border-radius:8px;overflow:hidden}
.bar-fill{height:100%;width:0;border-radius:8px;background:linear-gradient(90deg,var(--mint2),var(--mint));transition:width 1s cubic-bezier(.2,.8,.2,1)}
.bar-fill.cool{background:linear-gradient(90deg,#2b6f7f,#4a9aa8)}
.bar-fill.gold{background:linear-gradient(90deg,#caa24a,var(--gold))}
.bar-val{font-family:'Space Grotesk';font-weight:700;text-align:right;font-size:.95rem}
/* PODIUM */
.podium{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;align-items:end;margin-bottom:30px}
.pod{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:20px;text-align:center;position:relative}
.pod .medal{font-size:1.6rem}.pod .nm{font-weight:700;font-size:1.1rem;margin:6px 0 2px}
.pod .sc{font-family:'Space Grotesk';font-size:2rem;font-weight:700;color:var(--mint)}
.pod.p1{transform:translateY(-12px);box-shadow:0 0 0 1px rgba(122,252,208,.25),0 18px 50px rgba(122,252,208,.1)}
/* HEADLINE pair cards */
.duo{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:26px}
.duo .card{display:flex;flex-direction:column;gap:6px}
.duo .names{font-family:'Space Grotesk';font-size:1.5rem;font-weight:700}
.tag{display:inline-block;font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;
  padding:4px 9px;border-radius:999px;background:rgba(122,252,208,.12);color:var(--mint);width:fit-content}
.tag.warm{background:rgba(255,210,122,.14);color:var(--gold)}
.tag.cool{background:rgba(120,180,200,.14);color:#9cd0dd}
/* MATRIX */
.matrix-wrap{overflow-x:auto;padding-bottom:8px}
.matrix{display:inline-grid;gap:2px;margin-top:8px}
.mcell{width:19px;height:19px;border-radius:3px;background:var(--surface2);cursor:default}
.mlabel{font-size:.62rem;color:var(--muted);display:flex;align-items:center}
.mlabel.row{justify-content:flex-end;padding-right:6px;white-space:nowrap}
.mlabel.col{writing-mode:vertical-rl;transform:rotate(180deg);justify-content:flex-end;padding-bottom:6px;height:64px}
.mdiag{background:repeating-linear-gradient(45deg,#063a44,#063a44 3px,#0a4a56 3px,#0a4a56 6px)}
#mtip{position:fixed;pointer-events:none;z-index:90;background:#001a1f;border:1px solid var(--line);
  border-radius:10px;padding:8px 12px;font-size:.82rem;opacity:0;transition:opacity .12s;box-shadow:0 8px 30px rgba(0,0,0,.5)}
.legend{display:flex;align-items:center;gap:10px;margin-top:14px;font-size:.78rem;color:var(--muted)}
.legend .scale{height:10px;width:140px;border-radius:6px;background:linear-gradient(90deg,var(--surface2),var(--mint))}
/* match distribution */
.match-row{padding:14px 0;border-bottom:1px solid var(--line)}
.match-row:last-child{border:0}
.match-top{display:flex;justify-content:space-between;gap:10px;font-weight:600;margin-bottom:8px;font-size:.96rem}
.dist{display:flex;height:24px;border-radius:7px;overflow:hidden;font-size:.72rem;font-weight:700}
.dist span{display:flex;align-items:center;justify-content:center;color:#012;min-width:0;transition:width 1s cubic-bezier(.2,.8,.2,1)}
.dist .s1{background:var(--mint)}.dist .sx{background:#7fa8b2}.dist .s2{background:var(--gold)}
.modal-pill{color:var(--muted);font-size:.84rem;margin-top:6px}
/* chips teams */
.teamchip{display:inline-flex;align-items:center;gap:7px;background:var(--surface);border:1px solid var(--line);
  border-radius:999px;padding:7px 13px;font-size:.9rem;font-weight:600;margin:4px 4px 0 0}
.teamchip .n{color:var(--muted);font-family:'Space Grotesk';font-weight:700}
/* group cards */
.gcard{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:16px}
.gcard .gl{font-family:'Space Grotesk';color:var(--muted);font-size:.8rem;letter-spacing:.1em}
.gcard .fav{font-size:1.15rem;font-weight:700;margin:4px 0 10px;display:flex;align-items:center;gap:8px}
/* fichas */
.search{width:100%;max-width:340px;background:var(--surface);border:1px solid var(--line);color:var(--text);
  border-radius:12px;padding:11px 15px;font-size:.95rem;margin-bottom:20px;font-family:inherit}
.search:focus{outline:none;border-color:var(--mint)}
.ficha{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:18px;transition:.2s}
.ficha:hover{border-color:rgba(122,252,208,.4);transform:translateY(-3px)}
.ficha .fh{display:flex;justify-content:space-between;align-items:start;margin-bottom:10px}
.ficha .fn{font-size:1.2rem;font-weight:700}
.ficha .rk{font-family:'Space Grotesk';font-size:.8rem;color:var(--muted)}
.ficha .lab{font-size:.8rem;margin:2px 0 12px}
.fstats{display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:.84rem}
.fstats .v{font-family:'Space Grotesk';font-weight:700;color:var(--mint)}
.fline{font-size:.82rem;color:var(--muted);margin-top:10px;border-top:1px solid var(--line);padding-top:10px}
.fline b{color:var(--text)}
/* premios */
.award{background:linear-gradient(160deg,var(--surface2),var(--surface));border:1px solid var(--line);
  border-radius:18px;padding:22px;text-align:center}
.award .em{font-size:2.4rem}.award .ti{font-weight:700;margin:8px 0 2px}
.award .wn{font-family:'Space Grotesk';font-size:1.5rem;font-weight:700;color:var(--mint)}
.award .dt{font-size:.82rem;color:var(--muted);margin-top:2px}
/* teaser */
.teaser{text-align:center;padding:46px 22px}
.teaser .em{font-size:3rem}
/* reveal */
.reveal{opacity:0;transform:translateY(20px);transition:opacity .65s ease,transform .65s ease}
.reveal.in{opacity:1;transform:none}
footer{border-top:1px solid var(--line);padding:36px 0 60px;color:var(--muted);font-size:.82rem}
footer .brand{color:var(--mint);height:20px;display:inline-block;vertical-align:middle;margin-right:8px}
footer .brand svg{height:20px;width:auto}
@media(max-width:900px){
  nav.rail{flex-direction:row;width:100%;height:auto;top:auto;bottom:0;padding:8px;overflow-x:auto;gap:2px;
    border-right:0;border-top:1px solid var(--line)}
  nav.rail .brand{display:none}
  nav.rail a{font-size:.7rem;padding:7px 9px}nav.rail a .dot{display:none}
  .langtoggle{margin-bottom:0;flex:0 0 auto;order:-1}
  .wrap{padding:0 18px 70px}
  .g3,.g4{grid-template-columns:repeat(2,1fr)}.duo,.podium{grid-template-columns:1fr}
  .chips{grid-template-columns:repeat(2,1fr)}
}
@media(max-width:560px){.g2,.g3,.g4{grid-template-columns:1fr}}
"""

JS = r"""
const D = window.__PORRA__;
const N = D.hero.participants;
const ES2EN = D.es2en || {};
const logo = `__LOGO__`;
let LANG = 'es';
let wrap = null;

function L(es, en){ return LANG === 'es' ? es : en; }
function team(es){ return LANG === 'es' ? es : (ES2EN[es] || es); }
function esc(s){ return String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }
function fmt(v){
  let [a,b] = String(v).split('.');
  const th = LANG==='es' ? '.' : ',', dp = LANG==='es' ? ',' : '.';
  a = a.replace(/\B(?=(\d{3})+(?!\d))/g, th);
  return b !== undefined ? a + dp + b : a;
}
function pf(x){ let s = String(x); if (LANG==='es') s = s.replace('.', ','); return s + '%'; }
function animateCount(node){
  const to = parseFloat(node.dataset.count), dec = parseInt(node.dataset.decimals||'0'), suf = node.dataset.suffix||'';
  const dur = 900, t0 = performance.now();
  (function step(t){ let p = Math.min(1,(t-t0)/dur); p = 1-Math.pow(1-p,3);
    node.textContent = fmt((to*p).toFixed(dec)) + suf; if (p<1) requestAnimationFrame(step); })(t0);
}
function el(t,c,h){ const e = document.createElement(t); if(c) e.className=c; if(h!=null) e.innerHTML=h; return e; }

const NAV = [
  ['inicio','Inicio','Home'],['rebeldia','Rebeldía','Maverick'],['gemelas','Afinidad','Affinity'],
  ['estilo','Estilo','Style'],['favoritos','Favoritos','Favourites'],['partidos','Partidos','Matches'],
  ['lobo','Lobo solitario','Lone wolf'],['fichas','Fichas','Profiles'],['premios','Palmarés','Awards'],
  ['aciertos','En directo','Live'],
];
const LABELS = {
  goleador:['🥅 Goleador','🥅 Goal machine'], cerrojo:['🔒 Cerrojo','🔒 Parked bus'],
  empate:['🤝 Amigo del empate','🤝 Draw lover'], rebelde:['🐺 Rebelde','🐺 Maverick'],
  borrego:['🐑 Borrego','🐑 Sheep'], equilibrado:['⚖️ Equilibrado','⚖️ Balanced'],
};
function labelOf(k){ const x = LABELS[k] || LABELS.equilibrado; return L(x[0], x[1]); }

/* ---- NAV (built once, labels re-rendered on language change) ---- */
const rail = el('nav','rail');
document.body.appendChild(rail);
function renderRail(){
  rail.innerHTML = `<a class="brand" href="#inicio">${logo}</a>` +
    `<div class="langtoggle"><button data-l="es">ES</button><button data-l="en">EN</button></div>` +
    NAV.map(([id,es,en]) => `<a href="#${id}" data-id="${id}"><span class="dot"></span>${L(es,en)}</a>`).join('');
  const on = rail.querySelector('[data-l="'+LANG+'"]'); if (on) on.classList.add('on');
}
rail.addEventListener('click', e => {
  const b = e.target.closest('button[data-l]');
  if (b && b.dataset.l !== LANG){ LANG = b.dataset.l; rebuild(); }
});

/* shared tooltip for the matrix */
const tip = el('div'); tip.id = 'mtip'; document.body.appendChild(tip);

function section(id, kicker, title, sub){
  const s = el('section','sec'); s.id = id;
  s.appendChild(el('div','sec-head reveal',
    `<div class="kicker">${kicker}</div><h2>${title}</h2>${sub?`<p class="sub">${sub}</p>`:''}`));
  wrap.appendChild(s); return s;
}

/* ---- HERO ---- */
function chip(val, lab, dec){
  return `<div class="chip reveal"><div class="big" data-count="${val}"${dec?` data-decimals="${dec}"`:''}>0</div><div class="lab">${lab}</div></div>`;
}
function chipText(val, lab){
  return `<div class="chip reveal"><div class="big">${val}</div><div class="lab">${lab}</div></div>`;
}
function buildHero(){
  const h = el('header','hero'); h.id = 'inicio';
  const ts = D.hero.top_scoreline, bp = D.best_pair;
  h.innerHTML = `
    <div class="eyebrow reveal">${L('Reveni · Porra Mundial 2026','Reveni · World Cup 2026 Pool')}</div>
    <h1 class="reveal">${L('La quiniela,<br>bajo el microscopio.','The pool,<br>under the microscope.')}</h1>
    <p class="lead reveal">${L(
      'Hemos analizado las '+N+' quinielas de la porra del Mundial: quién se la juega, quién va a lo seguro, quién piensa igual que quién y qué partidos nos parten en dos.',
      'We crunched all '+N+' predictions in the office World Cup pool: who gambles, who plays it safe, who thinks like who, and which matches split us in two.')}</p>
    <div class="chips">
      ${chip(N, L('Participantes','Players'))}
      ${chip(D.hero.matches, L('Partidos de grupos','Group matches'))}
      ${chip(D.hero.groups, L('Grupos','Groups'))}
      ${chip(N*D.hero.matches, L('Marcadores analizados','Predictions analysed'))}
      ${chip(D.hero.total_goals, L('Goles apostados','Goals predicted'))}
      ${chip(D.hero.avg_goals_match, L('Goles/partido de media','Avg goals per match'), 1)}
      ${chipText(ts?ts.score:'–', L('Marcador estrella','Star scoreline')+' ('+pf(ts?ts.pct:0)+')')}
      ${chipText(pf(bp.pct), L('Máx. afinidad','Top affinity')+' · '+esc(bp.a)+' & '+esc(bp.b))}
    </div>
    <div class="scrollcue">${L('desliza para empezar','scroll to start')} ↓</div>`;
  wrap.appendChild(h);
}

/* ---- REBELDÍA ---- */
function buildRebeldia(){
  const s = section('rebeldia', L('01 · La estrella','01 · The star'),
    L('El ranking de rebeldía 🐺','The maverick ranking 🐺'),
    L('Combina dos cosas en cada partido: cuántas veces eliges el signo (1/X/2) en minoría, y cuánto se aleja tu marcador del marcador típico del grupo. 100 = el más inconformista. 0 = el más borrego.',
      'Two things per match: how often you back the (1/X/2) outcome in the minority, and how far your scoreline sits from the group typical one. 100 = the biggest maverick. 0 = the biggest sheep.'));
  const R = D.rebeldia, max = R[0].index || 100;
  const pod = el('div','podium reveal');
  const order = [1,0,2], medals = ['🥈','🥇','🥉'];
  order.forEach((oi,k) => { const r = R[oi]; if(!r) return;
    pod.appendChild(el('div','pod '+(oi===0?'p1':''),
      `<div class="medal">${medals[k]}</div><div class="nm">${esc(r.name)}</div>
       <div class="sc" data-count="${r.index}" data-decimals="1">0</div><div class="k">${L('índice','index')}</div>`)); });
  s.appendChild(pod);
  const list = el('div','reveal');
  R.forEach((r,i) => {
    const badge = i===0?'🐺 ':(i===R.length-1?'🐑 ':'');
    const cls = i===R.length-1?'gold':'';
    list.appendChild(el('div','bar-row',
      `<div class="bar-rank">${i+1}</div><div class="bar-name">${badge}${esc(r.name)}</div>
       <div class="bar-track"><div class="bar-fill ${cls}" data-w="${(r.index/max*100).toFixed(1)}"></div></div>
       <div class="bar-val" data-count="${r.index}" data-decimals="1">0</div>`));
  });
  s.appendChild(list);
}

/* ---- AFINIDAD ---- */
function buildAfinidad(){
  const s = section('gemelas', L('02 · Afinidad','02 · Affinity'),
    L('¿Quién piensa como quién? 💞','Who thinks like who? 💞'),
    L('Comparamos las quinielas marcador a marcador. Cuanto más brillante el cuadro, más coinciden esas dos personas en el resultado exacto.',
      'We compare predictions scoreline by scoreline. The brighter the cell, the more those two people match on the exact result.'));
  const bp = D.best_pair, wp = D.worst_pair;
  const duo = el('div','duo reveal');
  duo.innerHTML = `
    <div class="card glow"><span class="tag">${L('💞 Almas gemelas','💞 Soulmates')}</span>
      <div class="names">${esc(bp.a)} & ${esc(bp.b)}</div>
      <div class="muted">${bp.pct}% ${L('de marcadores','of scorelines')} <b class="mint">${L('idénticos','identical')}</b>${bp.pct>=100?L(' — sí, los 72. 👀 ¿alguien ha copiado?',' — yes, all 72. 👀 did someone copy?'):''}</div></div>
    <div class="card"><span class="tag cool">${L('🧊 Polos opuestos','🧊 Opposites')}</span>
      <div class="names">${esc(wp.a)} & ${esc(wp.b)}</div>
      <div class="muted">${L('los que más se contradicen (distancia media '+wp.dist+' goles por partido)','the most contradictory pair (avg distance '+wp.dist+' goals per match)')}</div></div>`;
  s.appendChild(duo);
  const names = D.matrix.names, sim = D.matrix.sim;
  let mx = 0; for(let i=0;i<N;i++) for(let j=0;j<N;j++) if(i!==j) mx = Math.max(mx, sim[i][j]);
  const wrapm = el('div','matrix-wrap reveal');
  const grid = el('div','matrix');
  grid.style.gridTemplateColumns = `120px repeat(${N},19px)`;
  grid.appendChild(el('div','mlabel'));
  names.forEach(n => grid.appendChild(el('div','mlabel col', esc(n))));
  for(let i=0;i<N;i++){
    grid.appendChild(el('div','mlabel row', esc(names[i])));
    for(let j=0;j<N;j++){
      if(i===j){ grid.appendChild(el('div','mcell mdiag')); continue; }
      const v = sim[i][j], a = mx ? Math.max(.05, v/mx) : 0;
      const c = el('div','mcell');
      c.style.background = `rgba(122,252,208,${a.toFixed(3)})`;
      c.dataset.t = `${esc(names[i])} ↔ ${esc(names[j])} · ${v}% ${L('idénticos','identical')}`;
      grid.appendChild(c);
    }
  }
  wrapm.appendChild(grid);
  wrapm.appendChild(el('div','legend',`<span>${L('menos afín','less alike')}</span><span class="scale"></span><span>${L('más afín','more alike')}</span>`));
  s.appendChild(wrapm);
  grid.addEventListener('mouseover', e => { const t = e.target.dataset.t; if(t){ tip.textContent = t; tip.style.opacity = 1; } });
  grid.addEventListener('mousemove', e => { tip.style.left = (e.clientX+14)+'px'; tip.style.top = (e.clientY+14)+'px'; });
  grid.addEventListener('mouseout', () => { tip.style.opacity = 0; });
}

/* ---- ESTILO ---- */
function buildEstilo(){
  const s = section('estilo', L('03 · Estilo','03 · Style'),
    L('Goleadores vs cerrojos 🥅','Goal-fests vs parked buses 🥅'),
    L('Media de goles que cada uno mete en sus marcadores, qué marcadores se repiten más en toda la porra y quién es el rey del empate.',
      'Average goals in everyone scorelines, the most repeated results across the pool, and who loves a draw.'));
  const st = D.style.slice().sort((a,b) => b.avg_goals-a.avg_goals);
  const gol = st[0], cer = st[st.length-1];
  const emp = D.style.slice().sort((a,b) => b.pct_draws-a.pct_draws)[0];
  const duo = el('div','grid g3 reveal');
  duo.innerHTML = `
    <div class="card glow"><span class="k">${L('🥅 El goleador','🥅 The goal machine')}</span><h3>${esc(gol.name)}</h3>
      <div class="big-num">${gol.avg_goals}</div><div class="muted">${L('goles de media por partido','avg goals per match')}</div></div>
    <div class="card"><span class="k">${L('🔒 El cerrojo','🔒 The parked bus')}</span><h3>${esc(cer.name)}</h3>
      <div class="big-num">${cer.avg_goals}</div><div class="muted">${L('el más tacaño con los goles','the stingiest with goals')}</div></div>
    <div class="card"><span class="k">${L('🤝 Amigo del empate','🤝 Draw lover')}</span><h3>${esc(emp.name)}</h3>
      <div class="big-num">${pf(emp.pct_draws)}</div><div class="muted">${L('de sus partidos, en tablas','of their matches drawn')}</div></div>`;
  s.appendChild(duo);
  const list = el('div','reveal'); list.style.marginTop = '22px';
  const mxg = st[0].avg_goals;
  list.appendChild(el('div','k', L('Media de goles por quiniela','Average goals per prediction')));
  st.forEach((r,i) => list.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${esc(r.name)}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${(r.avg_goals/mxg*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${r.avg_goals}" data-decimals="2">0</div>`)));
  s.appendChild(list);
  const cs = el('div','card reveal'); cs.style.marginTop = '22px';
  cs.innerHTML = `<span class="k">${L('Marcadores más apostados (de las '+fmt(N*D.hero.matches)+' apuestas)','Most-predicted scorelines (out of '+fmt(N*D.hero.matches)+' bets)')}</span>`;
  const mxc = D.common_scorelines[0].count;
  D.common_scorelines.forEach(c => cs.appendChild(el('div','bar-row',
    `<div class="bar-rank">${c.score}</div><div class="bar-name muted">${pf(c.pct)}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${(c.count/mxc*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${c.count}">0</div>`)));
  s.appendChild(cs);
}

/* ---- FAVORITOS ---- */
function buildFavoritos(){
  const s = section('favoritos', L('04 · Favoritos','04 · Favourites'),
    L('Los favoritos del torneo 🏆','Tournament favourites 🏆'),
    L('El consenso de la oficina: a quién ve todo el mundo ganando su grupo, qué selecciones se respetan más y qué grupo nos tiene más divididos.',
      'The office consensus: who everyone sees topping their group, the most respected teams, and the group that divides us most.'));
  const gc = D.group_consensus.slice().sort((a,b) => b.agreement-a.agreement);
  const grid = el('div','grid g4 reveal');
  gc.forEach(g => grid.appendChild(el('div','gcard',
    `<div class="gl">${L('GRUPO','GROUP')} ${g.group}</div>
     <div class="fav">${g.flag} ${esc(team(g.favorite))}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${g.agreement}"></div></div>
     <div class="muted" style="font-size:.8rem;margin-top:6px">${pf(g.agreement)} ${L('lo ve primero','see them top')}</div>`)));
  s.appendChild(grid);
  const mr = el('div','card reveal'); mr.style.marginTop = '22px';
  mr.innerHTML = `<span class="k">${L('Selecciones más respetadas (prestigio: 1.º=3, 2.º=2, 3.º=1)','Most respected teams (prestige: 1st=3, 2nd=2, 3rd=1)')}</span>`;
  const mxr = D.most_respected[0].count;
  D.most_respected.forEach((t,i) => mr.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${t.flag} ${esc(team(t.team))}</div>
     <div class="bar-track"><div class="bar-fill gold" data-w="${(t.count/mxr*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${t.count}">0</div>`)));
  s.appendChild(mr);
  const two = el('div','grid g2 reveal'); two.style.marginTop = '22px';
  const pol = D.polarizing.slice(0,4).map(p => {
    const top = Object.entries(p.dist).slice(0,3).map(([t,c]) => `${esc(team(t))} (${c})`).join(' · ');
    return `<div style="padding:8px 0;border-bottom:1px solid var(--line)"><b>${L('Grupo','Group')} ${p.group}</b> — ${pf(p.agreement)} ${L('de acuerdo','agreement')}<br><span class="muted" style="font-size:.84rem">${top}</span></div>`;
  }).join('');
  const th = D.top_thirds.slice(0,10).map(t => `<span class="teamchip">${t.flag} ${esc(team(t.team))} <span class="n">${t.count}</span></span>`).join('');
  two.innerHTML = `
    <div class="card"><span class="k">${L('🔥 Los grupos que más nos dividen','🔥 The groups that divide us most')}</span><div style="margin-top:10px">${pol}</div></div>
    <div class="card"><span class="k">${L('🎟️ Terceros más elegidos para colarse','🎟️ Most-picked third places')}</span><div style="margin-top:10px">${th}</div></div>`;
  s.appendChild(two);
}

/* ---- PARTIDOS ---- */
function buildPartidos(){
  const s = section('partidos', L('05 · Partidos','05 · Matches'),
    L('Los partidos que nos parten en dos 🔪','The matches that split us 🔪'),
    L('Dónde hay guerra de pronósticos (poco acuerdo en el 1/X/2) y dónde casi todos ponemos lo mismo.',
      'Where there is a prediction war (little agreement on 1/X/2) and where we nearly all agree.'));
  function matchBlock(m){
    const o = m.outcome_dist, tot = (o['1']||0)+(o['X']||0)+(o['2']||0);
    const p = k => tot ? Math.round((o[k]||0)/tot*100) : 0;
    return `<div class="match-row">
      <div class="match-top"><span>${m.home_flag} ${esc(team(m.home))} – ${esc(team(m.away))} ${m.away_flag}</span>
        <span class="muted">${pf(Math.round(m.outcome_agreement*100))} ${L('acuerdo','agree')}</span></div>
      <div class="dist">
        <span class="s1" style="width:${p('1')}%">${p('1')?p('1')+'%':''}</span>
        <span class="sx" style="width:${p('X')}%">${p('X')?'X '+p('X')+'%':''}</span>
        <span class="s2" style="width:${p('2')}%">${p('2')?p('2')+'%':''}</span></div>
      <div class="modal-pill">${L('marcador más repetido','most common scoreline')}: <b class="mint">${m.modal_scoreline}</b> (${pf(Math.round(m.modal_scoreline_share*100))})</div>
    </div>`;
  }
  const two = el('div','grid g2 reveal');
  two.innerHTML = `
    <div class="card"><span class="k">${L('😵‍💫 Máxima guerra de pronósticos','😵‍💫 Maximum disagreement')}</span><div style="margin-top:8px">${D.divisive.map(matchBlock).join('')}</div></div>
    <div class="card"><span class="k">${L('🤖 Casi unanimidad','🤖 Near unanimity')}</span><div style="margin-top:8px">${D.unanimous.map(matchBlock).join('')}</div></div>`;
  s.appendChild(two);
}

/* ---- LOBO SOLITARIO ---- */
function buildLobo(){
  const s = section('lobo', L('06 · Atrevimiento','06 · Boldness'),
    L('Lobo solitario 🐺','Lone wolf 🐺'),
    L('Marcadores que solo apostó UNA persona en todo el grupo. El que más tiene, es el que más se aleja del rebaño.',
      'Scorelines only ONE person predicted in the whole pool. The more you have, the further from the flock you roam.'));
  const mx = D.lobo[0].count;
  const list = el('div','reveal');
  D.lobo.forEach((r,i) => list.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${i===0?'🐺 ':''}${esc(r.name)}</div>
     <div class="bar-track"><div class="bar-fill" data-w="${(r.count/mx*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${r.count}">0</div>`)));
  s.appendChild(list);
  const ex = el('div','card reveal'); ex.style.marginTop = '22px';
  ex.innerHTML = `<span class="k">${L('Apuestas en solitario más salvajes','Wildest solo calls')}</span><div style="margin-top:10px">` +
    D.unique_examples.map(e => `<div style="padding:7px 0;border-bottom:1px solid var(--line)">
      <b>${esc(e.name)}</b> <span class="muted">${L('es el único que firmó','is the only one who called')}</span> ${e.flags} <b class="mint">${esc(team(e.home))} ${e.score} ${esc(team(e.away))}</b></div>`).join('') + '</div>';
  s.appendChild(ex);
}

/* ---- FICHAS ---- */
function buildFichas(){
  const s = section('fichas', L('07 · Uno a uno','07 · One by one'),
    L('La ficha de cada uno 🪪',"Everyone's card 🪪"),
    L('Resumen por persona: su puesto en rebeldía, su estilo, su gemelo y su apuesta más loca. Busca tu nombre.',
      'A summary per person: their maverick rank, style, twin and wildest bet. Search your name.'));
  const inp = el('input','search'); inp.placeholder = L('🔎 Busca tu nombre…','🔎 Search your name…');
  s.appendChild(inp);
  const grid = el('div','grid g3 reveal');
  D.cards.slice().sort((a,b) => a.rebel_rank-b.rebel_rank).forEach(c => {
    const big = c.biggest ? `${c.biggest.flags} ${esc(team(c.biggest.home))} ${c.biggest.score} ${esc(team(c.biggest.away))}` : '–';
    const card = el('div','ficha'); card.dataset.name = c.name.toLowerCase();
    card.innerHTML = `
      <div class="fh"><div><div class="fn">${esc(c.name)}</div><div class="lab">${labelOf(c.label)}</div></div>
        <div class="rk">#${c.rebel_rank}<br>${L('rebeldía','maverick')}</div></div>
      <div class="fstats">
        <div>${L('Índice rebeldía','Maverick index')}<br><span class="v">${c.rebel_index}</span></div>
        <div>${L('Goles/partido','Goals/match')}<br><span class="v">${c.avg_goals}</span></div>
        <div>${L('% empates','% draws')}<br><span class="v">${pf(c.pct_draws)}</span></div>
        <div>${L('Apuestas únicas','Unique bets')}<br><span class="v">${c.lobo}</span></div>
      </div>
      <div class="fline">${L('👯 Gemelo:','👯 Twin:')} <b>${esc(c.twin)}</b> (${pf(c.twin_pct)})</div>
      <div class="fline">${L('🎲 Su locura:','🎲 Wildest bet:')} <b>${big}</b></div>`;
    grid.appendChild(card);
  });
  s.appendChild(grid);
  inp.addEventListener('input', () => { const q = inp.value.toLowerCase().trim();
    grid.querySelectorAll('.ficha').forEach(f => { f.style.display = f.dataset.name.includes(q) ? '' : 'none'; }); });
}

/* ---- PREMIOS ---- */
function buildPremios(){
  const s = section('premios', L('08 · Palmarés','08 · Awards'),
    L('El palmarés de la porra 🏅',"The pool's hall of fame 🏅"),
    L('Los títulos honoríficos de esta edición.',"This edition's honorary titles."));
  const a = D.awards;
  const items = [
    ['🐺', L('El Rebelde','The Maverick'), a.rebelde.name, L('índice '+a.rebelde.index+'/100','index '+a.rebelde.index+'/100')],
    ['🐑', L('El Borrego','The Sheep'), a.borrego.name, L('el más previsible ('+a.borrego.index+'/100)','most predictable ('+a.borrego.index+'/100)')],
    ['💞', L('Almas Gemelas','Soulmates'), a.gemelas.a+' & '+a.gemelas.b, L(a.gemelas.pct+'% idénticos',a.gemelas.pct+'% identical')],
    ['🧊', L('Polos Opuestos','Opposites'), a.opuestos.a+' & '+a.opuestos.b, L('distancia '+a.opuestos.dist,'distance '+a.opuestos.dist)],
    ['🥅', L('El Goleador','The Goal Machine'), a.goleador.name, L(a.goleador.avg+' goles/partido',a.goleador.avg+' goals/match')],
    ['🔒', L('El Cerrojo','The Parked Bus'), a.cerrojo.name, L(a.cerrojo.avg+' goles/partido',a.cerrojo.avg+' goals/match')],
    ['🤝', L('Rey del Empate','Draw King'), a.empate.name, L(a.empate.pct+'% en tablas',a.empate.pct+'% drawn')],
    ['🎯', L('Lobo Solitario','Lone Wolf'), a.lobo.name, L(a.lobo.count+' apuestas únicas',a.lobo.count+' unique bets')],
  ];
  const grid = el('div','grid g4 reveal');
  items.forEach(([em,ti,wn,dt]) => grid.appendChild(el('div','award',
    `<div class="em">${em}</div><div class="ti">${ti}</div><div class="wn">${esc(wn)}</div><div class="dt">${dt}</div>`)));
  s.appendChild(grid);
}

/* ---- ACIERTOS ---- */
function buildAciertos(){
  const s = section('aciertos', L('09 · En directo','09 · Live'),
    L('Aciertos en directo ⚽','Live scoring ⚽'),
    L('Cuando empiece a rodar el balón y rellenes los resultados reales en el Excel, aquí saldrá el ranking de aciertos.',
      'Once the ball starts rolling and you fill in the real results in the Excel, the scoring ranking appears here.'));
  if(!D.live){
    s.appendChild(el('div','card teaser reveal',
      `<div class="em">⚽️</div><h3 style="margin:10px 0 6px">${L('El Mundial aún no ha empezado',"The World Cup hasn't kicked off yet")}</h3>
       <p class="muted">${L('Rellena los marcadores reales en la pestaña <b>Real results</b> del Excel y vuelve a generar el dashboard: aparecerá aquí el ranking de aciertos (signo +2, marcador exacto +4).','Fill in the real scores in the <b>Real results</b> tab of the Excel and regenerate the dashboard: the scoring ranking shows up here (correct outcome +2, exact score +4).')}</p>`));
    return;
  }
  const t = D.live.table, mx = t[0].pts || 1;
  s.appendChild(el('div','muted reveal', `${D.live.played} ${L('partidos jugados','matches played')}`));
  const list = el('div','reveal');
  t.forEach((r,i) => list.appendChild(el('div','bar-row',
    `<div class="bar-rank">${i+1}</div><div class="bar-name">${i===0?'👑 ':''}${esc(r.name)} <span class="muted" style="font-size:.78rem">(${r.exact} ${L('plenos','exact')} · ${r.sign} ${L('signos','outcomes')})</span></div>
     <div class="bar-track"><div class="bar-fill gold" data-w="${(r.pts/mx*100).toFixed(1)}"></div></div>
     <div class="bar-val" data-count="${r.pts}">0</div>`)));
  s.appendChild(list);
}

/* ---- FOOTER ---- */
function buildFooter(){
  const f = el('footer');
  f.innerHTML = `<span class="brand">${logo}</span> ${L('Porra Mundial 2026 · análisis de '+N+' quinielas · generado el __DATE__.','World Cup 2026 Pool · analysis of '+N+' predictions · generated on __DATE__.')}
    <br><span style="opacity:.6">${L('Hecho con cariño para la oficina. Tipografía corporativa real: Garnett (de pago); aquí usamos Space Grotesk + Inter como sustitutas libres.','Made with love for the office. Real brand typeface: Garnett (paid); here we use Space Grotesk + Inter as free stand-ins.')}</span>`;
  wrap.appendChild(f);
}

/* ---- OBSERVERS ---- */
function observeAll(){
  const io = new IntersectionObserver((es,o) => { es.forEach(e => { if(e.isIntersecting){
    e.target.classList.add('in');
    e.target.querySelectorAll && e.target.querySelectorAll('[data-count]').forEach(animateCount);
    e.target.querySelectorAll && e.target.querySelectorAll('.bar-fill').forEach(b => { b.style.width = (b.dataset.w||0)+'%'; });
    o.unobserve(e.target);
  }}); }, {threshold:.12});
  wrap.querySelectorAll('.reveal').forEach(n => io.observe(n));
  const navlinks = {}; rail.querySelectorAll('a[data-id]').forEach(a => navlinks[a.dataset.id] = a);
  const navio = new IntersectionObserver(es => { es.forEach(e => { if(e.isIntersecting){
    Object.values(navlinks).forEach(a => a.classList.remove('active'));
    if(navlinks[e.target.id]) navlinks[e.target.id].classList.add('active');
  }}); }, {rootMargin:'-45% 0px -50% 0px'});
  wrap.querySelectorAll('header.hero,section.sec').forEach(sec => navio.observe(sec));
}

/* ---- BUILD / REBUILD ---- */
function rebuild(){
  if(wrap) wrap.remove();
  wrap = el('div','wrap'); document.body.appendChild(wrap);
  renderRail();
  buildHero(); buildRebeldia(); buildAfinidad(); buildEstilo(); buildFavoritos();
  buildPartidos(); buildLobo(); buildFichas(); buildPremios(); buildAciertos(); buildFooter();
  observeAll();
}
rebuild();
"""

HTML_TEMPLATE = (
    '<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
    '<title>Porra Mundial 2026 · Reveni</title>\n'
    '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n'
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&'
    'family=Space+Grotesk:wght@500;700&display=swap" rel="stylesheet">\n'
    '<style>' + CSS + '</style>\n</head>\n<body>\n'
    '<div id="app"></div>\n'
    '<script>window.__PORRA__=__DATA__;</script>\n'
    '<script>' + JS + '</script>\n'
    '</body>\n</html>\n'
)


def load_logo():
    p = Path(LOGO_SVG_PATH)
    if not p.exists():
        return '<span class="wordmark">reveni</span>'
    svg = p.read_text(encoding="utf-8")
    return svg.replace('fill="#003C46"', 'fill="currentColor"')


def render_html(analytics, generated_on):
    data_json = json.dumps(analytics, ensure_ascii=False)
    logo = load_logo()
    return (
        HTML_TEMPLATE
        .replace("__DATA__", data_json)
        .replace("__LOGO__", logo)
        .replace("__DATE__", generated_on)
    )


if __name__ == "__main__":
    from datetime import date

    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    out = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT
    data = parse_workbook(xlsx)
    analytics = compute(data)
    html = render_html(analytics, date.today().strftime("%d/%m/%Y"))
    Path(out).write_text(html, encoding="utf-8")
    print(f"OK · {data['n']} participantes · {len(data['matches'])} partidos · "
          f"resultados cargados: {len(data['results'])}")
    print(f"Dashboard escrito en: {out}  ({len(html) // 1024} KB)")
