# Análisis de la fase de eliminatorias

Esta guía describe **qué datos hay**, **qué calcula el dashboard hoy** y **qué
queda por hacer** en la fase eliminatoria.

Relacionado: `PROTOTYPE_ko_metrics.md` (variantes `?ko=B|C` archivadas) y
`README.md` (flujo general del dashboard).

---

## 1. Dónde viven los datos

Todo sale del Excel **`Porra_Admin_v5_EN.xlsx`**. Cada participante ocupa **dos
columnas** (goles local / goles visitante o valor) en la hoja **`Raw data`**, con
su nombre en la fila 6.

| Bloque | Filas (hoja `Raw data`) | Estado (28-jun) |
|---|---|---|
| Marcadores de grupos | 7–78 | ✅ completo |
| Clasificación de grupos (1.º/2.º/3.º) | 80–115 | ✅ completo |
| 8 mejores terceros | 117–124 | ✅ completo |
| **R32** (16 cruces × score/desempate/ganador) | 126–173 | ⚠️ ~61 % celdas |
| **R16** (8 cruces) | 175–198 | ⚠️ ~46 % |
| **Cuartos** (4 cruces) | 200–211 | ⚠️ ~46 % |
| **Semis** (2 cruces) | 213–218 | ⚠️ ~49 % |
| 3.er puesto + Final | 220–224 | ⚠️ ~43 % |
| **Campeón / Subcampeón** | 225 / 226 | ⚠️ ~47 % |
| **Pichichi / Balón de Oro** | 228 / 229 | ⚠️ ~47 % |

Los **resultados reales** van en **`Real results`** (misma disposición de filas;
columnas C/D para marcador, C para ganador / desempate / campeón / premios).

### Cobertura actual (28 participantes)

| Métrica | Valor |
|---|---|
| Celdas KO rellenas | **74,6 %** (1400/1876) |
| Cuadro completo (R32→SF + campeón) | **18/28** |
| Sin campeón | Tere y Edu, Nadia, Juanorro, Meg, Andy, Ben, Mile, Cami, Jaime, Emilio |
| Resultados KO jugados | **1** — R32-M3: Sudáfrica 0-1 Canadá |
| Consenso campeón | España 8 · Francia 5 · Argentina 2… |

> Tras cada partido KO: rellenar `Real results` (p. ej. filas 132–134 para
> R32-M3) → `python3 generate_dashboard.py` → commit/push si toca publicar.

---

## 2. Qué calcula ya el Python

En `generate_dashboard.py`:

| Función | Qué hace | Estado |
|---|---|---|
| `parse_knockouts` | Lee porra KO por persona | ✅ |
| `parse_knockout_results` | Lee resultados reales | ✅ |
| `compute_knockout` | Consenso, cobertura, calendario, picks, stake | ✅ |
| `compute_knockout_scoring` | Ranking de puntos KO (marcador + pases + premios) | ✅ |
| `compute_knockout_metrics` | Métricas del relato (`D.knockout.metrics`) | ✅ Fase 1 |
| **`compute_ko_match_stake`** | Puntos y swing por persona en cruces sin jugar | ✅ **Fase 2** |
| **`compute_ko_progression`** | Ranking KO partido a partido (`D.knockout.progression`) | ✅ **Fase 2** |
| **`_ko_champion_fell_round`** | Supervivencia del campeón predicho | ✅ **Fase 2** |

### Vista producción (`?view=ko`)

Una sola vista: el **relato** con datos reales (actos 1–9) + calendario completo al final.
No hay variantes `?ko=A|B|C` ni barra flotante de prototipo.

---

## 3. Fase 1 — Métricas reales en el relato

Las métricas salen de **`D.knockout.metrics`**. El JS ya no genera datos demo.

| Campo | Origen |
|---|---|
| `champRank` | Distribución de `outright.champion` |
| `tsRank` | Distribución de `awards.top_scorer` |
| `twins` | Similitud de `winner_picks` + campeón/sub/pichichi |
| `chaosRank` / `people[].chaos` | Votos al peor prestigio del cruce |
| `grave` | Campeones cuyo equipo ya no puede ganar |
| `depth` | % de cuadros que llevan a cada selección a R16/QF/SF/Final/Campeón |
| `people[]` | Ficha: campeón, sub, pichichi, caos, riesgo, puntos KO |
| `pool` | Selecciones del dossier interactivo |

**Acto 7 “El profeta”** usa el ranking KO real (`scoring.table`) cuando hay resultados.

---

## 4. Fase 2 — Implementada

### 4.1 Supervivencia del campeón (`koSurvivalCard`)

- **`people[].fell`**: ronda en la que cae tu **campeón predicho** (no el primer pase fallido del cuadro).
- Rondas: Dieciseisavos → Octavos → Cuartos → Semis → Final → Campeón.
- Fallar un cruce intermedio **no** te elimina si tu campeón sigue en el torneo.
- Con R32-M3: quien tiene **Sudáfrica** como campeón cae en dieciseisavos; quien tiene España/Francia/etc. sigue verde.

### 4.2 “Lo que te juegas hoy” (`buildHoy` + `koMatchStakeHtml`)

Cada cruce KO en el JSON público incluye:

| Campo | Contenido |
|---|---|
| `picks[]` | Marcador + ganador predicho por persona |
| `outcome_dist` / `modal_scoreline` | Analítica agregada (como grupos) |
| `stake` | `max_one`, `max_swing`, `people[]` con pts y swing en ranking KO |

La pestaña **Hoy** muestra consenso de marcador, reparto 1/X/2, stake detallado y grid de apuestas.

### 4.3 Subidón / batacazo (`D.knockout.progression`)

- `compute_ko_progression()`: snapshot del ranking KO tras **cada partido jugado**.
- Expuesto en **`D.knockout.progression`** (`table`, `progression[]` con `delta`, `round_pts`, etc.).
- UI en **Acto 4** del relato: bloque “Subidón y batacazo del último cruce”.
- El ranking general (`D.live`) sigue acumulando grupos + bonus + KO en un solo paso virtual.

### 4.4 Últimos resultados

Los cruces KO en **Últimos** ya mostraban pleno/signo/pase; sin cambio estructural, pero ahora
conviven con la progression KO dedicada en el relato.

---

## 5. Qué sigue aproximado / pendiente (Fase 3)

| Pieza | Motivo |
|---|---|
| **Riesgo/recompensa** (scatter) | `variance` = picks contra consenso, no probabilidad implícita |
| **`exp` sin resultados KO** | Aproximación; con resultados = puntos reales |
| Progression KO multi-jornada rica | Crece sola; con 1 partido hay un solo paso |
| Bracket pueblo vs individuos, reventador, camino del campeón | Extras Fase 3 |
| Premios honoríficos KO | El agorero, el manual, etc. |

---

## 6. Checklist operativo durante el torneo

1. Pegar/completar porras KO en `Raw data`.
2. Rellenar `Real results` según se juega.
3. `python3 generate_dashboard.py` → reescribe `index.html`.
4. `git add -A && git commit -m "..." && git push` → GitHub Pages (~1 min).

---

## 7. Datos pendientes

- Completar cuadro de los **10 participantes** sin campeón / con picks parciales.
- Ir rellenando resultados KO en `Real results` partido a partido.
