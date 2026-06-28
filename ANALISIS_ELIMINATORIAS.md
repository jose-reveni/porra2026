# Análisis de la fase de eliminatorias

Esta guía describe **qué datos hay**, **qué calcula el dashboard hoy** y **qué
queda por hacer** en la fase eliminatoria.

Relacionado: `PROTOTYPE_ko_metrics.md` (variantes `?ko=B|C` del relato) y
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
| Resultados KO jugados | **1** — R32-M3: Canadá 0-1 Sudáfrica |
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
| `compute_knockout` | Consenso por cruce, cobertura, calendario FIFA | ✅ |
| `compute_knockout_scoring` | Ranking de puntos KO (marcador + pases + premios) | ✅ (1 partido) |
| **`compute_knockout_metrics`** | Métricas del relato (`D.knockout.metrics`) | ✅ **Fase 1** |

### Vista producción (`?view=ko`)

Una sola vista: el **relato** con datos reales (actos 1–9) + calendario completo al final.
No hay variantes `?ko=A|B|C` ni barra flotante de prototipo.

### Bracket (`koBracketBox` / Acto 1 del relato)

- Termómetro por cruce con **barra doble** local/visitante (% sobre N).
- Resultado real cuando existe.
- Resolución de `W##` a medida que entran resultados (p. ej. Canadá en R16-M2).

---

## 3. Fase 1 implementada — métricas reales en el relato

Las métricas salen de **`D.knockout.metrics`** (Python). El JS ya no genera datos demo.

Bloque `compute_knockout_metrics()` → expuesto en `D.knockout.metrics`:

| Campo | Origen |
|---|---|
| `champRank` | Distribución de `outright.champion` |
| `tsRank` | Distribución de `awards.top_scorer` |
| `twins` | Similitud de `winner_picks` + campeón/sub/pichichi |
| `chaosRank` / `people[].chaos` | Votos al **peor prestigio** del cruce (1.º=3, 2.º=2, 3.º=1) |
| `grave` | Campeones cuyo equipo **ya no puede ganar** (fuera de R32 o eliminado KO) |
| `depth` | % de cuadros que llevan a cada selección a R16/QF/SF/Final/Campeón |
| `people[]` | Ficha: campeón, sub, pichichi, caos, riesgo (`variance`), puntos KO |
| `pool` | Selecciones del dossier interactivo (top prestigio R32) |

**Acto 7 “El profeta”** usa el **ranking KO real** (`scoring.table`) cuando hay
resultados; si no, aproximación por alineación con el consenso.

### Qué sigue en demo / aproximado

| Pieza | Motivo |
|---|---|
| **Supervivencia** (`koSurvivalCard`) | Datos reales de `metrics.people` cuando hay métricas; sin pill demo |
| **Riesgo/recompensa** (scatter) | `variance` = picks contra consenso (aprox., no modelo probabilístico fino) |
| **`exp` sin resultados KO** | Aproximación; con resultados = puntos reales del ranking |
| **“Lo que te juegas hoy”** | Solo muestra techo de pts/partido; falta reparto por persona |

---

## 4. Checklist operativo durante el torneo

1. Pegar/completar porras KO en `Raw data`.
2. Rellenar `Real results` según se juega.
3. `python3 generate_dashboard.py` → reescribe `index.html`.
4. `git add -A && git commit -m "..." && git push` → GitHub Pages (~1 min).

---

## 5. Next steps (prioridad)

### Fase 2 — Con más partidos KO

1. **Supervivencia real** — sustituir hash de `koSurvivalCard()` por eliminación
   ronda a ronda del campeón/cuadro predicho.
2. **“Lo que te juegas hoy”** — quién gana/pierde cuántos pts en el cruce del
   día (ampliar `koMatchStakeHtml`).
3. **Subidón/batacazo del día** — delta del ranking KO tras cada jornada (como
   en grupos).

### Fase 3 — Extras

4. Bracket del pueblo vs individuos; el reventador; camino del campeón.
5. Riesgo/recompensa con probabilidad implícita del consenso.
6. Premios honoríficos KO (*El profeta* por aciertos, *El agorero*, *El manual*).

### Datos pendientes

- Completar cuadro de los **10 participantes** sin campeón / con picks parciales.
- Ir rellenando resultados KO en `Real results` partido a partido.