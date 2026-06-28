# Análisis de la fase de eliminatorias

Esta guía explica **cómo convertir el prototipo de métricas KO (datos demo) en
análisis real** una vez que cada participante haya subido su porra de
eliminatorias, y deja apuntadas las **métricas propuestas que aún no están
implementadas**.

Relacionado: `PROTOTYPE_ko_metrics.md` (qué es el prototipo y cómo verlo) y
`README.md` (flujo general del dashboard).

---

## 1. Dónde viven los datos

Todo sale del Excel `Porra_Admin_v4_EN.xlsx`. Cada participante ocupa **dos
columnas** (goles local / goles visitante o valor) en la hoja **`Raw data`**, con
su nombre en la fila 6.

| Bloque | Filas (hoja `Raw data`) | Qué hay |
|---|---|---|
| Marcadores de grupos | 7–78 | 72 partidos (ya cargado) |
| Clasificación de grupos (1.º/2.º/3.º) | 80–115 | quién pasa (ya cargado) |
| 8 mejores terceros | 117–124 | (ya cargado) |
| **R32** (16 cruces × score/desempate/ganador) | 126–173 | **porra KO** |
| **R16** (8 cruces) | 175–198 | **porra KO** |
| **Cuartos** (4 cruces) | 200–211 | **porra KO** |
| **Semis** (2 cruces) | 213–218 | **porra KO** |
| 3.er puesto + Final | 220–224 | **porra KO** |
| **Campeón / Subcampeón** | 225 / 226 | **porra KO** |
| **Pichichi / Balón de Oro** | 228 / 229 | **porra KO** |

Los **resultados reales** se rellenan en la hoja **`Real results`** (misma
disposición de filas; columnas C/D para el marcador, C para ganador / desempate /
campeón / premios) a medida que se juegan los partidos.

> Hoy mismo, las filas KO (126–229) están **vacías**: por eso el prototipo usa
> datos de ejemplo. Las de grupos y clasificación **sí** están rellenas.

---

## 2. Qué calcula ya el Python (no hay que tocarlo)

En `generate_dashboard.py`:

- `parse_knockouts(raw, participants)` — lee la porra KO de cada persona:
  `score_picks`, desempate si hay empate, `winner_picks` por cruce, y `outright`
  (campeón, subcampeón) y `awards` (pichichi, balón de oro).
- `parse_knockout_results(wb)` — lee los resultados reales de `Real results`.
- `compute_knockout(data)` — consenso por cruce (`_score_consensus`,
  `_text_consensus`): ganador más votado, marcador más repetido, % de acuerdo.
  Devuelve `D.knockout` con `rounds`, `final_matches`, `outright`, `awards`,
  `pct`/`filled` (cobertura de carga) y `ready` (true en cuanto hay ≥1 dato).
- `compute_knockout_scoring(data)` — cuando hay resultados reales, calcula el
  **ranking de puntos KO** (3 por signo a 90', +2 exacto, y los
  puntos por avance: R16 +1, QF +2, SF +4, finalista +6, 3.º +4, subcampeón +8,
  campeón +12, pichichi +8, balón de oro +8).

Cuando subas las porras y vuelvas a generar, el **bracket** (`buildBracket`) ya
mostrará el % de consenso del ganador en cada cruce, y la pestaña Eliminatorias
mostrará el campeón más votado, subcampeón, pichichi y el ranking de puntos.

---

## 3. El paso que falta: cablear las métricas nuevas a datos reales

Las métricas del relato (campeón del pueblo, índice de caos, riesgo/recompensa,
gemelos, supervivencia, dossier, fichas, "lo que te juegas hoy") **se dibujan con
`koDemoData()`** (JS, datos inventados). Plan para pasarlas a real:

1. **Calcular en Python** un bloque nuevo, p. ej. `compute_knockout_metrics(data)`,
   y exponerlo en `D.knockout.metrics`. Así viaja como JSON y el JS solo pinta.
2. **Sustituir `koDemoData()`** para que lea `D.knockout.metrics` cuando exista
   (`ready`) y caiga al demo solo si no hay datos. En la práctica: cambiar el
   cuerpo de `koDemoData()` por `return D.knockout.metrics || <demo>;`.
3. Quitar el banner "datos de ejemplo" (`koDemoBanner`) cuando `ready`.

### Receta por métrica (de dónde sale cada número con datos reales)

- **Campeón del pueblo / subcampeón / pichichi** → ya está en
  `D.knockout.outright` y `D.knockout.awards` (distribución `dist` + `count`).
  Reusar directamente, no hace falta recalcular.
- **Índice de caos** (sorpresas por persona) → para cada `winner_picks`, contar
  cuántas veces esa persona hace avanzar al equipo **peor clasificado** del cruce
  (el de "seed" más bajo según el ranking de grupos / prestigio). Más upsets =
  más caos.
- **Riesgo vs recompensa** → *recompensa* = puntos esperados = suma, sobre los
  cruces/premios, de `P(acierto) × puntos`, usando como `P` la **probabilidad
  implícita del consenso** (cuánta gente votó lo mismo que tú). *Riesgo* =
  varianza de esa distribución (apostar con la mayoría = bajo riesgo; ir solo =
  alto). Es el cálculo más "fino"; empezar con una aproximación simple.
- **Gemelos de cuadro** → similitud por par = % de coincidencia entre
  `winner_picks` (peso alto) + campeón + subcampeón + pichichi. Es el análogo KO
  de la afinidad de grupos (`compute` de afinidad ya existe como referencia).
- **Supervivencia / "¿cuándo cae?"** → necesita **resultados reales**: para cada
  persona, marcar la primera ronda en la que su ganador/campeón predicho ya ha
  sido eliminado de verdad. Hasta que no haya resultados, queda en demo.
- **Cementerio** → con resultados reales de grupos/KO: personas cuyo **campeón ya
  está eliminado**. (La parte de grupos se puede calcular **ya**, porque la fase
  de grupos terminó.)
- **Dossier de selección / profundidad esperada** → % de cuadros que llevan a
  cada equipo a R16/QF/SF/Final/Campeón, contando sobre los `winner_picks`.
- **Fichas por persona** → resumen por participante: su campeón, subcampeón,
  pichichi, índice de caos, gemelo y hasta dónde aguanta. Todo derivado de los
  campos anteriores.
- **Lo que te juegas hoy** (pestaña En directo) → con el calendario, identificar
  el/los cruce(s) de hoy y calcular el máximo vuelco de puntos posible según los
  `winner_picks`/`advance_points`. Hoy está fijo (demo Spain–Croatia).

### Checklist para actualizar durante el torneo

1. Pegar las porras KO de cada uno en `Raw data` (columnas de cada participante).
2. Ir rellenando `Real results` según se juega.
3. `python3 generate_dashboard.py` → reescribe `index.html`.
4. `git add -A && git commit -m "..." && git push` → GitHub Pages (~1 min).

---

## 4. Métricas propuestas AÚN NO implementadas

Del brainstorm inicial, quedan pendientes (candidatas a siguientes iteraciones):

- **Termómetro por cruce en el bracket** con barra de doble lado (p. ej.
  "Brasil 71% / Croacia 29%") en cada nodo. *Parcial*: el nodo ya muestra el % del
  ganador cuando hay datos, pero no el reparto completo a ambos lados ni grosor de
  arista por consenso.
- **El bracket del pueblo (cerebro colmena)** — montar el cuadro de consenso como
  un "participante" más y puntuarlo contra la realidad: ¿gana la sabiduría de la
  masa a los individuos?
- **El reventador (bracket-buster)** — el resultado real concreto que rompe más
  cuadros a la vez ("el partido que reventó la porra").
- **Camino del campeón** — resaltar en el bracket la ruta que necesita el campeón
  de cada uno y cuántos de esos rivales siguen vivos.
- **Subidón / batacazo del día** — mayores escaladores y caídas del ranking KO
  tras cada jornada (existe algo así en el ranking de grupos, falta el de KO).
- **Head-to-head de hoy** — para los cruces del día, quién votó qué y qué implica
  para el avance (ampliación de "lo que te juegas hoy").
- **Premios honoríficos KO** — *El profeta* (más aciertos de avance), *El agorero*
  (más sorpresas acertadas), *El manual* (cuadro más conservador). *Parcial*: el
  "profeta" aparece en el relato pero con puntos esperados demo, no con aciertos
  reales.
- **Riesgo/recompensa con modelo de probabilidad real** — sustituir la varianza
  demo por la probabilidad implícita del consenso (ver receta arriba).
- **"Muerto al llegar"** — tapados (campeón/sorpresa) ya eliminados en la fase de
  grupos. **Calculable hoy** con los resultados de grupos ya cargados.

### Decisión de diseño pendiente

- La variante **A** ("Cuadro vivo", `?ko=A`) quedó **redundante**: su bracket ya
  vive como Acto 1 del relato (`?ko=C`). Candidata a borrar; su único extra es el
  dossier interactivo, que ya está como acto del relato.
