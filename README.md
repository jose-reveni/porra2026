# Porra Mundial 2026 · Dashboard analítico (Reveni)

🔗 **En vivo:** https://jose-reveni.github.io/porra2026/

Genera un dashboard web **autocontenido** (`index.html`) a partir del
Excel de administración de la porra, con analíticas tipo "quién es el más rebelde",
"quién piensa como quién", estilo de apuesta, favoritos, partidos divisivos, etc.

## Uso

```bash
# 1) (solo la primera vez) instalar dependencia
pip3 install openpyxl

# 2) generar el dashboard
python3 generate_dashboard.py

# 3) abrir el resultado
open index.html        # macOS
```

Opcional, rutas personalizadas:

```bash
python3 generate_dashboard.py Porra_Admin_v4_EN.xlsx index.html
```

## Publicado en GitHub Pages

El sitio se sirve desde `index.html` en la raíz de la rama `main`. Para actualizarlo
tras regenerar:

```bash
python3 generate_dashboard.py    # reescribe index.html
git add -A && git commit -m "chore: actualizar dashboard" && git push
```

GitHub Pages tarda ~1 min en reflejar los cambios en
https://jose-reveni.github.io/porra2026/

El `.html` es un único archivo: lo puedes mandar por Slack/WhatsApp o subir a
cualquier sitio y se abre en cualquier navegador. (Para verlo con todo el estilo
necesita conexión la primera vez, porque carga las fuentes de Google Fonts; sin
internet usa fuentes del sistema.)

**Bilingüe:** arriba a la izquierda hay un selector **ES / EN** que traduce toda
la interfaz *y* los nombres de selección al vuelo. Arranca en español.

## Reejecutable durante el torneo

El dashboard se construye con las **apuestas** (72 marcadores de grupo,
clasificados, eliminatorias, campeón/subcampeón y premios).
Cuando empiece a rodar el balón:

1. Rellena los marcadores reales en la pestaña **`Real results`** del Excel.
2. Vuelve a correr `python3 generate_dashboard.py`.
3. La sección **"En directo"** mostrará el ranking de aciertos de grupos
   (signo correcto = +2, marcador exacto = +4) y **"Eliminatorias"** mostrará
   consenso del cuadro y marcador eliminatorio cuando haya datos de esa fase.

El calendario de eliminatorias (cruces, fechas, horas y sedes) se toma de la
página oficial de fixtures de FIFA y no del Excel. El Excel solo guarda
predicciones y resultados.

## Qué hay dentro

- **Índice de Rebeldía** — combina cuántas veces apuestas el signo (1/X/2) en
  minoría + cuánto se aleja tu marcador del marcador típico del grupo.
- **Afinidad** — matriz de quién coincide con quién (almas gemelas / polos opuestos).
- **Estilo** — goleadores vs cerrojos, marcadores más repetidos, rey del empate.
- **Favoritos** — consenso de clasificados por grupo, selecciones más respetadas.
- **Eliminatorias** — calendario oficial FIFA, consenso de campeón, subcampeón,
  ganadores por ronda, premios y ranking de aciertos de la fase eliminatoria.
- **Partidos** — los que más nos dividen vs los de casi unanimidad.
- **Lobo solitario** — marcadores que solo apostó una persona.
- **Fichas** — resumen por participante (buscador incluido).
- **Palmarés** — los premios honoríficos de la edición.

## Marca

Identidad visual de **Reveni**: teal oscuro `#022f36`, acento menta `#7afcd0`,
logo en `reveni-logo.svg`. La tipografía corporativa real es *Garnett* (de pago);
aquí se usan **Space Grotesk + Inter** como sustitutas libres.
