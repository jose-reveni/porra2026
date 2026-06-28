# Prototipo — métricas de eliminatorias (DESECHABLE)

**Pregunta que responde:** ¿cómo deberían verse las métricas "chulas y divertidas"
de la fase de eliminación en la web? (consenso de campeón, índice de caos,
riesgo/recompensa, gemelos de cuadro, cementerio, supervivencia, dossier por
selección, pichichi, lo que te juegas hoy…).

**Forma:** 3 variantes radicalmente distintas de la pestaña **Eliminatorias**,
conmutables con `?ko=A|B|C` + barra flotante (‹ ›, flechas ←/→). Es el mismo
patrón que el prototipo del ranking en directo (`?variant=`), pero aislado en su
propio parámetro para no pisarlo.

**Datos:** todo con **datos DEMO** generados en JS (`koDemoData()`), porque las
apuestas de eliminatorias aún no están subidas al Excel. Banner de aviso en cada
variante.

## Cómo verlo

```bash
python3 generate_dashboard.py
python3 -m http.server 8753   # file:// no sirve para el bracket
```

- `http://localhost:8753/index.html?view=ko&ko=A` — **Cuadro vivo**: el bracket
  como héroe + dossier interactivo por selección (hasta dónde la ve la oficina) +
  línea de supervivencia.
- `…?view=ko&ko=B` — **Sala de mandos**: rejilla densa de KPIs (campeón del
  pueblo, índice de caos, scatter riesgo/recompensa, gemelos, cementerio,
  pichichi, lo que te juegas hoy).
- `…?view=ko&ko=C` — **El relato ILUSTRADO** (favorito): scrollytelling editorial
  de 8 actos: **Acto 1 = el bracket** (el de la variante A, incrustado), luego
  reparto de campeón, índice de caos + scatter riesgo/recompensa (con **hover**
  que dice quién es cada punto), supervivencia, gemelos, cementerio,
  pichichi/profeta y **Acto 8 = directorio de fichas** por persona con buscador
  (estilo fase de grupos).
- `…?view=ko` (sin `ko`) — vista actual intacta (bracket + supervivencia + calendario).
- `…?view=live` — **En directo** ahora incluye **🔥 Lo que te juegas hoy** (sección
  `buildStakeToday`, entre los partidos de hoy y el ranking).

Las gráficas viven en helpers reutilizables (`koBracketBox`, `koChampBarsHtml`,
`koChaosBarsHtml`, `koScatter`, `koTwinsHtml`, `koGraveHtml`, `koPichichiHtml`,
`koStakeHtml`, `koSurvivalCard`) que comparten el cuadro, En directo, B y C.

## Dónde vive (para borrar luego)

Todo en `generate_dashboard.py`:
- CSS bajo el comentario `KO METRICS PROTOTYPE`.
- Bloque JS `KO METRICS PROTOTYPE` (funciones `*Ko*`, `buildKoVariantB/C`, etc.).
- Hooks en `renderView()` (rama `view === 'ko'`) y `rebuild()` (`renderKoSwitcher()`).

## Veredicto

> **Dirección elegida (28-jun):** gana el **formato relato (C)**, pero llevando
> dentro las gráficas chulas de **B**. Hecho: C es ahora el relato ilustrado.
>
> **Actualización (28-jun):** el bracket de A ya va incrustado como **Acto 1** del
> relato (C), así que la variante **A** queda redundante (candidata a borrar).
> **Lo que te juegas hoy** se ha movido a **En directo** (fuera de B).
>
> Pendiente de afinar antes de plegar en `buildEliminatorias`:
> - Borrar la variante **A** (su dossier "¿hasta dónde la ve la oficina?" podría
>   rescatarse como acto extra si gusta).
> - ¿Orden definitivo de los actos? ¿Algún acto extra (subcampeón)?
> - Cuando suban las apuestas KO, sustituir `koDemoData()` por datos reales y
>   conectar "Lo que te juegas hoy" al partido real del día.
