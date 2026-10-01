# Presupuesto SOFSA 2027 · Plurianual 2027–2029

Tablero en un único archivo (`index.html`, se abre directo en el navegador o por GitHub Pages).

## Pestañas

La pestaña inicial es **Red del presupuesto**. El tablero se adapta a celulares (en pantallas chicas la red se muestra como lista desplegable).

- **Presupuesto 2027** — presupuesto anual con los techos informados por el Ministerio de Economía: cascada plurianual → techo → presupuesto presentado, techos por inciso, cuota mensual, Gasto Habitual vs Emergencia Ferroviaria, ingresos propios, distribución por centro gestor (con compromisos 2027 ya asignados) y detalle por PosPre.
- **Mensualizado** — presupuesto 2027 por mes, inciso y PosPre, abierto en Gasto Habitual y Emergencia Ferroviaria, con techo, requerido y brecha, y también por área (centro gestor). Exporta a Excel (.xlsx): Resumen, Comparativo, Techo y Requerido mensual (GH, EF, Total) y Base.
- **Plurianual** (pestaña existente) — suma el plurianual mensualizado: 2027 por mes (AxI) y 2028/2029 anual, por inciso, PosPre y componente, comparado con el techo 2027 y exportable a Excel.
- **Histórico y ejecución** — techos y sobre techos 2020–2027 (pedidos vs otorgados), formulación vs presupuesto otorgado vs ejecución 2022–2026, en pesos corrientes o constantes de 2027. Datos cargados desde "Presupuesto y Techos.xlsx" en la constante `HIST` de `index.html`.
- **Inciso IV · SIFER** — Bienes de Uso por fuente de financiamiento (SIFER = Material Rodante DMU, PosPre 4.3.3.0.1; resto Tesoro), detalle de las DMU y proyectos con filtros, apertura mensual y CSV.
- **Ingresos y subsidio** — gasto operativo (gastos corrientes − ingresos), subsidio por mes (ingresos / gastos corrientes), subsidio por pasajero pago, ingresos y pasajeros pagos por línea y mes.
- **Lineamientos Presidencia** — cada proyecto/producto de la formulación clasificado en las 4 premisas de Presidencia (principal + secundarias), con cobertura por premisa, estado de cada compromiso de "Relación con el presupuesto 2027", interdependencia entre premisas, proyectos transversales y un árbol lineamiento › proyecto › producto que se puede exportar a CSV.
- **Red del presupuesto** — árbol deductivo interactivo: presupuesto → corriente/capital → inciso → emergencia/habitual → área → premisa → proyecto → producto, con niveles configurables, búsqueda y montos por techo o requerido.
- **Presentación** — slides generadas en vivo con los mismos datos (flechas ← → y pantalla completa).
- Resto de las pestañas: análisis del plurianual 2027–2029 (sin cambios).

## Datos
- `data/anual_2027.json` — extraído de "Ppto 2027 - V3" (versión final presentada). Va embebido en `index.html`.
- `data/meta_pluri.json` — datos del plurianual ya existentes en el tablero (se usan para rubro/criticidad por PosPre).

Para regenerar a partir de una nueva versión del Excel:

```
pip install openpyxl
python3 tools/build_anual_2027.py ruta/al/Ppto_2027_V3.xlsx ruta/a/PPTO_2027_Ingresos_Totales.xlsx
```

Las reglas de clasificación en lineamientos están en `classify()` dentro de `tools/build_anual_2027.py`.
