"""Genera data/anual_2027.json a partir del Excel del presupuesto 2027 y lo embebe en index.html.

Uso:  python3 tools/build_anual_2027.py <ruta/Ppto_2027.xlsx> [ruta/PPTO_2027_Ingresos_Totales.xlsx]

Toma:
  - "Formulación 27": líneas de requerimiento (proyecto, producto, prioridad, PosPre,
    requerido homogéneo ajustado, techo, sobre techo y cuota mensual del techo).
  - "Distribución de techos": techos informados por el Ministerio de Economía,
    pedido adicional a la S.T. y distribución de techos por Centro Gestor.
  - "Asignado 26 y 27": compromisos 2027 ya asignados (Solped / OC).
  - "Ingresos": ingresos propios 2027 actualizados.
  - "Variaciones": plurianual 2027 (AxI) por PosPre.
Cada línea se clasifica en los 4 lineamientos de Presidencia (ver classify()).
"""
import json, re, sys, os
from collections import defaultdict
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

L1, L2, L3, L4 = 'L1', 'L2', 'L3', 'L4'
LINEAMIENTOS = {
    L1: {'nombre': 'Experiencia del pasajero e imagen institucional', 'corto': 'Experiencia del pasajero',
         'premisa': 'Mejorar la experiencia del pasajero, cambiar la imagen institucional, garantizar la limpieza de las formaciones y flexibilizar la oferta de servicios (trenes rápidos / diferenciales).',
         'indicadores': ['Índice de satisfacción del pasajero (encuesta periódica), con meta de mejora interanual.',
                         '% de formaciones que cumplen el estándar de limpieza.',
                         'N° de líneas / franjas horarias con diagrama de servicios diferenciales implementado.']},
    L2: {'nombre': 'Seguridad operacional, mantenimiento y calidad del servicio', 'corto': 'Seguridad y mantenimiento',
         'premisa': 'Garantizar y fortalecer la seguridad operacional mediante procesos y procedimientos clave, cumplir la totalidad de las revisiones decanuales, mejorar la calidad del material rodante en talleres y desarrollar un área de control de los planes de mantenimiento.',
         'indicadores': ['% de revisiones decanuales cumplidas sobre las programadas.',
                         'Tasa de incidentes de seguridad operacional.']},
    L3: {'nombre': 'Gestión administrativa, normativa y de control', 'corto': 'Gestión y control',
         'premisa': 'Adecuar la normativa de gestión delimitando responsabilidades, modernizar la administración para mayor eficiencia en la generación de información, ordenar y armonizar los KPIs de las líneas y relevar el estado de situación legal de la empresa.',
         'indicadores': ['Relevamiento legal completo y actualizado.',
                         'Normativa de responsabilidades aprobada y difundida a las áreas operativas.']},
    L4: {'nombre': 'Organización, personas y eficiencia de recursos', 'corto': 'Personas y eficiencia',
         'premisa': 'Posibilitar el desarrollo de carrera del personal, mejorar la eficiencia financiera en la utilización de los recursos y armonizar la estructura organizativa de las líneas.',
         'indicadores': ['% de vacantes cubiertas mediante desarrollo de carrera interno vs. contratación externa.',
                         'N° de líneas con estructura organizativa armonizada respecto del total.',
                         'Indicador de eficiencia financiera con meta de mejora interanual.']},
}

IMP_POSPRE = {'IVA': '3.8.9.0.2', 'IDCB': '3.8.9.0.3', 'IIBB': '3.8.9.0.1'}
INC_MAP = {'Gastos en Personal': '1. Gastos en Personal', 'Bienes de Consumo': '2. Bienes de Consumo',
           'Servicios No Personales': '3. Servicios No Personales', 'Impuestos': '5. Impuestos',
           'Bienes de Uso - Inversiones': '4. Bienes de Uso'}
PRIO = {'alta': 'Alta', 'a': 'Alta', 'media': 'Media', 'm': 'Media', 'b': 'Media', 'baja': 'Baja', 'c': 'Baja'}


def num(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def has(t, pat):
    return re.search(pat, t, re.I) is not None


def classify(inc, cege, proy, prod, pp, ppdesc, det, efgh):
    """Devuelve (lineamiento principal, [lineamientos secundarios], criterio)."""
    t = ' '.join([proy, prod, det]).lower()
    c = cege.lower()
    if inc == 'Gastos en Personal':
        return L4, [], 'Dotación y salarios: base de la premisa de personas'
    if inc == 'Impuestos':
        return L3, [], 'Obligaciones fiscales (cumplimiento normativo)'
    if pp == '4.3.3.0.1':
        return L2, [L1], 'Renovación de flota: disponibilidad, seguridad y capacidad de oferta'
    if pp.startswith('4.2.1'):
        n = proy.lower()
        if has(n, r'estaci[oó]n|estaciones|anden|apeadero|puentes peatonales|bms|puesta en valor') \
                and not has(n, r'subestaci|se[ñn]alamiento|cerramiento'):
            return L1, [L2] if has(n, r'segur') else [], 'Obra en estaciones / entorno del pasajero'
        if has(n, r'dependencias (para )?personal|dependencias operativas'):
            return L4, [], 'Condiciones de trabajo del personal'
        if has(n, r'almac|dep[oó]sitos y pa[ñn]oles'):
            return L4, [L2], 'Logística y abastecimiento'
        if has(n, r'tratamiento|efluente|desag[uü]e|saneam'):
            return L3, [], 'Cumplimiento ambiental / normativo'
        if has(n, r'telecomunic'):
            return L2, [], 'Comunicaciones operativas'
        if has(n, r'taller'):
            return L2, [], 'Talleres de material rodante'
        return L2, [], 'Obra de vía, señalamiento, energía u obras de arte'
    if pp.startswith('4.') and efgh == 'EF':
        return L2, [], 'Equipamiento operativo (Emergencia Ferroviaria)'
    if 'material rodante' in c:
        if has(t, r'administrativ'):
            return L4, [], 'Soporte administrativo'
        return L2, [L1], 'Mantenimiento y reparación de material rodante (decanuales, talleres)'
    if 'prevenc' in c:
        return L1, [], 'Seguridad del pasajero en estaciones'
    if 'seg. operacional' in c:
        return L2, [L3], 'Gestión de la seguridad operacional'
    if 'serv m' in c:
        if has(t, r'factores humanos|psico|crpc'):
            return L2, [L4], 'Aptitud psicofísica del personal operativo'
        if has(t, r'pasajera|protegida|cardio'):
            return L1, [L4], 'Cobertura sanitaria de pasajeros y empleados'
        return L4, [], 'Salud ocupacional'
    if 'hig, seg' in c:
        if has(t, r'ambient'):
            return L3, [], 'Gestión ambiental regulatoria'
        if has(t, r'mobiliario|activos fijos'):
            return L4, [], 'Equipamiento'
        return L2, [L4], 'Prevención de riesgos y EPP'
    if 'planta combust' in c or pp == '3.1.1.0.1' or pp == '2.5.9.0.3':
        return L2, [], 'Energía / combustible de tracción: continuidad del servicio'
    if 'prensa' in c:
        return L1, [], 'Comunicación institucional e imagen'
    if 'comercial' in c:
        if has(t, r'ambientaci|institucional|se[ñn]al'):
            return L1, [L4], 'Imagen en espacios del pasajero'
        return L4, [], 'Explotación comercial de activos (ingresos propios)'
    if 'ggaj' in c or 'jur' in c:
        return L3, [], 'Estado legal: juicios y normativa'
    if 'auditor' in c or 'integr' in c:
        return L3, [], 'Control interno e integridad'
    if 'plan. y control' in c:
        return L3, [L4], 'Planificación, KPIs y control de gestión'
    if 'proy y proc' in c:
        return L3, [L2], 'Procesos, procedimientos y calidad (ISO 9001)'
    if 'riesgos' in c:
        return L3, [], 'Gestión de riesgos y seguros'
    if 'sistemas y proc' in c:
        return L3, [L4], 'Modernización administrativa y generación de información'
    if 'tec. inf' in c:
        if pp.startswith('3.3.') or pp.startswith('2.9.'):
            return L2, [], 'Mantenimiento de señales y comunicaciones'
        if has(t, r'ticketing'):
            return L1, [L3], 'Venta de pasajes / ticketing'
        return L3, [], 'Infraestructura digital y telecomunicaciones'
    if 'des. pers' in c or pp == '3.4.9.0.4':
        return L4, [L2], 'Formación y desarrollo de carrera'
    if 'automotores' in c or 'abast' in c:
        return L4, [], 'Eficiencia logística y de flota de apoyo'
    if pp.startswith('4.'):
        return L4, [], 'Equipamiento y activos de soporte'
    # Compras centralizadas, Ingeniería (gasto corriente), Coordinación Administrativa, líneas, etc.
    if has(t, r'limpieza de estaciones|mantenimiento y limpieza|estaciones y trazas|se[ñn]al[eé]tica informativa') \
            or pp in ('3.3.5.0.1', '2.5.9.0.4', '2.5.9.0.1', '3.2.6.0.7'):
        if has(t, r'zona de v[ií]a|obras civiles|obras de v[ií]a'):
            return L2, [], 'Mantenimiento de zona de vía'
        return L1, [], 'Limpieza y condiciones de estaciones y trazas'
    if has(t, r'\bv[ií]as?\b|se[ñn]alamiento|equipos pesados|ingenier[ií]a el[eé]ctrica|asistencia directa|energ[eé]tico cr[ií]tico') \
            or pp.startswith('2.9.') or pp.startswith('3.3.3') or pp.startswith('3.3.4'):
        return L2, [], 'Mantenimiento de infraestructura y repuestos'
    if has(t, r'formaci[oó]n operativa'):
        return L4, [L2], 'Formación operativa'
    return L4, [], 'Soporte administrativo y servicios generales (eficiencia de recursos)'


def rows_of(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def ingresos_detalle(xlsx):
    """Hoja "Detalle" del archivo de ingresos: ámbito, línea, tipo, concepto y 12 meses (incluye PPT)."""
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    out = []
    for r in list(wb['Detalle'].iter_rows(values_only=True))[1:]:
        if not r or not r[1]:
            continue
        m = [num(x) for x in r[4:16]]
        if not any(m):
            continue
        out.append(dict(amb=str(r[0]).strip(), linea=str(r[1]).strip(), tipo=str(r[2]).strip(),
                        c=str(r[3]).strip(), m=[round(x, 2) for x in m]))
    return out


def main(xlsx, xlsx_ing=None):
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    meta = json.load(open(os.path.join(ROOT, 'data', 'meta_pluri.json'), encoding='utf-8'))
    pmeta = {d['pospre']: d for d in meta['data']}

    # ---------- Formulación 27 (líneas) ----------
    F = rows_of(wb['Formulación 27'])
    hdr = F[7]
    ix = {h: i for i, h in enumerate(hdr) if isinstance(h, str)}
    iT = ix['Techos']; iS = ix['Sobre techos']; iReq = ix['27 Hom. Aj']; iHom = ix['27 Hom']
    iMes = [ix[m + ' T'] for m in ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
                                   'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']]
    lines = []
    for r in F[8:]:
        inc = (r[ix['Inciso']] or '').strip()
        if not inc:
            continue  # subtotales
        s = lambda k: str(r[ix[k]] or '').strip()
        pp = s('PosPre'); ppd = s('Pospre descripción')
        if not pp or pp == '(en blanco)':
            pp = IMP_POSPRE.get(ppd, pp)
        efgh = 'EF' if 'emergencia' in s('Emergencia Ferroviari - Gasto Habitual').lower() else 'GH'
        prio = PRIO.get(s('Prioridad (A-M-B)').lower(), 'Sin prioridad')
        proy = s('Nombre del Proyecto') or '(sin proyecto asignado)'
        prod = s('Producto') or '(sin producto)'
        det = ' '.join([s('Descripción general'), s('Detalle del Producto'), s('Concepto del insumo (Inciso II y III)')])
        lp, ls, crit = classify(inc, s('CEGE'), proy, prod, pp, ppd, det, efgh)
        lines.append(dict(cege=s('CEGE'), cod=s('Administrador'), proy=proy, prod=prod, prio=prio, pp=pp, ppd=ppd,
                          inc=INC_MAP[inc], efgh=efgh, req=num(r[iReq]), hom=num(r[iHom]), t=num(r[iT]), s=num(r[iS]),
                          m=[num(r[i]) for i in iMes], mr=[num(r[iReq + 1 + k]) for k in range(12)], L=lp, Ls=ls, crit=crit))

    # agregamos a nivel proyecto·producto·pospre (clave de análisis) para aliviar el tamaño
    agg = {}
    for l in lines:
        k = (l['cege'], l['proy'], l['prod'], l['pp'], l['efgh'], l['prio'], l['L'])
        a = agg.get(k)
        if a is None:
            a = agg[k] = dict(l, n=0, req=0.0, hom=0.0, t=0.0, s=0.0, m=[0.0] * 12, mr=[0.0] * 12)
            a['Ls'] = list(l['Ls'])
        a['n'] += 1
        for f_ in ('req', 'hom', 't', 's'):
            a[f_] += l[f_]
        a['m'] = [x + y for x, y in zip(a['m'], l['m'])]
        a['mr'] = [x + y for x, y in zip(a['mr'], l['mr'])]
    L = []
    for a in agg.values():
        pm = pmeta.get(a['pp'], {})
        rub = pm.get('rubro') or ('Gastos en Personal' if a['inc'].startswith('1') else 'Otros bienes esenciales de consumo')
        cri = pm.get('criticidad') or ('Gastos en Personal' if a['inc'].startswith('1') else 'Complementario')
        L.append(dict(cege=a['cege'], cod=str(a['cod']).split('.')[0], proy=a['proy'], prod=a['prod'], prio=a['prio'], pp=a['pp'], ppd=a['ppd'],
                      inc=a['inc'], rub=rub, cri=cri, efgh=a['efgh'], req=round(a['req'], 2), t=round(a['t'], 2),
                      s=round(a['s'], 2), m=[round(x, 2) for x in a['m']], mr=[round(x, 2) for x in a['mr']], L=a['L'], Ls=a['Ls'], crit=a['crit'], n=a['n']))

    # ---------- Distribución de techos ----------
    D = rows_of(wb['Distribución de techos'])
    lab = {str(r[0]).strip(): r + [None] * 12 for r in D[:20] if r and r[0]}
    def dist(name):
        r = lab[name]
        return dict(pluri=num(r[1]), techo_pluri=num(r[2]), req_orig=num(r[3]), req_axi=num(r[4]), techo=num(r[5]),
                    pedido_st=num(r[6]), sobre=num(r[7]), ts=num(r[8]), informado=num(r[10]))
    dist_rows = {k: dist(k) for k in ['Gastos en Personal', 'Bienes de Consumo', 'Servicios No Personales', 'Impuestos',
                                      'Gastos Operativos', 'Bienes de Uso - Inversiones', 'Gastos Totales']}
    ingresos_op = num(lab['Ingresos Operativos'][4])
    ceges = []
    for r in D[22:60]:
        if not r[1] or not str(r[0]).strip().isdigit():
            continue
        ceges.append(dict(cod=str(r[0]), nombre=str(r[1]).strip(), pluri=num(r[2]), req=num(r[3]), dev=num(r[6]),
                          techo=num(r[14]), solped=num(r[15]), oc=num(r[16]), asignado=num(r[17]),
                          disp=num(r[20]), sobre=num(r[22]), ts=num(r[23])))

    # ---------- Ingresos 2027 ----------
    I = rows_of(wb['Ingresos'])
    ing = []
    for r in I:
        if r[1] and isinstance(r[1], str) and r[1].strip() not in ('Ingresos Ppto 2027',) and isinstance(r[14], (int, float)):
            ing.append(dict(c=r[1].strip(), m=[num(x) for x in r[2:14]], t=num(r[14])))
    pax = next((r for r in I if r[1] and 'Pasajeros Pagos' in str(r[1])), None)

    # ---------- Plurianual 2027 por PosPre (Variaciones) ----------
    V = rows_of(wb['Variaciones'])
    pluri_pp = {}
    for r in V[3:]:
        pp = str(r[1] or '').strip(); d = str(r[2] or '').strip()
        if pp in ('(en blanco)', '') and d in IMP_POSPRE:
            pp = IMP_POSPRE[d]
        if pp and not str(r[0] or '').startswith('Total'):
            pluri_pp[pp] = dict(desc=d, pluri=num(r[3]), req=num(r[4]))
        if str(r[0] or '').startswith('Total Gastos en Personal'):
            pluri_pp['Personal'] = dict(desc='Gastos en Personal', pluri=num(r[3]), req=num(r[4]))

    # ---------- Plurianual 2027-2029 mensualizado (hoja "Pluri 27-28-29", valores AxI) ----------
    # 2027 viene por mes; 2028 y 2029 son anuales. Emergencia/Habitual: en Bienes de Uso según la marca de
    # cada fila; en el resto, proporción por PosPre del template SSTF (la misma que usa el tablero plurianual).
    P = rows_of(wb['Pluri 27-28-29'])
    imp_pp = {'IVA': '3.8.9.0.2', 'IDCB': '3.8.9.0.3', 'IIBB': '3.8.9.0.1'}
    pacc = {}
    for r in P[8:]:
        r = list(r) + [None] * 60
        inc = str(r[6] or '').strip()
        if inc not in INC_MAP:
            continue
        pp = str(r[2] or '').strip()
        pp = 'Personal' if inc == 'Gastos en Personal' else imp_pp.get(pp, pp)
        m = [num(r[41 + k]) for k in range(12)]
        y = [num(r[53]), num(r[54]), num(r[55])]
        pm = pmeta.get(pp, {})
        if inc == 'Bienes de Uso - Inversiones':
            parts = [('EF' if 'emergencia' in str(r[9] or '').lower() else 'GH', 1.0)]
        else:
            tot = (pm.get('gh27', 0) + pm.get('ef27', 0)) or 1
            parts = [('GH', pm.get('gh27', tot) / tot), ('EF', pm.get('ef27', 0) / tot)]
        for c, w in parts:
            if w <= 0:
                continue
            a = pacc.setdefault((c, pp), dict(pp=pp, desc=pm.get('desc') or str(r[3] or '').strip(), inc=INC_MAP[inc],
                                               rub=pm.get('rubro', ''), cri=pm.get('criticidad', ''), imp=pm.get('imputacion', ''),
                                               efgh=c, m=[0.0] * 12, y=[0.0, 0.0, 0.0]))
            a['m'] = [x + v * w for x, v in zip(a['m'], m)]
            a['y'] = [x + v * w for x, v in zip(a['y'], y)]
    pluri_m = [dict(v, m=[round(x, 2) for x in v['m']], y=[round(x, 2) for x in v['y']]) for v in pacc.values() if any(v['y'])]

    out = dict(
        version='Presupuesto 2027',
        lineamientos=LINEAMIENTOS, lines=L, dist=dist_rows, ingresos_op=ingresos_op, ceges=ceges,
        ingresos=ing, pluri_m=pluri_m, ing_det=ingresos_detalle(xlsx_ing) if xlsx_ing else [], pax=[num(x) for x in pax[2:14]] if pax else [], pluri_pp=pluri_pp,
    )
    json.dump(out, open(os.path.join(ROOT, 'data', 'anual_2027.json'), 'w', encoding='utf-8'), ensure_ascii=False)

    # ---------- embebido en index.html ----------
    p = os.path.join(ROOT, 'index.html')
    html = open(p, encoding='utf-8').read()
    blob = '<script id="anual-data">const ANUAL = ' + json.dumps(out, ensure_ascii=False, separators=(',', ':')) + ';</script>'
    html = re.sub(r'<script id="anual-data">.*?</script>', lambda _: blob, html, flags=re.S) if 'id="anual-data"' in html \
        else html.replace('</body>', blob + '\n</body>', 1)
    open(p, 'w', encoding='utf-8').write(html)
    tot = sum(l['t'] for l in L)
    print(f'lineas={len(lines)} agregadas={len(L)} techo={tot/1e6:,.0f}M')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
