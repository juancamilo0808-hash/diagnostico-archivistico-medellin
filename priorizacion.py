"""Matriz de priorización en Excel.

Genera un libro con el diagnóstico más reciente de cada entidad: resultado por
componente, nivel de madurez, clasificación (mapa de calor) y prioridad de atención,
más resúmenes que se recalculan con fórmulas (por sector, por naturaleza y por
componente) y una ficha interactiva por entidad.

La función principal recibe datos ya calculados (lista de diccionarios), así que no
depende de la base de datos y se puede probar por separado.
"""
from datetime import datetime
from io import BytesIO

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# Paleta institucional y del mapa de calor.
AZUL = '0B2D50'
AZUL_MEDIO = '123B65'
DORADO = 'C99A2E'
GRIS = '56687A'
FONDO = 'F5F8FB'
BORDE = 'DCE4EC'
BAJO, BAJO_TXT = 'FBE3E1', '9B1C1C'
MEDIO, MEDIO_TXT = 'FFF1C2', '7A5A00'
ALTO, ALTO_TXT = 'DDF3E4', '17633A'
BAJO_FUERTE, MEDIO_FUERTE, ALTO_FUERTE = 'D64545', 'E3B341', '2F9E5B'

FUENTE = 'Arial'
LINEA = Side(style='thin', color=BORDE)
CAJA = Border(left=LINEA, right=LINEA, top=LINEA, bottom=LINEA)

PRIMERA_FILA = 6  # encabezado de la tabla de la matriz


def _f(size=11, bold=False, color='1F2937', italic=False):
    return Font(name=FUENTE, size=size, bold=bold, color=color, italic=italic)


def _relleno(color):
    return PatternFill('solid', fgColor=color)


def _titulo(ws, texto, subtitulo, ancho_cols):
    ws.sheet_view.showGridLines = False
    ws['B2'] = texto
    ws['B2'].font = _f(18, True, AZUL)
    ws['B3'] = subtitulo
    ws['B3'].font = _f(10, color=GRIS)
    for col in range(2, ancho_cols + 2):
        ws.cell(row=4, column=col).border = Border(top=Side(style='medium', color=DORADO))


def _encabezado(celda, color=AZUL):
    celda.font = _f(10, True, 'FFFFFF')
    celda.fill = _relleno(color)
    celda.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    celda.border = CAJA


def _etiquetas(formato=None):
    """Etiquetas de datos solo con el valor (sin nombre de serie ni de categoría)."""
    e = DataLabelList()
    e.showVal = True
    e.showSerName = e.showCatName = e.showLegendKey = e.showPercent = e.showLeaderLines = False
    if formato:
        e.numFmt = formato
    return e


def _nivel_formula(ref, umbral_bajo, umbral_medio):
    return (f'=IF({ref}="","Sin dato",IF({ref}<={umbral_bajo},"Bajo",'
            f'IF({ref}<={umbral_medio},"Medio","Alto")))')


def _reglas_nivel(ws, rango):
    """Colorea las celdas de clasificación según el mapa de calor."""
    for texto, fondo, tinta in (('Bajo', BAJO, BAJO_TXT), ('Medio', MEDIO, MEDIO_TXT), ('Alto', ALTO, ALTO_TXT)):
        ws.conditional_formatting.add(rango, CellIsRule(operator='equal', formula=[f'"{texto}"'],
                                                       fill=_relleno(fondo), font=Font(color=tinta, bold=True)))


def _reglas_porcentaje(ws, rango, primera_celda, umbral_bajo, umbral_medio):
    """Fondo de mapa de calor para celdas con porcentaje (0 a 1)."""
    c = primera_celda
    ws.conditional_formatting.add(rango, FormulaRule(formula=[f'AND(ISNUMBER({c}),{c}<={umbral_bajo})'], fill=_relleno(BAJO), font=Font(color=BAJO_TXT)))
    ws.conditional_formatting.add(rango, FormulaRule(formula=[f'AND(ISNUMBER({c}),{c}>{umbral_bajo},{c}<={umbral_medio})'], fill=_relleno(MEDIO), font=Font(color=MEDIO_TXT)))
    ws.conditional_formatting.add(rango, FormulaRule(formula=[f'AND(ISNUMBER({c}),{c}>{umbral_medio})'], fill=_relleno(ALTO), font=Font(color=ALTO_TXT)))


def generar_matriz_priorizacion(entidades, componentes, umbral_bajo=40, umbral_medio=75):
    """Devuelve un BytesIO con el libro Excel.

    entidades: lista de dicts con nombre, nit, codigo, sector, naturaleza, fecha (datetime),
      version, ima (0-100), componentes {clave: 0..1 o None} y, si existen, conteos
      {clave: {'si','parcial','no','na'}}.
    componentes: lista de (clave, nombre) en orden.
    """
    ub, um = umbral_bajo / 100, umbral_medio / 100
    def _orden(e):
        """Nivel de madurez de menor a mayor. Desempate: más componentes en nivel bajo;
        si persiste, el componente más débil más bajo; por último, el nombre."""
        valores = [v for v in e['componentes'].values() if v is not None]
        en_bajo = sum(1 for v in valores if v <= ub)
        return (round(e['ima'], 4) if e['ima'] is not None else 999, -en_bajo,
                min(valores) if valores else 999, e['nombre'])

    ordenadas = sorted(entidades, key=_orden)
    n = len(ordenadas)
    nc = len(componentes)
    hoy = datetime.now()

    wb = Workbook()
    resumen = wb.active
    resumen.title = 'Resumen'
    matriz = wb.create_sheet('Matriz de priorización')
    ficha = wb.create_sheet('Ficha por entidad')
    sector = wb.create_sheet('Por sector y naturaleza')
    respuestas = wb.create_sheet('Respuestas por componente')
    metodo = wb.create_sheet('Metodología')

    # ------------------------------------------------------------------ Matriz
    m = matriz
    cols = ['Prioridad', 'Entidad', 'NIT', 'Sector', 'Naturaleza jurídica', 'Fecha del diagnóstico', 'Formulario']
    for _, nombre in componentes:
        cols += [f'{nombre} (%)', f'{nombre} · nivel']
    cols += ['Nivel de madurez (%)', 'Clasificación', 'Prioridad de atención', 'Componente más débil',
             'Componentes en nivel bajo', 'Brecha hasta 100 %', 'Acompañamiento']
    total_cols = len(cols)
    _titulo(m, 'Matriz de priorización de entidades',
            f'Consejo Distrital de Archivos de Medellín · Diagnóstico Territorial Archivístico · '
            f'Diagnóstico más reciente de cada entidad · Generada el {hoy:%d/%m/%Y %H:%M}', total_cols)
    m['B5'] = 'Ordenada de menor a mayor nivel de madurez: las primeras filas son las entidades que requieren atención prioritaria. Use los filtros del encabezado para ver un sector, una clasificación o una prioridad.'
    m['B5'].font = _f(9, italic=True, color=GRIS)
    fila0 = PRIMERA_FILA
    for j, nombre in enumerate(cols, start=2):
        _encabezado(m.cell(row=fila0, column=j, value=nombre), AZUL if 'nivel' not in nombre.lower() or nombre.startswith('Nivel de madurez') else AZUL_MEDIO)
    m.row_dimensions[fila0].height = 46

    col_comp = {}  # clave -> (letra %, letra nivel)
    for i, e in enumerate(ordenadas):
        r = fila0 + 1 + i
        m.cell(row=r, column=2, value=i + 1)
        m.cell(row=r, column=3, value=e['nombre'])
        m.cell(row=r, column=4, value=e.get('nit') or '')
        m.cell(row=r, column=5, value=e.get('sector') or 'Sin sector')
        m.cell(row=r, column=6, value=e.get('naturaleza') or 'Sin dato')
        c = m.cell(row=r, column=7, value=e['fecha'].replace(tzinfo=None) if e.get('fecha') else None)
        c.number_format = 'dd/mm/yyyy'
        m.cell(row=r, column=8, value=e.get('version') or '')
        col = 9
        for clave, _ in componentes:
            v = e['componentes'].get(clave)
            pc = m.cell(row=r, column=col, value=round(v, 4) if v is not None else None)
            pc.number_format = '0.0%'
            letra = get_column_letter(col)
            m.cell(row=r, column=col + 1, value=_nivel_formula(f'{letra}{r}', ub, um))
            col_comp[clave] = (letra, get_column_letter(col + 1))
            col += 2
        letra_ima = get_column_letter(col)
        ic = m.cell(row=r, column=col, value=round((e['ima'] or 0) / 100, 4))
        ic.number_format = '0.0%'
        m.cell(row=r, column=col + 1, value=_nivel_formula(f'{letra_ima}{r}', ub, um))
        letra_clas = get_column_letter(col + 1)
        m.cell(row=r, column=col + 2, value=f'=IF({letra_clas}{r}="Bajo","Alta",IF({letra_clas}{r}="Medio","Media","Baja"))')
        letras_pct = [col_comp[k][0] for k, _ in componentes]
        # Componente más débil: el de menor porcentaje (si hay empate, el primero en el orden del modelo).
        validos = [(e['componentes'].get(k), nm) for k, nm in componentes if e['componentes'].get(k) is not None]
        m.cell(row=r, column=col + 3, value=min(validos, key=lambda x: x[0])[1] if validos else 'Sin dato')
        m.cell(row=r, column=col + 4, value='=' + '+'.join(f'({l}{r}<={ub})*ISNUMBER({l}{r})' for l in letras_pct))
        bc = m.cell(row=r, column=col + 5, value=f'=1-{letra_ima}{r}')
        bc.number_format = '0.0%'
        m.cell(row=r, column=col + 6, value=e.get('acompanamiento') or 'Sin iniciar')
        for j in range(2, total_cols + 2):
            cel = m.cell(row=r, column=j)
            cel.border = CAJA
            cel.font = _f(10)
            cel.alignment = Alignment(vertical='center', horizontal='left' if j in (3, 5, 6) else 'center', wrap_text=j in (3, 5, 6))
        m.cell(row=r, column=3).font = _f(10, True, AZUL)

    ultima = fila0 + max(n, 1)
    letra_fin = get_column_letter(total_cols + 1)
    col_ima = get_column_letter(9 + 2 * nc)
    col_clas = get_column_letter(10 + 2 * nc)
    col_prio = get_column_letter(11 + 2 * nc)
    col_debil = get_column_letter(12 + 2 * nc)
    if n:
        tabla = Table(displayName='TablaPriorizacion', ref=f'B{fila0}:{letra_fin}{ultima}')
        tabla.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=True)
        m.add_table(tabla)
        for k, _ in componentes:
            pct, niv = col_comp[k]
            _reglas_porcentaje(m, f'{pct}{fila0 + 1}:{pct}{ultima}', f'{pct}{fila0 + 1}', ub, um)
            _reglas_nivel(m, f'{niv}{fila0 + 1}:{niv}{ultima}')
        _reglas_nivel(m, f'{col_clas}{fila0 + 1}:{col_clas}{ultima}')
        m.conditional_formatting.add(f'{col_ima}{fila0 + 1}:{col_ima}{ultima}',
                                     DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='4F81BD', showValue=True))
        for texto, fondo, tinta in (('Alta', BAJO, BAJO_TXT), ('Media', MEDIO, MEDIO_TXT), ('Baja', ALTO, ALTO_TXT)):
            m.conditional_formatting.add(f'{col_prio}{fila0 + 1}:{col_prio}{ultima}',
                                         CellIsRule(operator='equal', formula=[f'"{texto}"'], fill=_relleno(fondo), font=Font(color=tinta, bold=True)))
    anchos = {2: 10, 3: 42, 4: 14, 5: 24, 6: 22, 7: 13, 8: 12}
    for j in range(2, total_cols + 2):
        m.column_dimensions[get_column_letter(j)].width = anchos.get(j, 13)
    m.column_dimensions[col_debil].width = 26
    m.column_dimensions[get_column_letter(total_cols + 1)].width = 16
    m.column_dimensions['A'].width = 2
    m.freeze_panes = f'D{fila0 + 1}'
    m.sheet_properties.tabColor = AZUL

    ref = lambda letra: f"'Matriz de priorización'!${letra}${fila0 + 1}:${letra}${ultima}"
    R_IMA, R_CLAS, R_SECT, R_NAT, R_ENT = ref(col_ima), ref(col_clas), ref('E'), ref('F'), ref('C')

    # ----------------------------------------------------------------- Resumen
    s = resumen
    _titulo(s, 'Resumen de la priorización',
            f'Diagnóstico Territorial Archivístico · {n} entidades diagnosticadas · Generado el {hoy:%d/%m/%Y %H:%M}', 12)
    kpis = [
        ('Entidades diagnosticadas', f'=COUNTA({R_ENT})', '0'),
        ('Nivel de madurez promedio', f'=IFERROR(AVERAGE({R_IMA}),0)', '0.0%'),
        ('Prioridad alta (nivel bajo)', f'=COUNTIF({R_CLAS},"Bajo")', '0'),
        ('Nivel medio', f'=COUNTIF({R_CLAS},"Medio")', '0'),
        ('Nivel alto', f'=COUNTIF({R_CLAS},"Alto")', '0'),
    ]
    colores_kpi = [AZUL, AZUL_MEDIO, BAJO_FUERTE, MEDIO_FUERTE, ALTO_FUERTE]
    for i, (etq, formula, fmt) in enumerate(kpis):
        c0 = 2 + i * 2
        s.merge_cells(start_row=6, start_column=c0, end_row=6, end_column=c0 + 1)
        s.merge_cells(start_row=7, start_column=c0, end_row=7, end_column=c0 + 1)
        v = s.cell(row=6, column=c0, value=formula)
        v.number_format = fmt
        v.font = _f(22, True, 'FFFFFF')
        v.alignment = Alignment(horizontal='center', vertical='center')
        t = s.cell(row=7, column=c0, value=etq)
        t.font = _f(9, True, 'FFFFFF')
        t.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for rr in (6, 7):
            for cc in (c0, c0 + 1):
                s.cell(row=rr, column=cc).fill = _relleno(colores_kpi[i])
    s.row_dimensions[6].height = 40
    s.row_dimensions[7].height = 30

    # Resultado por componente
    s['B10'] = 'Resultado por componente'
    s['B10'].font = _f(13, True, AZUL)
    for j, h in enumerate(['Componente', 'Promedio', 'Clasificación', 'Entidades en nivel bajo', '% en nivel bajo'], start=2):
        _encabezado(s.cell(row=11, column=j, value=h))
    for i, (clave, nombre) in enumerate(componentes):
        r = 12 + i
        pct = col_comp[clave][0]
        s.cell(row=r, column=2, value=nombre)
        c = s.cell(row=r, column=3, value=f'=IFERROR(AVERAGE({ref(pct)}),"")')
        c.number_format = '0.0%'
        s.cell(row=r, column=4, value=_nivel_formula(f'C{r}', ub, um))
        s.cell(row=r, column=5, value=f'=COUNTIFS({ref(pct)},"<="&{ub})')
        p = s.cell(row=r, column=6, value=f'=IFERROR(E{r}/COUNT({ref(pct)}),0)')
        p.number_format = '0.0%'
    fin_comp = 11 + nc
    tr = fin_comp + 1
    s.cell(row=tr, column=2, value='Nivel de madurez')
    c = s.cell(row=tr, column=3, value=f'=IFERROR(AVERAGE({R_IMA}),"")')
    c.number_format = '0.0%'
    s.cell(row=tr, column=4, value=_nivel_formula(f'C{tr}', ub, um))
    s.cell(row=tr, column=5, value=f'=COUNTIF({R_CLAS},"Bajo")')
    p = s.cell(row=tr, column=6, value=f'=IFERROR(E{tr}/COUNTA({R_ENT}),0)')
    p.number_format = '0.0%'
    for r in range(12, tr + 1):
        for j in range(2, 7):
            cel = s.cell(row=r, column=j)
            cel.border = CAJA
            cel.font = _f(10, r == tr)
            cel.alignment = Alignment(horizontal='left' if j == 2 else 'center', vertical='center')
            if r == tr:
                cel.fill = _relleno(FONDO)
    _reglas_nivel(s, f'D12:D{tr}')
    _reglas_porcentaje(s, f'C12:C{tr}', 'C12', ub, um)

    # Distribución por clasificación
    s['H10'] = 'Distribución por nivel de madurez'
    s['H10'].font = _f(13, True, AZUL)
    for j, h in enumerate(['Nivel', 'Rango', 'Entidades', '%'], start=8):
        _encabezado(s.cell(row=11, column=j, value=h))
    rangos = [('Bajo', f'0 – {umbral_bajo} %'), ('Medio', f'más de {umbral_bajo} – {umbral_medio} %'), ('Alto', f'más de {umbral_medio} – 100 %')]
    for i, (niv, rango) in enumerate(rangos):
        r = 12 + i
        s.cell(row=r, column=8, value=niv)
        s.cell(row=r, column=9, value=rango)
        s.cell(row=r, column=10, value=f'=COUNTIF({R_CLAS},"{niv}")')
        p = s.cell(row=r, column=11, value=f'=IFERROR(J{r}/SUM($J$12:$J$14),0)')
        p.number_format = '0.0%'
        for j in range(8, 12):
            s.cell(row=r, column=j).border = CAJA
            s.cell(row=r, column=j).font = _f(10)
            s.cell(row=r, column=j).alignment = Alignment(horizontal='center')
    _reglas_nivel(s, 'H12:H14')

    # Entidades con prioridad (las 10 primeras de la matriz)
    top = min(10, n)
    base_top = tr + 3
    s.cell(row=base_top, column=2, value='Entidades que requieren atención prioritaria')
    s.cell(row=base_top, column=2).font = _f(13, True, AZUL)
    s.cell(row=base_top + 1, column=2, value='Las diez entidades con menor nivel de madurez, tomadas de la matriz.').font = _f(9, italic=True, color=GRIS)
    enc = ['Prioridad', 'Entidad', 'Sector', 'Nivel de madurez', 'Clasificación', 'Componente más débil']
    for j, h in enumerate(enc, start=2):
        _encabezado(s.cell(row=base_top + 2, column=j, value=h))
    hoja = "'Matriz de priorización'"
    for i in range(top):
        r = base_top + 3 + i
        fm = fila0 + 1 + i
        vals = [f'={hoja}!B{fm}', f'={hoja}!C{fm}', f'={hoja}!E{fm}', f'={hoja}!{col_ima}{fm}', f'={hoja}!{col_clas}{fm}', f'={hoja}!{col_debil}{fm}']
        for j, v in enumerate(vals, start=2):
            cel = s.cell(row=r, column=j, value=v)
            cel.border = CAJA
            cel.font = _f(10, j == 3, AZUL if j == 3 else '1F2937')
            cel.alignment = Alignment(horizontal='left' if j in (3, 4, 7) else 'center', vertical='center', wrap_text=j == 3)
        s.cell(row=r, column=5).number_format = '0.0%'
    if top:
        _reglas_nivel(s, f'F{base_top + 3}:F{base_top + 2 + top}')

    for col, w in {'A': 2, 'B': 30, 'C': 44, 'D': 26, 'E': 16, 'F': 16, 'G': 30, 'H': 12, 'I': 18, 'J': 12, 'K': 10, 'L': 4}.items():
        s.column_dimensions[col].width = w

    # Gráficos del resumen
    g1 = BarChart()
    g1.type = 'bar'
    g1.style = 10
    g1.title = 'Promedio por componente'
    g1.y_axis.title = None
    g1.y_axis.number_format = '0%'
    g1.y_axis.scaling.min = 0
    g1.y_axis.scaling.max = 1
    g1.y_axis.majorGridlines = None
    g1.add_data(Reference(s, min_col=3, min_row=11, max_row=fin_comp), titles_from_data=True)
    g1.set_categories(Reference(s, min_col=2, min_row=12, max_row=fin_comp))
    g1.legend = None
    g1.series[0].graphicalProperties.solidFill = AZUL_MEDIO
    g1.dataLabels = _etiquetas('0.0%')
    g1.height, g1.width = 7.5, 15
    s.add_chart(g1, f'B{base_top + 5 + top}')

    g2 = DoughnutChart()
    g2.title = 'Entidades por nivel de madurez'
    g2.add_data(Reference(s, min_col=10, min_row=11, max_row=14), titles_from_data=True)
    g2.set_categories(Reference(s, min_col=8, min_row=12, max_row=14))
    for idx, color in enumerate((BAJO_FUERTE, MEDIO_FUERTE, ALTO_FUERTE)):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = color
        g2.series[0].dPt.append(pt)
    g2.dataLabels = _etiquetas()
    g2.holeSize = 55
    g2.height, g2.width = 7.5, 11
    s.add_chart(g2, f'E{base_top + 5 + top}')
    s.sheet_properties.tabColor = DORADO

    # -------------------------------------------------------- Ficha por entidad
    f = ficha
    _titulo(f, 'Ficha por entidad', 'Elija una entidad en la celda amarilla: los resultados y el gráfico se actualizan solos.', 6)
    f['B6'] = 'Entidad'
    f['B6'].font = _f(11, True, AZUL)
    f.merge_cells('C6:F6')
    f['C6'] = ordenadas[0]['nombre'] if n else ''
    f['C6'].fill = _relleno('FFF7D6')
    f['C6'].font = _f(12, True, AZUL)
    f['C6'].border = Border(bottom=Side(style='medium', color=DORADO))
    if n:
        dv = DataValidation(type='list', formula1=f"='Matriz de priorización'!$C${fila0 + 1}:$C${ultima}", allow_blank=False)
        dv.error, dv.errorTitle = 'Elija una entidad de la lista.', 'Entidad no válida'
        dv.prompt, dv.promptTitle = 'Despliegue la lista y elija la entidad.', 'Ficha por entidad'
        f.add_data_validation(dv)
        dv.add('C6')
    fila_ent = f"MATCH($C$6,'Matriz de priorización'!$C${fila0 + 1}:$C${ultima},0)"

    def dato(letra):
        return f"=IFERROR(INDEX('Matriz de priorización'!${letra}${fila0 + 1}:${letra}${ultima},{fila_ent}),\"\")"

    datos = [('NIT', 'D', None), ('Sector', 'E', None), ('Naturaleza jurídica', 'F', None),
             ('Fecha del diagnóstico', 'G', 'dd/mm/yyyy'), ('Prioridad en la matriz', 'B', '0'),
             ('Prioridad de atención', col_prio, None)]
    for i, (etq, letra, fmt) in enumerate(datos):
        r = 8 + i
        f.cell(row=r, column=2, value=etq).font = _f(10, True, GRIS)
        c = f.cell(row=r, column=3, value=dato(letra))
        c.font = _f(10)
        c.alignment = Alignment(horizontal='left')
        if fmt:
            c.number_format = fmt
    _reglas_nivel(f, 'C13')
    for texto, fondo, tinta in (('Alta', BAJO, BAJO_TXT), ('Media', MEDIO, MEDIO_TXT), ('Baja', ALTO, ALTO_TXT)):
        f.conditional_formatting.add('C13', CellIsRule(operator='equal', formula=[f'"{texto}"'], fill=_relleno(fondo), font=Font(color=tinta, bold=True)))

    base = 16
    for j, h in enumerate(['Componente', 'Resultado', 'Clasificación', 'Promedio del Distrito', 'Diferencia'], start=2):
        _encabezado(f.cell(row=base, column=j, value=h))
    for i, (clave, nombre) in enumerate(componentes):
        r = base + 1 + i
        pct = col_comp[clave][0]
        f.cell(row=r, column=2, value=nombre)
        c = f.cell(row=r, column=3, value=dato(pct)); c.number_format = '0.0%'
        f.cell(row=r, column=4, value=_nivel_formula(f'C{r}', ub, um))
        c = f.cell(row=r, column=5, value=f"=Resumen!C{12 + i}"); c.number_format = '0.0%'
        c = f.cell(row=r, column=6, value=f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),C{r}-E{r},"")'); c.number_format = '+0.0%;-0.0%;0.0%'
    rt = base + 1 + nc
    f.cell(row=rt, column=2, value='Nivel de madurez')
    c = f.cell(row=rt, column=3, value=dato(col_ima)); c.number_format = '0.0%'
    f.cell(row=rt, column=4, value=_nivel_formula(f'C{rt}', ub, um))
    c = f.cell(row=rt, column=5, value=f'=Resumen!C{tr}'); c.number_format = '0.0%'
    c = f.cell(row=rt, column=6, value=f'=IF(AND(ISNUMBER(C{rt}),ISNUMBER(E{rt})),C{rt}-E{rt},"")'); c.number_format = '+0.0%;-0.0%;0.0%'
    for r in range(base + 1, rt + 1):
        for j in range(2, 7):
            cel = f.cell(row=r, column=j)
            cel.border = CAJA
            cel.font = _f(10, r == rt)
            cel.alignment = Alignment(horizontal='left' if j == 2 else 'center')
            if r == rt:
                cel.fill = _relleno(FONDO)
    _reglas_nivel(f, f'D{base + 1}:D{rt}')
    _reglas_porcentaje(f, f'C{base + 1}:C{rt}', f'C{base + 1}', ub, um)
    f.conditional_formatting.add(f'F{base + 1}:F{rt}', CellIsRule(operator='lessThan', formula=['0'], font=Font(color=BAJO_TXT, bold=True)))
    f.conditional_formatting.add(f'F{base + 1}:F{rt}', CellIsRule(operator='greaterThan', formula=['0'], font=Font(color=ALTO_TXT, bold=True)))

    g3 = BarChart()
    g3.type = 'col'
    g3.style = 10
    g3.title = 'Entidad frente al promedio del Distrito'
    g3.y_axis.number_format = '0%'
    g3.y_axis.scaling.min = 0
    g3.y_axis.scaling.max = 1
    g3.add_data(Reference(f, min_col=3, max_col=3, min_row=base, max_row=rt), titles_from_data=True)
    g3.add_data(Reference(f, min_col=5, max_col=5, min_row=base, max_row=rt), titles_from_data=True)
    g3.set_categories(Reference(f, min_col=2, min_row=base + 1, max_row=rt))
    g3.series[0].graphicalProperties.solidFill = AZUL_MEDIO
    g3.series[1].graphicalProperties.solidFill = DORADO
    g3.legend.position = 'b'
    g3.height, g3.width = 8, 17
    f.add_chart(g3, f'B{rt + 3}')
    for col, w in {'A': 2, 'B': 30, 'C': 22, 'D': 16, 'E': 20, 'F': 14}.items():
        f.column_dimensions[col].width = w
    f.sheet_properties.tabColor = AZUL_MEDIO

    # ---------------------------------------------------- Por sector y naturaleza
    p = sector
    _titulo(p, 'Resultados por sector y por naturaleza jurídica',
            'Resúmenes calculados con fórmulas a partir de la matriz (equivalen a una tabla dinámica).', 6 + nc)

    def bloque(fila, titulo, rango_grupo, grupos):
        p.cell(row=fila, column=2, value=titulo).font = _f(13, True, AZUL)
        enc = ['Grupo', 'Entidades', 'Nivel de madurez promedio'] + [nm for _, nm in componentes] + ['Nivel bajo', 'Nivel medio', 'Nivel alto']
        for j, h in enumerate(enc, start=2):
            _encabezado(p.cell(row=fila + 1, column=j, value=h))
        p.row_dimensions[fila + 1].height = 42
        for i, g in enumerate(grupos):
            r = fila + 2 + i
            p.cell(row=r, column=2, value=g)
            p.cell(row=r, column=3, value=f'=COUNTIF({rango_grupo},B{r})')
            c = p.cell(row=r, column=4, value=f'=IFERROR(AVERAGEIFS({R_IMA},{rango_grupo},B{r}),"")'); c.number_format = '0.0%'
            for k, (clave, _) in enumerate(componentes):
                c = p.cell(row=r, column=5 + k, value=f'=IFERROR(AVERAGEIFS({ref(col_comp[clave][0])},{rango_grupo},B{r}),"")')
                c.number_format = '0.0%'
            for k, niv in enumerate(('Bajo', 'Medio', 'Alto')):
                p.cell(row=r, column=5 + nc + k, value=f'=COUNTIFS({rango_grupo},B{r},{R_CLAS},"{niv}")')
            for j in range(2, 8 + nc):
                cel = p.cell(row=r, column=j)
                cel.border = CAJA
                cel.font = _f(10)
                cel.alignment = Alignment(horizontal='left' if j == 2 else 'center', wrap_text=j == 2)
        fin = fila + 1 + len(grupos)
        if grupos:
            _reglas_porcentaje(p, f'D{fila + 2}:{get_column_letter(4 + nc)}{fin}', f'D{fila + 2}', ub, um)
        return fin

    sectores = sorted({e.get('sector') or 'Sin sector' for e in ordenadas})
    naturalezas = sorted({e.get('naturaleza') or 'Sin dato' for e in ordenadas})
    fin1 = bloque(6, 'Por sector', R_SECT, sectores)
    bloque(fin1 + 3, 'Por naturaleza jurídica', R_NAT, naturalezas)
    p.column_dimensions['A'].width = 2
    p.column_dimensions['B'].width = 40
    for j in range(3, 8 + nc):
        p.column_dimensions[get_column_letter(j)].width = 15
    p.sheet_properties.tabColor = AZUL_MEDIO

    # ----------------------------------------------- Respuestas por componente
    q = respuestas
    _titulo(q, 'Respuestas por componente',
            'Cuántas respuestas Sí, Parcialmente, No y N/A dio cada entidad en cada componente (cuestionario V04).', 3 + 4 * nc)
    q.merge_cells(start_row=6, start_column=2, end_row=7, end_column=2)
    _encabezado(q.cell(row=6, column=2, value='Entidad'))
    _encabezado(q.cell(row=7, column=2))
    for k, (clave, nombre) in enumerate(componentes):
        c0 = 3 + k * 4
        q.merge_cells(start_row=6, start_column=c0, end_row=6, end_column=c0 + 3)
        _encabezado(q.cell(row=6, column=c0, value=nombre), AZUL if k % 2 == 0 else AZUL_MEDIO)
        for j, h in enumerate(('Sí', 'Parcial', 'No', 'N/A')):
            _encabezado(q.cell(row=7, column=c0 + j, value=h), AZUL if k % 2 == 0 else AZUL_MEDIO)
            q.column_dimensions[get_column_letter(c0 + j)].width = 8
    fila_q = 8
    for e in ordenadas:
        conteos = e.get('conteos') or {}
        if not conteos:
            continue
        q.cell(row=fila_q, column=2, value=e['nombre']).font = _f(10, True, AZUL)
        for k, (clave, _) in enumerate(componentes):
            cc = conteos.get(clave, {})
            for j, llave in enumerate(('si', 'parcial', 'no', 'na')):
                cel = q.cell(row=fila_q, column=3 + k * 4 + j, value=cc.get(llave, 0))
                cel.alignment = Alignment(horizontal='center')
                cel.font = _f(10)
        for j in range(2, 3 + 4 * nc):
            q.cell(row=fila_q, column=j).border = CAJA
        fila_q += 1
    if fila_q == 8:
        q.cell(row=8, column=2, value='No hay diagnósticos con el cuestionario V04.').font = _f(10, italic=True, color=GRIS)
    else:
        for k in range(nc):
            c0 = 3 + k * 4
            for j, color in ((0, ALTO_TXT), (2, BAJO_TXT)):
                letra = get_column_letter(c0 + j)
                q.conditional_formatting.add(f'{letra}8:{letra}{fila_q - 1}', DataBarRule(start_type='num', start_value=0, end_type='max', color=color, showValue=True))
    q.column_dimensions['A'].width = 2
    q.column_dimensions['B'].width = 42
    q.freeze_panes = 'C8'

    # -------------------------------------------------------------- Metodología
    t = metodo
    _titulo(t, 'Metodología', 'Cómo se calculan los resultados de esta matriz.', 3)
    t.column_dimensions['A'].width = 2
    t.column_dimensions['B'].width = 30
    t.column_dimensions['C'].width = 100
    filas = [
        ('Fuente', 'Diagnóstico Territorial Archivístico del Consejo Distrital de Archivos de Medellín. Se toma el diagnóstico más reciente de cada entidad.'),
        ('Puntaje de cada respuesta', 'Sí = 1; Parcialmente = 0,5; No = 0. «N/A – No aplica» no se cuenta. Las preguntas de caracterización (C1 a C9) y las preguntas complementarias no suman puntaje.'),
        ('Resultado por componente', 'Suma de los puntos del componente dividida entre las preguntas que cuentan, expresada en porcentaje.'),
        ('Nivel de madurez', 'Promedio de los cinco componentes del Modelo de Gestión Documental y Administración de Archivos (MGDA): estratégico, administración de archivos, procesos de la gestión documental, tecnológico y cultural. Cada componente pesa 20 %.'),
        ('Clasificación (mapa de calor)', f'Bajo: 0 a {umbral_bajo} %. Medio: más de {umbral_bajo} hasta {umbral_medio} %. Alto: más de {umbral_medio} hasta 100 %. Se aplica igual a cada componente y al nivel de madurez.'),
        ('Prioridad de atención', 'Alta para las entidades en nivel bajo, media para las de nivel medio y baja para las de nivel alto. La columna «Prioridad» ordena a las entidades de menor a mayor nivel de madurez.'),
        ('Acompañamiento', 'Estado del acompañamiento técnico registrado en la herramienta: sin iniciar, en curso o cerrado.'),
        ('Desempate', 'Si dos entidades tienen el mismo nivel de madurez, va primero la que tenga más componentes en nivel bajo; si siguen empatadas, la que tenga el componente más débil con menor porcentaje.'),
        ('Componente más débil', 'El componente con menor porcentaje en la entidad. Si hay empate, se toma el primero en el orden del modelo.'),
        ('Componentes en nivel bajo', f'Cuántos componentes de la entidad están en {umbral_bajo} % o menos.'),
        ('Formulario', 'V04 corresponde al cuestionario vigente (72 preguntas). «Anterior» corresponde al formulario de 10 preguntas; sus resultados no son directamente comparables con los del V04.'),
        ('Tablas dinámicas', 'La matriz está en formato de tabla de Excel («TablaPriorizacion»): para crear una tabla dinámica propia, ubíquese en la matriz y use Insertar > Tabla dinámica. Las hojas «Resumen», «Ficha por entidad» y «Por sector y naturaleza» se recalculan solas con fórmulas.'),
        ('Alcance', 'Los resultados se calculan con las respuestas de las entidades y no constituyen una certificación de cumplimiento normativo; deben validarse con las evidencias en la revisión técnica.'),
    ]
    for i, (a, b) in enumerate(filas):
        r = 6 + i
        ca = t.cell(row=r, column=2, value=a)
        cb = t.cell(row=r, column=3, value=b)
        ca.font = _f(10, True, AZUL)
        cb.font = _f(10)
        for c in (ca, cb):
            c.alignment = Alignment(vertical='top', wrap_text=True)
            c.border = CAJA
        t.row_dimensions[r].height = 34

    for hoja in wb.worksheets:
        hoja.sheet_view.zoomScale = 90
        hoja.page_setup.orientation = 'landscape'
        hoja.page_setup.fitToWidth = 1
        hoja.page_setup.fitToHeight = 0
        hoja.sheet_properties.pageSetUpPr.fitToPage = True
    wb.active = 0
    # Que Excel calcule todas las fórmulas al abrir el archivo, aunque esté en cálculo manual.
    wb.calculation.calcMode = 'auto'
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    # Imprimir la ficha solo con la entidad elegida: área de impresión de la ficha.
    ficha.print_area = f'B2:F{rt + 20}'

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    return salida
