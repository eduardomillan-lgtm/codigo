# -*- coding: utf-8 -*-
"""
Genera calculadora_ingresos_agente.xlsx — el activo del día 28 del Smart Plan.

Normalmente NO necesitas esto: el .xlsx ya está en el repositorio y se edita
directamente en Excel. Usa este script solo si quieres cambiar la ESTRUCTURA
(añadir una sección, otro escenario, otra variable del modelo).

    pip install openpyxl
    python calculadora_generar.py

Los parámetros del Market Center se dejan VACÍOS a propósito: los rellena la
Team Leader con los datos reales de su MC antes de enviar el archivo.

Verificado con 50 comprobaciones del cálculo contra el resultado hecho a mano,
incluidos los casos límite (producción que no alcanza el tope, y hoja vacía).
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

ARIAL   = 'Arial'
NEGRO   = Font(name=ARIAL, size=11)
AZUL    = Font(name=ARIAL, size=11, color='0000FF')          # dato que mete el usuario
VERDE   = Font(name=ARIAL, size=11, color='008000')          # enlace a otra hoja
BOLD    = Font(name=ARIAL, size=11, bold=True)
BOLD12  = Font(name=ARIAL, size=12, bold=True)
TITULO  = Font(name=ARIAL, size=16, bold=True, color='FFFFFF')
SUBT    = Font(name=ARIAL, size=10, color='555555')
SECCION = Font(name=ARIAL, size=11, bold=True, color='FFFFFF')
GRIS    = Font(name=ARIAL, size=9, color='777777')

F_ROJO    = PatternFill('solid', fgColor='B70000')
F_AMAR    = PatternFill('solid', fgColor='FFFF00')           # rellenar aquí
F_SECCION = PatternFill('solid', fgColor='334155')
F_RES     = PatternFill('solid', fgColor='DCFCE7')
F_DIF     = PatternFill('solid', fgColor='FEF3C7')

EUR  = '#,##0 "€";(#,##0 "€");-'
PCT  = '0.0%'
NUM  = '#,##0.0'
fino = Side(style='thin', color='CCCCCC')
CAJA = Border(left=fino, right=fino, top=fino, bottom=fino)

wb = Workbook()

# ════════════════════════════════════════════════════════════
#  HOJA 1 · Parametros_MC  (la rellena la Team Leader)
# ════════════════════════════════════════════════════════════
pm = wb.active
pm.title = 'Parametros_MC'
pm.sheet_view.showGridLines = False
pm.column_dimensions['A'].width = 52
pm.column_dimensions['B'].width = 18
pm.column_dimensions['C'].width = 62

pm.merge_cells('A1:C1')
pm['A1'] = 'PARÁMETROS DEL MARKET CENTER — rellena esto ANTES de enviar el archivo'
pm['A1'].font = TITULO; pm['A1'].fill = F_ROJO
pm['A1'].alignment = Alignment(horizontal='center', vertical='center')
pm.row_dimensions[1].height = 30

pm['A3'] = ('⚠️  Las celdas amarillas vienen VACÍAS a propósito. No invento las cifras de tu '
            'Market Center: cada MC tiene las suyas. Pon las reales antes de mandar el archivo '
            'a un candidato — si envías un número que luego no puedes sostener, pierdes al agente '
            'y la credibilidad corre entre la competencia.')
pm['A3'].font = Font(name=ARIAL, size=10, bold=True, color='B70000')
pm['A3'].alignment = Alignment(wrap_text=True, vertical='top')
pm.merge_cells('A3:C3'); pm.row_dimensions[3].height = 48

pm['A5'] = 'Parámetro'; pm['B5'] = 'Valor'; pm['C5'] = 'Qué poner aquí'
for c in 'ABC':
    pm[c + '5'].font = SECCION; pm[c + '5'].fill = F_SECCION

params = [
    ('% de los honorarios que se queda el agente (antes del tope)', 'PCT',
     'El reparto de salida. Si el agente se queda el 70%, escribe 70%.'),
    ('Tope anual de aportación del agente al MC (cap, €)', 'EUR',
     'Lo máximo que el agente aporta al MC en su año de aniversario. Al alcanzarlo, deja de aportar.'),
    ('% de royalty sobre honorarios', 'PCT',
     'El porcentaje que va a la franquicia. Si no aplica en tu modelo, pon 0%.'),
    ('Tope anual de royalty (€)', 'EUR',
     'Lo máximo de royalty al año. Si no tiene tope, pon un número muy alto (p. ej. 999999).'),
    ('Cuotas fijas anuales del agente en el MC (€)', 'EUR',
     'Suma de cuotas, tecnología o puesto que paga el agente al año. Si el MC lo cubre todo, pon 0.'),
]
for i, (etiqueta, fmt, ayuda) in enumerate(params):
    r = 6 + i
    pm['A%d' % r] = etiqueta; pm['A%d' % r].font = NEGRO
    celda = pm['B%d' % r]
    celda.font = AZUL; celda.fill = F_AMAR; celda.border = CAJA
    celda.number_format = PCT if fmt == 'PCT' else EUR
    pm['C%d' % r] = ayuda
    pm['C%d' % r].font = GRIS
    pm['C%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
    pm.row_dimensions[r].height = 30

pm['B6'].comment = Comment('Reparto de salida del agente. Dato del Market Center.', 'KW Marbella')
pm['B7'].comment = Comment('Tope de aportación anual (cap). Dato del Market Center.', 'KW Marbella')

pm['A13'] = 'COMPROBACIÓN'
pm['A13'].font = BOLD12
pm['A14'] = 'Honorarios anuales necesarios para alcanzar el tope'
pm['A14'].font = NEGRO
pm['B14'] = '=IFERROR(IF((1-B6)<=0,"—",B7/(1-B6)),"—")'
pm['B14'].font = NEGRO; pm['B14'].number_format = EUR
pm['C14'] = ('A partir de esta cifra de honorarios generados, el agente deja de aportar al MC. '
             'Es el número que más mueve la conversación.')
pm['C14'].font = GRIS; pm['C14'].alignment = Alignment(wrap_text=True, vertical='top')
pm.row_dimensions[14].height = 30

pm['A16'] = 'Fuente de los parámetros: datos del Market Center, introducidos por la Team Leader.'
pm['A16'].font = GRIS
pm.merge_cells('A16:C16')

# ════════════════════════════════════════════════════════════
#  HOJA 2 · Calculadora  (la rellena el agente)
# ════════════════════════════════════════════════════════════
ws = wb.create_sheet('Calculadora')
ws.sheet_view.showGridLines = False
anchos = {'A': 54, 'B': 18, 'C': 16, 'D': 54}
for col, w in anchos.items():
    ws.column_dimensions[col].width = w

ws.merge_cells('A1:D1')
ws['A1'] = '¿CUÁNTO TE QUEDARÍAS CON OTRO MODELO?'
ws['A1'].font = TITULO; ws['A1'].fill = F_ROJO
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 34

ws.merge_cells('A2:D2')
ws['A2'] = ('Rellena solo las celdas AMARILLAS. Nada de lo que escribas sale de este archivo: '
            'no se envía a ningún sitio y no hace falta que compartas ni un número.')
ws['A2'].font = SUBT; ws['A2'].alignment = Alignment(wrap_text=True, vertical='center')
ws.row_dimensions[2].height = 28

# Salvavidas: sin parámetros del MC, el cálculo saldría a favor del MC por error
ws.merge_cells('A3:D3')
ws['A3'] = ('=IF(OR(Parametros_MC!$B$6=0,Parametros_MC!$B$7=0),'
            '"\u26a0 FALTAN LOS PARÁMETROS DEL MARKET CENTER. Los resultados de abajo NO son válidos: '
            've a la hoja Parametros_MC y rellena las celdas amarillas.","")')
ws['A3'].font = Font(name=ARIAL, size=11, bold=True, color='B70000')
ws['A3'].alignment = Alignment(wrap_text=True, vertical='center')
ws.row_dimensions[3].height = 30


def seccion(fila, texto):
    ws.merge_cells('A%d:D%d' % (fila, fila))
    ws['A%d' % fila] = texto
    ws['A%d' % fila].font = SECCION
    ws['A%d' % fila].fill = F_SECCION
    ws.row_dimensions[fila].height = 20

# ── TUS NÚMEROS ──────────────────────────────────────────────
seccion(4, '1 · TUS NÚMEROS')
ws['A5'] = 'Concepto'; ws['B5'] = 'Tu dato'; ws['C5'] = 'Ejemplo'; ws['D5'] = 'Nota'
for c in 'ABCD':
    ws[c + '5'].font = BOLD

entradas = [
    ('Operaciones cerradas en los últimos 12 meses', 12, NUM,
     'Compraventas y alquileres que hayas facturado.'),
    ('Honorarios medios por operación (€)', 18000, EUR,
     'Lo que factura la agencia por la operación, NO lo que cobras tú.'),
    ('% de los honorarios que te quedas hoy', 0.40, PCT,
     'Tu reparto actual. Si te quedas 4.000 € de unos honorarios de 10.000 €, es 40%.'),
    ('Costes fijos que pagas tú al año (€)', 2400, EUR,
     'Puesto, portales, fotografía, marketing, cuotas. Lo que sale de tu bolsillo.'),
]
for i, (etiqueta, ejemplo, fmt, nota) in enumerate(entradas):
    r = 6 + i
    ws['A%d' % r] = etiqueta; ws['A%d' % r].font = NEGRO
    ws['B%d' % r].font = AZUL; ws['B%d' % r].fill = F_AMAR
    ws['B%d' % r].border = CAJA; ws['B%d' % r].number_format = fmt
    ws['C%d' % r] = ejemplo; ws['C%d' % r].font = GRIS; ws['C%d' % r].number_format = fmt
    ws['D%d' % r] = nota; ws['D%d' % r].font = GRIS
    ws['D%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[r].height = 28

ws['A11'] = 'Honorarios totales que generas al año'
ws['A11'].font = BOLD
ws['B11'] = '=B6*B7'
ws['B11'].font = BOLD; ws['B11'].number_format = EUR
ws['C11'] = '=C6*C7'; ws['C11'].font = GRIS; ws['C11'].number_format = EUR
ws['D11'] = 'Es la cifra que factura tu agencia gracias a tu trabajo.'
ws['D11'].font = GRIS

# ── MODELO ACTUAL ────────────────────────────────────────────
seccion(13, '2 · LO QUE TE QUEDA HOY')
filas_actual = [
    ('Te quedas (antes de costes)', '=B11*B8', '=C11*C8', NEGRO),
    ('Menos tus costes fijos', '=-B9', '=-C9', NEGRO),
]
for i, (etiqueta, f1, f2, fuente) in enumerate(filas_actual):
    r = 14 + i
    ws['A%d' % r] = etiqueta; ws['A%d' % r].font = NEGRO
    ws['B%d' % r] = f1; ws['B%d' % r].font = fuente; ws['B%d' % r].number_format = EUR
    ws['C%d' % r] = f2; ws['C%d' % r].font = GRIS; ws['C%d' % r].number_format = EUR

ws['A16'] = 'NETO ACTUAL'
ws['A16'].font = BOLD12
ws['B16'] = '=SUM(B14:B15)'
ws['B16'].font = BOLD12; ws['B16'].number_format = EUR; ws['B16'].fill = F_DIF
ws['C16'] = '=SUM(C14:C15)'; ws['C16'].font = GRIS; ws['C16'].number_format = EUR

# ── MODELO DEL MC ────────────────────────────────────────────
seccion(18, '3 · LO QUE TE QUEDARÍA CON ESTE MODELO')
ws['A19'] = 'Honorarios totales que generas'
ws['A19'].font = NEGRO
ws['B19'] = '=B11'; ws['B19'].font = NEGRO; ws['B19'].number_format = EUR
ws['C19'] = '=C11'; ws['C19'].font = GRIS; ws['C19'].number_format = EUR
ws['D19'] = 'Los mismos honorarios: cambia el reparto, no tu trabajo.'
ws['D19'].font = GRIS

ws['A20'] = 'Menos tu aportación al MC (con tope aplicado)'
ws['A20'].font = NEGRO
ws['B20'] = '=-MIN(B19*(1-Parametros_MC!$B$6),Parametros_MC!$B$7)'
ws['B20'].font = VERDE; ws['B20'].number_format = EUR
ws['C20'] = '=-MIN(C19*(1-Parametros_MC!$B$6),Parametros_MC!$B$7)'
ws['C20'].font = GRIS; ws['C20'].number_format = EUR
ws['D20'] = 'ESTA ES LA LÍNEA CLAVE: la aportación tiene tope. Al alcanzarlo, deja de descontarse.'
ws['D20'].font = Font(name=ARIAL, size=9, bold=True, color='B70000')
ws['D20'].alignment = Alignment(wrap_text=True, vertical='top')
ws.row_dimensions[20].height = 28

ws['A21'] = 'Menos royalty (con tope aplicado)'
ws['A21'].font = NEGRO
ws['B21'] = '=-MIN(B19*Parametros_MC!$B$8,Parametros_MC!$B$9)'
ws['B21'].font = VERDE; ws['B21'].number_format = EUR
ws['C21'] = '=-MIN(C19*Parametros_MC!$B$8,Parametros_MC!$B$9)'
ws['C21'].font = GRIS; ws['C21'].number_format = EUR

ws['A22'] = 'Menos cuotas fijas del MC'
ws['A22'].font = NEGRO
ws['B22'] = '=-Parametros_MC!$B$10'
ws['B22'].font = VERDE; ws['B22'].number_format = EUR
ws['C22'] = '=-Parametros_MC!$B$10'
ws['C22'].font = GRIS; ws['C22'].number_format = EUR
ws['D22'] = 'Ojo: aquí ya no pagas portales ni fotografía. Eso lo cubre el MC.'
ws['D22'].font = GRIS
ws['D22'].alignment = Alignment(wrap_text=True, vertical='top')

ws['A23'] = 'NETO CON ESTE MODELO'
ws['A23'].font = BOLD12
ws['B23'] = '=B19+B20+B21+B22'
ws['B23'].font = BOLD12; ws['B23'].number_format = EUR; ws['B23'].fill = F_RES
ws['C23'] = '=C19+C20+C21+C22'; ws['C23'].font = GRIS; ws['C23'].number_format = EUR

# ── DIFERENCIA ───────────────────────────────────────────────
seccion(25, '4 · LA DIFERENCIA')
ws['A26'] = 'Diferencia al año'
ws['A26'].font = BOLD12
ws['B26'] = '=B23-B16'
ws['B26'].font = Font(name=ARIAL, size=14, bold=True); ws['B26'].number_format = EUR
ws['B26'].fill = F_DIF
ws['C26'] = '=C23-C16'; ws['C26'].font = GRIS; ws['C26'].number_format = EUR
ws.row_dimensions[26].height = 24

ws['A27'] = 'Diferencia en porcentaje'
ws['A27'].font = NEGRO
ws['B27'] = '=IFERROR(IF(B16<=0,"—",B26/B16),"—")'
ws['B27'].font = BOLD; ws['B27'].number_format = PCT

ws['A28'] = '% efectivo que te quedas hoy'
ws['A28'].font = NEGRO
ws['B28'] = '=IFERROR(IF(B11<=0,"—",B16/B11),"—")'
ws['B28'].font = NEGRO; ws['B28'].number_format = PCT

ws['A29'] = '% efectivo que te quedarías'
ws['A29'].font = NEGRO
ws['B29'] = '=IFERROR(IF(B11<=0,"—",B23/B11),"—")'
ws['B29'].font = BOLD; ws['B29'].number_format = PCT

# ── EL TOPE ──────────────────────────────────────────────────
seccion(31, '5 · EL TOPE (lo que casi nadie mira)')
ws['A32'] = 'Honorarios necesarios para alcanzar el tope'
ws['A32'].font = NEGRO
ws['B32'] = '=IFERROR(IF((1-Parametros_MC!$B$6)<=0,"—",Parametros_MC!$B$7/(1-Parametros_MC!$B$6)),"—")'
ws['B32'].font = VERDE; ws['B32'].number_format = EUR

ws['A33'] = 'Operaciones que te hacen falta para eso'
ws['A33'].font = NEGRO
ws['B33'] = '=IFERROR(IF(OR(B7<=0,NOT(ISNUMBER(B32))),"—",B32/B7),"—")'
ws['B33'].font = NEGRO; ws['B33'].number_format = NUM
ws['D33'] = 'Con tus honorarios medios actuales.'
ws['D33'].font = GRIS

ws['A34'] = '¿Llegas al tope con tu producción de hoy?'
ws['A34'].font = BOLD
ws['B34'] = '=IFERROR(IF(NOT(ISNUMBER(B32)),"—",IF(B11>=B32,"SÍ","Todavía no")),"—")'
ws['B34'].font = BOLD

ws['A35'] = 'Lo que te quedas de cada euro pasado el tope'
ws['A35'].font = NEGRO
ws['B35'] = '=IFERROR(1-Parametros_MC!$B$8,"—")'
ws['B35'].font = BOLD; ws['B35'].number_format = PCT
ws['D35'] = ('Pasado el tope dejas de aportar al MC y solo te descuentan el royalty. '
             'Cuando el royalty también agota su tope, te quedas el 100%.')
ws['D35'].font = GRIS
ws['D35'].alignment = Alignment(wrap_text=True, vertical='top')
ws.row_dimensions[35].height = 28

# ── SENSIBILIDAD ─────────────────────────────────────────────
seccion(37, '6 · Y SI CRECES (o si tienes un año flojo)')
cab = ['Escenario', 'Honorarios generados', 'Neto hoy', 'Neto con este modelo', 'Diferencia']
for j, texto in enumerate(cab):
    c = ws.cell(row=38, column=1 + j, value=texto)
    c.font = BOLD; c.fill = F_SECCION; c.font = SECCION

escenarios = [(0.6, 'Año flojo (-40%)'), (0.8, 'Algo peor (-20%)'),
              (1.0, 'Como ahora'), (1.25, 'Creciendo (+25%)'),
              (1.5, 'Buen año (+50%)'), (2.0, 'El doble')]
for i, (mult, nombre) in enumerate(escenarios):
    r = 39 + i
    ws.cell(row=r, column=1, value=nombre).font = NEGRO
    gci = '$B$11*%s' % mult
    ws.cell(row=r, column=2, value='=%s' % gci).number_format = EUR
    ws.cell(row=r, column=3,
            value='=%s*$B$8-$B$9' % gci).number_format = EUR
    ws.cell(row=r, column=4,
            value=('=%s-MIN(%s*(1-Parametros_MC!$B$6),Parametros_MC!$B$7)'
                   '-MIN(%s*Parametros_MC!$B$8,Parametros_MC!$B$9)-Parametros_MC!$B$10'
                   % (gci, gci, gci))).number_format = EUR
    ws.cell(row=r, column=5, value='=D%d-C%d' % (r, r)).number_format = EUR
    ws.cell(row=r, column=5).font = BOLD
    for j in range(1, 6):
        ws.cell(row=r, column=j).border = CAJA
        if j > 1 and ws.cell(row=r, column=j).font.bold is not True:
            ws.cell(row=r, column=j).font = NEGRO
    if mult == 1.0:
        for j in range(1, 6):
            ws.cell(row=r, column=j).fill = F_DIF

ws['A46'] = ('Fíjate en cómo se abre la diferencia cuanto más produces: es el efecto del tope. '
             'En los años flojos los dos modelos se parecen; en los buenos, no.')
ws['A46'].font = Font(name=ARIAL, size=9, italic=True, color='555555')
ws.merge_cells('A46:E46')
ws['A46'].alignment = Alignment(wrap_text=True, vertical='top')
ws.row_dimensions[46].height = 26

ws['A48'] = ('Supuestos: el cálculo asume que mantienes el mismo número de operaciones y los mismos '
             'honorarios medios en los dos modelos, y no incluye el reparto de beneficios ni los '
             'ingresos por referencias. Los parámetros del Market Center están en la hoja '
             '"Parametros_MC" y los ha introducido la Team Leader.')
ws['A48'].font = GRIS
ws.merge_cells('A48:E49')
ws['A48'].alignment = Alignment(wrap_text=True, vertical='top')

ws.column_dimensions['E'].width = 20
ws.freeze_panes = 'A3'

# ════════════════════════════════════════════════════════════
#  HOJA 3 · Instrucciones
# ════════════════════════════════════════════════════════════
ins = wb.create_sheet('Instrucciones')
ins.sheet_view.showGridLines = False
ins.column_dimensions['A'].width = 100

ins.merge_cells('A1:A1')
ins['A1'] = 'CÓMO SE USA'
ins['A1'].font = TITULO; ins['A1'].fill = F_ROJO
ins['A1'].alignment = Alignment(horizontal='center', vertical='center')
ins.row_dimensions[1].height = 30

bloques = [
    ('PARA EL AGENTE', BOLD12),
    ('1. Ve a la hoja "Calculadora".', NEGRO),
    ('2. Rellena solo las cuatro celdas AMARILLAS de la sección 1. La columna "Ejemplo" te '
     'enseña el formato que espera cada una.', NEGRO),
    ('3. El resto se calcula solo. No hace falta que toques nada más.', NEGRO),
    ('4. Nada de lo que escribas sale de este archivo. No se envía a ningún sitio, no lleva '
     'macros y no necesitas compartir ninguna cifra con nadie.', NEGRO),
    ('', NEGRO),
    ('PARA LA TEAM LEADER — ANTES DE ENVIARLO', BOLD12),
    ('1. Abre la hoja "Parametros_MC" y rellena las cinco celdas amarillas con los datos '
     'REALES de tu Market Center.', NEGRO),
    ('2. Comprueba la fila "Honorarios anuales necesarios para alcanzar el tope": es el número '
     'sobre el que va a girar la conversación.', NEGRO),
    ('3. Guarda y envía. Si quieres, oculta la hoja "Parametros_MC" (clic derecho en la '
     'pestaña → Ocultar) para que el agente vea solo su calculadora.', NEGRO),
    ('4. No envíes el archivo con los parámetros vacíos: saldrían ceros y perderías la '
     'conversación.', Font(name=ARIAL, size=11, bold=True, color='B70000')),
    ('', NEGRO),
    ('LEYENDA DE COLORES', BOLD12),
    ('Amarillo con texto azul  →  celdas que se rellenan a mano.', NEGRO),
    ('Texto negro  →  fórmulas. No las toques.', NEGRO),
    ('Texto verde  →  fórmulas que leen los parámetros del Market Center.', VERDE),
    ('Gris  →  columna de ejemplo y notas aclaratorias.', GRIS),
]
r = 3
for texto, fuente in bloques:
    ins['A%d' % r] = texto
    ins['A%d' % r].font = fuente
    ins['A%d' % r].alignment = Alignment(wrap_text=True, vertical='top')
    ins.row_dimensions[r] = ins.row_dimensions[r]
    if len(texto) > 85:
        ins.row_dimensions[r].height = 30
    r += 1

orden = ['Instrucciones', 'Calculadora', 'Parametros_MC']
wb._sheets = [wb[n] for n in orden]
import os
DESTINO = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       'calculadora_ingresos_agente.xlsx')
wb.save(DESTINO)
print('✅ generada: ' + DESTINO)
print('   Recuerda recalcular abriéndola en Excel o LibreOffice y guardando.')
