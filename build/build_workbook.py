"""
Generador del Sistema Financiero Empresarial (Excel).

Reconstruye el libro completo desde cero a partir de este script.
No editar el .xlsx generado directamente para cambios estructurales:
editar aqui y volver a ejecutar, asi el diseno queda versionado y reproducible.

Uso:
    python build_workbook.py [ruta_salida.xlsx]
"""

import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle, Protection
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties

# ---------------------------------------------------------------------------
# Estilos / constantes visuales
# ---------------------------------------------------------------------------

COLOR_CONFIG = "1F4E78"   # azul oscuro - capa configuracion
COLOR_DATA = "2E7D32"     # verde - capa datos
COLOR_REPORT = "B7791F"   # dorado - capa reportes
COLOR_SYSTEM = "5A5A5A"   # gris - sistema (LOG, MENU)

FILL_HEADER_CONFIG = PatternFill("solid", fgColor=COLOR_CONFIG)
FILL_HEADER_DATA = PatternFill("solid", fgColor=COLOR_DATA)
FILL_HEADER_REPORT = PatternFill("solid", fgColor=COLOR_REPORT)
FILL_HEADER_SYSTEM = PatternFill("solid", fgColor=COLOR_SYSTEM)

FONT_HEADER = Font(color="FFFFFF", bold=True, size=10, name="Calibri")
FONT_TITLE = Font(color=COLOR_CONFIG, bold=True, size=16, name="Calibri")
FONT_SUBTITLE = Font(color="666666", italic=True, size=10, name="Calibri")

THIN = Side(style="thin", color="D9D9D9")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CURRENCY_FMT = '"$" #,##0'
# Para columnas de formula (Total) que devuelven 0 en filas sin datos: oculta el
# cero para que no ensucie visualmente filas vacias, pero mantiene un valor
# NUMERICO real (nunca texto ""), indispensable para que SUMPRODUCT en Dashboard
# y Reportes no falle con #VALOR! al multiplicar la columna completa.
CURRENCY_FMT_CALC = '"$" #,##0;-"$" #,##0;""'
DATE_FMT = "dd/mm/yyyy"
PCT_FMT = "0.00%"
DEFAULT_REPORT_YEAR = date.today().year

TABLE_STYLE = TableStyleInfo(
    name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False,
    showLastColumn=False, showColumnStripes=False,
)

# ---------------------------------------------------------------------------
# Registro de tablas: fila inicial/final de datos y encabezados por hoja.
#
# IMPORTANTE - por que no se usan referencias estructuradas de Excel (Tabla[Columna]
# o [@Columna]) en ninguna formula: al probar el archivo generado (abriendolo con
# Excel via COM y forzando recalculo completo), CUALQUIER formula de celda que
# contenga una referencia estructurada se corrompe a #REF! (o incluso impide abrir
# el archivo) cuando el libro se genera con openpyxl. Openpyxl no reproduce los
# metadatos internos que Excel necesita para resolver "Tabla[Columna]" dentro de
# una formula de celda. Las Tablas de Excel (Insertar > Tabla) SI se usan para
# estilo/autoexpansion/filtrado; simplemente no se referencian por nombre en
# formulas, que en su lugar usan direcciones de rango absolutas (hoja!$X$5:$X$64),
# calculadas aqui a partir de este registro. Esto es 100% compatible con Excel 2019.
# ---------------------------------------------------------------------------

SHEET_LAYOUT = {}

# Las fórmulas y los desplegables analizan hasta esta cantidad de filas.
# Las tablas comienzan con pocas filas y se expanden al registrar datos; separar
# ambas capacidades evita prellenar miles de filas vacías en el archivo.
ANALYTIC_ROW_CAPACITY = {
    "EMPRESA": 200,
    "PARAMETROS": 500,
    "CATEGORIAS": 1000,
    "IMPUESTOS": 1000,
    "TIPOS_DOCUMENTOS": 1000,
    "ESTADOS": 1000,
    "MEDIOS_PAGO": 1000,
    "USUARIOS": 1000,
    "CUENTAS": 500,
    "TERCEROS": 2000,
    "DOCUMENTOS": 10000,
    "MOVIMIENTOS": 10000,
    "PAGOS": 10000,
    "COBROS": 10000,
    "LOG": 20000,
    "REGLAS_IA": 100,
}


def register_layout(sheet_name, headers, first_data_row, last_data_row):
    SHEET_LAYOUT[sheet_name] = {
        "headers": headers,
        "first": first_data_row,
        "last": last_data_row,
    }


def colletter(sheet_name, colname):
    headers = SHEET_LAYOUT[sheet_name]["headers"]
    return get_column_letter(headers.index(colname) + 1)


def col_range(sheet_name, colname, sheet_prefix=True):
    """Rango analítico absoluto, ej. MOVIMIENTOS!$O$5:$O$10004."""
    info = SHEET_LAYOUT[sheet_name]
    letter = colletter(sheet_name, colname)
    rng = f"${letter}${info['first']}:${letter}${info['last']}"
    return f"{sheet_name}!{rng}" if sheet_prefix else rng


def data_col_range(sheet_name, colname):
    """Rango relativo para validación de datos y formato condicional."""
    info = SHEET_LAYOUT[sheet_name]
    letter = colletter(sheet_name, colname)
    return f"{letter}{info['first']}:{letter}{info['last']}"

def col_first_cell(sheet_name, colname, sheet_prefix=True):
    info = SHEET_LAYOUT[sheet_name]
    letter = colletter(sheet_name, colname)
    ref = f"${letter}${info['first']}"
    return f"{sheet_name}!{ref}" if sheet_prefix else ref


def cell_ref(sheet_name, colname, row, sheet_prefix=False, absolute_col=True):
    """Celda de una fila concreta, ej. cell_ref('MOVIMIENTOS','Total',5) -> 'O5'."""
    letter = colletter(sheet_name, colname)
    prefix = "$" if absolute_col else ""
    ref = f"{prefix}{letter}{row}"
    return f"{sheet_name}!{ref}" if sheet_prefix else ref


def style_header_row(ws, row, ncols, fill):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def autosize(ws, headers, min_width=12, max_width=40):
    for i, h in enumerate(headers, start=1):
        width = max(min_width, min(max_width, len(str(h)) + 6))
        ws.column_dimensions[get_column_letter(i)].width = width


def add_table(ws, name, ref, headers, header_row=1):
    style_header_row(ws, header_row, len(headers), FILL_HEADER_CONFIG)
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)
    autosize(ws, headers)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate

    table_last_row = int("".join(ch for ch in ref.split(":")[1] if ch.isdigit()))
    capacity = ANALYTIC_ROW_CAPACITY.get(ws.title, table_last_row - header_row)
    analytic_last_row = max(table_last_row, header_row + capacity)
    register_layout(ws.title, headers, header_row + 1, analytic_last_row)


def sheet_title(ws, title, subtitle=None):
    ws["A1"] = title
    ws["A1"].font = FONT_TITLE
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = FONT_SUBTITLE


# ---------------------------------------------------------------------------
# Construccion del libro
# ---------------------------------------------------------------------------

def build(out_path: Path):
    wb = Workbook()
    wb.remove(wb.active)

    ws_menu = build_menu(wb)

    ws_empresa = build_empresa(wb)
    ws_parametros = build_parametros(wb)
    ws_categorias = build_categorias(wb)
    ws_impuestos = build_impuestos(wb)
    ws_tipos_doc = build_tipos_documentos(wb)
    ws_estados = build_estados(wb)
    ws_medios_pago = build_medios_pago(wb)
    ws_usuarios = build_usuarios_roles(wb)

    ws_terceros = build_terceros(wb)
    ws_documentos = build_documentos(wb)
    ws_movimientos = build_movimientos(wb)
    ws_pagos = build_pagos(wb)
    ws_cobros = build_cobros(wb)
    # CUENTAS va despues de MOVIMIENTOS: su SaldoActual referencia el rango de
    # MOVIMIENTOS, que debe estar registrado en SHEET_LAYOUT antes de calcularlo.
    ws_cuentas = build_cuentas(wb)

    build_asistente(wb)
    build_reglas_ia(wb)

    build_dashboard(wb)
    build_ingresos_gastos(wb, "INGRESOS", "Ingreso", "Venta")
    build_ingresos_gastos(wb, "GASTOS", "Gasto", "Compra")
    build_flujo_caja(wb)
    build_cxc_cxp(wb, "CXC", "Ingreso", "Venta")
    build_cxc_cxp(wb, "CXP", "Gasto", "Compra")
    build_impuestos_reporte(wb)
    build_reporte_periodo(wb, "REPORTE_MENSUAL", "MES")
    build_reporte_periodo(wb, "REPORTE_ANUAL", "ANIO")

    build_log(wb)
    ws_inicio = build_inicio(wb)

    define_named_ranges(wb)
    apply_data_validations(wb)
    apply_conditional_formats(wb)
    protect_formula_columns(wb)

    # Excel recalculates formulas when the workbook opens or inputs change.
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    wb.active = wb.index(ws_inicio)
    wb.save(out_path)
    return out_path


# ---------------------------------------------------------------------------
# CAPA 1 - CONFIGURACION
# ---------------------------------------------------------------------------

def build_menu(wb):
    ws = wb.create_sheet("MENU")
    ws.sheet_properties.tabColor = COLOR_SYSTEM
    sheet_title(ws, "SISTEMA FINANCIERO EMPRESARIAL",
                "Plataforma configurable de gestion financiera - Colombia")

    ws["A4"] = "Navegacion rapida (Ctrl+clic en Excel Desktop, o clic simple segun version):"
    ws["A4"].font = Font(italic=True, size=9, color="666666")

    items = [
        ("INICIO", "INICIO"),
        ("CONFIGURAR EMPRESA", "EMPRESA"),
        ("PARAMETROS", "PARAMETROS"),
        ("CATEGORIAS", "CATEGORIAS"),
        ("IMPUESTOS", "IMPUESTOS"),
        ("CUENTAS", "CUENTAS"),
        ("TIPOS DE DOCUMENTO", "TIPOS_DOCUMENTOS"),
        ("TERCEROS", "TERCEROS"),
        ("ASISTENTE DE CLASIFICACION", "ASISTENTE"),
        ("REGLAS DEL ASISTENTE", "REGLAS_IA"),
        ("DOCUMENTOS", "DOCUMENTOS"),
        ("REGISTRAR MOVIMIENTO", "MOVIMIENTOS"),
        ("PAGOS", "PAGOS"),
        ("COBROS", "COBROS"),
        ("DASHBOARD", "DASHBOARD"),
        ("REPORTE MENSUAL", "REPORTE_MENSUAL"),
        ("REPORTE ANUAL", "REPORTE_ANUAL"),
        ("LOG DE AUDITORIA", "LOG"),
    ]
    row = 6
    for label, sheet in items:
        cell = ws.cell(row=row, column=1, value=label)
        cell.hyperlink = f"#{sheet}!A1"
        cell.font = Font(color="1155CC", underline="single", bold=True, size=12)
        ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor="EAF1FB")
        ws.row_dimensions[row].height = 22
        row += 1

    macro_row = row + 2
    ws.cell(row=macro_row, column=1, value="MACROS (Fases 4 y 5 - importar los .bas de la carpeta vba/)")
    ws.cell(row=macro_row, column=1).font = Font(bold=True, color=COLOR_REPORT)
    ws.cell(row=macro_row + 1, column=1, value=(
        "Ejecutar con Alt+F8, o asignar a un boton: Insertar > Formularios > Boton, dibujar, "
        "elegir el macro y pulsar Aceptar."
    ))
    ws.cell(row=macro_row + 1, column=1).font = Font(italic=True, size=9, color="808080")
    ws.merge_cells(start_row=macro_row + 1, start_column=1, end_row=macro_row + 1, end_column=3)

    macros = [
        ("Menu_ImportarDocumento", "Copia el soporte y extrae campos estructurados si es XML UBL"),
        ("Menu_ExtraerXMLDocumentoActivo", "Relee el XML UBL de la fila DOCUMENTOS seleccionada"),
        ("Menu_ConfirmarRevisionXMLDocumentoActivo", "Confirma la revisión humana del XML seleccionado"),
        ("Menu_RenombrarArchivoDocumentoActivo", "Renombra en disco el archivo de la fila DOCUMENTOS donde este el cursor"),
        ("Menu_NuevoMovimiento", "Inserta una fila nueva en MOVIMIENTOS con ID y fecha automaticos"),
        ("Menu_NuevoDocumento", "Inserta una fila nueva en DOCUMENTOS con ID automatico"),
        ("Menu_NuevoTercero", "Inserta una fila nueva en TERCEROS con ID automatico"),
        ("Menu_NuevoPago", "Inserta una fila nueva en PAGOS con ID automatico"),
        ("Menu_NuevoCobro", "Inserta una fila nueva en COBROS con ID automatico"),
        ("Menu_ValidarTodo", "Valida MOVIMIENTOS, DOCUMENTOS (incl. duplicados) y TERCEROS (NIT duplicado)"),
        ("Menu_ActualizarReportes", "Recalcula el libro y repuebla CXC/CXP desde MOVIMIENTOS pendientes"),
    ]
    header_row = macro_row + 3
    ws.cell(row=header_row, column=1, value="Macro").font = Font(bold=True)
    ws.cell(row=header_row, column=2, value="Que hace").font = Font(bold=True)
    for i, (nombre, desc) in enumerate(macros, start=header_row + 1):
        ws.cell(row=i, column=1, value=nombre).font = Font(name="Consolas", size=9)
        ws.cell(row=i, column=2, value=desc).font = Font(size=9)

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 70
    return ws


def build_asistente(wb):
    """Asistente local de clasificación, sin APIs ni transmisión de datos."""
    ws = wb.create_sheet("ASISTENTE")
    ws.sheet_properties.tabColor = COLOR_REPORT
    sheet_title(
        ws,
        "ASISTENTE LOCAL",
        "Sugerencias por reglas editables; revise cada resultado antes de registrarlo",
    )
    ws.sheet_view.showGridLines = False

    ws["A4"] = "DATOS DE LA OPERACIÓN"
    ws["A4"].font = Font(bold=True, color=COLOR_REPORT)
    ws["A5"] = "Tercero o proveedor"
    ws["A6"] = "Concepto o descripción"
    ws["B5"] = ""
    ws["B6"] = ""
    ws["A8"] = "SUGERENCIA"
    ws["A8"].font = Font(bold=True, color=COLOR_REPORT)

    labels = [
        ("Estado", 9),
        ("Tipo", 10),
        ("Categoría", 11),
        ("Subcategoría", 12),
        ("Medio de pago", 13),
        ("Cuenta", 14),
        ("Regla coincidente", 15),
        ("Prioridad de regla", 16),
    ]
    for label, row in labels:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)

    score_range = "REGLAS_IA!$K$5:$K$104"
    max_score = f"MAX({score_range})"
    ws["B9"] = (
        f'=IF($B$6="","Escriba un concepto",'
        f'IF({max_score}=0,"Sin coincidencia",'
        f'IF(COUNTIF({score_range},{max_score})>1,'
        '"Conflicto: ajuste prioridad","Sugerencia lista para revisar")))'
    )
    suggestion_columns = {
        10: "D", 11: "E", 12: "F", 13: "G", 14: "H", 15: "C", 16: "J"
    }
    for row, column in suggestion_columns.items():
        source_range = f"REGLAS_IA!${column}$5:${column}$104"
        ws.cell(
            row=row,
            column=2,
            value=(
                f'=IF($B$9<>"Sugerencia lista para revisar","",'
                f'IFERROR(INDEX({source_range},MATCH({max_score},{score_range},0)),""))'
            ),
        )

    ws["A18"] = "CÓMO USARLO"
    ws["A18"].font = Font(bold=True, color=COLOR_REPORT)
    ws["A19"] = "1. Escriba el tercero y el concepto. El tercero es opcional."
    ws["A20"] = "2. Revise la regla, el tipo y la categoría sugeridos."
    ws["A21"] = "3. Confirme los datos y regístrelos en MOVIMIENTOS."
    ws["A23"] = (
        "Funciona sin internet ni servicios pagos. La prioridad ordena reglas, "
        "no es una probabilidad de acierto. No calcula impuestos ni registra datos automáticamente."
    )
    ws["A23"].font = Font(italic=True, size=9, color="666666")
    ws.merge_cells("A23:F24")
    ws["A23"].alignment = Alignment(vertical="top", wrap_text=True)
    ws["A26"] = "Editar las reglas y crear patrones propios"
    ws["A26"].hyperlink = "#'REGLAS_IA'!A1"
    ws["A26"].font = Font(color="1155CC", underline="single", bold=True)
    ws["A27"] = "Volver al inicio"
    ws["A27"].hyperlink = "#'INICIO'!A1"
    ws["A27"].font = Font(color="1155CC", underline="single", bold=True)

    for cell in ("B5", "B6"):
        ws[cell].fill = PatternFill("solid", fgColor="FFF2CC")
        ws[cell].border = BORDER_THIN
    for row in range(9, 17):
        ws.cell(row=row, column=2).fill = PatternFill("solid", fgColor="E2F0D9")
        ws.cell(row=row, column=2).border = BORDER_THIN
    ws["B9"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 31
    ws.column_dimensions["B"].width = 48
    for col in "CDEF":
        ws.column_dimensions[col].width = 16
    ws.freeze_panes = "A5"
    return ws


def build_reglas_ia(wb):
    """Reglas iniciales genéricas que cada empresa puede editar o desactivar."""
    ws = wb.create_sheet("REGLAS_IA")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(
        ws,
        "REGLAS DEL ASISTENTE",
        "Catálogo editable de patrones; agregue reglas aprobadas por su empresa",
    )
    ws.sheet_view.showGridLines = False
    ws["A3"] = (
        "Las reglas solo proponen clasificación. Prioridad mayor gana; si hay empate, "
        "el asistente muestra conflicto. Ajuste palabras y prioridad según sus operaciones."
    )
    ws.merge_cells("A3:L3")
    ws["A3"].font = Font(italic=True, size=9, color="666666")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 30

    headers = [
        "ID", "Activa", "PalabraClave", "Tipo", "Categoria", "Subcategoria",
        "MedioPago", "Cuenta", "TerceroOpcional", "Prioridad",
        "Coincidencia", "Observaciones",
    ]
    start = 4
    for col, header in enumerate(headers, start=1):
        ws.cell(row=start, column=col, value=header)

    seeds = [
        ("IA-001", "Sí", "arriendo", "Gasto", "Administracion", "Arriendo", "", "", "", 20, "Ejemplo editable"),
        ("IA-002", "Sí", "servicios publicos", "Gasto", "Operacion", "Servicios publicos", "", "", "", 20, "Ejemplo sin tilde"),
        ("IA-003", "Sí", "servicios públicos", "Gasto", "Operacion", "Servicios publicos", "", "", "", 20, "Variante con tilde"),
        ("IA-004", "Sí", "suministros", "Gasto", "Operacion", "Suministros", "", "", "", 20, "Ejemplo editable"),
        ("IA-005", "Sí", "software", "Gasto", "Administracion", "Software", "", "", "", 20, "Ejemplo editable"),
        ("IA-006", "Sí", "nomina", "Gasto", "Nomina", "Salarios", "", "", "", 20, "Ejemplo sin tilde"),
        ("IA-007", "Sí", "nómina", "Gasto", "Nomina", "Salarios", "", "", "", 20, "Variante con tilde"),
        ("IA-008", "Sí", "venta de productos", "Ingreso", "Ventas", "Productos", "", "", "", 30, "Frase específica"),
        ("IA-009", "Sí", "venta productos", "Ingreso", "Ventas", "Productos", "", "", "", 25, "Frase breve"),
        ("IA-010", "Sí", "venta de servicios", "Ingreso", "Ventas", "Servicios", "", "", "", 30, "Frase específica"),
        ("IA-011", "Sí", "venta servicios", "Ingreso", "Ventas", "Servicios", "", "", "", 25, "Frase breve"),
        ("IA-012", "Sí", "rendimientos", "Ingreso", "Financiero", "Rendimientos", "", "", "", 20, "Ejemplo editable"),
        ("IA-013", "Sí", "compra mercancia", "Compra", "Inventario", "Mercancia", "", "", "", 25, "Ejemplo sin tilde"),
        ("IA-014", "Sí", "compra mercancía", "Compra", "Inventario", "Mercancia", "", "", "", 25, "Variante con tilde"),
        ("IA-015", "Sí", "tecnologia", "Compra", "Activos", "Tecnologia", "", "", "", 20, "Ejemplo sin tilde"),
        ("IA-016", "Sí", "tecnología", "Compra", "Activos", "Tecnologia", "", "", "", 20, "Variante con tilde"),
    ]
    last_row = start + ANALYTIC_ROW_CAPACITY["REGLAS_IA"]
    for row in range(start + 1, last_row + 1):
        seed_index = row - start - 1
        if seed_index < len(seeds):
            for col, value in enumerate(seeds[seed_index], start=1):
                if col <= 10:
                    ws.cell(row=row, column=col, value=value)
                elif col == 11:
                    ws.cell(row=row, column=12, value=value)
        ws.cell(
            row=row,
            column=11,
            value=(
                f'=IF(OR(B{row}<>"Sí",C{row}="",ASISTENTE!$B$6=""),0,'
                f'IFERROR(IF(AND(OR(I{row}="",ISNUMBER(SEARCH(I{row},ASISTENTE!$B$5))),'
                f'ISNUMBER(SEARCH(C{row},ASISTENTE!$B$5&" "&ASISTENTE!$B$6))),J{row},0),0))'
            ),
        )
        ws.cell(row=row, column=10).number_format = "0"
        ws.cell(row=row, column=11).number_format = "0"

    add_table(ws, "REGLAS_IA", f"A{start}:L{last_row}", headers, header_row=start)
    widths = [12, 10, 24, 14, 22, 24, 18, 22, 24, 12, 16, 28]
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A5"
    return ws


def build_inicio(wb):
    """Página de entrada con tareas frecuentes y pasos de puesta en marcha."""
    ws = wb.create_sheet("INICIO", 0)
    ws.sheet_properties.tabColor = COLOR_SYSTEM
    ws.sheet_view.showGridLines = False
    sheet_title(
        ws,
        "SISTEMA FINANCIERO",
        "Inicio rápido | Configure la empresa, registre operaciones y revise resultados",
    )
    ws["A4"] = "TAREAS FRECUENTES"
    ws["A4"].font = Font(bold=True, color=COLOR_REPORT)

    cards = [
        (6, 1, 4, "1. CONFIGURAR EMPRESA", "Actualice datos, moneda y parámetros.", "EMPRESA"),
        (6, 6, 9, "2. REGISTRAR MOVIMIENTO", "Capture ingreso, gasto, compra o venta.", "MOVIMIENTOS"),
        (10, 1, 4, "3. ASISTENTE LOCAL", "Reciba una sugerencia y revísela.", "ASISTENTE"),
        (10, 6, 9, "4. REVISAR DOCUMENTOS", "Consulte soportes y complete datos.", "DOCUMENTOS"),
        (14, 1, 4, "5. VER INDICADORES", "Filtre ingresos, gastos y pendientes.", "DASHBOARD"),
        (14, 6, 9, "6. ABRIR REPORTES", "Consulte periodos mensual y anual.", "REPORTE_MENSUAL"),
    ]
    for row, col_start, col_end, title, description, target in cards:
        first = get_column_letter(col_start)
        last = get_column_letter(col_end)
        title_range = f"{first}{row}:{last}{row}"
        body_range = f"{first}{row + 1}:{last}{row + 2}"
        ws.merge_cells(title_range)
        ws.merge_cells(body_range)
        title_cell = ws.cell(row=row, column=col_start, value=title)
        title_cell.hyperlink = f"#'{target}'!A1"
        title_cell.font = Font(color="FFFFFF", bold=True, size=11, underline="single")
        title_cell.fill = PatternFill("solid", fgColor=COLOR_CONFIG)
        title_cell.alignment = Alignment(vertical="center")
        body_cell = ws.cell(row=row + 1, column=col_start, value=description)
        body_cell.fill = PatternFill("solid", fgColor="EAF1FB")
        body_cell.font = Font(size=10, color="333333")
        body_cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[row].height = 26
        ws.row_dimensions[row + 1].height = 24
        ws.row_dimensions[row + 2].height = 18

    ws["A18"] = "PRIMEROS PASOS"
    ws["A18"].font = Font(bold=True, color=COLOR_REPORT)
    ws["A19"] = "1. Configure empresa, categorías, cuentas y terceros."
    ws["A20"] = "2. Registre una operación o use el asistente para clasificar el concepto."
    ws["A21"] = "3. Revise los indicadores y conserve el soporte relacionado."
    ws["A23"] = (
        "El libro XLSX permite captura, fórmulas y sugerencias locales. La importación "
        "de archivos, lectura XML y altas automáticas requieren la copia XLSM con los módulos VBA instalados."
    )
    ws.merge_cells("A23:I24")
    ws["A23"].font = Font(italic=True, size=9, color="666666")
    ws["A23"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["A26"] = "Ver todos los módulos"
    ws["A26"].hyperlink = "#'MENU'!A1"
    ws["A26"].font = Font(color="1155CC", underline="single", bold=True)
    ws["F26"] = "Editar reglas del asistente"
    ws["F26"].hyperlink = "#'REGLAS_IA'!A1"
    ws["F26"].font = Font(color="1155CC", underline="single", bold=True)

    for col in "ABCDEFGHI":
        ws.column_dimensions[col].width = 18
    ws.freeze_panes = "A4"
    return ws


def build_empresa(wb):
    ws = wb.create_sheet("EMPRESA")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "EMPRESA", "Datos de la empresa - editar libremente, no usar en formulas de otras hojas")

    headers = ["Campo", "Valor"]
    start = 4
    ws.append([])
    for r in range(1, start):
        pass
    ws.cell(row=start, column=1, value=headers[0])
    ws.cell(row=start, column=2, value=headers[1])

    data = [
        ("NombreEmpresa", "Empresa Ejemplo S.A.S."),
        ("NIT", "900000000-1"),
        ("RegimenTributario", "Configurar (Comun / No responsable de IVA / Simple / etc.)"),
        ("ActividadEconomica", "Configurar codigo CIIU"),
        ("ResponsableIVA", "No"),
        ("DireccionPrincipal", ""),
        ("Ciudad", ""),
        ("Telefono", ""),
        ("Email", ""),
        ("MonedaPrincipal", "COP"),
        ("AnioFiscalInicioMes", "Enero"),
        ("RutaRaizArchivos", ""),
        ("LogoRuta", ""),
    ]
    for i, (campo, valor) in enumerate(data, start=start + 1):
        ws.cell(row=i, column=1, value=campo).font = Font(bold=True)
        ws.cell(row=i, column=2, value=valor)

    add_table(ws, "EMPRESA", f"A{start}:B{start + len(data)}", headers, header_row=start)
    ws.column_dimensions["B"].width = 45
    return ws


def build_parametros(wb):
    ws = wb.create_sheet("PARAMETROS")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "PARAMETROS", "Parametros generales del sistema - configurables sin tocar formulas")

    headers = ["Codigo", "Parametro", "Valor", "Tipo", "Descripcion"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    data = [
        ("PAR-001", "DECIMALES_MONEDA", 0, "Numero", "Decimales para valores monetarios"),
        ("PAR-002", "FORMATO_FECHA", "dd/mm/aaaa", "Texto", "Formato de fecha estandar"),
        ("PAR-003", "PREFIJO_MOVIMIENTO", "MOV-", "Texto", "Prefijo para ID de movimientos"),
        ("PAR-004", "PREFIJO_DOCUMENTO", "DOC-", "Texto", "Prefijo para ID de documentos"),
        ("PAR-005", "UMBRAL_DIAS_DUPLICADO", 5, "Numero", "Dias de tolerancia para sugerir posible duplicado"),
        ("PAR-006", "RUTA_DOCUMENTOS", "Documentos", "Texto", "Ruta relativa a la carpeta del libro o ruta absoluta de soportes"),
        ("PAR-007", "MODO_COMPATIBILIDAD", "COMPATIBLE", "Texto", "COMPATIBLE (Excel 2019) o MODERNO"),
    ]
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals, 1):
            ws.cell(row=i, column=c, value=v)

    add_table(ws, "PARAMETROS", f"A{start}:E{start + len(data)}", headers, header_row=start)
    return ws


CATEGORIAS_DATA = [
    ("ING-001", "Ingreso", "Ventas", "Productos", "Si", ""),
    ("ING-002", "Ingreso", "Ventas", "Servicios", "Si", ""),
    ("ING-003", "Ingreso", "Financiero", "Rendimientos", "Si", ""),
    ("GAS-001", "Gasto", "Operacion", "Servicios publicos", "Si", ""),
    ("GAS-002", "Gasto", "Operacion", "Suministros", "Si", ""),
    ("GAS-003", "Gasto", "Administracion", "Software", "Si", ""),
    ("GAS-004", "Gasto", "Administracion", "Arriendo", "Si", ""),
    ("GAS-005", "Gasto", "Nomina", "Salarios", "Si", ""),
    ("COM-001", "Compra", "Inventario", "Mercancia", "Si", ""),
    ("COM-002", "Compra", "Activos", "Tecnologia", "Si", ""),
    ("VEN-001", "Venta", "Productos", "General", "Si", ""),
    ("OTR-001", "Otro", "Sin clasificar", "Sin clasificar", "Si", "Usar temporalmente y reclasificar"),
]


def build_categorias(wb):
    ws = wb.create_sheet("CATEGORIAS")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "CATEGORIAS", "Clasificacion financiera - editable, no ligada a formulas ni VBA")

    headers = ["Codigo", "Tipo", "Categoria", "Subcategoria", "Activa", "Descripcion"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    data = CATEGORIAS_DATA
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals, 1):
            ws.cell(row=i, column=c, value=v)

    add_table(ws, "CATEGORIAS", f"A{start}:F{start + len(data)}", headers, header_row=start)
    return ws


def build_impuestos(wb):
    ws = wb.create_sheet("IMPUESTOS")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "IMPUESTOS", "Impuestos y retenciones - VERIFICAR PORCENTAJES VIGENTES ANTES DE USAR")

    headers = ["Codigo", "Impuesto", "Porcentaje", "Tipo", "Base", "CuentaAsociada", "Activo", "Observaciones"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    nota = "Configurar segun normativa DIAN/municipal vigente"
    data = [
        ("IVA-001", "IVA general", 0, "Impuesto", "Subtotal", "", "Si", nota),
        ("RET-001", "Retencion en la fuente - Servicios", 0, "Retencion", "Subtotal", "", "Si", nota),
        ("RET-002", "Retencion en la fuente - Compras", 0, "Retencion", "Subtotal", "", "Si", nota),
        ("RET-003", "ReteICA", 0, "Retencion", "Subtotal", "", "Si", nota),
        ("IC-001", "Impuesto de Industria y Comercio (ICA)", 0, "Impuesto", "Ingresos", "", "Si", nota),
    ]
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals, 1):
            ws.cell(row=i, column=c, value=v)
        ws.cell(row=i, column=3).number_format = PCT_FMT

    add_table(ws, "IMPUESTOS", f"A{start}:H{start + len(data)}", headers, header_row=start)
    return ws


def build_cuentas(wb):
    ws = wb.create_sheet("CUENTAS")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "CUENTAS", "Cuentas de la empresa - el saldo actual se calcula automaticamente")

    headers = ["ID", "Cuenta", "Tipo", "Entidad", "NumeroReferencia", "SaldoInicial", "SaldoActual", "Activa", "Observaciones"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    data = [
        ("CTA-001", "Caja General", "Caja", "", "", 0, "Efectivo", "Si", ""),
        ("CTA-002", "Banco Principal", "Banco", "Configurar entidad", "", 0, "Banco", "Si", ""),
        ("CTA-003", "Cuentas por Cobrar", "Cuenta por cobrar", "", "", 0, "CxC", "Si", ""),
        ("CTA-004", "Cuentas por Pagar", "Cuenta por pagar", "", "", 0, "CxP", "Si", ""),
    ]
    n = len(data)
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals[:6], 1):
            ws.cell(row=i, column=c, value=v)
        ws.cell(row=i, column=6).number_format = CURRENCY_FMT
        ws.cell(row=i, column=8, value=row_vals[7])
        ws.cell(row=i, column=9, value=row_vals[8])

    register_layout("CUENTAS", headers, start + 1, start + n)
    mov_total = col_range("MOVIMIENTOS", "Total")
    mov_cuenta = col_range("MOVIMIENTOS", "Cuenta")
    mov_tipo = col_range("MOVIMIENTOS", "Tipo")
    mov_estado = col_range("MOVIMIENTOS", "Estado")
    for i in range(start + 1, start + n + 1):
        cta_ref = cell_ref("CUENTAS", "Cuenta", i)
        saldo_ini = cell_ref("CUENTAS", "SaldoInicial", i)
        formula = (
            f"={saldo_ini}"
            f'+SUMIFS({mov_total},{mov_cuenta},{cta_ref},{mov_tipo},"Ingreso",{mov_estado},"Pagado")'
            f'+SUMIFS({mov_total},{mov_cuenta},{cta_ref},{mov_tipo},"Venta",{mov_estado},"Pagado")'
            f'-SUMIFS({mov_total},{mov_cuenta},{cta_ref},{mov_tipo},"Gasto",{mov_estado},"Pagado")'
            f'-SUMIFS({mov_total},{mov_cuenta},{cta_ref},{mov_tipo},"Compra",{mov_estado},"Pagado")'
        )
        ws.cell(row=i, column=7, value=formula)
        ws.cell(row=i, column=7).number_format = CURRENCY_FMT

    add_table(ws, "CUENTAS", f"A{start}:I{start + n}", headers, header_row=start)
    return ws


def build_tipos_documentos(wb):
    ws = wb.create_sheet("TIPOS_DOCUMENTOS")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "TIPOS_DOCUMENTOS", "Tipos de documento soportados - ampliable")

    headers = ["Codigo", "TipoDocumento", "Prefijo", "RequiereCUFE", "Activo", "Descripcion"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    data = [
        ("TD-01", "Factura electronica", "FE", "Si", "Si", ""),
        ("TD-02", "Documento equivalente", "DE", "No", "Si", ""),
        ("TD-03", "Documento soporte", "DS", "No", "Si", "Compras a no obligados a facturar"),
        ("TD-04", "Nota credito", "NC", "Si", "Si", ""),
        ("TD-05", "Nota debito", "ND", "Si", "Si", ""),
        ("TD-06", "Recibo", "RC", "No", "Si", ""),
        ("TD-07", "Comprobante", "CO", "No", "Si", ""),
        ("TD-08", "Extracto bancario", "EB", "No", "Si", ""),
        ("TD-09", "Nomina", "NM", "No", "Si", ""),
        ("TD-10", "Contrato", "CT", "No", "Si", ""),
        ("TD-11", "Otro", "OT", "No", "Si", ""),
    ]
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals, 1):
            ws.cell(row=i, column=c, value=v)

    add_table(ws, "TIPOS_DOCUMENTOS", f"A{start}:F{start + len(data)}", headers, header_row=start)
    return ws


def build_estados(wb):
    ws = wb.create_sheet("ESTADOS")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "ESTADOS", "Catalogo de estados por ambito - referencia; los desplegables usan listas fijas por ambito")

    headers = ["Codigo", "Estado", "Ambito", "Activo", "Descripcion"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    data = [
        ("EST-01", "Pendiente", "Movimiento", "Si", ""),
        ("EST-02", "Pagado", "Movimiento", "Si", ""),
        ("EST-03", "Anulado", "Movimiento", "Si", ""),
        ("EST-04", "Validado", "Documento", "Si", ""),
        ("EST-05", "Pendiente", "Documento", "Si", ""),
        ("EST-06", "Error", "Documento", "Si", ""),
        ("EST-07", "Revision Manual", "Documento", "Si", ""),
        ("EST-08", "Duplicado", "Documento", "Si", ""),
        ("EST-09", "Pendiente", "Pago/Cobro", "Si", ""),
        ("EST-10", "Aplicado", "Pago/Cobro", "Si", ""),
        ("EST-11", "Anulado", "Pago/Cobro", "Si", ""),
    ]
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals, 1):
            ws.cell(row=i, column=c, value=v)

    add_table(ws, "ESTADOS", f"A{start}:E{start + len(data)}", headers, header_row=start)
    return ws


def build_medios_pago(wb):
    ws = wb.create_sheet("MEDIOS_PAGO")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "MEDIOS_PAGO")

    headers = ["Codigo", "MedioPago", "Activo", "Descripcion"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    data = [
        ("MP-01", "Efectivo", "Si", ""),
        ("MP-02", "Transferencia bancaria", "Si", ""),
        ("MP-03", "Tarjeta debito", "Si", ""),
        ("MP-04", "Tarjeta credito", "Si", ""),
        ("MP-05", "Cheque", "Si", ""),
        ("MP-06", "Billetera digital", "Si", ""),
        ("MP-07", "Otro", "Si", ""),
    ]
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals, 1):
            ws.cell(row=i, column=c, value=v)

    add_table(ws, "MEDIOS_PAGO", f"A{start}:D{start + len(data)}", headers, header_row=start)
    return ws


def build_usuarios_roles(wb):
    ws = wb.create_sheet("USUARIOS_ROLES")
    ws.sheet_properties.tabColor = COLOR_CONFIG
    sheet_title(ws, "USUARIOS_ROLES", "Administracion de usuarios - control de acceso logico (no reemplaza proteccion de Excel)")

    headers = ["ID", "Usuario", "Rol", "Email", "Activo", "Observaciones"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    data = [
        ("USR-001", "Administrador", "Administrador", "", "Si", "Acceso total"),
        ("USR-002", "Contador", "Contador", "", "Si", "Registro y validacion contable"),
        ("USR-003", "Auxiliar", "Auxiliar", "", "Si", "Digitacion y digitalizacion"),
        ("USR-004", "Consulta", "Consulta", "", "Si", "Solo lectura de reportes"),
    ]
    for i, row_vals in enumerate(data, start=start + 1):
        for c, v in enumerate(row_vals, 1):
            ws.cell(row=i, column=c, value=v)

    add_table(ws, "USUARIOS", f"A{start}:F{start + len(data)}", headers, header_row=start)
    return ws


# ---------------------------------------------------------------------------
# CAPA 2 - DATOS
# ---------------------------------------------------------------------------

def build_terceros(wb):
    ws = wb.create_sheet("TERCEROS")
    ws.sheet_properties.tabColor = COLOR_DATA
    sheet_title(ws, "TERCEROS", "Base unica de clientes, proveedores, empleados, socios y entidades")

    headers = ["ID", "Tipo", "NombreRazonSocial", "NIT_ID", "Contacto", "Email", "Telefono",
               "Direccion", "Ciudad", "Estado", "Observaciones", "ClaveDuplicado"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    n_example_rows = 30
    last_row = start + n_example_rows
    register_layout("TERCEROS", headers, start + 1, last_row)
    nit_col = colletter("TERCEROS", "NIT_ID")
    for i in range(start + 1, last_row + 1):
        ws.cell(row=i, column=10, value="Activo")
        ws.cell(row=i, column=12, value=f'=IF({nit_col}{i}="","",TRIM(UPPER({nit_col}{i})))')

    add_table(ws, "TERCEROS", f"A{start}:L{last_row}", headers, header_row=start)
    return ws


def build_documentos(wb):
    ws = wb.create_sheet("DOCUMENTOS")
    ws.sheet_properties.tabColor = COLOR_DATA
    sheet_title(ws, "DOCUMENTOS", "Soportes digitalizados: facturas, recibos, extractos y demas documentos")

    headers = ["ID", "Fecha", "TipoDocumento", "Numero", "Tercero", "Identificacion",
               "Subtotal", "Impuestos", "Retenciones", "Total", "CUFE_Referencia", "Estado",
               "RutaArchivo", "NombreArchivo", "TipoArchivo", "FechaExtraccion",
               "MetodoExtraccion", "Confianza", "Observaciones"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    n_example_rows = 40
    last_row = start + n_example_rows
    register_layout("DOCUMENTOS", headers, start + 1, last_row)
    for i in range(start + 1, last_row + 1):
        ws.cell(row=i, column=2).number_format = DATE_FMT
        ws.cell(row=i, column=7).number_format = CURRENCY_FMT
        ws.cell(row=i, column=8).number_format = CURRENCY_FMT
        ws.cell(row=i, column=9).number_format = CURRENCY_FMT
        # Total es el valor final del documento, que puede diferir de una
        # suma simple por descuentos, cargos, anticipos o reglas del emisor.
        ws.cell(row=i, column=10).number_format = CURRENCY_FMT
        ws.cell(row=i, column=16).number_format = DATE_FMT

    add_table(ws, "DOCUMENTOS", f"A{start}:S{last_row}", headers, header_row=start)
    return ws


def build_movimientos(wb):
    ws = wb.create_sheet("MOVIMIENTOS")
    ws.sheet_properties.tabColor = COLOR_DATA
    sheet_title(ws, "MOVIMIENTOS", "Tabla principal: ingresos, gastos, compras, ventas y otros movimientos")

    headers = ["ID", "Fecha", "Tipo", "Documento", "Tercero", "NIT_ID", "Concepto",
               "Categoria", "Subcategoria", "MedioPago", "Cuenta", "Subtotal", "Impuestos",
               "Retenciones", "Total", "Estado", "SoporteID", "Observaciones"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    n_example_rows = 60
    last_row = start + n_example_rows
    register_layout("MOVIMIENTOS", headers, start + 1, last_row)
    sub_c = colletter("MOVIMIENTOS", "Subtotal")
    imp_c = colletter("MOVIMIENTOS", "Impuestos")
    ret_c = colletter("MOVIMIENTOS", "Retenciones")
    for i in range(start + 1, last_row + 1):
        ws.cell(row=i, column=2).number_format = DATE_FMT
        ws.cell(row=i, column=12).number_format = CURRENCY_FMT
        ws.cell(row=i, column=13).number_format = CURRENCY_FMT
        ws.cell(row=i, column=14).number_format = CURRENCY_FMT
        ws.cell(row=i, column=15, value=(
            f'=IF({sub_c}{i}="",0,{sub_c}{i}+{imp_c}{i}-{ret_c}{i})'
        ))
        ws.cell(row=i, column=15).number_format = CURRENCY_FMT_CALC

    add_table(ws, "MOVIMIENTOS", f"A{start}:R{last_row}", headers, header_row=start)
    return ws


def build_pagos(wb):
    ws = wb.create_sheet("PAGOS")
    ws.sheet_properties.tabColor = COLOR_DATA
    sheet_title(ws, "PAGOS", "Salidas de caja aplicadas contra movimientos de Gasto/Compra (permite pagos parciales)")

    headers = ["ID", "Fecha", "MovimientoID", "Tercero", "Cuenta", "MedioPago", "ValorPagado", "Estado", "Observaciones"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    n_example_rows = 20
    last_row = start + n_example_rows
    for i in range(start + 1, last_row + 1):
        ws.cell(row=i, column=2).number_format = DATE_FMT
        ws.cell(row=i, column=7).number_format = CURRENCY_FMT

    add_table(ws, "PAGOS", f"A{start}:I{last_row}", headers, header_row=start)
    return ws


def build_cobros(wb):
    ws = wb.create_sheet("COBROS")
    ws.sheet_properties.tabColor = COLOR_DATA
    sheet_title(ws, "COBROS", "Entradas de caja aplicadas contra movimientos de Ingreso/Venta (permite cobros parciales)")

    headers = ["ID", "Fecha", "MovimientoID", "Tercero", "Cuenta", "MedioPago", "ValorCobrado", "Estado", "Observaciones"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    n_example_rows = 20
    last_row = start + n_example_rows
    for i in range(start + 1, last_row + 1):
        ws.cell(row=i, column=2).number_format = DATE_FMT
        ws.cell(row=i, column=7).number_format = CURRENCY_FMT

    add_table(ws, "COBROS", f"A{start}:I{last_row}", headers, header_row=start)
    return ws


# ---------------------------------------------------------------------------
# CAPA 3 - REPORTES
# ---------------------------------------------------------------------------

def build_dashboard(wb):
    ws = wb.create_sheet("DASHBOARD", 0)
    ws.sheet_properties.tabColor = COLOR_REPORT
    sheet_title(ws, "DASHBOARD", "Indicadores financieros - se actualiza automaticamente desde MOVIMIENTOS y DOCUMENTOS")

    ws["A4"] = "FILTROS"
    ws["A4"].font = Font(bold=True, color=COLOR_REPORT)
    filtros = [
        ("Anio (vacio = todos)", "B5"),
        ("Mes (1-12, vacio = todos)", "B6"),
        ("Tipo (vacio = todos)", "B7"),
        ("Categoria (vacio = todos)", "B8"),
        ("Tercero (vacio = todos)", "B9"),
        ("Cuenta (vacio = todos)", "B10"),
        ("Estado (vacio = todos)", "B11"),
    ]
    for i, (label, cellref) in enumerate(filtros):
        ws.cell(row=5 + i, column=1, value=label)
        ws[cellref].fill = PatternFill("solid", fgColor="FFF2CC")
        ws[cellref].border = BORDER_THIN

    dv_anio = DataValidation(type="whole", operator="between", formula1="1900", formula2="9999", allow_blank=True)
    ws.add_data_validation(dv_anio)
    dv_anio.add("B5")
    dv_mes = DataValidation(type="whole", operator="between", formula1="1", formula2="12", allow_blank=True)
    ws.add_data_validation(dv_mes)
    dv_mes.add("B6")
    for celda, nombre_lista in [
        ("B7", "Lst_TipoMovimiento"),
        ("B8", "Lst_Categoria"),
        ("B9", "Lst_TerceroNombre"),
        ("B10", "Lst_Cuenta"),
        ("B11", "Lst_EstadoMovimiento"),
    ]:
        validacion = DataValidation(type="list", formula1=f"={nombre_lista}", allow_blank=True)
        ws.add_data_validation(validacion)
        validacion.add(celda)

    ws["A13"] = "INDICADORES"
    ws["A13"].font = Font(bold=True, color=COLOR_REPORT)

    mov_tipo = col_range("MOVIMIENTOS", "Tipo")
    mov_estado = col_range("MOVIMIENTOS", "Estado")
    mov_fecha = col_range("MOVIMIENTOS", "Fecha")
    mov_categoria = col_range("MOVIMIENTOS", "Categoria")
    mov_tercero = col_range("MOVIMIENTOS", "Tercero")
    mov_cuenta = col_range("MOVIMIENTOS", "Cuenta")
    mov_total = col_range("MOVIMIENTOS", "Total")
    mov_impuestos = col_range("MOVIMIENTOS", "Impuestos")
    doc_estado = col_range("DOCUMENTOS", "Estado")
    doc_fecha = col_range("DOCUMENTOS", "Fecha")
    doc_tercero = col_range("DOCUMENTOS", "Tercero")

    filtro_periodo_cat = (
        f'(($B$5="")+(YEAR({mov_fecha})=$B$5))*(($B$6="")+(MONTH({mov_fecha})=$B$6))'
        f'*(($B$7="")+({mov_tipo}=$B$7))*(($B$8="")+({mov_categoria}=$B$8))'
        f'*(($B$9="")+({mov_tercero}=$B$9))'
        f'*(($B$10="")+({mov_cuenta}=$B$10))'
    )
    filtro_estado = (
        f'((($B$11="")*({mov_estado}="Pagado"))'
        f'+(($B$11<>"")*({mov_estado}=$B$11)))'
    )
    filtro_pendientes = f'((($B$11="")+($B$11="Pendiente"))*({mov_estado}="Pendiente"))'
    filtro_periodo_doc = (
        f'(($B$5="")+(YEAR({doc_fecha})=$B$5))*(($B$6="")+(MONTH({doc_fecha})=$B$6))'
        f'*(($B$9="")+({doc_tercero}=$B$9))'
    )

    kpis = [
        ("Ingresos", f'=SUMPRODUCT((({mov_tipo}="Ingreso")+({mov_tipo}="Venta"))*{filtro_estado}*{filtro_periodo_cat}*{mov_total})'),
        ("Gastos", f'=SUMPRODUCT((({mov_tipo}="Gasto")+({mov_tipo}="Compra"))*{filtro_estado}*{filtro_periodo_cat}*{mov_total})'),
        ("Flujo Neto", "=B14-B15"),
        ("Cuentas por Cobrar", f'=SUMPRODUCT((({mov_tipo}="Ingreso")+({mov_tipo}="Venta"))*{filtro_periodo_cat}*{filtro_pendientes}*{mov_total})'),
        ("Cuentas por Pagar", f'=SUMPRODUCT((({mov_tipo}="Gasto")+({mov_tipo}="Compra"))*{filtro_periodo_cat}*{filtro_pendientes}*{mov_total})'),
        ("Impuestos Recaudados/Pagados", f'=SUMPRODUCT({filtro_estado}*{filtro_periodo_cat}*{mov_impuestos})'),
        ("Documentos Registrados", f'=SUMPRODUCT({filtro_periodo_doc}*({doc_estado}="Validado"))'),
        ("Documentos Pendientes", f'=SUMPRODUCT({filtro_periodo_doc}*((({doc_estado}="Pendiente")+({doc_estado}="Revision Manual"))))'),
        ("Documentos con Error", f'=SUMPRODUCT({filtro_periodo_doc}*((({doc_estado}="Error")+({doc_estado}="Duplicado"))))'),
    ]
    for i, (label, formula) in enumerate(kpis):
        r = 14 + i
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = CURRENCY_FMT if "Documento" not in label else "#,##0"
        c.fill = PatternFill("solid", fgColor="E2EFDA")
        c.border = BORDER_THIN

    ws["A24"] = "Tipo/Categoria/Cuenta filtran movimientos; los conteos de documentos usan Año, Mes y Tercero."
    ws["A24"].font = Font(italic=True, size=9, color="808080")
    ws["A25"] = "Para Anio/Mes vacios se muestran todos los periodos."
    ws["A25"].font = Font(italic=True, size=9, color="808080")

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 22

    chart = BarChart()
    chart.title = "Ingresos vs Gastos"
    chart.y_axis.title = "COP"
    chart.x_axis.title = "Indicador"
    data_ref = Reference(ws, min_col=2, min_row=14, max_row=15)
    cats_ref = Reference(ws, min_col=1, min_row=14, max_row=15)
    chart.add_data(data_ref, titles_from_data=False)
    chart.set_categories(cats_ref)
    chart.height = 8
    chart.width = 14
    ws.add_chart(chart, "D5")

    return ws


def build_ingresos_gastos(wb, sheet_name, tipo_a, tipo_b):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = COLOR_REPORT
    sheet_title(ws, sheet_name, f"Detalle de {tipo_a}/{tipo_b} agrupado por categoria y mes")
    ws["A3"] = "Año del reporte"
    ws["B3"] = DEFAULT_REPORT_YEAR
    ws["B3"].fill = PatternFill("solid", fgColor="FFF2CC")

    headers = ["Categoria", "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic", "Total"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)
    style_header_row(ws, start, len(headers), FILL_HEADER_REPORT)

    categorias = []
    for _codigo, tipo, categoria, _sub, _activa, _desc in CATEGORIAS_DATA:
        if tipo in (tipo_a, tipo_b) and categoria not in categorias:
            categorias.append(categoria)

    mov_categoria = col_range("MOVIMIENTOS", "Categoria")
    mov_tipo = col_range("MOVIMIENTOS", "Tipo")
    mov_fecha = col_range("MOVIMIENTOS", "Fecha")
    mov_estado = col_range("MOVIMIENTOS", "Estado")
    mov_total = col_range("MOVIMIENTOS", "Total")

    for offset, categoria in enumerate(categorias):
        r = start + 1 + offset
        cat_ref = f"$A{r}"
        ws.cell(row=r, column=1, value=categoria)
        for m in range(1, 13):
            col = 1 + m
            formula = (
                f'=SUMPRODUCT(({mov_categoria}={cat_ref})'
                f'*(({mov_tipo}="{tipo_a}")+({mov_tipo}="{tipo_b}"))'
                f'*(YEAR({mov_fecha})=$B$3)'
                f'*(MONTH({mov_fecha})={m})*({mov_estado}="Pagado")*{mov_total})'
            )
            ws.cell(row=r, column=col, value=formula).number_format = CURRENCY_FMT
        ws.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})").number_format = CURRENCY_FMT

    total_row = start + 1 + len(categorias)
    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
    for col in range(2, 15):
        cl = get_column_letter(col)
        ws.cell(
            row=total_row, column=col,
            value=f"=SUM({cl}{start + 1}:{cl}{total_row - 1})",
        ).number_format = CURRENCY_FMT
        ws.cell(row=total_row, column=col).font = Font(bold=True)

    autosize(ws, headers, min_width=9, max_width=14)
    ws.column_dimensions["A"].width = 22
    return ws


def build_flujo_caja(wb):
    ws = wb.create_sheet("FLUJO_CAJA")
    ws.sheet_properties.tabColor = COLOR_REPORT
    sheet_title(ws, "FLUJO_CAJA", "Entradas y salidas de caja por mes (base: PAGOS y COBROS aplicados)")
    ws["A3"] = "Año del reporte"
    ws["B3"] = DEFAULT_REPORT_YEAR
    ws["B3"].fill = PatternFill("solid", fgColor="FFF2CC")

    headers = ["Mes", "Cobros", "Pagos", "Flujo Neto Mensual", "Flujo Acumulado"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)
    style_header_row(ws, start, len(headers), FILL_HEADER_REPORT)

    cob_fecha = col_range("COBROS", "Fecha")
    cob_estado = col_range("COBROS", "Estado")
    cob_valor = col_range("COBROS", "ValorCobrado")
    pag_fecha = col_range("PAGOS", "Fecha")
    pag_estado = col_range("PAGOS", "Estado")
    pag_valor = col_range("PAGOS", "ValorPagado")

    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
             "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    for i, mes in enumerate(meses, start=1):
        r = start + i
        ws.cell(row=r, column=1, value=mes)
        ws.cell(row=r, column=2, value=f'=SUMPRODUCT((YEAR({cob_fecha})=$B$3)*(MONTH({cob_fecha})={i})*({cob_estado}="Aplicado")*{cob_valor})').number_format = CURRENCY_FMT
        ws.cell(row=r, column=3, value=f'=SUMPRODUCT((YEAR({pag_fecha})=$B$3)*(MONTH({pag_fecha})={i})*({pag_estado}="Aplicado")*{pag_valor})').number_format = CURRENCY_FMT
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = CURRENCY_FMT
        if i == 1:
            ws.cell(row=r, column=5, value=f"=D{r}").number_format = CURRENCY_FMT
        else:
            ws.cell(row=r, column=5, value=f"=E{r - 1}+D{r}").number_format = CURRENCY_FMT

    autosize(ws, headers)
    return ws


def build_cxc_cxp(wb, sheet_name, tipo_a, tipo_b):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = COLOR_REPORT
    label = "Cobrar" if sheet_name == "CXC" else "Pagar"
    sheet_title(ws, sheet_name, f"Movimientos pendientes de {label.lower()} ({tipo_a}/{tipo_b} con Estado=Pendiente)")

    headers = ["ID", "Fecha", "Tercero", "Documento", "Total", "DiasVencido", "Estado"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    ws["A" + str(start + 1)] = (
        "Ejecute Menu_ActualizarReportes después de importar los módulos VBA; mientras tanto, "
        "filtre MOVIMIENTOS por Estado=Pendiente (Datos > Filtro)."
    )
    ws["A" + str(start + 1)].font = Font(italic=True, size=9, color="808080")
    ws.merge_cells(start_row=start + 1, start_column=1, end_row=start + 1, end_column=7)

    autosize(ws, headers)
    return ws


def build_impuestos_reporte(wb):
    ws = wb.create_sheet("IMPUESTOS_REPORTE")
    ws.sheet_properties.tabColor = COLOR_REPORT
    sheet_title(ws, "IMPUESTOS_REPORTE", "Impuestos y retenciones causados por mes - base para declaraciones")
    ws["A3"] = "Año del reporte"
    ws["B3"] = DEFAULT_REPORT_YEAR
    ws["B3"].fill = PatternFill("solid", fgColor="FFF2CC")

    headers = ["Mes", "Impuestos Causados (Ingresos)", "Impuestos Pagados (Gastos)", "Retenciones Practicadas", "Retenciones Sufridas"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)
    style_header_row(ws, start, len(headers), FILL_HEADER_REPORT)

    mov_fecha = col_range("MOVIMIENTOS", "Fecha")
    mov_tipo = col_range("MOVIMIENTOS", "Tipo")
    mov_impuestos = col_range("MOVIMIENTOS", "Impuestos")
    mov_retenciones = col_range("MOVIMIENTOS", "Retenciones")

    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
             "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    for i, mes in enumerate(meses, start=1):
        r = start + i
        ws.cell(row=r, column=1, value=mes)
        ws.cell(row=r, column=2, value=f'=SUMPRODUCT((YEAR({mov_fecha})=$B$3)*(MONTH({mov_fecha})={i})*(({mov_tipo}="Ingreso")+({mov_tipo}="Venta"))*{mov_impuestos})').number_format = CURRENCY_FMT
        ws.cell(row=r, column=3, value=f'=SUMPRODUCT((YEAR({mov_fecha})=$B$3)*(MONTH({mov_fecha})={i})*(({mov_tipo}="Gasto")+({mov_tipo}="Compra"))*{mov_impuestos})').number_format = CURRENCY_FMT
        ws.cell(row=r, column=4, value=f'=SUMPRODUCT((YEAR({mov_fecha})=$B$3)*(MONTH({mov_fecha})={i})*(({mov_tipo}="Gasto")+({mov_tipo}="Compra"))*{mov_retenciones})').number_format = CURRENCY_FMT
        ws.cell(row=r, column=5, value=f'=SUMPRODUCT((YEAR({mov_fecha})=$B$3)*(MONTH({mov_fecha})={i})*(({mov_tipo}="Ingreso")+({mov_tipo}="Venta"))*{mov_retenciones})').number_format = CURRENCY_FMT

    autosize(ws, headers, max_width=24)
    return ws


def build_reporte_periodo(wb, sheet_name, period_label):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = COLOR_REPORT
    sheet_title(ws, sheet_name, "Resumen ejecutivo del periodo - ingresos, gastos y resultado")

    headers = [period_label, "Ingresos", "Gastos", "Resultado", "Documentos Validados", "Documentos con Error"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)
    style_header_row(ws, start, len(headers), FILL_HEADER_REPORT)

    if period_label == "MES":
        etiquetas = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
                     "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        ws["A3"] = "Año del reporte"
        ws["B3"] = DEFAULT_REPORT_YEAR
    else:
        etiquetas = list(range(12))
        ws["A3"] = "Año inicial del rango"
        ws["B3"] = DEFAULT_REPORT_YEAR - 5
    ws["B3"].fill = PatternFill("solid", fgColor="FFF2CC")

    mov_fecha = col_range("MOVIMIENTOS", "Fecha")
    mov_tipo = col_range("MOVIMIENTOS", "Tipo")
    mov_estado = col_range("MOVIMIENTOS", "Estado")
    mov_total = col_range("MOVIMIENTOS", "Total")
    doc_fecha = col_range("DOCUMENTOS", "Fecha")
    doc_estado = col_range("DOCUMENTOS", "Estado")

    for i, label in enumerate(etiquetas, start=1):
        r = start + i
        if period_label == "MES":
            ws.cell(row=r, column=1, value=label)
            cond_mov = f'(YEAR({mov_fecha})=$B$3)*(MONTH({mov_fecha})={i})'
            cond_doc = f'(YEAR({doc_fecha})=$B$3)*(MONTH({doc_fecha})={i})'
        else:
            ws.cell(row=r, column=1, value=f'=$B$3+{i - 1}')
            cond_mov = f'YEAR({mov_fecha})=A{r}'
            cond_doc = f'YEAR({doc_fecha})=A{r}'
        ws.cell(row=r, column=2, value=f'=SUMPRODUCT(({cond_mov})*(({mov_tipo}="Ingreso")+({mov_tipo}="Venta"))*({mov_estado}="Pagado")*{mov_total})').number_format = CURRENCY_FMT
        ws.cell(row=r, column=3, value=f'=SUMPRODUCT(({cond_mov})*(({mov_tipo}="Gasto")+({mov_tipo}="Compra"))*({mov_estado}="Pagado")*{mov_total})').number_format = CURRENCY_FMT
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = CURRENCY_FMT
        ws.cell(row=r, column=5, value=f'=SUMPRODUCT(({cond_doc})*({doc_estado}="Validado"))')
        ws.cell(row=r, column=6, value=f'=SUMPRODUCT(({cond_doc})*({doc_estado}="Error"))')

    autosize(ws, headers, max_width=22)
    return ws


def build_log(wb):
    ws = wb.create_sheet("LOG")
    ws.sheet_properties.tabColor = COLOR_SYSTEM
    sheet_title(ws, "LOG", "Registro de auditoria - se completa automaticamente por VBA (Fase 4); estructura lista desde ya")

    headers = ["Fecha", "Hora", "Usuario", "Accion", "Modulo", "RegistroAfectado", "ValorAnterior", "ValorNuevo", "Resultado"]
    start = 4
    for c, h in enumerate(headers, 1):
        ws.cell(row=start, column=c, value=h)

    last_row = start + 200
    for i in range(start + 1, last_row + 1):
        ws.cell(row=i, column=1).number_format = DATE_FMT

    add_table(ws, "LOG", f"A{start}:I{last_row}", headers, header_row=start)
    return ws


# ---------------------------------------------------------------------------
# Nombres definidos, validaciones, formato condicional, protecciones
# ---------------------------------------------------------------------------

def define_named_ranges(wb):
    """Nombres definidos a nivel de libro.

    Los desplegables (Data Validation) de una hoja no pueden referenciar
    directamente una tabla ubicada en OTRA hoja (Excel lo rechaza al abrir el
    archivo). La solucion estandar es envolver esa referencia en un Nombre
    definido a nivel de libro y usar el Nombre como origen del desplegable.
    Por eso toda columna de tabla que alimenta un desplegable entre hojas
    tiene aqui su propio nombre "Lst_*", aunque ya exista como tabla.
    """
    names = {
        "Lst_TipoTercero": "\"Cliente,Proveedor,Empleado,Socio,Entidad publica,Otro\"",
        "Lst_TipoMovimiento": "\"Ingreso,Gasto,Compra,Venta,Otro\"",
        "Lst_EstadoMovimiento": "\"Pendiente,Pagado,Anulado\"",
        "Lst_EstadoDocumento": "\"Validado,Pendiente,Error,Revision Manual,Duplicado\"",
        "Lst_EstadoPagoCobro": "\"Pendiente,Aplicado,Anulado\"",
        "Lst_MetodoExtraccion": "\"XML,PDF,OCR,Manual,IA\"",
        "Lst_TipoArchivo": "\"PDF,JPG,PNG,XML,CSV,Otro\"",
        "Lst_Categoria": col_range("CATEGORIAS", "Categoria"),
        "Lst_MedioPago": col_range("MEDIOS_PAGO", "MedioPago"),
        "Lst_Cuenta": col_range("CUENTAS", "Cuenta"),
        "Lst_TerceroNombre": col_range("TERCEROS", "NombreRazonSocial"),
        "Lst_TipoDocumentoNombre": col_range("TIPOS_DOCUMENTOS", "TipoDocumento"),
    }
    for name, formula in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=formula)


def _dv(formula1, allow_blank=True):
    return DataValidation(type="list", formula1=formula1, allow_blank=allow_blank, showDropDown=False)


def apply_data_validations(wb):
    ws = wb["TERCEROS"]
    dv_tipo = _dv("Lst_TipoTercero")
    ws.add_data_validation(dv_tipo)
    dv_tipo.add(data_col_range("TERCEROS", "Tipo"))
    dv_estado = _dv('"Activo,Inactivo"')
    ws.add_data_validation(dv_estado)
    dv_estado.add(data_col_range("TERCEROS", "Estado"))

    ws = wb["MOVIMIENTOS"]
    dv_tipo = _dv("Lst_TipoMovimiento")
    ws.add_data_validation(dv_tipo)
    dv_tipo.add(data_col_range("MOVIMIENTOS", "Tipo"))
    dv_cat = _dv("Lst_Categoria")
    ws.add_data_validation(dv_cat)
    dv_cat.add(data_col_range("MOVIMIENTOS", "Categoria"))
    dv_mp = _dv("Lst_MedioPago")
    ws.add_data_validation(dv_mp)
    dv_mp.add(data_col_range("MOVIMIENTOS", "MedioPago"))
    dv_cta = _dv("Lst_Cuenta")
    ws.add_data_validation(dv_cta)
    dv_cta.add(data_col_range("MOVIMIENTOS", "Cuenta"))
    dv_est = _dv("Lst_EstadoMovimiento")
    ws.add_data_validation(dv_est)
    dv_est.add(data_col_range("MOVIMIENTOS", "Estado"))
    dv_terc = _dv("Lst_TerceroNombre")
    ws.add_data_validation(dv_terc)
    dv_terc.add(data_col_range("MOVIMIENTOS", "Tercero"))

    ws = wb["DOCUMENTOS"]
    dv_tipodoc = _dv("Lst_TipoDocumentoNombre")
    ws.add_data_validation(dv_tipodoc)
    dv_tipodoc.add(data_col_range("DOCUMENTOS", "TipoDocumento"))
    dv_est_doc = _dv("Lst_EstadoDocumento")
    ws.add_data_validation(dv_est_doc)
    dv_est_doc.add(data_col_range("DOCUMENTOS", "Estado"))
    dv_tipoarch = _dv("Lst_TipoArchivo")
    ws.add_data_validation(dv_tipoarch)
    dv_tipoarch.add(data_col_range("DOCUMENTOS", "TipoArchivo"))
    dv_metodo = _dv("Lst_MetodoExtraccion")
    ws.add_data_validation(dv_metodo)
    dv_metodo.add(data_col_range("DOCUMENTOS", "MetodoExtraccion"))
    dv_terc2 = _dv("Lst_TerceroNombre")
    ws.add_data_validation(dv_terc2)
    dv_terc2.add(data_col_range("DOCUMENTOS", "Tercero"))

    ws = wb["PAGOS"]
    dv_cta2 = _dv("Lst_Cuenta")
    ws.add_data_validation(dv_cta2)
    dv_cta2.add(data_col_range("PAGOS", "Cuenta"))
    dv_mp2 = _dv("Lst_MedioPago")
    ws.add_data_validation(dv_mp2)
    dv_mp2.add(data_col_range("PAGOS", "MedioPago"))
    dv_estpc = _dv("Lst_EstadoPagoCobro")
    ws.add_data_validation(dv_estpc)
    dv_estpc.add(data_col_range("PAGOS", "Estado"))

    ws = wb["COBROS"]
    dv_cta3 = _dv("Lst_Cuenta")
    ws.add_data_validation(dv_cta3)
    dv_cta3.add(data_col_range("COBROS", "Cuenta"))
    dv_mp3 = _dv("Lst_MedioPago")
    ws.add_data_validation(dv_mp3)
    dv_mp3.add(data_col_range("COBROS", "MedioPago"))
    dv_estpc2 = _dv("Lst_EstadoPagoCobro")
    ws.add_data_validation(dv_estpc2)
    dv_estpc2.add(data_col_range("COBROS", "Estado"))

    ws = wb["REGLAS_IA"]
    dv_activa = _dv('"Sí,No"')
    ws.add_data_validation(dv_activa)
    dv_activa.add(data_col_range("REGLAS_IA", "Activa"))
    dv_tipo_ia = _dv("Lst_TipoMovimiento")
    ws.add_data_validation(dv_tipo_ia)
    dv_tipo_ia.add(data_col_range("REGLAS_IA", "Tipo"))
    dv_categoria_ia = _dv("Lst_Categoria")
    ws.add_data_validation(dv_categoria_ia)
    dv_categoria_ia.add(data_col_range("REGLAS_IA", "Categoria"))
    dv_medio_ia = _dv("Lst_MedioPago")
    ws.add_data_validation(dv_medio_ia)
    dv_medio_ia.add(data_col_range("REGLAS_IA", "MedioPago"))
    dv_cuenta_ia = _dv("Lst_Cuenta")
    ws.add_data_validation(dv_cuenta_ia)
    dv_cuenta_ia.add(data_col_range("REGLAS_IA", "Cuenta"))
    dv_prioridad_ia = DataValidation(
        type="whole", operator="between", formula1="1", formula2="100", allow_blank=True
    )
    ws.add_data_validation(dv_prioridad_ia)
    dv_prioridad_ia.add(data_col_range("REGLAS_IA", "Prioridad"))


def apply_conditional_formats(wb):
    ws = wb["ASISTENTE"]
    ws.conditional_formatting.add(
        "B9",
        FormulaRule(
            formula=['$B$9="Sugerencia lista para revisar"'],
            fill=PatternFill("solid", fgColor="C6EFCE"),
        ),
    )
    ws.conditional_formatting.add(
        "B9",
        FormulaRule(
            formula=['OR($B$9="Conflicto: ajuste prioridad",$B$9="Sin coincidencia")'],
            fill=PatternFill("solid", fgColor="FFEB9C"),
        ),
    )

    ws = wb["TERCEROS"]
    tercero_last = SHEET_LAYOUT["TERCEROS"]["last"]
    ws.conditional_formatting.add(
        data_col_range("TERCEROS", "ClaveDuplicado"),
        FormulaRule(formula=[f'AND($L5<>"",COUNTIF($L$5:$L${tercero_last},$L5)>1)'], fill=PatternFill("solid", fgColor="FFC7CE")),
    )

    ws = wb["DOCUMENTOS"]
    ws.conditional_formatting.add(
        data_col_range("DOCUMENTOS", "Estado"),
        FormulaRule(formula=['$L5="Duplicado"'], fill=PatternFill("solid", fgColor="FFC7CE")),
    )
    ws.conditional_formatting.add(
        data_col_range("DOCUMENTOS", "Estado"),
        FormulaRule(formula=['$L5="Error"'], fill=PatternFill("solid", fgColor="FFEB9C")),
    )
    ws.conditional_formatting.add(
        data_col_range("DOCUMENTOS", "Estado"),
        FormulaRule(formula=['$L5="Validado"'], fill=PatternFill("solid", fgColor="C6EFCE")),
    )

    ws = wb["MOVIMIENTOS"]
    ws.conditional_formatting.add(
        data_col_range("MOVIMIENTOS", "Estado"),
        FormulaRule(formula=['$P5="Pendiente"'], fill=PatternFill("solid", fgColor="FFEB9C")),
    )
    ws.conditional_formatting.add(
        data_col_range("MOVIMIENTOS", "Estado"),
        FormulaRule(formula=['$P5="Anulado"'], fill=PatternFill("solid", fgColor="FFC7CE")),
    )


def protect_formula_columns(wb):
    """Bloquea solo las celdas de formula (Total, SaldoActual, ClaveDuplicado, dashboard).
    El resto de celdas queda desbloqueado para edicion libre. La hoja se protege sin
    contrasena (proteccion de estructura, no de seguridad) para permitir edicion total
    por el propietario desde el menu Revisar > Desproteger hoja."""
    formula_ranges = {
        "CUENTAS": ["G"],
        "TERCEROS": ["L"],
        "MOVIMIENTOS": ["O"],
        "REGLAS_IA": ["K"],
    }
    for sheet_name, cols in formula_ranges.items():
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                cell.protection = Protection(locked=False)
        for col in cols:
            for r in range(1, ws.max_row + 1):
                ws[f"{col}{r}"].protection = Protection(locked=True)
        # Nota: la proteccion de hoja queda desactivada por defecto. Si se activa
        # (Revisar > Proteger hoja), Excel bloquea el autoexpandido de la Tabla al
        # escribir una fila nueva debajo (limitacion nativa de Excel), lo que rompe
        # el requisito de extensibilidad sin fricciones. Las celdas de formula ya
        # quedan marcadas como "bloqueadas" para cuando el propietario decida
        # proteger puntualmente (por ejemplo, antes de compartir el archivo).


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "dist" / "Sistema_Financiero_Empresa_2026.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    result = build(out)
    print(f"Libro generado: {result}")
