from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
from dateutil.relativedelta import relativedelta


def calcular_cuota(capital, tasa_mensual, plazo):
    if tasa_mensual == 0:
        return capital / plazo
    return capital * tasa_mensual / (1 - (1 + tasa_mensual) ** -plazo)


def generar_excel(filas, capital, tasa, plazo, cuota, moneda, nombre):
    wb = Workbook()
    ws = wb.active
    ws.title = "Amortización"

    borde   = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"),  bottom=Side(style="thin"))
    centro  = Alignment(horizontal="center")
    derecha = Alignment(horizontal="right")

    # Título
    ws.merge_cells("A1:G1")
    ws["A1"] = f"TABLA DE AMORTIZACIÓN — {nombre.upper()}"
    ws["A1"].font      = Font(name="Arial", bold=True, size=13, color="1F3864")
    ws["A1"].alignment = centro

    ws.merge_cells("A2:G2")
    ws["A2"] = (f"Capital: {moneda} {capital:,.0f}   |   Tasa anual: {tasa:.2f}%"
                f"   |   Plazo: {plazo} meses   |   Cuota fija: {moneda} {cuota:,.0f}")
    ws["A2"].font      = Font(name="Arial", size=9, color="555555")
    ws["A2"].alignment = centro

    # Encabezados
    cols = [("N°",6), ("Fecha",13), ("Saldo Inicial",17), ("Cuota",14),
            ("Interés",13), ("Capital",13), ("Saldo Final",14)]
    for i, (h, w) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(3, i, h)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill      = PatternFill("solid", fgColor="1F3864")
        c.alignment = centro
        c.border    = borde

    # Filas de datos
    for mes, fecha, s_ini, cuota_r, interes, cap, s_fin in filas:
        row     = mes + 3
        relleno = PatternFill("solid", fgColor="D6E4F0" if mes % 2 == 0 else "FFFFFF")
        datos   = [mes, fecha, s_ini, cuota_r, interes, cap, s_fin]
        fmts    = ["0", "DD/MM/YYYY", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0"]
        alns    = [centro, centro, derecha, derecha, derecha, derecha, derecha]
        for col, (v, f, a) in enumerate(zip(datos, fmts, alns), 1):
            c = ws.cell(row, col, v)
            c.font, c.fill, c.alignment, c.border = (
                Font(name="Arial", size=10), relleno, a, borde)
            c.number_format = f

    # Totales
    tr = plazo + 4
    ws.merge_cells(f"A{tr}:B{tr}")
    for col in range(1, 8):
        c = ws.cell(tr, col)
        c.fill   = PatternFill("solid", fgColor="D9D9D9")
        c.border = borde
        c.font   = Font(name="Arial", bold=True, size=10)
        if col == 1:
            c.value, c.alignment = "TOTAL", centro
        elif col in [4, 5, 6]:
            letra           = get_column_letter(col)
            c.value         = f"=SUM({letra}4:{letra}{tr-1})"
            c.number_format = "#,##0"
            c.alignment     = derecha

    ws.freeze_panes = "A4"
    archivo = f"amortizacion_{nombre.replace(' ', '_')}.xlsx"
    wb.save(archivo)
    return archivo


def run():
    print("\n=== TABLA DE AMORTIZACIÓN (Sistema Francés) ===\n")

    capital = float(input("  Capital            : ").replace(",", ""))
    tasa    = float(input("  Tasa anual (%)     : "))
    plazo   = int(input(  "  Plazo (meses)      : "))
    nombre  = input(      "  Nombre préstamo    : ") or "Prestamo"
    moneda  = input(      "  Moneda (CLP/USD)   : ") or "CLP"

    tasa_mensual = tasa / 100 / 12
    cuota = calcular_cuota(capital, tasa_mensual, plazo)

    print(f"\n  Cuota mensual      : {moneda} {cuota:,.0f}")
    print(f"  Total a pagar      : {moneda} {cuota * plazo:,.0f}")
    print(f"  Total intereses    : {moneda} {cuota * plazo - capital:,.0f}\n")

    filas = []
    saldo = capital
    fecha = date.today().replace(day=1) + relativedelta(months=1)

    for mes in range(1, plazo + 1):
        interes = saldo * tasa_mensual
        cap     = cuota - interes
        if mes == plazo:
            cap     = saldo
            cuota_r = cap + interes
        else:
            cuota_r = cuota
        s_fin = max(saldo - cap, 0)
        filas.append((mes, fecha, saldo, cuota_r, interes, cap, s_fin))
        saldo  = s_fin
        fecha += relativedelta(months=1)

    archivo = generar_excel(filas, capital, tasa, plazo, cuota, moneda, nombre)
    print(f"  ✅ Excel generado: {archivo}")