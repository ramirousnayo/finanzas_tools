from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
from dateutil.relativedelta import relativedelta


# ── Estilos reutilizables ──────────────────────────────────────────

def _borde():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex):
    return PatternFill("solid", fgColor=hex)

CENTRO  = Alignment(horizontal="center")
DERECHA = Alignment(horizontal="right")


# ── Cálculos ───────────────────────────────────────────────────────

def cuota_frances(capital, tasa_mensual, plazo):
    """Cuota fija — Sistema Francés."""
    if tasa_mensual == 0:
        return capital / plazo
    return capital * tasa_mensual / (1 - (1 + tasa_mensual) ** -plazo)


def filas_frances(capital, tasa_mensual, plazo, fecha_inicio):
    cuota = cuota_frances(capital, tasa_mensual, plazo)
    filas, saldo, fecha = [], capital, fecha_inicio
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
    return filas, cuota


def filas_aleman(capital, tasa_mensual, plazo, fecha_inicio):
    """Cuota de capital fija — Sistema Alemán."""
    cap_fijo = capital / plazo
    filas, saldo, fecha = [], capital, fecha_inicio
    for mes in range(1, plazo + 1):
        interes = saldo * tasa_mensual
        cuota_r = cap_fijo + interes
        s_fin   = max(saldo - cap_fijo, 0)
        filas.append((mes, fecha, saldo, cuota_r, interes, cap_fijo, s_fin))
        saldo  = s_fin
        fecha += relativedelta(months=1)
    return filas


def filas_americano(capital, tasa_mensual, plazo, fecha_inicio):
    """Intereses periódicos + capital al final — Sistema Americano."""
    interes = capital * tasa_mensual
    filas, fecha = [], fecha_inicio
    for mes in range(1, plazo + 1):
        if mes < plazo:
            cuota_r = interes
            cap     = 0
            s_fin   = capital
        else:
            cuota_r = capital + interes
            cap     = capital
            s_fin   = 0
        filas.append((mes, fecha, capital if mes < plazo else capital, cuota_r, interes, cap, s_fin))
        fecha += relativedelta(months=1)
    return filas


def filas_bullet(capital, tasa_mensual, plazo, fecha_inicio):
    """Pago único al vencimiento — Sistema Bullet."""
    interes_total = capital * ((1 + tasa_mensual) ** plazo - 1)
    filas, fecha  = [], fecha_inicio
    for mes in range(1, plazo + 1):
        interes_acum = capital * ((1 + tasa_mensual) ** mes - 1)
        if mes < plazo:
            filas.append((mes, fecha, capital, 0, 0, 0, capital))
        else:
            pago = capital + interes_total
            filas.append((mes, fecha, capital, pago, interes_total, capital, 0))
        fecha += relativedelta(months=1)
    return filas


# ── Excel ──────────────────────────────────────────────────────────

def generar_excel(filas, capital, tasa, plazo, moneda, nombre, sistema):
    wb = Workbook()
    ws = wb.active
    ws.title = sistema

    # Título
    ws.merge_cells("A1:G1")
    ws["A1"] = f"TABLA DE AMORTIZACIÓN — {sistema.upper()} — {nombre.upper()}"
    ws["A1"].font      = Font(name="Arial", bold=True, size=13, color="1F3864")
    ws["A1"].alignment = CENTRO

    total_cuotas = sum(f[3] for f in filas)
    total_int    = sum(f[4] for f in filas)

    ws.merge_cells("A2:G2")
    ws["A2"] = (f"Capital: {moneda} {capital:,.0f}   |   Tasa anual: {tasa:.2f}%   |   "
                f"Plazo: {plazo} meses   |   Total pagado: {moneda} {total_cuotas:,.0f}   |   "
                f"Total intereses: {moneda} {total_int:,.0f}")
    ws["A2"].font      = Font(name="Arial", size=9, color="555555")
    ws["A2"].alignment = CENTRO

    # Encabezados
    cols = [("N°",6), ("Fecha",13), ("Saldo Inicial",17), ("Cuota",14),
            ("Interés",13), ("Capital",13), ("Saldo Final",14)]
    for i, (h, w) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(3, i, h)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill      = _fill("1F3864")
        c.alignment = CENTRO
        c.border    = _borde()

    # Filas
    for mes, fecha, s_ini, cuota_r, interes, cap, s_fin in filas:
        row     = mes + 3
        relleno = _fill("D6E4F0" if mes % 2 == 0 else "FFFFFF")
        datos   = [mes, fecha, s_ini, cuota_r, interes, cap, s_fin]
        fmts    = ["0", "DD/MM/YYYY", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0"]
        alns    = [CENTRO, CENTRO, DERECHA, DERECHA, DERECHA, DERECHA, DERECHA]
        for col, (v, f, a) in enumerate(zip(datos, fmts, alns), 1):
            c = ws.cell(row, col, v)
            c.font, c.fill, c.alignment, c.border = (
                Font(name="Arial", size=10), relleno, a, _borde())
            c.number_format = f

    # Totales
    tr = plazo + 4
    ws.merge_cells(f"A{tr}:B{tr}")
    for col in range(1, 8):
        c = ws.cell(tr, col)
        c.fill, c.border = _fill("D9D9D9"), _borde()
        c.font = Font(name="Arial", bold=True, size=10)
        if col == 1:
            c.value, c.alignment = "TOTAL", CENTRO
        elif col in [4, 5, 6]:
            letra           = get_column_letter(col)
            c.value         = f"=SUM({letra}4:{letra}{tr-1})"
            c.number_format = "#,##0"
            c.alignment     = DERECHA

    ws.freeze_panes = "A4"
    archivo = f"amortizacion_{sistema.lower()}_{nombre.replace(' ', '_')}.xlsx"
    wb.save(archivo)
    return archivo


# ── Menú ───────────────────────────────────────────────────────────

def run():
    print("\n=== TABLA DE AMORTIZACIÓN ===\n")
    print("  Sistemas disponibles:")
    print("  [1] Francés   — cuota fija")
    print("  [2] Alemán    — capital fijo")
    print("  [3] Americano — intereses periódicos + capital al final")
    print("  [4] Bullet    — pago único al vencimiento\n")

    sistema = input("  Selecciona sistema [1-4]: ").strip()
    if sistema not in ["1", "2", "3", "4"]:
        print("  Opción no válida.")
        return

    capital = float(input("  Capital            : ").replace(",", ""))
    tasa    = float(input("  Tasa anual (%)     : "))
    plazo   = int(input(  "  Plazo (meses)      : "))
    nombre  = input(      "  Nombre préstamo    : ") or "Prestamo"
    moneda  = input(      "  Moneda (CLP/USD)   : ") or "CLP"

    tasa_mensual = tasa / 100 / 12
    fecha_inicio = date.today().replace(day=1) + relativedelta(months=1)

    nombres_sistema = {
        "1": "Francés", "2": "Alemán", "3": "Americano", "4": "Bullet"
    }
    nombre_sistema = nombres_sistema[sistema]

    if sistema == "1":
        filas, cuota = filas_frances(capital, tasa_mensual, plazo, fecha_inicio)
        print(f"\n  Cuota mensual      : {moneda} {cuota:,.0f}")
    elif sistema == "2":
        filas = filas_aleman(capital, tasa_mensual, plazo, fecha_inicio)
        print(f"\n  Capital por cuota  : {moneda} {capital/plazo:,.0f}")
    elif sistema == "3":
        filas = filas_americano(capital, tasa_mensual, plazo, fecha_inicio)
        print(f"\n  Interés mensual    : {moneda} {capital * tasa_mensual:,.0f}")
    elif sistema == "4":
        filas = filas_bullet(capital, tasa_mensual, plazo, fecha_inicio)
        pago_final = capital * ((1 + tasa_mensual) ** plazo)
        print(f"\n  Pago único final   : {moneda} {pago_final:,.0f}")

    total = sum(f[3] for f in filas)
    interes_total = sum(f[4] for f in filas)
    print(f"  Total a pagar      : {moneda} {total:,.0f}")
    print(f"  Total intereses    : {moneda} {interes_total:,.0f}\n")

    archivo = generar_excel(filas, capital, tasa, plazo, moneda, nombre, nombre_sistema)
    print(f"  ✅ Excel generado: {archivo}\n")