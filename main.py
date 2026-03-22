from modules.amortizacion import run as run_amortizacion
from modules.cobranzas import run as run_cobranzas
from modules.flujo_caja import run as run_flujo
from modules.conciliacion import run as run_conciliacion
from modules.gestor_archivos import run as run_gestor
from modules.estimador_cobranzas import run as run_estimador
from modules.analisis_proveedores import run as run_proveedores
from modules.indicadores_financieros import run as run_indicadores
from modules.lineas_credito import run as run_lineas
from modules.reporte_ejecutivo import run as run_reporte

while True:
    print("\n╔══════════════════════════════════╗")
    print("║       FINANZAS TOOLS  v1.0       ║")
    print("╚══════════════════════════════════╝")
    print("\n  [1] Tabla de amortización")
    print("  [2] Aging de cobranzas")
    print("  [3] Flujo de caja")
    print("  [4] Conciliación bancaria")
    print("  [5] Gestor de archivos")
    print("  [6] Estimador de cobranzas")
    print("  [7] Análisis de proveedores")
    print("  [8] Indicadores financieros")
    print("  [9] Líneas de crédito")
    print("  [10] Reporte ejecutivo")
    print("  [0] Salir\n")

    opcion = input("  Selecciona una opción: ").strip()

    if opcion == "1":
        run_amortizacion()
    elif opcion == "2":
        run_cobranzas()
    elif opcion == "3":
        run_flujo()
    elif opcion == "4":
        run_conciliacion()
    elif opcion == "5":
        run_gestor()
    elif opcion == "6":
        run_estimador()
    elif opcion == "7":
        run_proveedores()
    elif opcion == "8":
        run_indicadores()
    elif opcion == "9":
        run_lineas()
    elif opcion == "10":
        run_reporte()
    elif opcion == "0":
        print("\n  Hasta luego 👋\n")
        break
    else:
        print("\n  Opción no válida.\n")