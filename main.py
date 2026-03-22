from modules.amortizacion import run as run_amortizacion
from modules.cobranzas import run as run_cobranzas
from modules.flujo_caja import run as run_flujo
from modules.conciliacion import run as run_conciliacion
from modules.gestor_archivos import run as run_gestor

while True:
    print("\n╔══════════════════════════════════╗")
    print("║       FINANZAS TOOLS  v1.0       ║")
    print("╚══════════════════════════════════╝")
    print("\n  [1] Tabla de amortización")
    print("  [2] Aging de cobranzas")
    print("  [3] Flujo de caja")
    print("  [4] Conciliación bancaria")
    print("  [5] Gestor de archivos")
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
    elif opcion == "0":
        print("\n  Hasta luego 👋\n")
        break
    else:
        print("\n  Opción no válida.\n")