from modules.amortizacion import run as run_amortizacion
from modules.cobranzas import run as run_cobranzas

while True:
    print("\n╔══════════════════════════════════╗")
    print("║       FINANZAS TOOLS  v1.0       ║")
    print("╚══════════════════════════════════╝")
    print("\n  [1] Tabla de amortización")
    print("  [2] Aging de cobranzas")
    print("  [0] Salir\n")

    opcion = input("  Selecciona una opción: ").strip()

    if opcion == "1":
        run_amortizacion()
    elif opcion == "2":
        run_cobranzas()
    elif opcion == "0":
        print("\n  Hasta luego 👋\n")
        break
    else:
        print("\n  Opción no válida.\n")