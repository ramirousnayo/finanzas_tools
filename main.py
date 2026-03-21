from modules.amortizacion import run

print("\n╔══════════════════════════════════╗")
print("║       FINANZAS TOOLS  v1.0       ║")
print("╚══════════════════════════════════╝")
print("\n  [1] Tabla de amortización")
print("  [0] Salir\n")

opcion = input("  Selecciona una opción: ").strip()

if opcion == "1":
    run()
elif opcion == "0":
    print("\n  Hasta luego 👋\n")
else:
    print("\n  Opción no válida.\n")