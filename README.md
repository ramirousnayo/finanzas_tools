# 💰 Finanzas Tools

Herramientas financieras de uso interno desarrolladas en Python.
Diseñadas para automatizar tareas manuales del área de finanzas.

---

## 🚀 Instalación
```bash
# Clonar el repositorio
git clone https://github.com/ramirousnayo/finanzas_tools.git
cd finanzas_tools

# Crear entorno virtual
python -m venv venv
source venv/bin/activate

# Instalar dependencias
pip install -r requirements.txt
```

## ▶️ Uso
```bash
python main.py
```

---

## 🧩 Módulos disponibles

### 1. Tabla de Amortización

Genera una tabla de amortización en Excel con los siguientes sistemas:

| Sistema | Descripción |
|---|---|
| Francés | Cuota fija durante todo el plazo |
| Alemán | Capital fijo, intereses decrecientes |
| Americano | Intereses periódicos, capital al final |
| Bullet | Pago único al vencimiento |
| Leasing | Cuota fija con opción de compra al final |

### 2. Aging de Cobranzas

Lee un archivo Excel exportado del ERP y genera reporte por tramos de vencimiento.

| Tramo | Descripción |
|---|---|
| Al día | Sin vencimiento |
| 0-30 días | Vencimiento reciente |
| 31-60 días | Seguimiento requerido |
| 61-90 días | ⚠️ Atención |
| +90 días | 🔴 Crítico |

**Columnas requeridas en el archivo fuente:**

| Columna | Descripción |
|---|---|
| cliente | Nombre o RUT del cliente |
| factura | Número de factura |
| fecha_venc | Fecha de vencimiento (DD/MM/YYYY) |
| monto | Monto adeudado |

### 3. Flujo de Caja Mensual

Genera una plantilla Excel para proyectar el flujo de caja mensual y detecta meses con saldo negativo.

| Categoría | Conceptos |
|---|---|
| Ingresos | Cobranza clientes, Otros ingresos |
| Egresos | Pago proveedores, Sueldos, Préstamos, Otros fijos |

**Funciones:**
- Opción 1 — Genera plantilla en blanco para completar
- Opción 2 — Procesa plantilla completada y alerta saldos negativos

### 4. Conciliación Bancaria

Cruza el extracto del banco contra el libro interno por fecha y monto.

**Columnas requeridas en ambos archivos:**

| Columna | Descripción |
|---|---|
| fecha | Fecha del movimiento (DD/MM/YYYY) |
| monto | Monto del movimiento |
| descripcion | Descripción o glosa |

**Output — Excel con 4 hojas:**
- Resumen general con totales
- Movimientos conciliados ✅
- Solo en banco ⚠️
- Solo en libro interno 🔴### 4. Conciliación Bancaria

Cruza el extracto del banco contra el libro interno por fecha y monto.

**Columnas requeridas en ambos archivos:**

| Columna | Descripción |
|---|---|
| fecha | Fecha del movimiento (DD/MM/YYYY) |
| monto | Monto del movimiento |
| descripcion | Descripción o glosa |

**Output — Excel con 4 hojas:**
- Resumen general con totales
- Movimientos conciliados ✅
- Solo en banco ⚠️
- Solo en libro interno 🔴

### 5. Gestor de Archivos

Organiza, renombra y genera índice Excel de archivos financieros.

**Estructura generada:**
```
destino/
└── 2025/
    └── 03_Marzo/
        ├── Contratos/
        ├── Bancos/
        ├── Respaldos/
        ├── Prestamos/
        └── Facturas/
```

**Convención de nombres:**
```
CATEGORIA_YYYYMM_nombre_original.ext
```

**Output:**
- Archivos organizados en carpetas por categoría
- Índice Excel con detalle y resumen por categoría

---

## 📁 Estructura del proyecto
```
finanzas_tools/
├── data/
│   ├── archivos_prueba/
│   ├── cobranzas_prueba.xlsx
│   ├── extracto_banco_marzo.xlsx
│   ├── libro_interno_marzo.xlsx
│   └── flujo_caja_2025.xlsx
├── modules/
│   ├── __init__.py
│   ├── amortizacion.py
│   ├── cobranzas.py
│   ├── conciliacion.py
│   ├── flujo_caja.py
│   └── gestor_archivos.py
├── main.py
├── requirements.txt
└── README.md
```
---

## 🛠️ Stack

- Python 3
- openpyxl
- pandas
- python-dateutil