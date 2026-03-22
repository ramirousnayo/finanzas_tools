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

### 6. Estimador de Cobranzas y Recaudación

Proyecta la recaudación esperada aplicando probabilidades de cobro
por tramo de vencimiento sobre la cartera actual.

**Probabilidades aplicadas:**

| Tramo | Probabilidad | Interpretación |
|---|---|---|
| Al día | 95% | Alta probabilidad — cliente al corriente |
| 0-30 días | 85% | Buena probabilidad — atraso leve |
| 31-60 días | 65% | Riesgo moderado — seguimiento requerido |
| 61-90 días | 40% | Riesgo alto — gestión activa necesaria |
| +90 días | 15% | Riesgo crítico — posible incobrable |

**Output — Excel con 4 hojas:**
- Detalle con monto estimado por factura
- Proyección mensual de recaudación
- Estimación por cliente con probabilidad promedio
- Tabla de probabilidades aplicadas

### 7. Análisis de Proveedores

Ranking, concentración, proveedores críticos e historial mensual de pagos.

**Columnas requeridas en el archivo Excel:**

| Columna | Descripción |
|---|---|
| proveedor | Nombre del proveedor |
| rut | RUT del proveedor |
| factura | Número de factura |
| fecha_factura | Fecha emisión (DD/MM/YYYY) |
| fecha_pago | Fecha de pago (DD/MM/YYYY) — vacío si pendiente |
| monto | Monto de la factura |
| credito_dias | Días de crédito pactado (default 30) |

**Estados detectados automáticamente:**
- ✅ Pagado a tiempo
- ⚠️ Pagado con atraso
- 🔵 Pendiente vigente
- 🔴 Pendiente vencido

**Output — Excel con 4 hojas:**
- Ranking por monto con concentración acumulada
- Proveedores críticos con facturas vencidas
- Historial mensual por estado
- Detalle completo de facturas

---

## 📁 Estructura del proyecto
```
finanzas_tools/
├── data/
│   ├── archivos_prueba/
│   ├── cobranzas_prueba.xlsx
│   ├── extracto_banco_marzo.xlsx
│   ├── libro_interno_marzo.xlsx
│   ├── flujo_caja_2025.xlsx
│   └── proveedores_prueba.xlsx
├── modules/
│   ├── __init__.py
│   ├── amortizacion.py
│   ├── analisis_proveedores.py
│   ├── cobranzas.py
│   ├── conciliacion.py
│   ├── estimador_cobranzas.py
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