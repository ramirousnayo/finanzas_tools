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

---

## 📁 Estructura del proyecto
```
finanzas_tools/
├── data/
│   └── cobranzas_prueba.xlsx
├── modules/
│   ├── __init__.py
│   ├── amortizacion.py
│   └── cobranzas.py
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