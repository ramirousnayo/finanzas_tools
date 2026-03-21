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

---

## 📁 Estructura del proyecto
```
finanzas_tools/
├── modules/
│   ├── __init__.py
│   └── amortizacion.py
├── main.py
├── requirements.txt
└── README.md
```

---

## 🛠️ Stack

- Python 3
- openpyxl
- python-dateutil