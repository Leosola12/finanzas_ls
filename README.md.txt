# 📊 Finanzas Personales - MVP

Aplicación en Streamlit para analizar gastos e ingresos personales a partir de un Excel, con una capa de inteligencia basada en mapeo dinámico de subcategorías.

---

## 🚀 Características

### 📥 Input simple
- Subida de archivo Excel
- Detección automática de hojas (Gastos / Ingresos)

---

### 🧠 Mapeo inteligente (feature principal)
- Clasificación de subcategorías en:
  - Área (ALIM, VIV, TRANS, etc.)
  - Naturaleza (FIJO, NEC, DISC)
  - Controlable (Sí / No)
- Editor interactivo en sidebar
- Detección automática de subcategorías nuevas
- Persistencia mediante CSV

---

### 💰 Análisis financiero

#### KPIs:
- Ingresos
- Gastos
- Ahorro teórico
- Ahorro real (input manual)

#### Ajustes manuales:
- Ahorro real del mes
- Pago de deudas (ej: tarjeta)

---

### 📊 Dashboard

- Gasto por naturaleza
- Gasto controlable vs no controlable
- Top gastos
- Fugas (gastos controlables)

---

### 🧠 Insights automáticos
- Detección de subcategorías sin clasificar
- Alertas sobre diferencias entre ahorro teórico y real

---

## 🏗️ Estructura del Excel esperado

### Hoja 1: Gastos
Debe contener al menos:
- Subcategoria
- Monto (o Importe)

---

### Hoja 2: Ingresos
Debe contener:
- Monto (o Importe)

---

## ⚙️ Instalación

```bash
pip install -r requirements.txt