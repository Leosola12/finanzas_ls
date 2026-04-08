# finanzas_ls

## 📥 Formato de datos esperado

Esta herramienta está diseñada para trabajar con el formato de reportes exportados por **TOTO BOT**. Aunque puede readaptarse (ver al final del readme)

👉 Sitio oficial de la herramienta: https://soytoto.com/ 

## DISCLAIMER: 

No soy asociado a TOTO BOT, ni es una recomendación de uso de la misma. Sólo se explica a fines de poder reutilizar el script para otro tipo de formatos. Desarrollé esta herramienta sólo a fines de mi uso personal


El archivo Excel base debe contener al menos dos hojas:

---

### 📄 Hoja 1: Gastos

| Fecha | Subcategoría | Monto |
| ----- | ------------ | ----- |

* **Fecha**: fecha del gasto
* **Subcategoría**: categoría específica (ej: Supermercado, Delivery, Transporte)
* **Monto**: importe del gasto

---

### 📄 Hoja 2: Ingresos

| Fecha | Tipo de ingreso | Monto |
| ----- | --------------- | ----- |

* **Fecha**: fecha del ingreso
* **Tipo de ingreso**: origen del ingreso (ej: Sueldo, Freelance, Venta)
* **Monto**: importe

---

⚠️ Importante:

* Los nombres de columnas pueden variar levemente (la app intenta detectarlos automáticamente)
* Las subcategorías deben ser consistentes para aprovechar el análisis

---

## 🗂️ Archivos de mapeo (opcional pero recomendado)

Los archivos de mapeo permiten **agregar una capa de interpretación** sobre los datos.

Sin mapping → ves datos
Con mapping → entendés tus finanzas

---

# 📊 mapeo_gastos.csv

## 📁 Estructura

| Subcategoria | Area | Naturaleza | Controlable |
| ------------ | ---- | ---------- | ----------- |

---

## 🧠 Definición de cada campo

### 🔹 Subcategoria

Debe coincidir con la subcategoría del Excel.

Ejemplos:

* Supermercado
* Delivery
* Transporte

---

### 🔹 Area

Agrupación más general (opcional, más organizativa que analítica)

Ejemplos:

* Alimentación
* Transporte
* Vivienda

---

### 🔹 Naturaleza (CLAVE)

Define el tipo de gasto desde el punto de vista financiero:

#### 🧱 FIJO

Gastos recurrentes y difíciles de modificar en el corto plazo

Ejemplos:

* Alquiler
* Internet
* Suscripciones

---

#### 🧠 NEC (Necesario)

Gastos variables pero necesarios para vivir

Ejemplos:

* Supermercado
* Transporte
* Farmacia

---

#### 🎯 DISC (Discrecional)

Gastos opcionales o evitables

Ejemplos:

* Delivery
* Salidas
* Compras impulsivas

---

👉 Esta clasificación se usa para calcular métricas como:

* costo de vida
* nivel de gasto discrecional

---

### 🔹 Controlable

Indica si el gasto puede ser reducido activamente:

* **Si** → se puede ajustar (ej: delivery, compras)
* **No** → difícil de modificar (ej: alquiler)

---

## 🎯 Ejemplo

```csv
Subcategoria,Area,Naturaleza,Controlable
Supermercado,Alimentacion,NEC,Si
Delivery,Alimentacion,DISC,Si
Alquiler,Vivienda,FIJO,No
Internet,Vivienda,FIJO,No
```

---

# 💰 mapeo_ingresos.csv

## 📁 Estructura

| Subcategoria | Origen | Estabilidad |
| ------------ | ------ | ----------- |

---

## 🧠 Definición

### 🔹 Subcategoria

Debe coincidir con el tipo de ingreso del Excel

Ej:

* Sueldo
* Freelance

---

### 🔹 Origen

Clasificación general del ingreso

Ej:

* Laboral
* Inversiones
* Otros

---

### 🔹 Estabilidad (CLAVE)

Define la previsibilidad del ingreso:

#### 🟢 FIJO

Ingreso estable y recurrente

Ej:

* Sueldo

---

#### 🟡 VARIABLE

Ingreso irregular o no garantizado

Ej:

* Freelance
* Ventas

---

👉 Esto permite analizar:

* dependencia de ingresos variables
* estabilidad financiera

---

## 🎯 Ejemplo

```csv
Subcategoria,Origen,Estabilidad
Sueldo,Laboral,FIJO
Freelance,Laboral,VARIABLE
Intereses,Inversiones,VARIABLE
```

---

# ⚠️ Notas importantes

* El mapping no es obligatorio, pero mejora mucho el análisis (es una traducción entre el archivo excel de toto bot y el script de python)
* Si no existe mapping, las subcategorías quedan como "SIN_MAPEAR" - Lo que significa que perdés cierto nivel de análisis
* Los textos no son sensibles a mayúsculas/minúsculas (la app los normaliza)

---

# 🧠 Concepto clave

👉 TOTO BOT organiza los datos y los categoriza (categorías y subcategorías)
👉 Esta herramienta los interpreta (en función de mis decisiones, intenciones y necesidades de información).

El mapping es el puente entre ambos.



## 🔄 Compatibilidad con otras planillas

Si bien esta herramienta está pensada para trabajar con reportes de TOTO BOT, también puede adaptarse a otras fuentes de datos.

Para que el análisis funcione correctamente, la planilla debe contener como mínimo:

### Para gastos:

* Fecha
* Subcategoría
* Monto

### Para ingresos:

* Fecha
* Tipo de ingreso
* Monto

---

⚠️ Importante:

La calidad del análisis depende directamente de la calidad de la categorización.

Subcategorías inconsistentes o genéricas (ej: "varios", "otros") reducen significativamente el valor de la herramienta.

---

💡 Recomendación:

Mantener una estructura clara y consistente de subcategorías para aprovechar al máximo las capacidades de análisis.

