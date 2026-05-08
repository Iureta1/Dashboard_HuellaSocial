# Dashboard Huella Social Cooperativas CAC Chile

Dashboard interactivo para el análisis financiero, territorial y social de cooperativas de ahorro y crédito (CAC) en Chile.

El proyecto utiliza información de:

- DAES
- CMF
- SII

y genera un dashboard HTML interactivo completamente offline.

---

# Requisitos

Instalar Python 3.10 o superior.

Instalar librerías necesarias:

```bash
pip install pandas numpy plotly openpyxl pyshp
```

---

# Archivos necesarios

El repositorio incluye:

- Script principal Python
- Bases de datos Excel
- Dashboard HTML generado
- Shapefile regional comprimido (`regional_shp.zip`)

---

# Ejecución

Desde la carpeta del proyecto ejecutar:

```bash
python Dashboard_huella_v5.py
```

El script:

1. Extrae automáticamente el shapefile regional
2. Genera el GeoJSON
3. Construye el dashboard interactivo
4. Exporta:

```text
dashboard_huella_v5.html
```

---

# Visualización

Abrir el archivo:

```text
dashboard_huella_v5.html
```

en cualquier navegador web.

---

# Tecnologías utilizadas

- Python
- Pandas
- Plotly
- pyshp (lectura de shapefiles)
- HTML/CSS/JavaScript

---

# Autor

Ignacio Ureta

Memoria de título — Ingeniería Civil Industrial  
Universidad de los Andes