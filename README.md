# Huella Social  Dashboard Cuenta Satélite CAC

Plataforma de análisis para el sector de Cooperativas de Ahorro y Crédito (CAC) en Chile.
Desarrollada como parte de la memoria de título *'Plataforma de análisis para cooperativas de ahorro y crédito en Chile'*,
Universidad de los Andes  Ingeniería Civil Industrial.

**Autores:** Ignacio Ureta  Antonio Ruiz Tagle  **Prof. guía:** Sebastián Cea  **Prof. co-guía:** Joaquín Fernández

---

## Requisitos

Tener Python instalado (3.9 o superior). Luego instalar las librerías necesarias:

```bash
pip install plotly pandas openpyxl geopandas
```

> Si no tienes `geopandas` instalado y no hay `Regional.geojson` en la carpeta,
> el mapa no se renderizará pero el resto del dashboard funciona igual.

---

## Archivos necesarios

Todos los archivos deben estar en la misma carpeta `Dashboard/`:

| Archivo | Descripción |
|---|---|
| `dashboard_HuellaSocial.py` | Script principal del dashboard |
| `Consolidado_cooperativas.xlsx` | Panel DAES + SII + CMF + Ley 19.862 (39 CAC activas) |
| `consolidado_CAC_balances_resultados_v2.xlsx` | Estados de resultados y balances históricos + CMF Feb-2026 |
| `regional_shp.zip` | Shapefile de regiones de Chile (mapa coroplético) |

---

## Cómo ejecutar

Desde la carpeta `Dashboard/`, ejecutar en la terminal:

```bash
python dashboard_HuellaSocial.py
```

El script genera el archivo `dashboard_HuellaSocial.html` y lo abre automáticamente en el navegador.

---

## Contenido del dashboard

| Pestaña | Descripción |
|---|---|
|  P1B1g Producción y VAB | Evolución del Valor Agregado Bruto del sector |
|  D1B2g Generación Ingreso | Remuneraciones y excedente bruto de explotación |
|  F2F4 Cuenta Financiera | Depósitos, colocaciones, mora y solvencia |
|  Por Entidad | Series históricas por cooperativa |
|  Mapa Chile | Distribución territorial por región |
|  Geografía | Barras regionales: socios, VAB, trabajadores |
|  Panel DAES | Tabla integrada de las 39 CAC activas |
|  Glosario de Variables | Definición de cada código SCN/ONU TSE/CIRIEC |
|  Brechas de Datos | Cuadro 4.5: disponibilidad de variables vs. España, Portugal y Polonia |

---

## Fuentes de datos

- **DAES**  Directorio de cooperativas activas
- **SII**  Series de actividad económica por RUT
- **CMF BEST Platform**  Estados financieros auditados (7 CAC supervisadas, Feb-2026)
- **Ley 19.862**  Transferencias públicas históricas
- **MIP Chile 2018**  Matriz Insumo-Producto (α = 0,3776, sector 94)
- **BCN**  Shapefile de regiones de Chile

---

## Marco metodológico

El dashboard implementa la metodología de Cuentas Satélite del **SCN 2025**,
el **Manual ONU TSE (2018)** y el **Manual CIRIEC (2007)**,
estimando el Valor Agregado Bruto (VAB) del sector CAC mediante un método híbrido:
método directo para las 7 CAC supervisadas por CMF y método indirecto por tramos SII para el resto.
