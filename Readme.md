# Huella Social — Dashboard Cuenta Satélite CAC Chile

Dashboard interactivo que estima el aporte de las Cooperativas de Ahorro y Crédito (CAC) al PIB chileno, construido como parte de una memoria de título en Ingeniería Civil Industrial en la Universidad de los Andes (2026).

**Autores:** Ignacio Ureta · Antonio Ruiz Tagle · **Supervisores:** Sebastián Cea · Joaquín Fernández

---

## Contenido

| Archivo | Descripción |
|---|---|
| `Dashboard_cuentasatelite.py` | Script Python que genera el dashboard |
| `dashboard_huellasocial.html` | Dashboard listo para abrir en el navegador |
| `HuellaSocial_Consolidado.xlsx` | Base de datos financiera (CMF 2013–2025 · DAES) |

## Origen de los datos y cálculos (Cuadros 5.2–5.9 de la memoria)

Las tablas de la cuenta financiera (Cuadros 5.4, 5.6, 5.8) y las de producción
(Cuadros 5.2, 5.5, 5.7) se calculan como fórmulas en `HuellaSocial_Consolidado.xlsx`
(hojas "Agregados DAES", "Agregados CMF", "Agregado Total"), y luego se
transcriben manualmente a las tablas LaTeX de la memoria. No existe un script
que genere esos cuadros directamente — solo este dashboard, que sí lee el
Excel en vivo para el panel interactivo (§5.4 de la memoria).

Cruce de entidades entre fuentes: por RUT normalizado (sin puntos, guiones
ni espacios), nunca por nombre, ya que hay cooperativas que cambiaron de
razón social durante el período (ej. Lautaro Rosas → Coonfía).

## Uso

```bash
pip install pandas numpy openpyxl
python Dashboard_cuentasatelite.py
```

El script lee el Excel y genera el HTML en el mismo directorio. No requiere servidor ni conexión a internet.
## Entorno de ejecución

- Python 3.11.5
- Sistema operativo: Windows

## Metodología

Sigue el Manual ONU-TSE 2018. El VAB bruto (B1g) se estima como `P1 − P2`, donde P1 son los ingresos operacionales y P2 = α × P1 con α = 0,3776 (MIP Chile 2018, Sector 94). Fuentes: CMF, DAES, Banco Central.