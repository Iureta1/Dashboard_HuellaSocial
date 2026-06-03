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

## Uso

```bash
pip install pandas numpy openpyxl
python Dashboard_cuentasatelite.py
```

El script lee el Excel y genera el HTML en el mismo directorio. No requiere servidor ni conexión a internet.

## Metodología

Sigue el Manual ONU-TSE 2018. El VAB bruto (B1g) se estima como `P1 − P2`, donde P1 son los ingresos operacionales y P2 = α × P1 con α = 0,3776 (MIP Chile 2018, Sector 94). Fuentes: CMF, DAES, Banco Central.