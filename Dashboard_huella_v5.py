"""
╔══════════════════════════════════════════════════════════════════════╗
║  HUELLA SOCIAL — Dashboard Cuenta Satélite CAC  v5                  ║
║  Sector CAC completo: histórico 2014-2024 + CMF Feb-2026            ║
║  Mapa coroplético de Chile por región integrado                      ║
║  Variables alineadas SCN 2025 / Manual ONU TSE / CIRIEC             ║
║                                                                      ║
║  Ejecutar:  python dashboard_huella_v5.py                            ║
║  Genera:    dashboard_huella_v5.html  (se abre en el navegador)      ║
║  Requiere:  pip install plotly pandas openpyxl                       ║
║  Requiere:  Regional.geojson en la misma carpeta (BCN)               ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import pandas as pd
import numpy as np
import json, os, webbrowser

# Posicionarse en la carpeta del script para que los paths relativos funcionen
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# ── ARCHIVOS ─────────────────────────────────────────────────────────────────
F_CAC     = "consolidado_CAC_balances_resultados_v2.xlsx"
F_PANEL   = "Consolidado_cooperativas.xlsx"
F_GEOJSON = "Regional.geojson"
F_SHP     = "Regional.shp"
OUT       = "dashboard_huella_v5.html"

# ── PARÁMETROS SCN (§4.3.2 memoria) ─────────────────────────────────────────
ALPHA_CI = 0.3776   # CI/P — MIP Chile 2018 sector 94 (Intermediación financiera)

# ── PALETA ───────────────────────────────────────────────────────────────────
C = dict(
    azul_osc="#1F3864", azul_med="#2E75B6", azul_clr="#BDD7EE",
    verde_osc="#375623", verde_med="#70AD47",
    amarillo="#FFC000",  rojo="#C00000",
    naranja="#ED7D31",   morado="#7030A0",
    gris_bg="#F0F4FA",   texto_sub="#6B7280",
)
SEQ = [C["azul_med"], C["verde_med"], C["amarillo"], C["rojo"],
       C["naranja"], C["morado"], "#06B6D4", "#10B981", "#3B82F6", "#F59E0B",
       "#84CC16", "#EC4899", "#14B8A6", "#8B5CF6", "#F97316"]

# ══════════════════════════════════════════════════════════════════════════════
# 1. CARGA
# ══════════════════════════════════════════════════════════════════════════════
print("Cargando datos...")

raw    = pd.read_excel(F_CAC, sheet_name=None, header=0)
er_h   = raw["Estados de Resultados"].copy()
bg_h   = raw["Balances Generales"].copy()
cmf_er = raw["CMF ER - Feb 2026"].copy()
cmf_bg = raw["CMF BG - Feb 2026"].copy()
panel  = pd.read_excel(F_PANEL, sheet_name="Panel Cooperativas", header=0)

# ── nombre corto ─────────────────────────────────────────────────────────────
def shorten(s):
    if pd.isna(s): return "Sin nombre"
    for r in ["Cooperativa de Ahorro y Crédito ", "Cooperativa de Ahorro, Crédito y Servicios Financieros ",
              "Coop. de Ahorro y Crédito ", "Coop. de Ahorro, Crédito y Servicios Financieros ",
              "COOPERATIVA DE AHORRO Y CREDITO ", "COOPERATIVA DE AHORRO, CREDITO Y SERVICIOS FINANCIEROS "]:
        s = s.replace(r, "")
    return s.replace(" Ltda.","").replace(" Limitada","").replace(" LTDA.","").replace(" LIMITADA","")[:34]

for df in [er_h, bg_h, cmf_er, cmf_bg]:
    df["nombre_corto"] = df["Nombre de la entidad"].apply(shorten)

# ══════════════════════════════════════════════════════════════════════════════
# 2. UNIFICAR CMF (MM$→CLP) CON HISTÓRICO
# ══════════════════════════════════════════════════════════════════════════════
NUM_ER = ["Ingresos por Intereses y Reajustes","Ingresos por Inversiones",
          "Otros Ingresos de Operación","Total Ingresos de Operación",
          "Gastos por Interés, Reajustes y Comisión","Margen Bruto",
          "Remuneraciones y Gastos del Personal","Gastos de Administración y Otros",
          "Depreciaciones y Amortizaciones","Margen Neto",
          "Provisiones sobre Activos Riesgosos","Recuperación de Colocaciones Castigadas",
          "Resultado Operacional","Ingresos no Operacionales","Gastos no Operacionales",
          "Corrección Monetaria","Resultado Antes de Impuesto","Remanente del Período"]

NUM_BG = ["Disponible","Préstamos Comerciales","Préstamos de Consumo",
          "Otras Colocaciones Vigentes","Cartera Vencida","Total Colocaciones",
          "Provisión sobre Colocaciones","Total Colocaciones Netas",
          "Inversiones Financieras","Otros Activos","Total Activo Fijo Neto",
          "Inversiones en Sociedades","Total Activos","Depósitos a la Vista",
          "Depósitos a Plazo","Total Depósitos",
          "Obligaciones con Bancos e Inst. Finan.","Otros Pasivos","Total Pasivos",
          "Capital Pagado","Reservas","Remanente Acumulado","Total Patrimonio"]

# Convertir CMF MM$ → CLP (*1e6) y renombrar columna año
cmf_er_clp = cmf_er.copy()
cmf_er_clp["Año"] = 2026
cmf_er_clp["fuente_cmf"] = True
for c in NUM_ER:
    if c in cmf_er_clp.columns:
        cmf_er_clp[c] = pd.to_numeric(cmf_er_clp[c], errors="coerce") * 1e6

cmf_bg_clp = cmf_bg.copy()
cmf_bg_clp["Año"] = 2026
cmf_bg_clp["fuente_cmf"] = True
for c in NUM_BG:
    if c in cmf_bg_clp.columns:
        cmf_bg_clp[c] = pd.to_numeric(cmf_bg_clp[c], errors="coerce") * 1e6

er_h["fuente_cmf"] = False
bg_h["fuente_cmf"] = False

# Concat con columnas en común + Año
cols_er_common = ["nombre_corto","RUT","Año","fuente_cmf"] + NUM_ER
cols_bg_common = ["nombre_corto","RUT","Año","fuente_cmf"] + NUM_BG

er = pd.concat([
    er_h[cols_er_common],
    cmf_er_clp[[c for c in cols_er_common if c in cmf_er_clp.columns]],
], ignore_index=True)

bg = pd.concat([
    bg_h[cols_bg_common],
    cmf_bg_clp[[c for c in cols_bg_common if c in cmf_bg_clp.columns]],
], ignore_index=True)

print(f"  ER unificado: {len(er)} registros · {er['nombre_corto'].nunique()} entidades · años {sorted(er['Año'].unique())}")
print(f"  BG unificado: {len(bg)} registros · {bg['nombre_corto'].nunique()} entidades")

# ══════════════════════════════════════════════════════════════════════════════
# 3. VARIABLES SCN
# ══════════════════════════════════════════════════════════════════════════════
er["P1"]   = er["Total Ingresos de Operación"]
er["P2"]   = er["P1"] * ALPHA_CI
er["B1g"]  = er["P1"] * (1 - ALPHA_CI)
er["D1"]   = er["Remuneraciones y Gastos del Personal"].abs()
er["P51d"] = er["Depreciaciones y Amortizaciones"].abs()
er["GAdm"] = er["Gastos de Administración y Otros"].abs()
er["B2g"]  = er["B1g"] - er["D1"].fillna(0)
er["B1n"]  = er["B1g"] - er["P51d"].fillna(0)
er["Rem"]  = er["Remanente del Período"]
er["Prov"] = er["Provisiones sobre Activos Riesgosos"]

# Balance
bg["F2"]   = bg["Total Depósitos"]
bg["F4"]   = bg["Total Colocaciones Netas"]
bg["mora"] = (bg["Cartera Vencida"] / bg["Total Colocaciones"] * 100).where(
    bg["Cartera Vencida"].notna() & bg["Total Colocaciones"].notna())
bg["solv"] = (bg["Total Patrimonio"] / bg["Total Activos"] * 100).where(
    bg["Total Patrimonio"].notna() & bg["Total Activos"].notna())

# MM$ helpers
for c in ["P1","P2","B1g","D1","P51d","GAdm","B2g","B1n","Rem"]:
    er[c+"_MM"] = er[c] / 1e9
for c in ["F2","F4","Total Activos","Total Patrimonio","Total Pasivos"]:
    bg[c+"_MM"] = bg[c] / 1e9

# ══════════════════════════════════════════════════════════════════════════════
# 4. AGREGADOS SECTORIALES POR AÑO (SECTOR COMPLETO)
# ══════════════════════════════════════════════════════════════════════════════
def agg_yr(df, cols_sum, cols_mean=None):
    ops = {c: (c, "sum") for c in cols_sum}
    if cols_mean:
        ops.update({c+"_avg": (c, lambda x: x.dropna().mean()) for c in cols_mean})
    ops["N"] = ("RUT", "count")
    ops["N_cmf"] = ("fuente_cmf", "sum")
    ops["N_def"] = ("Rem", lambda x: (x < 0).sum())
    return df.groupby("Año").agg(**ops).reset_index()

sector = er.groupby("Año").agg(
    P1_MM    =("P1_MM",   "sum"),
    P2_MM    =("P2_MM",   "sum"),
    B1g_MM   =("B1g_MM",  "sum"),
    D1_MM    =("D1_MM",   "sum"),
    P51d_MM  =("P51d_MM", "sum"),
    B2g_MM   =("B2g_MM",  "sum"),
    B1n_MM   =("B1n_MM",  "sum"),
    Rem_MM   =("Rem_MM",  "sum"),
    N        =("RUT",     "count"),
    N_cmf    =("fuente_cmf", "sum"),
    N_def    =("Rem",     lambda x: (x.fillna(0) < 0).sum()),
).reset_index()
sector["pct_def"]  = (sector["N_def"] / sector["N"] * 100).round(1)
sector["D1_pct"]   = (sector["D1_MM"] / sector["B1g_MM"] * 100).round(1)
sector["B2g_pct"]  = (sector["B2g_MM"] / sector["B1g_MM"] * 100).round(1)

mora_yr = bg[bg["mora"].notna()].groupby("Año")["mora"].mean().round(2).reset_index()
solv_yr = bg[bg["solv"].notna()].groupby("Año")["solv"].mean().round(1).reset_index()
f2_yr   = bg.groupby("Año")["F2_MM"].sum().round(1).reset_index()
f4_yr   = bg.groupby("Año")["F4_MM"].sum().round(1).reset_index()
act_yr  = bg.groupby("Año")["Total Activos_MM"].sum().round(1).reset_index()

sector = sector.merge(mora_yr, on="Año", how="left") \
               .merge(solv_yr, on="Año", how="left") \
               .merge(f2_yr,   on="Año", how="left") \
               .merge(f4_yr,   on="Año", how="left") \
               .merge(act_yr,  on="Año", how="left")

# ══════════════════════════════════════════════════════════════════════════════
# 5. SERIES POR ENTIDAD (para gráfico comparativo)
# ══════════════════════════════════════════════════════════════════════════════
entidades_list = sorted(er["nombre_corto"].unique())
cmf_names = set(cmf_er["nombre_corto"].tolist())

er_js = {}
for ent in entidades_list:
    d = er[er["nombre_corto"] == ent].sort_values("Año")
    er_js[ent] = {
        "años":  d["Año"].astype(str).tolist(),
        "B1g":   d["B1g_MM"].round(2).tolist(),
        "P1":    d["P1_MM"].round(2).tolist(),
        "P2":    d["P2_MM"].round(2).tolist(),
        "D1":    d["D1_MM"].round(2).tolist(),
        "B2g":   d["B2g_MM"].round(2).tolist(),
        "P51d":  d["P51d_MM"].round(2).tolist(),
        "B1n":   d["B1n_MM"].round(2).tolist(),
        "Rem":   d["Rem_MM"].round(2).tolist(),
        "cmf":   d["fuente_cmf"].tolist(),
    }

# scatter mora/solv por año
scatter_js = {}
for año in sorted(bg["Año"].unique()):
    d = bg[(bg["Año"] == año) & bg["mora"].notna() & bg["solv"].notna()].copy()
    scatter_js[int(año)] = {
        "ents":  d["nombre_corto"].tolist(),
        "mora":  d["mora"].round(2).tolist(),
        "solv":  d["solv"].round(1).tolist(),
        "act":   d["Total Activos_MM"].fillna(1).round(1).tolist(),
        "cmf":   d["fuente_cmf"].tolist(),
    }

# ══════════════════════════════════════════════════════════════════════════════
# 6. GEOGRAFÍA — Panel DAES (40 CAC activas) + GeoJSON coroplético
# ══════════════════════════════════════════════════════════════════════════════
cac = panel[panel["Subrubro"] == "Ahorro y Crédito"].copy()
cac["nombre_corto"] = cac["Razón Social"].apply(shorten)

# Tramo ventas → producción representativa (punto medio SII)
TRAMOS = {1:0, 2:800e3, 3:2.4e6, 4:7.4e6, 5:17.4e6, 6:47.4e6,
          7:112.4e6, 8:312.4e6, 9:812.4e6, 10:1812.4e6,
          11:3812.4e6, 12:8812.4e6, 13:25e9}
cac["prod_sii"] = cac["sii_tramo_ventas"].map(TRAMOS)
cac["vab_sii"]  = cac["prod_sii"] * (1 - ALPHA_CI)
cac["prod_MM"]  = cac["prod_sii"] / 1e9
cac["vab_MM"]   = cac["vab_sii"]  / 1e9

# Convertir socios a numérico
for c in ["Total Socios","Socios Hombres","Socios Mujeres","sii_trabajadores"]:
    cac[c] = pd.to_numeric(cac[c], errors="coerce")

# ── Tabla de match: nombre DAES (MAYÚSCULAS) → nombre GeoJSON (BCN) ──────────
# GeoJSON usa: "Región de Antofagasta", "Región Metropolitana de Santiago", etc.
REGION_MAP = {
    "REGIÓN DE ANTOFAGASTA":          "Región de Antofagasta",
    "REGIÓN DE ARICA Y PARINACOTA":   "Región de Arica y Parinacota",
    "REGIÓN DE COQUIMBO":             "Región de Coquimbo",
    "REGIÓN DE LA ARAUCANÍA":         "Región de La Araucanía",
    "REGIÓN DE LOS LAGOS":            "Región de Los Lagos",
    "REGIÓN DE VALPARAISO":           "Región de Valparaíso",
    "REGIÓN DEL BÍO-BÍO":             "Región del Bío-Bío",
    "REGIÓN DEL MAULE":               "Región del Maule",
    "REGIÓN METROPOLITANA":           "Región Metropolitana de Santiago",
}
cac["region_geojson"] = cac["Región"].map(REGION_MAP).fillna("Sin datos")

# Nombre corto para labels en barras
def short_reg(s):
    if pd.isna(s): return "Sin datos"
    return (str(s)
        .replace("REGIÓN DE ","").replace("REGIÓN DEL ","")
        .replace("REGIÓN METROPOLITANA","R. Metropolitana")
        .title())
cac["region_short"] = cac["Región"].apply(short_reg)

# Agregado por región (con nombre GeoJSON para join)
geo = cac.groupby(["region_geojson","region_short"]).agg(
    N        =("Razón Social",    "count"),
    Socios   =("Total Socios",    "sum"),
    SocH     =("Socios Hombres",  "sum"),
    SocM     =("Socios Mujeres",  "sum"),
    Trab_SII =("sii_trabajadores","sum"),
    Prod_MM  =("prod_MM",         "sum"),
    Vab_MM   =("vab_MM",          "sum"),
).reset_index().sort_values("Socios", ascending=False).reset_index(drop=True)

# Socios hombre/mujer total
soc_h = float(cac["Socios Hombres"].sum())
soc_m = float(cac["Socios Mujeres"].sum())

# ── Cargar GeoJSON (o convertir desde shapefile) y embeber en HTML ───────────
import json as _json
import json
geojson_str = "{}"
geojson_ok  = False
_gj         = None

if os.path.exists(F_GEOJSON):
    with open(F_GEOJSON, encoding="utf-8") as _f:
        _gj = _json.load(_f)
    print(f"  GeoJSON cargado: {len(_gj['features'])} regiones")

elif os.path.exists(F_SHP):
    print("  Shapefile encontrado — convirtiendo a GeoJSON...")

    try:
        import geopandas as gpd

        gdf = gpd.read_file(F_SHP)

        # convertir a coordenadas geográficas reales
        gdf = gdf.to_crs(epsg=4326)

        # detectar nombre región
        CANDS = ["NOM_REG", "Region", "REGION", "NOMBRE", "NOM_REGION"]
        name_f = next((c for c in CANDS if c in gdf.columns), gdf.columns[0])

        gdf["Region"] = gdf[name_f].astype(str).str.strip()

        _gj = json.loads(gdf.to_json())

        with open(F_GEOJSON, "w", encoding="utf-8") as f:
            json.dump(_gj, f, ensure_ascii=False)

        print(f"  ✅ Regional.geojson generado ({len(_gj['features'])} regiones)")

    except Exception as e:
        print(f"  ⚠ Error convirtiendo shapefile: {e}")
    except ImportError:
        print("  ⚠ Instala pyshp:  pip install pyshp")
    except Exception as _e:
        print(f"  ⚠ Error convirtiendo shapefile: {_e}")

else:
    print("  ⚠ Regional.geojson / Regional.shp no encontrados — mapa no disponible")

if _gj is not None:
    # Diagnóstico: mostrar nombres reales para verificar coincidencias
    _actual = sorted(set(f["properties"].get("Region","") for f in _gj["features"] if f["properties"].get("Region","")))
    _expected = sorted(set(REGION_MAP.values()))
    _matches = sorted(set(_actual) & set(_expected))
    print(f"  Nombres GeoJSON:   {_actual}")
    print(f"  Nombres esperados: {_expected}")
    print(f"  Coincidencias ({len(_matches)}): {_matches}")

    # Simplificar coordenadas para reducir tamaño del HTML embebido
    def _simp_ring(ring, step=5, prec=4):
        if len(ring) <= 4:
            return [[round(c, prec) for c in p[:2]] for p in ring]
        return [[round(c, prec) for c in p[:2]] for p in
                [ring[0]] + [ring[i] for i in range(1, len(ring)-1, step)] + [ring[-1]]]
    def _simp_geom(g):
        if g["type"] == "Polygon":
            return {"type":"Polygon","coordinates":[_simp_ring(r) for r in g["coordinates"]]}
        if g["type"] == "MultiPolygon":
            return {"type":"MultiPolygon","coordinates":[[_simp_ring(r) for r in p] for p in g["coordinates"]]}
        return g

    for feat in _gj["features"]:
        feat["id"] = feat["properties"].get("Region", "")
        feat["geometry"] = _simp_geom(feat["geometry"])
    geojson_str = _json.dumps(_gj, ensure_ascii=False)
    geojson_ok  = True
    print(f"  GeoJSON embebido: {len(geojson_str)//1024} KB")

# ── Generar paths SVG simplificados para mapa coroplético offline ─────────────
SVG_W, SVG_H = 380, 860
PAD = 8
LON_MIN_SVG, LON_MAX_SVG = -76.0, -64.5
LAT_MIN_SVG, LAT_MAX_SVG = -56.5, -17.0

def proj_svg(lon, lat):
    x = (lon - LON_MIN_SVG) / (LON_MAX_SVG - LON_MIN_SVG) * (SVG_W - 2*PAD) + PAD
    y = (LAT_MAX_SVG - lat) / (LAT_MAX_SVG - LAT_MIN_SVG) * (SVG_H - 2*PAD) + PAD
    return round(x, 1), round(y, 1)

def simplify_ring_svg(ring, step):
    pts = []
    for i, c in enumerate(ring):
        if i % step == 0:
            lon, lat = c[0], c[1]
            if LON_MIN_SVG-1 <= lon <= LON_MAX_SVG+1 and LAT_MIN_SVG-1 <= lat <= LAT_MAX_SVG+1:
                x, y = proj_svg(lon, lat)
                pts.append(f"{x},{y}")
    return ("M " + " L ".join(pts) + " Z") if len(pts) > 2 else ""

STEP_MAP = {
    "Región de Los Lagos": 6,
    "Región de Aysén del Gral.Ibañez del Campo": 8,
    "Región de Magallanes y Antártica Chilena": 10,
}

regions_svg, centroids_svg = {}, {}
if geojson_ok:
    for feat in _gj["features"]:
        name = feat["properties"]["Region"]
        if "sin demarcar" in name: continue
        geom = feat["geometry"]
        step = STEP_MAP.get(name, 1)
        if geom["type"] == "Polygon":
            parts = [simplify_ring_svg(r, step) for r in geom["coordinates"]]
        else:
            parts = [simplify_ring_svg(r, step) for poly in geom["coordinates"] for r in poly]
        path = " ".join(p for p in parts if p)
        if path:
            regions_svg[name] = path
            # centroide: polígono más grande
            if geom["type"] == "Polygon":
                main_ring = geom["coordinates"][0]
            else:
                main_ring = max(geom["coordinates"], key=lambda p: len(p[0]))[0]
            lons = [c[0] for c in main_ring[::3] if LON_MIN_SVG <= c[0] <= LON_MAX_SVG]
            lats = [c[1] for c in main_ring[::3] if LAT_MIN_SVG <= c[1] <= LAT_MAX_SVG]
            if lons:
                cx, cy = proj_svg(sum(lons)/len(lons), sum(lats)/len(lats))
                centroids_svg[name] = [cx, cy]

# Panel tabla
panel_cols = ["nombre_corto","region_short","Total Socios","Socios Hombres","Socios Mujeres",
              "sii_trabajadores","sii_tramo_ventas","prod_MM","vab_MM",
              "cmf_empleados","cmf_oficinas","en_cmf","reg19862_monto_total"]
pt = cac[panel_cols].copy()
pt["en_cmf"] = pt["en_cmf"].map({True:"✔ CMF", False:"—"})

def fmt(v, dec=1):
    if pd.isna(v): return "—"
    try:
        f = float(v)
        if dec == 0: return f"{int(f):,}"
        return f"{f:,.{dec}f}"
    except: return str(v)

panel_rows = []
for _, r in pt.iterrows():
    panel_rows.append({
        "Entidad":    r["nombre_corto"],
        "Región":     r["region_short"],
        "Socios":     fmt(r["Total Socios"], 0),
        "Soc.H":      fmt(r["Socios Hombres"], 0),
        "Soc.M":      fmt(r["Socios Mujeres"], 0),
        "Trab.SII":   fmt(r["sii_trabajadores"], 0),
        "Prod.MM$":   fmt(r["prod_MM"], 2),
        "VAB.MM$":    fmt(r["vab_MM"], 2),
        "Empl.CMF":   fmt(r["cmf_empleados"], 0),
        "Of.CMF":     fmt(r["cmf_oficinas"], 0),
        "CMF":        r["en_cmf"],
        "Transf.19862": fmt(r["reg19862_monto_total"], 0),
    })

# ── KPIs ─────────────────────────────────────────────────────────────────────
yr_ref  = int(sector.loc[sector["N"].idxmax(), "Año"])
row_r   = sector[sector["Año"] == yr_ref].iloc[0]
total_socios  = int(cac["Total Socios"].sum())
total_trab_cmf = int(cac["cmf_empleados"].dropna().sum())

print(f"  Sector unificado: {er['nombre_corto'].nunique()} entidades · {len(er)} registros")
print(f"  CMF en comparativos: {len(cmf_names)} entidades (punto 2026)")
print(f"  Geografía: {len(geo)} regiones · {len(cac)} CAC DAES")
print("Construyendo HTML...")

# ══════════════════════════════════════════════════════════════════════════════
# 7. SERIALIZACIÓN SECTOR PARA JS
# ══════════════════════════════════════════════════════════════════════════════
def jl(col, d=2):
    return json.dumps([round(float(v), d) if pd.notna(v) else None for v in sector[col]])

def ji(col):
    return json.dumps([int(v) for v in sector[col]])

HTML = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Huella Social · Cuenta Satélite CAC</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:{C['gris_bg']};font-family:'Segoe UI',Arial,sans-serif;color:#1A1A2E;font-size:13px}}

.hdr{{background:linear-gradient(135deg,{C['azul_osc']} 0%,{C['azul_med']} 100%);
  padding:18px 34px;display:flex;align-items:center;gap:14px;
  box-shadow:0 4px 20px rgba(31,56,100,.3)}}
.hdr h1{{color:white;font-size:21px;font-weight:800;letter-spacing:.07em}}
.hdr p{{color:rgba(255,255,255,.72);font-size:11.5px;margin-top:2px}}
.badge{{background:rgba(255,255,255,.18);color:white;font-size:9.5px;font-weight:700;
  padding:2px 8px;border-radius:20px;letter-spacing:.05em;display:inline-block;margin-top:4px}}

.wrap{{padding:20px 30px 48px}}

.kpi-row{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}}
.kpi{{background:white;border-radius:10px;padding:13px 16px;flex:1;min-width:140px;
  border-top:4px solid {C['azul_med']};box-shadow:0 2px 10px rgba(0,0,0,.07)}}
.kpi-lbl{{font-size:9px;text-transform:uppercase;letter-spacing:.07em;
  color:{C['texto_sub']};font-weight:700;margin-bottom:4px}}
.kpi-val{{font-size:22px;font-weight:800;line-height:1}}
.kpi-sub{{font-size:9px;color:{C['texto_sub']};margin-top:3px}}

.card{{background:white;border-radius:12px;padding:18px 20px;margin-bottom:16px;
  box-shadow:0 2px 12px rgba(0,0,0,.07)}}

.g2{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}}
@media(max-width:900px){{.g2,.g3{{grid-template-columns:1fr}}}}

.tabs{{display:flex;gap:3px;border-bottom:2px solid {C['azul_clr']};flex-wrap:wrap}}
.tab{{padding:7px 14px;font-size:11.5px;font-weight:600;border:none;background:transparent;
  cursor:pointer;border-radius:7px 7px 0 0;color:{C['texto_sub']};transition:all .15s}}
.tab.on{{background:{C['azul_med']};color:white}}
.tab:hover:not(.on){{background:{C['azul_clr']}}}
.tab-pane{{display:none;padding-top:14px}}.tab-pane.on{{display:block}}

.ctrl-row{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:11px;align-items:flex-end}}
.ctrl{{display:flex;flex-direction:column;gap:3px}}
.ctrl label{{font-size:9px;font-weight:700;text-transform:uppercase;
  letter-spacing:.05em;color:{C['texto_sub']}}}
select{{padding:6px 9px;border:1.5px solid #D1D5DB;border-radius:7px;font-size:12px;
  background:white;cursor:pointer;font-family:inherit}}
select:focus{{outline:none;border-color:{C['azul_med']}}}
select[multiple]{{height:105px}}

.info-box{{border-left:4px solid {C['azul_med']};background:{C['azul_clr']};
  padding:9px 13px;border-radius:0 7px 7px 0;font-size:11px;color:{C['azul_osc']};
  margin-bottom:12px;line-height:1.55}}
.warn-box{{border-left:4px solid {C['amarillo']};background:#FEF9E7;
  padding:9px 13px;border-radius:0 7px 7px 0;font-size:11px;color:#78350F;
  margin-bottom:12px;line-height:1.55}}
.cmf-dot{{display:inline-block;width:8px;height:8px;border-radius:50%;
  background:{C['naranja']};margin-right:4px;vertical-align:middle}}

.tbl-wrap{{overflow-x:auto;max-height:420px;overflow-y:auto}}
table{{width:100%;border-collapse:collapse;font-size:11px}}
th{{background:{C['azul_osc']};color:white;padding:7px 9px;font-weight:700;
  position:sticky;top:0;text-align:left;white-space:nowrap;font-size:10px}}
td{{padding:6px 9px;border-bottom:1px solid #E5E7EB;white-space:nowrap}}
tr:nth-child(even) td{{background:{C['gris_bg']}}}
.cmf-badge{{background:{C['naranja']};color:white;font-size:8.5px;
  padding:1px 5px;border-radius:8px;font-weight:700}}

footer{{text-align:center;font-size:10px;color:{C['texto_sub']};
  padding:16px;border-top:1px solid {C['azul_clr']}}}
</style>
</head>
<body>

<div class="hdr">
  <span style="font-size:30px">🌱</span>
  <div>
    <h1>HUELLA SOCIAL</h1>
    <p>Cuenta Satélite · Sector CAC Completo (histórico 2014–2024 + CMF Feb-2026) · Chile</p>
    <span class="badge">SCN 2025 · ONU TSE · CIRIEC · α={ALPHA_CI} MIP Chile 2018</span>
  </div>
</div>

<div class="wrap">

<!-- KPIs -->
<div class="kpi-row">
  <div class="kpi" style="border-color:{C['azul_med']}">
    <div class="kpi-lbl">Entidades en consolidado</div>
    <div class="kpi-val" style="color:{C['azul_med']}">{er['nombre_corto'].nunique()}</div>
    <div class="kpi-sub">{int(row_r['N_cmf'])} supervisadas CMF · {int(row_r['N']-row_r['N_cmf'])} históricas · Año {yr_ref}</div>
  </div>
  <div class="kpi" style="border-color:{C['verde_osc']}">
    <div class="kpi-lbl">VAB Bruto [B1g]</div>
    <div class="kpi-val" style="color:{C['verde_osc']}">MM${row_r['B1g_MM']:,.0f}</div>
    <div class="kpi-sub">Sector completo · Año {yr_ref} · P1×{1-ALPHA_CI:.4f}</div>
  </div>
  <div class="kpi" style="border-color:{C['azul_med']}">
    <div class="kpi-lbl">Remuneraciones [D1]</div>
    <div class="kpi-val" style="color:{C['azul_med']}">MM${row_r['D1_MM']:,.0f}</div>
    <div class="kpi-sub">{row_r['D1_pct']:.0f}% del VAB · Año {yr_ref}</div>
  </div>
  <div class="kpi" style="border-color:{C['naranja']}">
    <div class="kpi-lbl">Supervisadas CMF</div>
    <div class="kpi-val" style="color:{C['naranja']}">{len(cmf_names)}</div>
    <div class="kpi-sub">Incluidas en todos los gráficos · Feb-2026</div>
  </div>
  <div class="kpi" style="border-color:{C['verde_med']}">
    <div class="kpi-lbl">Total Socios DAES</div>
    <div class="kpi-val" style="color:{C['verde_med']}">{total_socios:,}</div>
    <div class="kpi-sub">Universo 40 CAC activas [CIRIEC]</div>
  </div>
  <div class="kpi" style="border-color:{C['morado']}">
    <div class="kpi-lbl">Empleados CMF [PEP]</div>
    <div class="kpi-val" style="color:{C['morado']}">{total_trab_cmf:,}</div>
    <div class="kpi-sub">Variable PEP ONU TSE · supervisadas</div>
  </div>
</div>

<!-- TABS -->
<div class="card" style="padding-bottom:22px">
<div class="tabs">
  <button class="tab on"  onclick="tab('produccion','t1')" id="t1">📊 P1→B1g Producción y VAB</button>
  <button class="tab"     onclick="tab('ingreso','t2')"    id="t2">💼 D1·B2g Generación Ingreso</button>
  <button class="tab"     onclick="tab('financiera','t3')" id="t3">🏦 F2·F4 Cuenta Financiera</button>
  <button class="tab"     onclick="tab('entidad','t4')"    id="t4">🔍 Por Entidad</button>
  <button class="tab"     onclick="tab('mapa','t5')"       id="t5">🗺️ Mapa Chile</button>
  <button class="tab"     onclick="tab('geografia','t6')"  id="t6">📍 Geografía</button>
  <button class="tab"     onclick="tab('panel','t7')"      id="t7">📋 Panel DAES</button>
</div>

<!-- ══ TAB: PRODUCCIÓN Y VAB ══════════════════════════════════════════════ -->
<div class="tab-pane on" id="tab-produccion">
  <div class="info-box">
    <b>SCN 2025 §7.82 · §4.3.3 memoria:</b>
    P1 (Producción) = Total Ingresos de Operación.
    P2 (Consumo Intermedio) = α×P1, α={ALPHA_CI} (MIP Chile 2018, sector 94).
    B1g (VAB Bruto) = P1−P2 = {1-ALPHA_CI:.4f}×P1.
    B1n (VAB Neto) = B1g−P51d. &nbsp;·&nbsp;
    <span class="cmf-dot"></span><b>Puntos naranjos</b> = datos CMF Feb-2026 (método directo).
  </div>
  <div class="g2">
    <div id="g-b1g-año"  style="min-height:310px"></div>
    <div id="g-vab-comp" style="min-height:310px"></div>
  </div>
  <div class="g2">
    <div id="g-p1-p2"    style="min-height:290px"></div>
    <div id="g-cobertura" style="min-height:290px"></div>
  </div>
</div>

<!-- ══ TAB: GENERACIÓN DEL INGRESO ════════════════════════════════════════ -->
<div class="tab-pane" id="tab-ingreso">
  <div class="info-box">
    <b>SCN 2025 §8.5 · Cuenta de Generación del Ingreso:</b>
    D1 = Remuneraciones (proxy: Remuneraciones y Gastos del Personal).
    B2g = Excedente Bruto de Explotación = B1g − D1.
  </div>
  <div class="warn-box">
    <b>⚠ Brechas documentadas (§4.2.3 memoria):</b>
    D11/D12 (sueldos/cotizaciones) sin desagregación en fuentes admin.
    PEP (empleo equiv. tiempo completo, ONU TSE) disponible solo en CMF.
    VOV (trabajo voluntario) sin fuente en Chile.
  </div>
  <div class="g2">
    <div id="g-d1-b2g"   style="min-height:310px"></div>
    <div id="g-d1-pct"   style="min-height:310px"></div>
  </div>
  <div class="g2">
    <div id="g-remanente" style="min-height:290px"></div>
    <div id="g-deficit"   style="min-height:290px"></div>
  </div>
</div>

<!-- ══ TAB: CUENTA FINANCIERA ═════════════════════════════════════════════ -->
<div class="tab-pane" id="tab-financiera">
  <div class="info-box">
    <b>SCN 2025 Cap.12 · CIRIEC §4:</b>
    F2 = Depósitos captados · F4 = Colocaciones Netas.
    Mora = Cartera Vencida / Total Colocaciones · Solvencia = Patrimonio / Activos.
  </div>
  <div class="g2">
    <div id="g-f2f4"     style="min-height:310px"></div>
    <div id="g-mora-solv" style="min-height:310px"></div>
  </div>
  <div class="ctrl-row">
    <div class="ctrl">
      <label>Año (scatter mora vs solvencia)</label>
      <select id="sel-scat-año" onchange="updateScatter()">
        {''.join(f'<option value="{int(a)}">{int(a)}</option>' for a in sorted(bg["Año"].unique(), reverse=True))}
      </select>
    </div>
  </div>
  <div id="g-scatter" style="min-height:370px"></div>
</div>

<!-- ══ TAB: POR ENTIDAD ═══════════════════════════════════════════════════ -->
<div class="tab-pane" id="tab-entidad">
  <div class="info-box">
    Todas las entidades del consolidado incluyendo CMF supervisadas (2026).
    <span class="cmf-dot"></span> Marcadores naranjos = punto CMF Feb-2026.
    Usar Ctrl+clic para seleccionar varias entidades.
  </div>
  <div class="ctrl-row">
    <div class="ctrl">
      <label>Entidad/es (Ctrl+clic para comparar)</label>
      <select id="sel-ent" multiple onchange="updateEntidad()" style="height:110px;min-width:270px">
        {''.join(f'<option value="{e}" {"selected" if i<4 else ""}>{("★ " if e in cmf_names else "") + e}</option>' for i,e in enumerate(entidades_list))}
      </select>
    </div>
    <div class="ctrl">
      <label>Variable SCN</label>
      <select id="sel-ent-var" onchange="updateEntidad()">
        <option value="B1g">VAB Bruto [B1g] (MM$)</option>
        <option value="P1">Producción [P1] (MM$)</option>
        <option value="P2">Consumo Intermedio [P2] (MM$)</option>
        <option value="D1">Remuneraciones [D1] (MM$)</option>
        <option value="B2g">Excedente Bruto [B2g] (MM$)</option>
        <option value="P51d">Depreciación [P51d] (MM$)</option>
        <option value="B1n">VAB Neto [B1n] (MM$)</option>
        <option value="Rem">Remanente del Período (MM$)</option>
      </select>
    </div>
  </div>
  <div id="g-entidad" style="min-height:390px"></div>
</div>

<!-- ══ TAB: MAPA CHILE ════════════════════════════════════════════════════ -->
<div class="tab-pane" id="tab-mapa">
  <div class="info-box">
    <b>CIRIEC §4 · Distribución territorial del sector CAC.</b>
    Polígonos oficiales BCN (Regional.geojson). Haz clic en una región para ver el detalle.
    Hover sobre la región para ver todas las variables.
  </div>
  <div style="display:flex;gap:16px;align-items:flex-start;flex-wrap:wrap">
    <!-- Controles -->
    <div style="display:flex;flex-direction:column;gap:10px;min-width:160px">
      <div class="ctrl">
        <label>Variable</label>
        <select id="map-var" onchange="renderMapa()" style="min-width:155px">
          <option value="N">N° Cooperativas</option>
          <option value="socios">Total Socios</option>
          <option value="vab_mm">VAB estimado (MM$)</option>
          <option value="trab">Trabajadores SII</option>
        </select>
      </div>
      <div class="ctrl">
        <label>Paleta</label>
        <select id="map-pal" onchange="renderMapa()" style="min-width:155px">
          <option value="blues">Azules</option>
          <option value="greens">Verdes</option>
          <option value="oranges">Naranjo</option>
          <option value="reds">Rojos</option>
          <option value="purples">Morados</option>
        </select>
      </div>
      <!-- Leyenda dinámica -->
      <div id="map-legend" style="margin-top:8px"></div>
      <!-- Panel detalle región clickeada -->
      <div id="map-detail" style="margin-top:12px;display:none;
        background:#F0F4FA;border-radius:8px;padding:12px;font-size:11.5px;line-height:1.8">
      </div>
    </div>
    <!-- Mapa Plotly -->
    <div style="flex:1;min-width:300px">
      <div id="g-mapa" style="min-height:680px"></div>
    </div>
  </div>
</div>

<!-- ══ TAB: GEOGRAFÍA ════════════════════════════════════════════════════ -->
<div class="tab-pane" id="tab-geografia">
  <div class="info-box">
    <b>CIRIEC §4 / Manual ONU TSE (Tabla 4.5):</b>
    Distribución territorial del sector · Fuente: Panel DAES · {len(cac)} CAC activas.
    VAB estimado por método indirecto (tramos SII → punto medio → α MIP 2018).
  </div>
  <div class="g2">
    <div id="g-geo-n"      style="min-height:360px"></div>
    <div id="g-geo-socios" style="min-height:360px"></div>
  </div>
  <div class="g2">
    <div id="g-geo-vab"    style="min-height:340px"></div>
    <div id="g-geo-genero" style="min-height:340px"></div>
  </div>
</div>

<!-- ══ TAB: PANEL DAES ════════════════════════════════════════════════════ -->
<div class="tab-pane" id="tab-panel">
  <div class="info-box">
    Panel integrado DAES + SII + CMF + Ley 19.862 · {len(cac)} CAC activas.
    <span class="cmf-badge">★ CMF</span> = supervisada con datos auditados.
    Producción y VAB estimados por método indirecto SII (tramos).
  </div>
  <div class="tbl-wrap" id="panel-tabla"></div>
</div>

</div><!-- card -->
</div><!-- wrap -->

<footer>
  Huella Social · Universidad de los Andes / INAC · Prof. guía: Sebastián Cea ·
  Fuentes: DAES · SII · CMF BEST · Ley 19.862 · MIP Chile 2018 · SCN 2025 · Manual ONU TSE
</footer>

<script>
// ── Datos pre-calculados ──────────────────────────────────────────────────────
const ER   = {json.dumps(er_js, ensure_ascii=False)};
const SCAT = {json.dumps(scatter_js, ensure_ascii=False)};
const CMF_NAMES = new Set({json.dumps(list(cmf_names), ensure_ascii=False)});
const PANEL_ROWS = {json.dumps(panel_rows, ensure_ascii=False)};
const SEQ   = {json.dumps(SEQ)};

const SECTOR = {{
  años:    {json.dumps(sector["Año"].astype(str).tolist())},
  P1:      {jl("P1_MM")},
  P2:      {jl("P2_MM")},
  B1g:     {jl("B1g_MM")},
  D1:      {jl("D1_MM")},
  P51d:    {jl("P51d_MM")},
  B2g:     {jl("B2g_MM")},
  B1n:     {jl("B1n_MM")},
  Rem:     {jl("Rem_MM")},
  N:       {ji("N")},
  N_cmf:   {ji("N_cmf")},
  N_def:   {ji("N_def")},
  pct_def: {jl("pct_def",1)},
  D1_pct:  {jl("D1_pct",1)},
  B2g_pct: {jl("B2g_pct",1)},
  mora:    {jl("mora",2)},
  solv:    {jl("solv",1)},
  F2:      {jl("F2_MM")},
  F4:      {jl("F4_MM")},
  Activos: {jl("Total Activos_MM")},
}};

const GEO = {{
  regiones:     {json.dumps(geo["region_short"].tolist(),    ensure_ascii=False)},
  reg_geojson:  {json.dumps(geo["region_geojson"].tolist(),  ensure_ascii=False)},
  N:            {json.dumps(geo["N"].tolist())},
  socios:       {json.dumps([int(v) if pd.notna(v) else 0 for v in geo["Socios"]])},
  vab_mm:       {json.dumps([round(float(v),2) if pd.notna(v) else 0 for v in geo["Vab_MM"]])},
  trab:         {json.dumps([int(v) if pd.notna(v) else 0 for v in geo["Trab_SII"]])},
  soc_h:        {soc_h},
  soc_m:        {soc_m},
}};
const GEOJSON_OK = {'true' if geojson_ok else 'false'};
const GEOJSON    = {geojson_str};

const AZ={json.dumps(C['azul_med'])},GN={json.dumps(C['verde_med'])},RJ={json.dumps(C['rojo'])};
const AM={json.dumps(C['amarillo'])},NA={json.dumps(C['naranja'])},AZ2={json.dumps(C['azul_osc'])};

const CFG = {{displayModeBar:false, responsive:true}};
const LAY = {{
  font:{{family:"Segoe UI,Arial,sans-serif",color:"#1A1A2E",size:12}},
  paper_bgcolor:"rgba(0,0,0,0)", plot_bgcolor:"rgba(0,0,0,0)",
  margin:{{t:48,b:58,l:62,r:20}},
  legend:{{orientation:"h",yanchor:"bottom",y:-0.38,xanchor:"center",x:0.5}},
  yaxis:{{gridcolor:"#E5E7EB",zeroline:true,zerolinecolor:"#D1D5DB"}},
  xaxis:{{type:"category"}},
}};

function tit(t){{return {{text:t,font:{{size:13,color:AZ2}}}}}}

// Barra simple
function bar(x,y,name,color,extra={{}}){{
  return {{type:"bar",x,y,name,marker:{{color}},
    hovertemplate:"<b>%{{x}}</b><br>"+name+": %{{y:.2f}}<extra></extra>",...extra}};
}}
// Línea simple
function ln(x,y,name,color,extra={{}}){{
  return {{type:"scatter",mode:"lines+markers",x,y,name,
    line:{{color,width:2.5}},marker:{{size:7}},
    hovertemplate:"<b>%{{x}}</b><br>"+name+": %{{y:.2f}}<extra></extra>",...extra}};
}}

// ── TABS ──────────────────────────────────────────────────────────────────────
const rendered={{}};
function tab(id,btnId){{
  document.querySelectorAll(".tab-pane").forEach(p=>p.classList.remove("on"));
  document.querySelectorAll(".tab").forEach(b=>b.classList.remove("on"));
  document.getElementById("tab-"+id).classList.add("on");
  document.getElementById(btnId).classList.add("on");
  if(!rendered[id]){{ rendered[id]=true; RENDER[id](); }}
}}
const RENDER={{
  produccion: renderProduccion,
  ingreso:    renderIngreso,
  financiera: ()=>{{renderFinanciera();updateScatter();}},
  entidad:    updateEntidad,
  mapa:       renderMapa,
  geografia:  renderGeo,
  panel:      renderPanel,
}};

// ══ PRODUCCIÓN Y VAB ════════════════════════════════════════════════════════
function renderProduccion(){{
  // B1g por año — barras con N entidades en hover
  Plotly.newPlot("g-b1g-año",[
    bar(SECTOR.años, SECTOR.B1g, "VAB Bruto [B1g]", AZ, {{
      customdata: SECTOR.N,
      hovertemplate:"<b>%{{x}}</b><br>B1g: MM$%{{y:.1f}}<br>N entidades: %{{customdata}}<extra></extra>",
      text: SECTOR.B1g.map(v=>v!=null?v.toFixed(1):""), textposition:"outside",
    }}),
  ],{{...LAY, title:tit("VAB Bruto del Sector [B1g] — Sector CAC Completo (MM$ CLP)"),
     yaxis:{{...LAY.yaxis,title:"MM$"}}}}, CFG);

  // Descomposición apilada del VAB
  Plotly.newPlot("g-vab-comp",[
    bar(SECTOR.años,SECTOR.D1,  "Remuneraciones [D1]", AZ,  {{stackgroup:"s"}}),
    bar(SECTOR.años,SECTOR.P51d,"Depreciación [P51d]",  AM,  {{stackgroup:"s"}}),
    bar(SECTOR.años,SECTOR.B2g, "Excedente [B2g]",      GN,  {{stackgroup:"s"}}),
  ],{{...LAY, barmode:"relative",
     title:tit("Descomposición B1g = D1 + P51d + B2g (MM$)"),
     yaxis:{{...LAY.yaxis,title:"MM$"}}}}, CFG);

  // P1 vs P2
  const p2neg = SECTOR.P2.map(v=>v!=null?-v:null);
  Plotly.newPlot("g-p1-p2",[
    bar(SECTOR.años,SECTOR.P1, "Producción [P1]",        AZ, {{opacity:0.9}}),
    bar(SECTOR.años,p2neg,     "Cons. Intermedio −[P2]",  RJ, {{opacity:0.9}}),
  ],{{...LAY, barmode:"overlay",
     title:tit("Producción [P1] y Consumo Intermedio [P2] (MM$)"),
     yaxis:{{...LAY.yaxis,title:"MM$"}},
     annotations:[{{x:1,y:1,xref:"paper",yref:"paper",xanchor:"right",yanchor:"top",
       text:"α={ALPHA_CI} — MIP Chile 2018",showarrow:false,
       font:{{size:10,color:"#374151"}},bgcolor:"rgba(255,255,255,0.85)",borderpad:4}}]}}, CFG);

  // Cobertura: superavit + deficit
  const sin = SECTOR.N.map((n,i)=>n - SECTOR.N_def[i]);
  Plotly.newPlot("g-cobertura",[
    bar(SECTOR.años, sin,           "Con superávit",          GN, {{stackgroup:"s"}}),
    bar(SECTOR.años, SECTOR.N_def,  "Con déficit (rem.<0)",    RJ, {{stackgroup:"s"}}),
    {{type:"scatter",mode:"lines+markers",x:SECTOR.años,y:SECTOR.N_cmf,
      name:"De ellas: CMF", yaxis:"y2",
      line:{{color:NA,width:2,dash:"dot"}},marker:{{size:6}},
      hovertemplate:"<b>%{{x}}</b><br>CMF: %{{y}}<extra></extra>"}},
  ],{{...LAY, barmode:"stack",
     title:tit("Cobertura del Consolidado (N° entidades)"),
     yaxis:{{...LAY.yaxis,title:"N° entidades"}},
     yaxis2:{{title:"N° CMF",overlaying:"y",side:"right",showgrid:false}}}}, CFG);
}}

// ══ GENERACIÓN DEL INGRESO ═══════════════════════════════════════════════════
function renderIngreso(){{
  Plotly.newPlot("g-d1-b2g",[
    bar(SECTOR.años, SECTOR.D1,  "Remun. [D1]",         AZ),
    bar(SECTOR.años, SECTOR.B2g, "Excedente bruto [B2g]",GN),
  ],{{...LAY, barmode:"group",
     title:tit("Remuneraciones [D1] y Excedente Bruto [B2g] (MM$)"),
     yaxis:{{...LAY.yaxis,title:"MM$"}}}}, CFG);

  Plotly.newPlot("g-d1-pct",[
    ln(SECTOR.años, SECTOR.D1_pct,  "% D1 / B1g", AZ),
    ln(SECTOR.años, SECTOR.B2g_pct, "% B2g / B1g",GN),
  ],{{...LAY,
     title:tit("Participación D1 y B2g en el VAB (%)"),
     yaxis:{{...LAY.yaxis,title:"%",range:[0,100]}},
     shapes:[{{type:"line",x0:0,x1:1,xref:"paper",y0:100,y1:100,
       line:{{color:"#9CA3AF",width:1,dash:"dot"}}}}]}}, CFG);

  const colors_rem = SECTOR.Rem.map(v=>v!=null&&v>=0?GN:RJ);
  Plotly.newPlot("g-remanente",[
    {{type:"bar",x:SECTOR.años,y:SECTOR.Rem,
      marker:{{color:colors_rem}},
      name:"Remanente del Período",
      hovertemplate:"<b>%{{x}}</b><br>Remanente: MM$%{{y:.2f}}<extra></extra>"}},
  ],{{...LAY,
     title:tit("Remanente Agregado del Sector (MM$)"),
     yaxis:{{...LAY.yaxis,title:"MM$"}}}}, CFG);

  Plotly.newPlot("g-deficit",[
    bar(SECTOR.años, SECTOR.pct_def, "% entidades con déficit", AM, {{
      text:SECTOR.pct_def.map(v=>v!=null?v.toFixed(1)+"%":""),textposition:"outside"}}),
  ],{{...LAY,
     title:tit("% Entidades con Remanente Negativo"),
     yaxis:{{...LAY.yaxis,title:"%",range:[0,80]}}}}, CFG);
}}

// ══ CUENTA FINANCIERA ════════════════════════════════════════════════════════
function renderFinanciera(){{
  Plotly.newPlot("g-f2f4",[
    bar(SECTOR.años, SECTOR.F2, "Depósitos captados [F2]",  AZ),
    bar(SECTOR.años, SECTOR.F4, "Colocaciones netas [F4]",  GN),
  ],{{...LAY, barmode:"group",
     title:tit("Cuenta Financiera: F2 Depósitos y F4 Colocaciones (MM$)"),
     yaxis:{{...LAY.yaxis,title:"MM$"}}}}, CFG);

  Plotly.newPlot("g-mora-solv",[
    ln(SECTOR.años, SECTOR.mora, "Mora promedio [CIRIEC]",    RJ),
    ln(SECTOR.años, SECTOR.solv, "Solvencia promedio [CIRIEC]",GN,
       {{yaxis:"y2",line:{{color:GN,width:2.5,dash:"dash"}},marker:{{size:7}}}}),
  ],{{...LAY,
     title:tit("Mora y Solvencia Promedio del Sector (%)"),
     yaxis:{{...LAY.yaxis,title:"Mora (%)"}},
     yaxis2:{{title:"Solvencia (%)",overlaying:"y",side:"right",showgrid:false}}}}, CFG);
}}

function updateScatter(){{
  const año = parseInt(document.getElementById("sel-scat-año").value);
  const d   = SCAT[año];
  if(!d||!d.ents.length){{Plotly.react("g-scatter",[],LAY,CFG);return;}}
  const ma = d.mora.reduce((a,b)=>a+b,0)/d.mora.length;
  const sa = d.solv.reduce((a,b)=>a+b,0)/d.solv.length;

  // Separar CMF y no CMF para colorear distinto
  const xCmf=[],yCmf=[],tCmf=[],aCmf=[];
  const xReg=[],yReg=[],tReg=[],aReg=[];
  d.ents.forEach((e,i)=>{{
    if(d.cmf[i]){{ xCmf.push(d.mora[i]);yCmf.push(d.solv[i]);tCmf.push(e);aCmf.push(d.act[i]); }}
    else         {{ xReg.push(d.mora[i]);yReg.push(d.solv[i]);tReg.push(e);aReg.push(d.act[i]); }}
  }});

  Plotly.react("g-scatter",[
    {{type:"scatter",mode:"markers",name:"Históricas",
      x:xReg,y:yReg,text:tReg,
      marker:{{size:aReg.map(a=>Math.max(9,Math.min(44,a*0.9))),color:AZ,
               opacity:0.7,line:{{width:1,color:"white"}}}},
      hovertemplate:"<b>%{{text}}</b><br>Mora: %{{x:.1f}}%<br>Solvencia: %{{y:.1f}}%<extra></extra>"}},
    {{type:"scatter",mode:"markers",name:"CMF supervisadas",
      x:xCmf,y:yCmf,text:tCmf,
      marker:{{size:aCmf.map(a=>Math.max(12,Math.min(50,a*0.9))),color:NA,symbol:"diamond",
               opacity:0.85,line:{{width:1.5,color:"white"}}}},
      hovertemplate:"<b>★ %{{text}}</b><br>Mora: %{{x:.1f}}%<br>Solvencia: %{{y:.1f}}%<extra></extra>"}},
  ],{{...LAY,
     title:tit("Mora vs. Solvencia [CIRIEC] · "+año+" (tamaño ∝ activos)"),
     xaxis:{{...LAY.xaxis,type:"linear",title:"Mora (%)"}},
     yaxis:{{...LAY.yaxis,title:"Solvencia (%)"}},
     shapes:[
       {{type:"line",x0:ma,x1:ma,y0:0,y1:1,yref:"paper",line:{{color:RJ,width:1.5,dash:"dot"}}}},
       {{type:"line",x0:0,x1:1,xref:"paper",y0:sa,y1:sa,line:{{color:GN,width:1.5,dash:"dot"}}}},
     ],
     annotations:[
       {{x:ma,y:1,yref:"paper",text:"Mora prom "+ma.toFixed(1)+"%",showarrow:false,
         font:{{color:RJ,size:10}},xanchor:"left",yanchor:"top"}},
       {{x:1,xref:"paper",y:sa,text:"Solvencia prom "+sa.toFixed(1)+"%",showarrow:false,
         font:{{color:GN,size:10}},xanchor:"right"}},
     ]}}, CFG);
}}

// ══ POR ENTIDAD ═══════════════════════════════════════════════════════════════
const VAR_LBL={{
  B1g:"VAB Bruto [B1g] (MM$)", P1:"Producción [P1] (MM$)",
  P2:"Consumo Intermedio [P2] (MM$)", D1:"Remuneraciones [D1] (MM$)",
  B2g:"Excedente Bruto [B2g] (MM$)", P51d:"Depreciación [P51d] (MM$)",
  B1n:"VAB Neto [B1n] (MM$)", Rem:"Remanente del Período (MM$)",
}};

function updateEntidad(){{
  const sel = Array.from(document.getElementById("sel-ent").selectedOptions).map(o=>o.value);
  const v   = document.getElementById("sel-ent-var").value;
  if(!sel.length) return;

  const traces = sel.map((ent,i)=>{{
    const d = ER[ent]; if(!d) return null;
    const color = CMF_NAMES.has(ent) ? NA : SEQ[i%SEQ.length];
    // Separar puntos CMF (2026) del resto para marcador diferente
    const xHist=[],yHist=[],xCmf=[],yCmf=[];
    d.años.forEach((a,j)=>{{
      if(d.cmf[j]){{ xCmf.push(a); yCmf.push(d[v][j]); }}
      else          {{ xHist.push(a);yHist.push(d[v][j]);}}
    }});
    const traces_ent = [{{
      type:"scatter",mode:"lines+markers",name:ent,x:xHist,y:yHist,
      line:{{color,width:2.5}},marker:{{size:7}},showlegend:true,
      hovertemplate:"<b>"+ent+"</b><br>%{{x}}: MM$%{{y:.2f}}<extra></extra>",
    }}];
    if(xCmf.length){{
      traces_ent.push({{
        type:"scatter",mode:"markers",name:ent+" (CMF Feb-2026)",
        x:xCmf,y:yCmf,
        marker:{{size:12,color:NA,symbol:"star",line:{{width:1.5,color:"white"}}}},
        showlegend:true,
        hovertemplate:"<b>★ "+ent+"</b><br>CMF Feb-2026: MM$%{{y:.2f}}<extra></extra>",
      }});
    }}
    return traces_ent;
  }}).filter(Boolean).flat();

  Plotly.react("g-entidad", traces, {{...LAY,
    title:tit(VAR_LBL[v]+" — Comparativo por Entidad (★ = CMF Feb-2026)"),
    yaxis:{{...LAY.yaxis,title:"MM$"}}}}, CFG);
}}

// ══ MAPA SVG COROPLÉTICO ══════════════════════════════════════════════════════
const SVG_PATHS    = {json.dumps(regions_svg, ensure_ascii=False)};
const SVG_CENTROIDS= {json.dumps(centroids_svg, ensure_ascii=False)};

const PALETTES = {{
  blues:   ["#DBEAFE","#93C5FD","#3B82F6","#1D4ED8","#1E3A8A"],
  greens:  ["#DCFCE7","#86EFAC","#22C55E","#16A34A","#14532D"],
  oranges: ["#FEF3C7","#FCD34D","#F97316","#C2410C","#7C2D12"],
  reds:    ["#FEE2E2","#FCA5A5","#EF4444","#B91C1C","#7F1D1D"],
  purples: ["#EDE9FE","#C4B5FD","#8B5CF6","#6D28D9","#3B0764"],
}};

const VAR_GEO_LBL = {{
  N:"N° Cooperativas", socios:"Total Socios",
  vab_mm:"VAB estimado (MM$)", trab:"Trabajadores SII",
}};

// Map: nombre GeoJSON → índice en GEO arrays
const GEO_IDX = {{}};
GEO.reg_geojson.forEach((r,i) => GEO_IDX[r] = i);

let selectedRegion = null;

function getColor(val, vals, palette){{
  const pal  = PALETTES[palette] || PALETTES.blues;
  const vmax = Math.max(...vals.filter(v=>v>0));
  if(!val || vmax === 0) return "#E2E8F0";
  const t = Math.sqrt(val / vmax);            // raíz para mejor contraste
  const idx = Math.min(Math.floor(t * (pal.length-1)), pal.length-1);
  return pal[idx];
}}

let mapaInitialized = false;

function renderMapa(){{
  const mapDiv = document.getElementById("g-mapa");

  if (!GEOJSON_OK) {{
    mapDiv.innerHTML =
      '<div style="padding:40px;text-align:center;color:#6B7280;font-size:13px">' +
      '<div style="font-size:32px;margin-bottom:12px">🗺️</div>' +
      '<b>Regional.geojson no encontrado.</b><br><br>' +
      'Coloque el archivo en la misma carpeta que el script y regenere el dashboard.' +
      '</div>';
    document.getElementById("map-legend").innerHTML = "";
    return;
  }}

  const varKey    = document.getElementById("map-var").value;
  const pal       = document.getElementById("map-pal").value;
  const vals      = GEO[varKey];
  const label     = VAR_GEO_LBL[varKey];
  const palColors = PALETTES[pal];
  const cs        = palColors.map((c, i) => [i / (palColors.length - 1), c]);

  const hoverText = GEO.reg_geojson.map((r, i) =>
    "<b>" + GEO.regiones[i] + "</b><br>" +
    "CAC activas: " + GEO.N[i] + "<br>" +
    "Socios: " + GEO.socios[i].toLocaleString("es-CL") + "<br>" +
    "VAB est.: MM$" + GEO.vab_mm[i].toFixed(2) + "<br>" +
    "Trab. SII: " + GEO.trab[i]
  );

  // IDs reales del GeoJSON → garantiza que Chile siempre sea visible
  const allIds = GEOJSON.features ? GEOJSON.features.map(f => f.id).filter(Boolean) : [];
  const vmax   = Math.max(...vals.filter(v => v > 0), 1);

  Plotly.react(mapDiv, [
    // Capa 1: fondo gris para todas las regiones (siempre visible)
    {{
      type: "choropleth", geojson: GEOJSON, featureidkey: "id",
      locations: allIds, z: allIds.map(() => 0),
      colorscale: [["0","#CBD5E0"],["1","#CBD5E0"]],
      showscale: false, hoverinfo: "skip",
      marker: {{ line: {{ color: "white", width: 0.8 }} }},
    }},
    // Capa 2: regiones con datos coloreadas por variable
    {{
      type: "choropleth", geojson: GEOJSON, featureidkey: "id",
      locations: GEO.reg_geojson, z: vals,
      text: hoverText, hovertemplate: "%{{text}}<extra></extra>",
      colorscale: cs, showscale: false, zmin: 0, zmax: vmax,
      marker: {{ line: {{ color: "white", width: 0.8 }} }},
    }},
  ], {{
    font: {{ family: "Segoe UI,Arial,sans-serif", color: "#1A1A2E", size: 12 }},
    paper_bgcolor: "rgba(0,0,0,0)",
    title: tit("Distribución CAC por Región · " + label),
    geo: {{
      visible: false,
      bgcolor: "#EFF6FF",
      lonaxis: {{ range: [-77, -63] }},
      lataxis: {{ range: [-58, -16] }},
    }},
    margin: {{ t: 50, b: 10, l: 10, r: 10 }},
    height: 660,
  }}, CFG);

  renderLeyenda(vals, label, pal);

  if (!mapaInitialized) {{
    mapaInitialized = true;
    mapDiv.on("plotly_click", function(data) {{
      if (data.points && data.points.length > 0) {{
        const loc = data.points[0].location;
        const i = GEO.reg_geojson.indexOf(loc);
        if (i >= 0) showDetail(loc, i);
      }}
    }});
  }}
}}

function renderLeyenda(vals, label, pal){{
  const pal_colors = PALETTES[pal] || PALETTES.blues;
  const vmax = Math.max(...vals.filter(v=>v>0));
  const steps = 5;
  let html = `<div style="font-size:9.5px;font-weight:700;text-transform:uppercase;
    letter-spacing:.05em;color:#6B7280;margin-bottom:6px">${{label}}</div>`;
  for(let i=steps-1;i>=0;i--){{
    const v = Math.round(vmax * i/(steps-1));
    html += `<div style="display:flex;align-items:center;gap:6px;margin-bottom:3px">
      <div style="width:16px;height:12px;border-radius:2px;background:${{pal_colors[i]}};
        border:1px solid rgba(0,0,0,.1)"></div>
      <span style="font-size:10px;color:#374151">${{v.toLocaleString("es-CL")}}</span>
    </div>`;
  }}
  html += `<div style="display:flex;align-items:center;gap:6px;margin-top:4px">
    <div style="width:16px;height:12px;border-radius:2px;background:#E2E8F0;
      border:1px solid rgba(0,0,0,.1)"></div>
    <span style="font-size:10px;color:#374151">Sin datos</span>
  </div>`;
  document.getElementById("map-legend").innerHTML = html;
}}

function showDetail(regName, idx){{
  const detail = document.getElementById("map-detail");
  if(idx === undefined){{ detail.style.display="none"; return; }}
  const shortName = regName.replace("Región de ","").replace("Región del ","")
                           .replace("Región Metropolitana de Santiago","R. Metropolitana");
  detail.style.display = "block";
  detail.innerHTML = `
    <b style="color:#1F3864;font-size:12px">${{shortName}}</b><br>
    <hr style="border:none;border-top:1px solid #CBD5E0;margin:5px 0">
    🏢 CAC activas: <b>${{GEO.N[idx]}}</b><br>
    👥 Total socios: <b>${{GEO.socios[idx].toLocaleString("es-CL")}}</b><br>
    💰 VAB est.: <b>MM$${{GEO.vab_mm[idx].toFixed(2)}}</b><br>
    👷 Trab. SII: <b>${{GEO.trab[idx]}}</b>
  `;
}}

// ══ GEOGRAFÍA (barras + pie) ══════════════════════════════════════════════════
function renderGeo(){{
  const idx = GEO.N.map((_,i)=>i).sort((a,b)=>GEO.N[b]-GEO.N[a]);
  const regs = idx.map(i=>GEO.regiones[i]);

  Plotly.newPlot("g-geo-n",[{{
    type:"bar",orientation:"h",
    x:idx.map(i=>GEO.N[i]), y:regs,
    marker:{{color:AZ}},
    text:idx.map(i=>GEO.N[i]),textposition:"outside",
    hovertemplate:"<b>%{{y}}</b><br>CAC: %{{x}}<extra></extra>",
  }}],{{...LAY,
    title:tit("CAC por Región [CIRIEC — Directorio]"),
    xaxis:{{...LAY.xaxis,type:"linear",title:"N° entidades"}},
    yaxis:{{autorange:"reversed"}},margin:{{...LAY.margin,l:160}},showlegend:false}},CFG);

  const idx2 = GEO.socios.map((_,i)=>i).sort((a,b)=>GEO.socios[b]-GEO.socios[a]);
  Plotly.newPlot("g-geo-socios",[{{
    type:"bar",orientation:"h",
    x:idx2.map(i=>GEO.socios[i]), y:idx2.map(i=>GEO.regiones[i]),
    marker:{{color:GN}},
    text:idx2.map(i=>GEO.socios[i].toLocaleString("es-CL")),textposition:"outside",
    hovertemplate:"<b>%{{y}}</b><br>Socios: %{{x:,}}<extra></extra>",
  }}],{{...LAY,
    title:tit("Total Socios por Región [CIRIEC / ONU TSE §4]"),
    xaxis:{{...LAY.xaxis,type:"linear",title:"N° socios"}},
    yaxis:{{autorange:"reversed"}},margin:{{...LAY.margin,l:160}},showlegend:false}},CFG);

  const idx3 = GEO.vab_mm.map((_,i)=>i).sort((a,b)=>GEO.vab_mm[b]-GEO.vab_mm[a]);
  Plotly.newPlot("g-geo-vab",[{{
    type:"bar",orientation:"h",
    x:idx3.map(i=>GEO.vab_mm[i]), y:idx3.map(i=>GEO.regiones[i]),
    marker:{{color:AM}},
    text:idx3.map(i=>GEO.vab_mm[i].toFixed(2)),textposition:"outside",
    hovertemplate:"<b>%{{y}}</b><br>VAB: MM$%{{x:.2f}}<extra></extra>",
  }}],{{...LAY,
    title:tit("VAB estimado [B1g] por Región — método SII (MM$)"),
    xaxis:{{...LAY.xaxis,type:"linear",title:"MM$"}},
    yaxis:{{autorange:"reversed"}},margin:{{...LAY.margin,l:160}},showlegend:false}},CFG);

  Plotly.newPlot("g-geo-genero",[{{
    type:"pie",
    labels:["Socios Hombres","Socias Mujeres"],
    values:[GEO.soc_h, GEO.soc_m],
    marker:{{colors:[AZ,NA]}},
    textinfo:"label+percent",
    hovertemplate:"<b>%{{label}}</b><br>%{{value:,.0f}} (%{{percent}})<extra></extra>",
    hole:0.42,textfont:{{size:12}},
  }}],{{...LAY,
    title:tit("Distribución de Socios por Género [CIRIEC §4 / ONU TSE]"),
    showlegend:true,legend:{{orientation:"h",y:-0.1}},
    margin:{{t:48,b:40,l:20,r:20}}}},CFG);
}}

// ══ PANEL DAES ════════════════════════════════════════════════════════════════
function renderPanel(){{
  const cols=["Entidad","Región","Socios","Soc.H","Soc.M","Trab.SII",
              "Prod.MM$","VAB.MM$","Empl.CMF","Of.CMF","CMF","Transf.19862"];
  let html="<table><thead><tr>"+cols.map(c=>"<th>"+c+"</th>").join("")+"</tr></thead><tbody>";
  PANEL_ROWS.forEach(r=>{{
    const cmf = r.CMF==="✔ CMF" ? '<span class="cmf-badge">★ CMF</span>' : "—";
    html+="<tr>"+
      "<td><b>"+r.Entidad+"</b></td><td>"+r.Región+"</td>"+
      "<td style='text-align:right'>"+r.Socios+"</td>"+
      "<td style='text-align:right'>"+r["Soc.H"]+"</td>"+
      "<td style='text-align:right'>"+r["Soc.M"]+"</td>"+
      "<td style='text-align:right'>"+r["Trab.SII"]+"</td>"+
      "<td style='text-align:right'>"+r["Prod.MM$"]+"</td>"+
      "<td style='text-align:right'><b>"+r["VAB.MM$"]+"</b></td>"+
      "<td style='text-align:right'>"+r["Empl.CMF"]+"</td>"+
      "<td style='text-align:right'>"+r["Of.CMF"]+"</td>"+
      "<td>"+cmf+"</td>"+
      "<td style='text-align:right'>"+r["Transf.19862"]+"</td>"+
    "</tr>";
  }});
  html+="</tbody></table>";
  document.getElementById("panel-tabla").innerHTML=html;
}}

// ── Inicializar ───────────────────────────────────────────────────────────────
window.addEventListener("load",()=>RENDER["produccion"]());
</script>
</body>
</html>"""

with open(OUT, "w", encoding="utf-8") as f:
    f.write(HTML)
print(f"\n✅  Generado: {OUT}")
webbrowser.open(os.path.abspath(OUT))
