"""
╔══════════════════════════════════════════════════════════════════════╗
║  HUELLA SOCIAL — Dashboard Cuenta Satélite CAC  v5                  ║
║  Sector CAC completo: histórico 2014-2024 + CMF Feb-2026            ║
║  Mapa coroplético de Chile por región integrado                      ║
║  Variables alineadas SCN 2025 / Manual ONU TSE / CIRIEC             ║
║                                                                      ║
║  Ejecutar:  python dashboard_HuellaSocial.py                            ║
║  Genera:    dashboard_HuellaSocial.html  (se abre en el navegador)      ║
║  Requiere:  pip install plotly pandas openpyxl                       ║
║  Requiere:  Regional.geojson en la misma carpeta (BCN)               ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import pandas as pd
import numpy as np
import json, os, webbrowser
import zipfile

# Posicionarse en la carpeta del script para que los paths relativos funcionen
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# ── ARCHIVOS ─────────────────────────────────────────────────────────────────
F_CAC     = "consolidado_CAC_balances_resultados_v2.xlsx"
F_PANEL   = "Consolidado_cooperativas.xlsx"
F_GEOJSON = "Regional.geojson"
F_SHP     = "Regional.shp"
ZIP_GEO   = "regional_shp.zip"
OUT       = "dashboard_HuellaSocial.html"

# Extraer shapefile si no existe pero hay un zip (útil al clonar desde GitHub)
if not os.path.exists(F_SHP) and os.path.exists(ZIP_GEO):
    print("Extrayendo shapefile regional...")
    with zipfile.ZipFile(ZIP_GEO, "r") as z:
        z.extractall(".")

# ── PARÁMETROS SCN (§4.3.2 memoria) ─────────────────────────────────────────
ALPHA_CI = 0.3776   # CI/P — MIP Chile 2018 sector 94 (Intermediación financiera)
N_MIN    = 3        # Umbral mínimo de entidades para incluir un año en gráficos sectoriales

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

# ── Agregado SIN Coopeuch para gráficos de descomposición (escala homogénea) ─
# Coopeuch (RUT 82878900-7) distorsiona los promedios 2015/2016 por ser ~100x
# más grande que el resto del consolidado histórico
RUT_COOPEUCH = "82878900-7"
er_sin_coopeuch = er[er["RUT"].astype(str).str.strip() != RUT_COOPEUCH].copy()
sector_comp = er_sin_coopeuch.groupby("Año").agg(
    D1_MM    =("D1_MM",   "sum"),
    P51d_MM  =("P51d_MM", "sum"),
    B2g_MM   =("B2g_MM",  "sum"),
    B1g_MM   =("B1g_MM",  "sum"),
    Rem_MM   =("Rem_MM",  "sum"),
    N        =("RUT",     "count"),
    N_cmf    =("fuente_cmf", "sum"),
    N_def    =("Rem",     lambda x: (x.fillna(0) < 0).sum()),
).reset_index()
sector_comp["pct_def"] = (sector_comp["N_def"] / sector_comp["N"] * 100).round(1)
for col in ["D1_MM","P51d_MM","B2g_MM","B1g_MM","Rem_MM"]:
    sector_comp[col.replace("_MM","_avg")] = (sector_comp[col] / sector_comp["N"]).round(2)
sector_comp_filt = sector_comp[sector_comp["N"] >= N_MIN].copy()

mora_yr = bg[bg["mora"].notna()].groupby("Año")["mora"].mean().round(2).reset_index()
# Solvencia: separar histórico (sin 2026) para no distorsionar el eje
solv_yr     = bg[bg["solv"].notna()].groupby("Año")["solv"].mean().round(1).reset_index()
solv_hist   = solv_yr[solv_yr["Año"] != 2026].copy()
solv_cmf26  = solv_yr[solv_yr["Año"] == 2026].copy()
f2_yr   = bg.groupby("Año")["F2_MM"].sum().round(1).reset_index()
f4_yr   = bg.groupby("Año")["F4_MM"].sum().round(1).reset_index()
act_yr  = bg.groupby("Año")["Total Activos_MM"].sum().round(1).reset_index()

sector = sector.merge(mora_yr, on="Año", how="left") \
               .merge(solv_yr, on="Año", how="left") \
               .merge(f2_yr,   on="Año", how="left") \
               .merge(f4_yr,   on="Año", how="left") \
               .merge(act_yr,  on="Año", how="left")

# ── Promedios por entidad (eliminan sesgo de cobertura) ───────────────────────
for col in ["P1_MM","P2_MM","B1g_MM","D1_MM","P51d_MM","B2g_MM","B1n_MM","Rem_MM"]:
    sector[col.replace("_MM","_avg")] = (sector[col] / sector["N"]).round(2)

# ── Sector filtrado: solo años con N ≥ N_MIN (después de todos los merges) ────
sector_filt = sector[sector["N"] >= N_MIN].copy()
print(f"  Años excluidos por N<{N_MIN}: {sorted(set(sector['Año']) - set(sector_filt['Año']))}")

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
def jl(col, d=2, src=None):
    df = src if src is not None else sector_filt
    return json.dumps([round(float(v), d) if pd.notna(v) else None for v in df[col]])

def ji(col, src=None):
    df = src if src is not None else sector_filt
    return json.dumps([int(v) for v in df[col]])

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
  <button class="tab"     onclick="tab('glosario','t8')"  id="t8">📖 Glosario de Variables</button>
  <button class="tab"     onclick="tab('brechas','t9')"   id="t9">🔍 Brechas de Datos</button>
</div>

<!-- ══ TAB: PRODUCCIÓN Y VAB ══════════════════════════════════════════════ -->
<div class="tab-pane on" id="tab-produccion">
  <div class="info-box">
    <b>SCN 2025 §7.82 · §4.3.3 memoria:</b>
    P1 (Producción) = Total Ingresos de Operación.
    P2 (Consumo Intermedio) = α×P1, α={ALPHA_CI} (MIP Chile 2018, sector 94).
    B1g (VAB Bruto) = P1−P2 = {1-ALPHA_CI:.4f}×P1. B1n (VAB Neto) = B1g−P51d. &nbsp;·&nbsp;
    <span class="cmf-dot"></span><b>Puntos naranjos</b> = datos CMF Feb-2026 (método directo).
  </div>
  <div class="warn-box">
    <b>⚠ Nota metodológica (cobertura variable):</b>
    Los gráficos muestran <b>promedios por entidad</b> para hacer comparables los años.
    El número de entidades con datos varía por año (N se indica en cada barra).
    Solo se muestran años con N ≥ {N_MIN} entidades. El gráfico de cobertura muestra la evolución de N.
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
    Los valores de D1 y B2g se expresan como <b>promedio por entidad</b> para hacer comparables los años con distinta cobertura.
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

<!-- ══ TAB: GLOSARIO DE VARIABLES ══════════════════════════════════════════ -->
<div class="tab-pane" id="tab-glosario">
  <div class="info-box">
    <b>Guía de referencia metodológica:</b> cada código de variable utilizado en el dashboard,
    su nombre completo, fórmula de cálculo, marco teórico de referencia y fuente de datos primaria.
    El dashboard sigue la nomenclatura del <b>SCN 2025</b>, el <b>Manual ONU TSE (2014)</b> y el marco <b>CIRIEC</b>.
  </div>

  <!-- Sección: Cuenta de Producción -->
  <div style="margin-bottom:22px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;padding-bottom:6px;
      border-bottom:2px solid {C['azul_med']}">
      <span style="font-size:18px">📊</span>
      <span style="font-size:13px;font-weight:800;color:{C['azul_osc']};text-transform:uppercase;
        letter-spacing:.05em">Cuenta de Producción · SCN 2025 §7</span>
    </div>
    <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th style="width:80px">Código</th>
        <th style="width:220px">Nombre Completo</th>
        <th>Fórmula / Definición</th>
        <th style="width:150px">Marco Teórico</th>
        <th style="width:160px">Fuente de Datos</th>
      </tr></thead>
      <tbody>
        <tr>
          <td><b style="color:{C['azul_med']};font-size:14px">P1</b></td>
          <td><b>Producción</b></td>
          <td>Total Ingresos de Operación del Estado de Resultados.
              Para cooperativas financieras, la producción se aproxima por los ingresos de su actividad principal
              (intereses, reajustes, inversiones, otros ingresos de operación).</td>
          <td>SCN 2025 §7.82<br><span style="color:{C['texto_sub']};font-size:10px">§4.3.2 memoria</span></td>
          <td>Estados de Resultados CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['azul_med']};font-size:14px">P2</b></td>
          <td><b>Consumo Intermedio</b></td>
          <td>α × P1 &nbsp;·&nbsp; donde α = {ALPHA_CI} es la razón CI/Producción del sector de
              Intermediación Financiera (sector 94) extraída de la Matriz Insumo-Producto de Chile 2018.
              Representa los bienes y servicios consumidos en el proceso productivo.</td>
          <td>SCN 2025 §6.147<br><span style="color:{C['texto_sub']};font-size:10px">MIP Chile 2018</span></td>
          <td>MIP Chile 2018, sector 94<br><span style="color:{C['texto_sub']};font-size:10px">α = {ALPHA_CI} (parámetro fijo)</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['verde_osc']};font-size:14px">B1g</b></td>
          <td><b>Valor Agregado Bruto (VAB)</b></td>
          <td>P1 − P2 &nbsp;=&nbsp; (1 − α) × P1 &nbsp;=&nbsp; {1-ALPHA_CI:.4f} × P1.
              Mide el valor creado por el sector en el proceso de producción, antes de deducir el consumo de capital fijo.
              Es la variable central de la Cuenta Satélite.</td>
          <td>SCN 2025 §6.222<br><span style="color:{C['texto_sub']};font-size:10px">CIRIEC §3 / §4.3.3 memoria</span></td>
          <td>Calculado sobre ER CMF<br><span style="color:{C['texto_sub']};font-size:10px">+ método indirecto SII (tramos)</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['verde_osc']};font-size:14px">B1n</b></td>
          <td><b>Valor Agregado Neto (VAN)</b></td>
          <td>B1g − P51d. El VAB después de deducir la depreciación del capital fijo.
              Representa la riqueza neta creada por el sector.</td>
          <td>SCN 2025 §6.222<br><span style="color:{C['texto_sub']};font-size:10px">§4.3.3 memoria</span></td>
          <td>Calculado (B1g − P51d)<br><span style="color:{C['texto_sub']};font-size:10px">ER CMF</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['texto_sub']};font-size:14px">P51d</b></td>
          <td><b>Consumo de Capital Fijo (Depreciación)</b></td>
          <td>Depreciaciones y Amortizaciones del Estado de Resultados.
              Mide el valor del capital fijo consumido durante el período contable.</td>
          <td>SCN 2025 §6.179<br><span style="color:{C['texto_sub']};font-size:10px">§4.3.3 memoria</span></td>
          <td>Estados de Resultados CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
      </tbody>
    </table>
    </div>
  </div>

  <!-- Sección: Cuenta de Generación del Ingreso -->
  <div style="margin-bottom:22px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;padding-bottom:6px;
      border-bottom:2px solid {C['verde_med']}">
      <span style="font-size:18px">💼</span>
      <span style="font-size:13px;font-weight:800;color:{C['azul_osc']};text-transform:uppercase;
        letter-spacing:.05em">Cuenta de Generación del Ingreso · SCN 2025 §8</span>
    </div>
    <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th style="width:80px">Código</th>
        <th style="width:220px">Nombre Completo</th>
        <th>Fórmula / Definición</th>
        <th style="width:150px">Marco Teórico</th>
        <th style="width:160px">Fuente de Datos</th>
      </tr></thead>
      <tbody>
        <tr>
          <td><b style="color:{C['azul_med']};font-size:14px">D1</b></td>
          <td><b>Remuneraciones de los Asalariados</b></td>
          <td>Remuneraciones y Gastos del Personal del Estado de Resultados (valor absoluto).
              Proxy del componente D1 del SCN. Incluye sueldos, cotizaciones previsionales y otros gastos laborales.
              <b>Limitación:</b> no disponible la desagregación D11 (sueldos) / D12 (cotizaciones).</td>
          <td>SCN 2025 §8.5<br><span style="color:{C['texto_sub']};font-size:10px">§4.2.3 memoria</span></td>
          <td>Estados de Resultados CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['naranja']};font-size:14px">B2g</b></td>
          <td><b>Excedente Bruto de Explotación</b></td>
          <td>B1g − D1. Lo que queda del VAB después de pagar las remuneraciones.
              En cooperativas equivale al excedente disponible para reservas, remanentes y provisiones.
              Valores negativos indican que las remuneraciones superan el VAB generado.</td>
          <td>SCN 2025 §8.8<br><span style="color:{C['texto_sub']};font-size:10px">§4.3.3 memoria</span></td>
          <td>Calculado (B1g − D1)<br><span style="color:{C['texto_sub']};font-size:10px">ER CMF</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['texto_sub']};font-size:14px">Rem</b></td>
          <td><b>Remanente del Período</b></td>
          <td>Resultado neto del ejercicio de la cooperativa (equivalente a la utilidad/pérdida en empresas).
              Las cooperativas no distribuyen dividendos: el remanente se destina a reservas legales y fondos sociales.
              Valores negativos = cooperativa con déficit en el período.</td>
          <td>Ley 19.832 cooperativas<br><span style="color:{C['texto_sub']};font-size:10px">§4.2.1 memoria</span></td>
          <td>Estados de Resultados CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['texto_sub']};font-size:14px">GAdm</b></td>
          <td><b>Gastos de Administración</b></td>
          <td>Gastos de Administración y Otros del Estado de Resultados (valor absoluto).
              Incluye gastos operacionales no clasificados en remuneraciones ni depreciación.</td>
          <td>SCN 2025 §6.147<br><span style="color:{C['texto_sub']};font-size:10px">(componente de P2)</span></td>
          <td>Estados de Resultados CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
      </tbody>
    </table>
    </div>
  </div>

  <!-- Sección: Cuenta Financiera -->
  <div style="margin-bottom:22px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;padding-bottom:6px;
      border-bottom:2px solid {C['amarillo']}">
      <span style="font-size:18px">🏦</span>
      <span style="font-size:13px;font-weight:800;color:{C['azul_osc']};text-transform:uppercase;
        letter-spacing:.05em">Cuenta Financiera · SCN 2025 Cap. 12 / CIRIEC §4</span>
    </div>
    <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th style="width:80px">Código</th>
        <th style="width:220px">Nombre Completo</th>
        <th>Fórmula / Definición</th>
        <th style="width:150px">Marco Teórico</th>
        <th style="width:160px">Fuente de Datos</th>
      </tr></thead>
      <tbody>
        <tr>
          <td><b style="color:{C['azul_med']};font-size:14px">F2</b></td>
          <td><b>Depósitos (Pasivo)</b></td>
          <td>Total Depósitos del Balance General: suma de Depósitos a la Vista + Depósitos a Plazo.
              Representa el ahorro captado de los socios. Es el principal pasivo de una CAC.</td>
          <td>SCN 2025 §12.36 (AF.2)<br><span style="color:{C['texto_sub']};font-size:10px">CIRIEC §4</span></td>
          <td>Balances Generales CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['verde_med']};font-size:14px">F4</b></td>
          <td><b>Colocaciones Netas (Activo)</b></td>
          <td>Total Colocaciones Netas del Balance General (colocaciones brutas menos provisiones).
              Representa los préstamos vigentes otorgados a los socios. Es el principal activo de una CAC.</td>
          <td>SCN 2025 §12.38 (AF.4)<br><span style="color:{C['texto_sub']};font-size:10px">CIRIEC §4</span></td>
          <td>Balances Generales CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['rojo']};font-size:14px">mora</b></td>
          <td><b>Índice de Mora</b></td>
          <td>Cartera Vencida / Total Colocaciones × 100 (%).
              Mide la calidad de la cartera crediticia. Valores altos indican mayor riesgo de crédito.
              Calculado como promedio simple entre las entidades con datos disponibles.</td>
          <td>Regulación CMF<br><span style="color:{C['texto_sub']};font-size:10px">§4.5 memoria</span></td>
          <td>Balances Generales CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['verde_osc']};font-size:14px">solv</b></td>
          <td><b>Índice de Solvencia</b></td>
          <td>Total Patrimonio / Total Activos × 100 (%).
              Mide la proporción de activos financiados con recursos propios.
              Calculado como promedio simple entre las entidades con datos disponibles.</td>
          <td>Regulación CMF<br><span style="color:{C['texto_sub']};font-size:10px">§4.5 memoria</span></td>
          <td>Balances Generales CMF<br><span style="color:{C['texto_sub']};font-size:10px">consolidado_CAC_balances_resultados_v2.xlsx</span></td>
        </tr>
      </tbody>
    </table>
    </div>
  </div>

  <!-- Sección: Empleo y Variables Estructurales -->
  <div style="margin-bottom:22px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;padding-bottom:6px;
      border-bottom:2px solid {C['morado']}">
      <span style="font-size:18px">👷</span>
      <span style="font-size:13px;font-weight:800;color:{C['azul_osc']};text-transform:uppercase;
        letter-spacing:.05em">Empleo y Variables Estructurales · ONU TSE / CIRIEC / DAES</span>
    </div>
    <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th style="width:80px">Código</th>
        <th style="width:220px">Nombre Completo</th>
        <th>Fórmula / Definición</th>
        <th style="width:150px">Marco Teórico</th>
        <th style="width:160px">Fuente de Datos</th>
      </tr></thead>
      <tbody>
        <tr>
          <td><b style="color:{C['morado']};font-size:14px">PEP</b></td>
          <td><b>Puestos de Empleo Remunerado</b></td>
          <td>Número de empleados reportados por la cooperativa al regulador CMF.
              En el dashboard: columna "Empl.CMF". Equivalente al concepto PEP (Puestos de Empleo Pagado)
              del Manual ONU TSE. Solo disponible para las 7 CAC supervisadas por CMF.</td>
          <td>ONU TSE §4, Tabla 4.5<br><span style="color:{C['texto_sub']};font-size:10px">§4.5.4 memoria</span></td>
          <td>CMF BEST Platform<br><span style="color:{C['texto_sub']};font-size:10px">Informes feb-2026</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['texto_sub']};font-size:14px">Trab.SII</b></td>
          <td><b>Trabajadores SII</b></td>
          <td>Número de trabajadores declarados ante el Servicio de Impuestos Internos.
              Fuente: cruce del directorio DAES con el archivo de empresas SII (timeseries).
              Disponible para las entidades con información en el SII.</td>
          <td>Registro administrativo<br><span style="color:{C['texto_sub']};font-size:10px">§4.2.2 memoria</span></td>
          <td>SII — Empresas timeseries<br><span style="color:{C['texto_sub']};font-size:10px">Consolidado_cooperativas.xlsx</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['texto_sub']};font-size:14px">Of.CMF</b></td>
          <td><b>Oficinas / Sucursales</b></td>
          <td>Número de oficinas o sucursales reportadas por la cooperativa a la CMF.
              Indicador de presencia territorial. Solo disponible para las 7 CAC supervisadas por CMF.</td>
          <td>ONU TSE §4 (presencia)<br><span style="color:{C['texto_sub']};font-size:10px">§4.5.4 memoria</span></td>
          <td>CMF BEST Platform<br><span style="color:{C['texto_sub']};font-size:10px">Informes feb-2026</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['texto_sub']};font-size:14px">Socios / Soc.H / Soc.M</b></td>
          <td><b>Total Socios (y por Género)</b></td>
          <td>Total de socios registrados en la cooperativa, desagregados por género (Hombres/Mujeres).
              Fuente: Directorio DAES (SUBDERE). Indicador de cobertura social del sector.</td>
          <td>ONU TSE §4.3 (membresía)<br><span style="color:{C['texto_sub']};font-size:10px">CIRIEC §4 / §4.1 memoria</span></td>
          <td>DAES — Directorio CAC<br><span style="color:{C['texto_sub']};font-size:10px">Consolidado_cooperativas.xlsx</span></td>
        </tr>
      </tbody>
    </table>
    </div>
  </div>

  <!-- Sección: Parámetros y Método -->
  <div style="margin-bottom:22px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;padding-bottom:6px;
      border-bottom:2px solid {C['naranja']}">
      <span style="font-size:18px">⚙️</span>
      <span style="font-size:13px;font-weight:800;color:{C['azul_osc']};text-transform:uppercase;
        letter-spacing:.05em">Parámetros Metodológicos y Fuentes</span>
    </div>
    <div class="tbl-wrap">
    <table>
      <thead><tr>
        <th style="width:160px">Parámetro / Fuente</th>
        <th>Descripción</th>
        <th style="width:160px">Referencia</th>
      </tr></thead>
      <tbody>
        <tr>
          <td><b style="color:{C['azul_med']}">α = {ALPHA_CI}</b><br><span style="font-size:10px;color:{C['texto_sub']}">alpha CI</span></td>
          <td>Razón Consumo Intermedio / Producción del sector de Intermediación Financiera (código 94)
              de la Matriz Insumo-Producto de Chile 2018, publicada por el Banco Central.
              Se usa para estimar el P2 cuando no se dispone de datos directos de compras.</td>
          <td>MIP Chile 2018<br>Banco Central<br><span style="font-size:10px;color:{C['texto_sub']}">§4.3.2 memoria</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['azul_med']}">Método Indirecto</b><br><span style="font-size:10px;color:{C['texto_sub']}">Tramos SII</span></td>
          <td>Para las CAC sin estados financieros CMF, la producción se estima usando el tramo de ventas
              del SII → punto medio del rango → P1 estimada → B1g = {1-ALPHA_CI:.4f} × P1 estimada.
              Tramos van del 1 (sin ventas) al 13 (ventas &gt; $25.000 MM).</td>
          <td>SII — Empresas timeseries<br><span style="font-size:10px;color:{C['texto_sub']}">§4.3.3 memoria</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['naranja']}">★ CMF supervisadas</b><br><span style="font-size:10px;color:{C['texto_sub']}">Método Directo</span></td>
          <td>7 CAC bajo supervisión CMF: Coopeuch, Coocretal, Oriencoop, Capual, Detacoop,
              Ahorrocoop y Confía. Cuentan con estados financieros auditados y reportados mensualmente.
              Sus datos provienen directamente de los informes del BEST Platform CMF (feb-2026).</td>
          <td>CMF BEST Platform<br><span style="font-size:10px;color:{C['texto_sub']}">§4.5.4 / §4.3.4 memoria</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['azul_med']}">Transf.19862</b><br><span style="font-size:10px;color:{C['texto_sub']}">Ley 19.862</span></td>
          <td>Monto total de transferencias públicas recibidas por la cooperativa registradas en el
              sistema de la Ley 19.862 (Registro de transferencias de fondos públicos).
              Solo 10 de 38 CAC registran transferencias históricas.</td>
          <td>Ley 19.862<br>Ministerio de Hacienda<br><span style="font-size:10px;color:{C['texto_sub']}">§4.4 memoria</span></td>
        </tr>
        <tr>
          <td><b style="color:{C['texto_sub']}">N mínimo = {N_MIN}</b><br><span style="font-size:10px;color:{C['texto_sub']}">Umbral cobertura</span></td>
          <td>Los gráficos de series sectoriales solo incluyen años con al menos {N_MIN} entidades con datos.
              Esto evita que promedios basados en 1-2 entidades distorsionen las tendencias del sector.</td>
          <td>Decisión metodológica<br><span style="font-size:10px;color:{C['texto_sub']}">§4.3.1 memoria</span></td>
        </tr>
      </tbody>
    </table>
    </div>
  </div>

  <!-- Flujo visual -->
  <div style="background:white;border-radius:10px;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,.07)">
    <div style="font-size:12px;font-weight:800;color:{C['azul_osc']};margin-bottom:12px;
      text-transform:uppercase;letter-spacing:.05em">🔗 Flujo de estimación del VAB</div>
    <div style="display:flex;align-items:center;gap:0;flex-wrap:wrap;font-size:11.5px">
      <div style="background:{C['azul_clr']};border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
        <div style="font-weight:800;color:{C['azul_osc']}">P1</div>
        <div style="color:{C['texto_sub']};font-size:10px">Ingresos de<br>Operación</div>
      </div>
      <div style="font-size:18px;color:{C['texto_sub']};padding:0 6px">−</div>
      <div style="background:#FEF9E7;border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
        <div style="font-weight:800;color:#78350F">P2</div>
        <div style="color:{C['texto_sub']};font-size:10px">α×P1<br>({ALPHA_CI}×P1)</div>
      </div>
      <div style="font-size:18px;color:{C['texto_sub']};padding:0 6px">=</div>
      <div style="background:#DCFCE7;border-radius:8px;padding:8px 12px;text-align:center;min-width:110px;border:2px solid {C['verde_med']}">
        <div style="font-weight:800;color:{C['verde_osc']}">B1g</div>
        <div style="color:{C['texto_sub']};font-size:10px">VAB Bruto<br>({1-ALPHA_CI:.4f}×P1)</div>
      </div>
      <div style="font-size:18px;color:{C['texto_sub']};padding:0 6px">−</div>
      <div style="background:{C['azul_clr']};border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
        <div style="font-weight:800;color:{C['azul_osc']}">D1</div>
        <div style="color:{C['texto_sub']};font-size:10px">Remuneraciones<br>(personal)</div>
      </div>
      <div style="font-size:18px;color:{C['texto_sub']};padding:0 6px">=</div>
      <div style="background:#FFF7ED;border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
        <div style="font-weight:800;color:{C['naranja']}">B2g</div>
        <div style="color:{C['texto_sub']};font-size:10px">Excedente Bruto<br>de Explotación</div>
      </div>
      <div style="font-size:18px;color:{C['texto_sub']};padding:0 6px">&nbsp;&nbsp;&nbsp;</div>
      <div style="background:#DCFCE7;border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
        <div style="font-weight:800;color:{C['verde_osc']}">B1g</div>
        <div style="color:{C['texto_sub']};font-size:10px">VAB Bruto</div>
      </div>
      <div style="font-size:18px;color:{C['texto_sub']};padding:0 6px">−</div>
      <div style="background:{C['azul_clr']};border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
        <div style="font-weight:800;color:{C['azul_osc']}">P51d</div>
        <div style="color:{C['texto_sub']};font-size:10px">Depreciación<br>(CKF)</div>
      </div>
      <div style="font-size:18px;color:{C['texto_sub']};padding:0 6px">=</div>
      <div style="background:#EDE9FE;border-radius:8px;padding:8px 12px;text-align:center;min-width:110px">
        <div style="font-weight:800;color:{C['morado']}">B1n</div>
        <div style="color:{C['texto_sub']};font-size:10px">VAB Neto<br>(VAN)</div>
      </div>
    </div>
  </div>

</div>

<!-- ══ TAB: BRECHAS DE DATOS (Cuadro 4.5 memoria) ══════════════════════════ -->
<div class="tab-pane" id="tab-brechas">
  <div class="info-box">
    <b>Cuadro 4.5 · §4.2.3 memoria:</b> Inventario comparado de variables requeridas para la Cuenta Satélite de CAC.
    Contrasta los requerimientos del SCN 2025 / ONU TSE / CIRIEC con la disponibilidad efectiva en Chile,
    España, Portugal y Polonia. Los colores indican el nivel de cobertura de cada variable en cada país.
  </div>

  <!-- Leyenda -->
  <div style="display:flex;gap:16px;flex-wrap:wrap;margin-bottom:14px;align-items:center">
    <span style="font-size:11px;font-weight:700;color:{C['texto_sub']};text-transform:uppercase;letter-spacing:.05em">Nivel de cobertura:</span>
    <span style="display:flex;align-items:center;gap:5px;font-size:11.5px">
      <span style="width:14px;height:14px;border-radius:3px;background:#22C55E;display:inline-block"></span> Disponible directamente
    </span>
    <span style="display:flex;align-items:center;gap:5px;font-size:11.5px">
      <span style="width:14px;height:14px;border-radius:3px;background:#F59E0B;display:inline-block"></span> Disponible como proxy
    </span>
    <span style="display:flex;align-items:center;gap:5px;font-size:11.5px">
      <span style="width:14px;height:14px;border-radius:3px;background:#EF4444;display:inline-block"></span> No disponible / brecha
    </span>
    <span style="display:flex;align-items:center;gap:5px;font-size:11.5px">
      <span style="width:14px;height:14px;border-radius:3px;background:#94A3B8;display:inline-block"></span> Parcial / con limitaciones
    </span>
  </div>

  <!-- Tabla principal Cuadro 4.5 -->
  <div class="tbl-wrap" style="max-height:none;overflow-x:auto">
  <table style="min-width:750px">
    <thead>
      <tr>
        <th style="width:180px">Variable (SCN 2025)</th>
        <th style="width:60px;text-align:center">Código</th>
        <th style="width:50px;text-align:center">Marco</th>
        <th style="width:140px;text-align:center">🇪🇸 España</th>
        <th style="width:140px;text-align:center">🇵🇹 Portugal</th>
        <th style="width:140px;text-align:center">🇵🇱 Polonia</th>
        <th style="width:200px;text-align:center">🇨🇱 Chile (este estudio)</th>
      </tr>
    </thead>
    <tbody>
      <!-- Grupo: Cuenta de Producción -->
      <tr>
        <td colspan="7" style="background:{C['azul_osc']};color:white;font-weight:800;font-size:10.5px;
          text-transform:uppercase;letter-spacing:.06em;padding:7px 9px">
          📊 Cuenta de Producción · SCN 2025 §7
        </td>
      </tr>
      <tr>
        <td><b>Producción total</b></td>
        <td style="text-align:center"><b style="color:{C['azul_med']}">P1</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">SCN</td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Declaración IS<br><span style="font-size:9.5px;color:#166534">cifra de negocios</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ IES + estados<br><span style="font-size:9.5px">financieros</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Encuesta SP<br><span style="font-size:9.5px">+ CIT-8</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ Proxy: tramo ventas SII<br><span style="font-size:9.5px">punto medio del rango</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>Consumo intermedio</b></td>
        <td style="text-align:center"><b style="color:{C['azul_med']}">P2</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">SCN</td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Declaración IS<br><span style="font-size:9.5px">gastos explotación</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ IES + costos<br><span style="font-size:9.5px">operacionales</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Encuesta SP<br><span style="font-size:9.5px">costos</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ MIP BCCh 2018<br><span style="font-size:9.5px">α = {ALPHA_CI} (sector 94)</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>VAB Bruto</b></td>
        <td style="text-align:center"><b style="color:{C['verde_osc']}">B1g</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">SCN</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Calculado</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Calculado</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Calculado</div></td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ Directo (CMF)<br><span style="font-size:9.5px">+ Indirecto (SII) resto</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>Depreciación</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">P51d</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">SCN</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Declaración IS</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ IES</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Encuesta SP</div></td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ Solo CMF (7 CAC)<br><span style="font-size:9.5px">ER auditados</span>
          </div>
        </td>
      </tr>

      <!-- Grupo: Generación del Ingreso -->
      <tr>
        <td colspan="7" style="background:{C['verde_osc']};color:white;font-weight:800;font-size:10.5px;
          text-transform:uppercase;letter-spacing:.06em;padding:7px 9px">
          💼 Cuenta de Generación del Ingreso · SCN 2025 §8
        </td>
      </tr>
      <tr>
        <td><b>Remuneraciones</b></td>
        <td style="text-align:center"><b style="color:{C['azul_med']}">D1</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">SCN</td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ TGSS<br><span style="font-size:9.5px">cotizaciones</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Relatório Único<br><span style="font-size:9.5px">masa salarial</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ ZUS<br><span style="font-size:9.5px">seguridad social</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ Proxy: workers SII<br><span style="font-size:9.5px">sin masa salarial</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>Excedente bruto explot.</b></td>
        <td style="text-align:center"><b style="color:{C['naranja']}">B2g</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">SCN</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Calculado</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Calculado</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Calculado</div></td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ Solo CMF (7 CAC)<br><span style="font-size:9.5px">aproximado para resto</span>
          </div>
        </td>
      </tr>

      <!-- Grupo: Fuentes de ingreso ONU TSE -->
      <tr>
        <td colspan="7" style="background:{C['azul_med']};color:white;font-weight:800;font-size:10.5px;
          text-transform:uppercase;letter-spacing:.06em;padding:7px 9px">
          🌐 Fuentes de Ingreso Desagregadas · ONU TSE Tabla 4.3
        </td>
      </tr>
      <tr>
        <td><b>Transferencias del Estado</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">D3 / D75g</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">ONU TSE</td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Base Nac.<br><span style="font-size:9.5px">Subvenciones</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Dir. Gral.<br><span style="font-size:9.5px">Presupuesto</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Min.<br><span style="font-size:9.5px">Fondos Regionales</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Ley 19.862<br><span style="font-size:9.5px">0,57% faltantes</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>Donaciones privadas</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">D75</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">ONU TSE</td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Declaraciones<br><span style="font-size:9.5px">tributarias</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Encuesta<br><span style="font-size:9.5px">CASES + INE</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Encuesta<br><span style="font-size:9.5px">sectorial</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#E2E8F0;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#475569">
            ≈ Parcial: MDS<br><span style="font-size:9.5px">57,7% sin región</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>Cuotas de socios</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">MDR</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">ONU TSE</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Encuesta directa</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Encuesta CASES</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Encuesta SP</div></td>
        <td style="text-align:center">
          <div style="background:#FEE2E2;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#991B1B">
            ✗ No disponible<br><span style="font-size:9.5px">sin encuesta directa</span>
          </div>
        </td>
      </tr>

      <!-- Grupo: Empleo y Voluntariado ONU TSE -->
      <tr>
        <td colspan="7" style="background:{C['morado']};color:white;font-weight:800;font-size:10.5px;
          text-transform:uppercase;letter-spacing:.06em;padding:7px 9px">
          👷 Empleo y Voluntariado · ONU TSE Tabla 4.5
        </td>
      </tr>
      <tr>
        <td><b>Empleo remunerado</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">PEP</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">ONU TSE</td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ TGSS<br><span style="font-size:9.5px">asalariados</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Relatório Único<br><span style="font-size:9.5px">+ ZUS</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Encuesta SP<br><span style="font-size:9.5px">+ ZUS</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ CMF (7 CAC)<br><span style="font-size:9.5px">+ Proxy workers SII</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>Trabajo voluntario</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">VOV</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">ONU TSE</td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Enc. Voluntariado<br><span style="font-size:9.5px">+ Enc. Uso Tiempo</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Encuesta CASES<br><span style="font-size:9.5px">directa</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">
            ✓ Encuesta PNZ<br><span style="font-size:9.5px">módulo EFL</span>
          </div>
        </td>
        <td style="text-align:center">
          <div style="background:#FEE2E2;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#991B1B">
            ✗ No disponible<br><span style="font-size:9.5px">sin encuesta equivalente</span>
          </div>
        </td>
      </tr>

      <!-- Grupo: Directorio CIRIEC -->
      <tr>
        <td colspan="7" style="background:{C['naranja']};color:white;font-weight:800;font-size:10.5px;
          text-transform:uppercase;letter-spacing:.06em;padding:7px 9px">
          🏛️ Directorio y Estructura Social · CIRIEC §4
        </td>
      </tr>
      <tr>
        <td><b>Universo de entidades</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">—</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">CIRIEC</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ DIRCE</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ CASES</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ KRS</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ DAES<br><span style="font-size:9.5px">39 CAC activas</span></div></td>
      </tr>
      <tr>
        <td><b>Socios (membresía)</b></td>
        <td style="text-align:center"><b style="color:{C['texto_sub']}">—</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">CIRIEC</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Encuesta directa</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Encuesta CASES</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Encuesta SP</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ DAES<br><span style="font-size:9.5px">con género</span></div></td>
      </tr>
      <tr>
        <td><b>Depósitos captados</b></td>
        <td style="text-align:center"><b style="color:{C['azul_med']}">F2</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">CIRIEC</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Balance</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Balance</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Balance</div></td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ Solo CMF (7 CAC)<br><span style="font-size:9.5px">+ hist. 37 entidades</span>
          </div>
        </td>
      </tr>
      <tr>
        <td><b>Colocaciones netas</b></td>
        <td style="text-align:center"><b style="color:{C['verde_med']}">F4</b></td>
        <td style="text-align:center;font-size:10px;color:{C['texto_sub']}">CIRIEC</td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Balance</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Balance</div></td>
        <td style="text-align:center"><div style="background:#DCFCE7;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#166534">✓ Balance</div></td>
        <td style="text-align:center">
          <div style="background:#FEF9C3;border-radius:6px;padding:5px 7px;font-size:10.5px;color:#854D0E">
            ⚡ Solo CMF (7 CAC)<br><span style="font-size:9.5px">+ hist. 37 entidades</span>
          </div>
        </td>
      </tr>
    </tbody>
  </table>
  </div>

  <!-- Resumen de brechas -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:18px">
    <div style="background:white;border-radius:10px;padding:15px 18px;box-shadow:0 2px 8px rgba(0,0,0,.07);
      border-left:4px solid #EF4444">
      <div style="font-size:12px;font-weight:800;color:#991B1B;margin-bottom:8px;text-transform:uppercase;
        letter-spacing:.04em">✗ Brechas críticas Chile (§4.2.3)</div>
      <ul style="font-size:11.5px;color:#374151;line-height:1.8;padding-left:16px">
        <li><b>Cuotas de socios (MDR):</b> sin encuesta directa a organizaciones</li>
        <li><b>Trabajo voluntario (VOV):</b> no existe registro equivalente en Chile</li>
        <li><b>Masa salarial desagregada:</b> SII entrega conteo de trabajadores sin D11/D12</li>
        <li><b>CI sin fuente directa:</b> exención tributaria art. 78 impide declaración de costos</li>
        <li><b>Fecha constitución CAC:</b> 49% faltantes en campo SII</li>
      </ul>
    </div>
    <div style="background:white;border-radius:10px;padding:15px 18px;box-shadow:0 2px 8px rgba(0,0,0,.07);
      border-left:4px solid #22C55E">
      <div style="font-size:12px;font-weight:800;color:#166534;margin-bottom:8px;text-transform:uppercase;
        letter-spacing:.04em">✓ Fortalezas de Chile</div>
      <ul style="font-size:11.5px;color:#374151;line-height:1.8;padding-left:16px">
        <li><b>Universo DAES:</b> registro sólido de 39 CAC activas con RUT</li>
        <li><b>Ley 19.862:</b> transferencias del Estado con alta completitud (0,57% faltantes)</li>
        <li><b>CMF (7 CAC):</b> estados financieros auditados permiten método directo</li>
        <li><b>MIP BCCh 2018:</b> permite estimar α = {ALPHA_CI} para consumo intermedio</li>
        <li><b>Socios con género:</b> DAES entrega desagregación hombre/mujer</li>
      </ul>
    </div>
  </div>

  <div class="warn-box" style="margin-top:14px">
    <b>📌 Nota metodológica (§4.3.4 memoria):</b>
    La estrategia adoptada sigue la práctica internacional estándar: construir una versión inicial con los datos disponibles,
    documentar las brechas explícitamente, y proponer una agenda de mejora para versiones posteriores (Archambault, 2022).
    Las brechas identificadas — cuotas de socios, voluntariado, masa salarial desagregada — constituyen la agenda de datos
    pendiente para una Cuenta Satélite más completa.
  </div>
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

const N_MIN  = {N_MIN};   // umbral mínimo de entidades por año
const SECTOR = {{
  años:    {json.dumps(sector_filt["Año"].astype(str).tolist())},
  // Totales (referencia — sensibles a cobertura)
  P1:      {jl("P1_MM")},
  P2:      {jl("P2_MM")},
  B1g:     {jl("B1g_MM")},
  D1:      {jl("D1_MM")},
  P51d:    {jl("P51d_MM")},
  B2g:     {jl("B2g_MM")},
  B1n:     {jl("B1n_MM")},
  Rem:     {jl("Rem_MM")},
  // Promedios por entidad (comparables entre años)
  P1_avg:   {jl("P1_avg")},
  P2_avg:   {jl("P2_avg")},
  B1g_avg:  {jl("B1g_avg")},
  D1_avg:   {jl("D1_avg")},
  P51d_avg: {jl("P51d_avg")},
  B2g_avg:  {jl("B2g_avg")},
  B1n_avg:  {jl("B1n_avg")},
  Rem_avg:  {jl("Rem_avg")},
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

// SECTOR_COMP: promedios SIN Coopeuch — escala homogénea para descomposición
// Coopeuch (RUT 82878900-7) es ~100x más grande que el resto en 2015/2016
const SECTOR_COMP = {{
  años:     {json.dumps(sector_comp_filt["Año"].astype(str).tolist())},
  D1_avg:   {json.dumps([round(float(v),2) if pd.notna(v) else None for v in sector_comp_filt["D1_avg"]])},
  P51d_avg: {json.dumps([round(float(v),2) if pd.notna(v) else None for v in sector_comp_filt["P51d_avg"]])},
  B2g_avg:  {json.dumps([round(float(v),2) if pd.notna(v) else None for v in sector_comp_filt["B2g_avg"]])},
  B1g_avg:  {json.dumps([round(float(v),2) if pd.notna(v) else None for v in sector_comp_filt["B1g_avg"]])},
  N:        {json.dumps([int(v) for v in sector_comp_filt["N"]])},
  N_cmf:    {json.dumps([int(v) for v in sector_comp_filt["N_cmf"]])},
  N_def:    {json.dumps([int(v) for v in sector_comp_filt["N_def"]])},
  pct_def:  {json.dumps([round(float(v),1) if pd.notna(v) else None for v in sector_comp_filt["pct_def"]])},
  Rem_avg:  {json.dumps([round(float(v),2) if pd.notna(v) else None for v in sector_comp_filt["Rem_avg"]])},
}};

// Solvencia histórica sin 2026 (para eje consistente) + punto CMF 2026 aparte
const SOLV_HIST = {{
  años: {json.dumps(solv_hist["Año"].astype(str).tolist())},
  solv: {json.dumps([round(float(v),1) if pd.notna(v) else None for v in solv_hist["solv"]])},
}};
const SOLV_CMF26 = {round(float(solv_cmf26["solv"].values[0]),1) if len(solv_cmf26) else "null"};

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
  xaxis:{{type:"linear", tickformat:"d"}},
}};

function tit(t){{return {{text:t,font:{{size:13,color:AZ2}}}}}}

// Convertir todos los años a números para eje lineal
SECTOR.años       = SECTOR.años.map(Number);
SECTOR_COMP.años  = SECTOR_COMP.años.map(Number);
SOLV_HIST.años    = SOLV_HIST.años.map(Number);

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
  glosario:   ()=>{{}},
  brechas:    ()=>{{}},
}};

// ══ PRODUCCIÓN Y VAB ════════════════════════════════════════════════════════
function renderProduccion(){{
  // Helper: anotaciones de N encima de cada barra
  function nAnnot(años, N, vals){{
    return años.map((a,i)=>{{
      const v = vals[i];
      return {{
        x:a, y:v!=null?v:0,
        xref:"x", yref:"y",
        text:"n="+N[i],
        showarrow:false,
        font:{{size:9,color:"#6B7280"}},
        yanchor:"bottom",
        yshift:3,
      }};
    }});
  }}

  // B1g promedio por entidad
  Plotly.newPlot("g-b1g-año",[
    bar(SECTOR.años, SECTOR.B1g_avg, "VAB Bruto [B1g] prom./entidad", AZ, {{
      customdata: SECTOR.N,
      hovertemplate:"<b>%{{x}}</b><br>B1g prom: MM$%{{y:.2f}}<br>N entidades: %{{customdata}}<extra></extra>",
      text: SECTOR.N.map(n=>"n="+n), textposition:"outside",
    }}),
  ],{{...LAY,
    title:tit("VAB Bruto Promedio por Entidad [B1g/N] — CAC (MM$ CLP)"),
    yaxis:{{...LAY.yaxis,title:"MM$ / entidad"}},
    annotations:[{{
      x:1,y:1,xref:"paper",yref:"paper",xanchor:"right",yanchor:"top",
      text:"Promedio por entidad · años N≥{N_MIN} incluidos",
      showarrow:false,font:{{size:9.5,color:"#374151"}},
      bgcolor:"rgba(255,255,255,0.85)",borderpad:4
    }}]
  }}, CFG);

  // Descomposición apilada del VAB — promedios SIN Coopeuch (escala homogénea)
  Plotly.newPlot("g-vab-comp",[
    {{type:"bar",x:SECTOR_COMP.años,y:SECTOR_COMP.D1_avg,  name:"Remuneraciones [D1]", marker:{{color:AZ}},
      hovertemplate:"<b>%{{x}}</b><br>D1: MM$%{{y:.2f}}<extra></extra>"}},
    {{type:"bar",x:SECTOR_COMP.años,y:SECTOR_COMP.P51d_avg,name:"Depreciación [P51d]",  marker:{{color:AM}},
      hovertemplate:"<b>%{{x}}</b><br>P51d: MM$%{{y:.2f}}<extra></extra>"}},
    {{type:"bar",x:SECTOR_COMP.años,y:SECTOR_COMP.B2g_avg, name:"Excedente [B2g]",      marker:{{color:GN}},
      hovertemplate:"<b>%{{x}}</b><br>B2g: MM$%{{y:.2f}}<extra></extra>"}},
  ],{{...LAY, barmode:"relative",
     title:tit("Descomposición B1g = D1+P51d+B2g — Prom. por Entidad (MM$, excl. Coopeuch)"),
     yaxis:{{...LAY.yaxis,title:"MM$ / entidad"}},
     annotations:[{{x:1,y:1,xref:"paper",yref:"paper",xanchor:"right",yanchor:"top",
       text:"Excluye Coopeuch (outlier de escala)",
       showarrow:false,font:{{size:9.5,color:"#374151"}},
       bgcolor:"rgba(255,255,255,0.85)",borderpad:4}}]}}, CFG);

  // P1 vs P2 promedios
  const p2neg = SECTOR.P2_avg.map(v=>v!=null?-v:null);
  Plotly.newPlot("g-p1-p2",[
    bar(SECTOR.años,SECTOR.P1_avg, "Producción [P1]",        AZ, {{opacity:0.9}}),
    bar(SECTOR.años,p2neg,          "Cons. Intermedio −[P2]",  RJ, {{opacity:0.9}}),
  ],{{...LAY, barmode:"overlay",
     title:tit("Producción [P1] y Consumo Intermedio [P2] — Prom./Entidad (MM$)"),
     yaxis:{{...LAY.yaxis,title:"MM$ / entidad"}},
     annotations:[{{x:1,y:1,xref:"paper",yref:"paper",xanchor:"right",yanchor:"top",
       text:"α={ALPHA_CI} — MIP Chile 2018",showarrow:false,
       font:{{size:10,color:"#374151"}},bgcolor:"rgba(255,255,255,0.85)",borderpad:4}}]}}, CFG);

  // Cobertura: superavit + deficit — usa SECTOR_COMP (sin Coopeuch, escala homogénea)
  const sin = SECTOR_COMP.N.map((n,i)=>n - SECTOR_COMP.N_def[i]);
  Plotly.newPlot("g-cobertura",[
    bar(SECTOR_COMP.años, sin,                "Con superávit",          GN, {{stackgroup:"s"}}),
    bar(SECTOR_COMP.años, SECTOR_COMP.N_def,  "Con déficit (rem.<0)",    RJ, {{stackgroup:"s"}}),
    {{type:"scatter",mode:"lines+markers",x:SECTOR_COMP.años,y:SECTOR_COMP.N_cmf,
      name:"De ellas: CMF", yaxis:"y2",
      line:{{color:NA,width:2,dash:"dot"}},marker:{{size:6}},
      hovertemplate:"<b>%{{x}}</b><br>CMF: %{{y}}<extra></extra>"}},
  ],{{...LAY, barmode:"stack",
     title:tit("Cobertura del Consolidado (N° entidades — solo años N≥{N_MIN})"),
     yaxis:{{...LAY.yaxis,title:"N° entidades"}},
     yaxis2:{{title:"N° CMF",overlaying:"y",side:"right",showgrid:false}}}}, CFG);
}}

// ══ GENERACIÓN DEL INGRESO ═══════════════════════════════════════════════════
function renderIngreso(){{
  Plotly.newPlot("g-d1-b2g",[
    bar(SECTOR.años, SECTOR.D1_avg,  "Remun. [D1] prom./entidad",         AZ),
    bar(SECTOR.años, SECTOR.B2g_avg, "Excedente bruto [B2g] prom./entidad",GN),
  ],{{...LAY, barmode:"group",
     title:tit("Remuneraciones [D1] y Excedente [B2g] — Prom. por Entidad (MM$)"),
     yaxis:{{...LAY.yaxis,title:"MM$ / entidad"}}}}, CFG);

  // D1_pct y B2g_pct ya son cocientes internos (D1/B1g), no dependen de N
  Plotly.newPlot("g-d1-pct",[
    ln(SECTOR.años, SECTOR.D1_pct,  "% D1 / B1g", AZ),
    ln(SECTOR.años, SECTOR.B2g_pct, "% B2g / B1g",GN),
  ],{{...LAY,
     title:tit("Participación D1 y B2g en el VAB (%) — invariante a N"),
     yaxis:{{...LAY.yaxis,title:"%",range:[0,100]}},
     shapes:[{{type:"line",x0:0,x1:1,xref:"paper",y0:100,y1:100,
       line:{{color:"#9CA3AF",width:1,dash:"dot"}}}}]}}, CFG);

  // Remanente promedio — usa SECTOR_COMP (sin Coopeuch)
  const colors_rem = SECTOR_COMP.Rem_avg.map(v=>v!=null&&v>=0?GN:RJ);
  Plotly.newPlot("g-remanente",[
    {{type:"bar",x:SECTOR_COMP.años,y:SECTOR_COMP.Rem_avg,
      marker:{{color:colors_rem}},
      name:"Remanente prom./entidad",
      customdata:SECTOR_COMP.N,
      hovertemplate:"<b>%{{x}}</b><br>Rem prom: MM$%{{y:.2f}}<br>N: %{{customdata}}<extra></extra>"}},
  ],{{...LAY,
     title:tit("Remanente Promedio por Entidad (MM$, excl. Coopeuch)"),
     yaxis:{{...LAY.yaxis,title:"MM$ / entidad"}}}}, CFG);

  // % con déficit — usa SECTOR_COMP
  Plotly.newPlot("g-deficit",[
    bar(SECTOR_COMP.años, SECTOR_COMP.pct_def, "% entidades con déficit", AM, {{
      text:SECTOR_COMP.pct_def.map(v=>v!=null?v.toFixed(1)+"%":""),textposition:"outside"}}),
  ],{{...LAY,
     title:tit("% Entidades con Remanente Negativo — invariante a N"),
     yaxis:{{...LAY.yaxis,title:"%",range:[0,80]}}}}, CFG);
}}

// ══ CUENTA FINANCIERA ════════════════════════════════════════════════════════
function renderFinanciera(){{
  // F2 y F4 son totales sensibles a N → normalizar por entidad
  const F2_avg = SECTOR.F2.map((v,i)=>v!=null&&SECTOR.N[i]>0?+(v/SECTOR.N[i]).toFixed(2):null);
  const F4_avg = SECTOR.F4.map((v,i)=>v!=null&&SECTOR.N[i]>0?+(v/SECTOR.N[i]).toFixed(2):null);
  const añosNum = SECTOR.años.map(Number);

  Plotly.newPlot("g-f2f4",[
    {{type:"bar", x:añosNum, y:F2_avg, name:"Depósitos captados [F2] prom./entidad",
      marker:{{color:AZ}}, hovertemplate:"<b>%{{x}}</b><br>F2: MM$%{{y:.2f}}<extra></extra>"}},
    {{type:"bar", x:añosNum, y:F4_avg, name:"Colocaciones netas [F4] prom./entidad",
      marker:{{color:GN}}, hovertemplate:"<b>%{{x}}</b><br>F4: MM$%{{y:.2f}}<extra></extra>"}},
  ],{{...LAY, barmode:"group",
     title:tit("Cuenta Financiera: F2 y F4 — Prom. por Entidad (MM$)"),
     xaxis:{{...LAY.xaxis, type:"linear", tickformat:"d"}},
     yaxis:{{...LAY.yaxis,title:"MM$ / entidad"}}}}, CFG);

  // Mora: solo histórico (2026 no tiene datos de mora en CMF)
  // Solvencia: histórico en eje continuo + punto CMF 2026 separado con anotación
  const moraAños = SECTOR.años.map(Number);
  const solvHistAños = SOLV_HIST.años.map(Number);
  Plotly.newPlot("g-mora-solv",[
    {{type:"scatter",mode:"lines+markers",
      x:moraAños, y:SECTOR.mora,
      name:"Mora promedio [CIRIEC]",
      line:{{color:RJ,width:2.5}}, marker:{{size:7,color:RJ}},
      hovertemplate:"<b>%{{x}}</b><br>Mora: %{{y:.2f}}%<extra></extra>"}},
    {{type:"scatter", mode:"lines+markers",
      x: solvHistAños, y: SOLV_HIST.solv,
      name:"Solvencia histórica [CIRIEC]", yaxis:"y2",
      line:{{color:GN,width:2.5,dash:"dash"}}, marker:{{size:7,color:GN}},
      hovertemplate:"<b>%{{x}}</b><br>Solvencia: %{{y:.1f}}%<extra></extra>"}},
    {{type:"scatter", mode:"markers",
      x:[2026], y:[SOLV_CMF26],
      name:"Solvencia CMF 2026", yaxis:"y2",
      marker:{{color:NA, size:12, symbol:"star"}},
      hovertemplate:"<b>2026 (CMF)</b><br>Solvencia: %{{y:.1f}}%<br>⚠ Solo 7 CAC supervisadas<extra></extra>"}},
  ],{{...LAY,
     title:tit("Mora y Solvencia Promedio del Sector (%)"),
     xaxis:{{...LAY.xaxis, type:"linear", tickformat:"d", title:"Año"}},
     yaxis:{{...LAY.yaxis,title:"Mora (%)"}},
     yaxis2:{{title:"Solvencia (%)",overlaying:"y",side:"right",showgrid:false,range:[0,85]}},
     annotations:[{{
       x:2026, y:SOLV_CMF26, xref:"x", yref:"y2",
       text:"⚠ CMF 2026<br>(7 CAC, mayor<br>apalancamiento)",
       showarrow:true, arrowhead:2, ax:45, ay:-40,
       font:{{size:9, color:NA}}, bgcolor:"rgba(255,255,255,0.85)", borderpad:3
     }}]}}, CFG);
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