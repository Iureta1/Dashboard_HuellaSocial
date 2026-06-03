#!/usr/bin/env python3
"""
Dashboard Huella Social — Cooperativas de Ahorro y Crédito (CAC)
v4.0  |  Proyecto Huella Social — Universidad de los Andes, 2026
Autores: Ignacio Ureta · Antonio Ruiz Tagle
Supervisor: Sebastián Cea

CORRECCIONES v3.2:
  - DAES agg: filtro doble Año+Aporte(%) para excluir fila de totales
  - Feb 2026: excluye fila "TOTAL SEGMENTO CMF"; nombres de entidad desde RUT
  - Match CMF↔coop: usa cmf_nombre (columna directa) en vez de búsqueda fuzzy inversa
  - cmf_colors: clave normalizada a nombre real de entidad del panel CMF
  - Separadores de miles en JSON: usa json.dumps con ensure_ascii=False
  - Seguridad template literal JS: escapa backticks en DATA_JSON
"""

import pandas as pd
import numpy as np
import json
import os
import webbrowser

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_HS   = os.path.join(BASE_DIR, "HuellaSocial_Consolidado.xlsx")
EXCEL_COOP = os.path.join(BASE_DIR, "Consolidado_cooperativas.xlsx")
OUTPUT     = os.path.join(BASE_DIR, "dashboard_huellasocial.html")

# ─── LEER DATOS ───────────────────────────────────────────────────────────────
xls_hs = pd.read_excel(EXCEL_HS, sheet_name=None, header=None)

def parse_sheet(xls, name, hrow):
    df = xls[name].copy()
    df.columns = df.iloc[hrow]
    return df.iloc[hrow+1:].reset_index(drop=True)

NUM = ["P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","B2g (MM$)","Rem (MM$)",
       "Activos (MM$)","Patrimonio (MM$)","F2 (MM$)","F4 (MM$)","PIB_MM","Aporte (%)"]

def numify(df, cols=None):
    for c in (cols or NUM):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

# ── CMF panel ──────────────────────────────────────────────────────────────────
df_cmf = numify(parse_sheet(xls_hs, "📁 Panel CMF", 1))
df_cmf["Año"] = pd.to_numeric(df_cmf["Año"], errors="coerce").astype("Int64")
df_cmf = df_cmf.dropna(subset=["Entidad","Año"]).reset_index(drop=True)

# ── DAES panel ─────────────────────────────────────────────────────────────────
df_daes = numify(parse_sheet(xls_hs, "📁 Panel DAES", 1))
df_daes["Año"] = pd.to_numeric(df_daes["Año"], errors="coerce").astype("Int64")
df_daes = df_daes.dropna(subset=["Entidad","Año"]).reset_index(drop=True)

# ── CMF agg ────────────────────────────────────────────────────────────────────
df_cmf_agg = parse_sheet(xls_hs, "🏦 Agregados CMF", 2)
numify(df_cmf_agg, ["P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","B2g (MM$)","Rem (MM$)","N","PIB Chile (MM$)","Aporte (%)"])
df_cmf_agg["Año"] = pd.to_numeric(df_cmf_agg["Año"], errors="coerce")
# FIX: filtrar por Año notna (excluye filas de totales/cabecera)
df_cmf_agg = df_cmf_agg[df_cmf_agg["Año"].notna()].sort_values("Año").reset_index(drop=True)

# ── DAES agg ───────────────────────────────────────────────────────────────────
df_daes_agg_raw = parse_sheet(xls_hs, "📊 Agregados DAES", 2)
numify(df_daes_agg_raw, ["P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","B2g (MM$)","Rem (MM$)","D1/B1g","N","PIB Chile (MM$)","Aporte (%)"])
df_daes_agg_raw["Año"] = pd.to_numeric(df_daes_agg_raw["Año"], errors="coerce")
# FIX: filtrar por Año Y Aporte(%) notna — excluye fila de totales (Año=NaN, Aporte=0.195)
df_daes_agg = df_daes_agg_raw[
    df_daes_agg_raw["Aporte (%)"].notna() & df_daes_agg_raw["Año"].notna()
].sort_values("Año").reset_index(drop=True)

# ── Agregado Total (CMF + DAES) ───────────────────────────────────────────────
df_total_raw = parse_sheet(xls_hs, "🏦 Agregado Total", 2)
numify(df_total_raw, ["P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","B2g (MM$)","Rem (MM$)","D1/B1g","N","PIB Chile (MM$)","Aporte (%)"])
df_total_raw["Año"] = pd.to_numeric(df_total_raw["Año"], errors="coerce")
df_total_agg = df_total_raw[
    df_total_raw["Año"].notna() & df_total_raw["Aporte (%)"].notna()
].sort_values("Año").reset_index(drop=True)

# ── Feb 2026 ───────────────────────────────────────────────────────────────────
df_feb_raw = parse_sheet(xls_hs, "📅 CMF Feb-2026", 2)
numify(df_feb_raw, ["P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","Rem (MM$)","B1g Anualiz. (MM$)","PIB 2025 (MM$)","Aporte (%)"])
# FIX: excluir fila "TOTAL SEGMENTO CMF" y filas sin RUT
df_feb = df_feb_raw[
    df_feb_raw["Entidad"].notna() &
    df_feb_raw["RUT"].notna() &
    (~df_feb_raw["Entidad"].astype(str).str.upper().str.contains("TOTAL"))
].reset_index(drop=True)

# Mapa RUT (formato Feb-2026) → nombre canónico del Panel CMF
# RUTs verificados desde el Excel CMF Feb-2026
FEB_RUT_TO_ENT = {
    "81836800-3": "AHORROCOOP",
    "84156800-1": "CAPUAL",
    "70015260-K": "Coocretal",
    "82878900-7": "Coopeuch Ltda.",
    "70017860-9": "Detacoop Ltda.",
    "70286300-7": "CONFIA Ltda.",
    "70010920-8": "Oriencoop Ltda.",
}

def feb_nombre(row):
    rut = str(row["RUT"]).strip() if pd.notna(row["RUT"]) else ""
    # Intentar match directo
    if rut in FEB_RUT_TO_ENT:
        return FEB_RUT_TO_ENT[rut]
    # Fallback: nombre truncado del Excel
    return str(row["Entidad"])[:50]

df_feb["Nombre"] = df_feb.apply(feb_nombre, axis=1)

# ── Cooperativas — región ──────────────────────────────────────────────────────
df_coop = pd.read_excel(EXCEL_COOP, sheet_name="Panel Cooperativas")
cac_reg = df_coop[df_coop["Subrubro"] == "Ahorro y Crédito"].copy()
for c in ["Total Socios","sii_trabajadores","cmf_empleados","cmf_oficinas"]:
    cac_reg[c] = pd.to_numeric(cac_reg[c], errors="coerce")

def short_region(r):
    if pd.isna(r) or str(r).strip() in ("0",""):
        return "Sin dato"
    r = str(r).upper()
    MAP = {
        "ARICA Y PARINACOTA": "Arica y Parinacota",
        "ANTOFAGASTA": "Antofagasta",
        "ATACAMA": "Atacama",
        "COQUIMBO": "Coquimbo",
        "VALPARAISO": "Valparaíso",
        "VALPARAÍSO": "Valparaíso",
        "O'HIGGINS": "O'Higgins",
        "OHIGGINS": "O'Higgins",
        "MAULE": "Maule",
        "ÑUBLE": "Ñuble",
        "BÍO-BÍO": "Biobío",
        "BIO-BIO": "Biobío",
        "BIOBIO": "Biobío",
        "ARAUCANÍA": "Araucanía",
        "ARAUCANIA": "Araucanía",
        "LOS RÍOS": "Los Ríos",
        "LOS RIOS": "Los Ríos",
        "LOS LAGOS": "Los Lagos",
        "AYSÉN": "Aysén",
        "AYSEN": "Aysén",
        "MAGALLANES": "Magallanes",
        "METROPOLITANA": "R. Metropolitana",
    }
    for k,v in MAP.items():
        if k in r:
            return v
    return r.replace("REGIÓN DE ","").replace("REGIÓN DEL ","").replace("REGIÓN","").strip().title()

cac_reg = cac_reg.copy()
cac_reg["Región_short"] = cac_reg["Región"].apply(short_region)

# ─── COLORES ──────────────────────────────────────────────────────────────────
# FIX: claves deben coincidir EXACTAMENTE con valores de df_cmf["Entidad"]
cmf_colors = {
    "Coopeuch Ltda.": "#1a56db",
    "Oriencoop Ltda.": "#e74c3c",
    "AHORROCOOP": "#27ae60",
    "CAPUAL": "#f39c12",
    "Coocretal": "#8e44ad",
    "Detacoop Ltda.": "#16a085",
    "CONFIA Ltda.": "#d35400",
}

# ─── SERIALIZE ────────────────────────────────────────────────────────────────
def clean(v):
    if v is None: return None
    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)): return None
    if hasattr(v, "item"): return v.item()
    return v

# ── CMF por entidad ────────────────────────────────────────────────────────────
cmf_entity_data = {}
for ent, grp in df_cmf.groupby("Entidad"):
    grp = grp.sort_values("Año")
    cmf_entity_data[ent] = {
        "años":       [clean(x) for x in grp["Año"].tolist()],
        "P1":         [clean(x) for x in grp["P1 (MM$)"].tolist()],
        "P2":         [clean(x) for x in grp["P2 (MM$)"].tolist()],
        "B1g":        [clean(x) for x in grp["B1g (MM$)"].tolist()],
        "D1":         [clean(x) for x in grp["D1 (MM$)"].tolist()],
        "Rem":        [clean(x) for x in grp["Rem (MM$)"].tolist()],
        "Activos":    [clean(x) for x in grp["Activos (MM$)"].tolist()],
        "Patrimonio": [clean(x) for x in grp["Patrimonio (MM$)"].tolist()],
        "Aporte":     [clean(x) for x in grp["Aporte (%)"].tolist()],
        "color":      cmf_colors.get(ent, "#555"),
    }

# ── CMF agg ────────────────────────────────────────────────────────────────────
cmf_agg_data = {}
for c in ["Año","N","P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","B2g (MM$)","Rem (MM$)","PIB Chile (MM$)","Aporte (%)"]:
    if c in df_cmf_agg.columns:
        cmf_agg_data[c] = [clean(x) for x in df_cmf_agg[c].tolist()]

# ── DAES agg ───────────────────────────────────────────────────────────────────
daes_agg_data = {}
for c in ["Año","N","P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","B2g (MM$)","Rem (MM$)","D1/B1g","PIB Chile (MM$)","Aporte (%)"]:
    if c in df_daes_agg.columns:
        daes_agg_data[c] = [clean(x) for x in df_daes_agg[c].tolist()]

# ── DAES por entidad (top 10 B1g acumulado) ───────────────────────────────────
daes_colors = ["#1a56db","#e74c3c","#27ae60","#f39c12","#8e44ad","#16a085","#d35400","#2980b9","#c0392b","#1abc9c"]
top10 = df_daes.groupby("Entidad")["B1g (MM$)"].sum().nlargest(10).index.tolist()
daes_entity_data = {}
for i, ent in enumerate(top10):
    grp = df_daes[df_daes["Entidad"] == ent].sort_values("Año")
    daes_entity_data[ent] = {
        "años":   [clean(x) for x in grp["Año"].tolist()],
        "B1g":    [clean(x) for x in grp["B1g (MM$)"].tolist()],
        "P1":     [clean(x) for x in grp["P1 (MM$)"].tolist()],
        "Aporte": [clean(x) for x in grp["Aporte (%)"].tolist()],
        "color":  daes_colors[i % len(daes_colors)],
    }

# ── DAES panel completo por año ────────────────────────────────────────────────
daes_panel_records = []
for _, row in df_daes.iterrows():
    daes_panel_records.append({
        "Entidad":    str(row["Entidad"])[:50] if pd.notna(row["Entidad"]) else "–",
        "RUT":        str(row["RUT"]) if pd.notna(row["RUT"]) else "–",
        "Año":        clean(row["Año"]),
        "P1":         clean(row["P1 (MM$)"]),
        "P2":         clean(row["P2 (MM$)"]),
        "B1g":        clean(row["B1g (MM$)"]),
        "D1":         clean(row["D1 (MM$)"]),
        "B2g":        clean(row["B2g (MM$)"]),
        "Rem":        clean(row["Rem (MM$)"]),
        "Activos":    clean(row["Activos (MM$)"]),
        "Patrimonio": clean(row["Patrimonio (MM$)"]),
        "Aporte":     clean(row["Aporte (%)"]),
    })

# (feb2026 eliminado — reemplazado por pestaña Serie Histórica CMF)

# ── Región ─────────────────────────────────────────────────────────────────────
reg_counts = cac_reg["Región_short"].value_counts().reset_index()
reg_counts.columns = ["Región","N_CAC"]
reg_socios = cac_reg.groupby("Región_short")["Total Socios"].sum().reset_index()
reg_socios.columns = ["Región","Total_Socios"]
reg_df = reg_counts.merge(reg_socios, on="Región").sort_values("N_CAC", ascending=False)
reg_json = {
    "regiones": reg_df["Región"].tolist(),
    "n_cac":    [clean(x) for x in reg_df["N_CAC"].tolist()],
    "socios":   [clean(x) for x in reg_df["Total_Socios"].tolist()],
}

# ── Agregado Total (CMF + DAES) ────────────────────────────────────────────────
total_agg_data = {}
for c in ["Año","N","P1 (MM$)","P2 (MM$)","B1g (MM$)","D1 (MM$)","B2g (MM$)","Rem (MM$)","D1/B1g","PIB Chile (MM$)","Aporte (%)"]:
    if c in df_total_agg.columns:
        total_agg_data[c] = [clean(x) for x in df_total_agg[c].tolist()]

# ── CMF table (todos los años, para selector) ──────────────────────────────────
pib_by_year = dict(zip(df_cmf_agg["Año"].tolist(), df_cmf_agg["PIB Chile (MM$)"].tolist()))

# FIX: match por cmf_nombre (columna directa) → Entidad CMF
# Construir mapa cmf_nombre → fila de cac_reg
cmf_nombre_to_row = {}
for _, row in cac_reg[cac_reg["cmf_nombre"].notna()].iterrows():
    cmf_nombre_to_row[str(row["cmf_nombre"]).strip()] = row

# Mapa de entidad CMF → cmf_nombre (invertido de cmf_colors keys)
# Necesitamos saber qué cmf_nombre corresponde a cada entidad del panel CMF
# Construimos un mapa manual usando RUT
ent_to_cmfnombre = {
    "AHORROCOOP":     "Ahorrocoop",
    "CAPUAL":         "Capual",
    "CONFIA Ltda.":   "Coonfía",
    "Coocretal":      "Coocretal",
    "Coopeuch Ltda.": "Coopeuch",
    "Detacoop Ltda.": "Detacoop",
    "Oriencoop Ltda.":"Oriencoop",
}

cmf_table_all = []
for ent, grp in df_cmf.groupby("Entidad"):
    nombre_cmf = ent_to_cmfnombre.get(ent, "")
    row_reg = cmf_nombre_to_row.get(nombre_cmf)
    region    = short_region(row_reg["Región"])         if row_reg is not None else "–"
    empleados = int(row_reg["cmf_empleados"])           if row_reg is not None and pd.notna(row_reg["cmf_empleados"]) else None
    oficinas  = int(row_reg["cmf_oficinas"])            if row_reg is not None and pd.notna(row_reg["cmf_oficinas"]) else None
    socios    = int(row_reg["Total Socios"])            if row_reg is not None and pd.notna(row_reg["Total Socios"]) else None
    for _, r in grp.iterrows():
        yr = clean(r["Año"])
        pib = pib_by_year.get(yr) if yr else None
        cmf_table_all.append({
            "entidad":    ent,
            "año":        yr,
            "region":     region,
            "socios":     socios,
            "empleados":  empleados,
            "oficinas":   oficinas,
            "B1g":        clean(r["B1g (MM$)"]),
            "P1":         clean(r["P1 (MM$)"]),
            "Activos":    clean(r["Activos (MM$)"]),
            "Patrimonio": clean(r["Patrimonio (MM$)"]),
            "pib":        clean(pib),
            "color":      cmf_colors.get(ent, "#555"),
        })

# ─── JSON FINAL ───────────────────────────────────────────────────────────────
DATA_JSON = json.dumps({
    "cmf_entities":   cmf_entity_data,
    "cmf_agg":        cmf_agg_data,
    "daes_agg":       daes_agg_data,
    "total_agg":      total_agg_data,
    "daes_entities":  daes_entity_data,
    "daes_panel":     daes_panel_records,
    "region":         reg_json,
    "cmf_table_all":  cmf_table_all,
}, ensure_ascii=False)

# FIX: escapar backticks para que no rompan el template literal de JS
DATA_JSON_SAFE = DATA_JSON.replace("`", "\\`").replace("</script>", "<\\/script>")

# ─── HTML ─────────────────────────────────────────────────────────────────────
HTML = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Huella Social — CAC Chile</title>
<script src="https://cdn.plot.ly/plotly-2.30.0.min.js"></script>
<style>
  :root {{
    --bg:#0f1117; --surface:#1a1d27; --surface2:#232736;
    --accent:#3b82f6; --accent2:#10b981; --accent3:#f59e0b;
    --text:#e2e8f0; --muted:#94a3b8; --border:#2d3347;
    --danger:#ef4444; --r:12px;
  }}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;}}

  .header{{background:linear-gradient(135deg,#1e2d5a 0%,#0f1117 60%);padding:28px 40px 20px;border-bottom:1px solid var(--border);}}
  .header h1{{font-size:24px;font-weight:700;color:#fff;letter-spacing:-.5px;}}
  .header-sub{{color:var(--muted);font-size:12px;margin-top:4px;}}
  .authors{{margin-top:10px;font-size:12px;color:var(--muted);}}
  .authors span{{color:#a5b4fc;font-weight:600;}}
  .badges{{display:flex;gap:10px;margin-top:12px;flex-wrap:wrap;}}
  .badge{{background:var(--surface2);border:1px solid var(--border);padding:3px 10px;border-radius:20px;font-size:11px;color:var(--muted);}}
  .badge b{{color:var(--accent);}}

  nav{{background:var(--surface);border-bottom:1px solid var(--border);padding:0 40px;display:flex;gap:4px;overflow-x:auto;}}
  .nb{{background:none;border:none;color:var(--muted);cursor:pointer;padding:13px 16px;font-size:13px;font-weight:500;
       border-bottom:2px solid transparent;white-space:nowrap;transition:all .2s;}}
  .nb:hover{{color:var(--text);}}
  .nb.active{{color:var(--accent);border-bottom-color:var(--accent);}}

  main{{padding:24px 40px 60px;max-width:1400px;margin:0 auto;}}
  .section{{display:none;}}
  .section.active{{display:block;}}
  .stitle{{font-size:17px;font-weight:700;margin-bottom:4px;}}
  .ssub{{font-size:12px;color:var(--muted);margin-bottom:20px;}}

  .kgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;margin-bottom:24px;}}
  .kcard{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px;position:relative;overflow:hidden;}}
  .kcard::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;}}
  .kcard.blue::before{{background:var(--accent);}}
  .kcard.green::before{{background:var(--accent2);}}
  .kcard.amber::before{{background:var(--accent3);}}
  .kcard.red::before{{background:var(--danger);}}
  .klabel{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.8px;margin-bottom:6px;}}
  .kvalue{{font-size:26px;font-weight:700;line-height:1;}}
  .ksub{{font-size:11px;color:var(--muted);margin-top:5px;}}

  .cgrid{{display:grid;gap:16px;margin-bottom:16px;}}
  .cgrid.c2{{grid-template-columns:1fr 1fr;}}
  .cgrid.c3{{grid-template-columns:1fr 1fr 1fr;}}
  .ccard{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:20px;}}
  .ctitle{{font-size:13px;font-weight:600;margin-bottom:3px;}}
  .csub{{font-size:11px;color:var(--muted);margin-bottom:12px;}}

  .yr-row{{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px;align-items:center;}}
  .yr-lbl{{font-size:11px;color:var(--muted);margin-right:2px;}}
  .ybtn{{background:var(--surface2);border:1px solid var(--border);color:var(--muted);
         padding:4px 11px;border-radius:16px;cursor:pointer;font-size:12px;font-weight:500;transition:all .2s;}}
  .ybtn:hover{{border-color:var(--accent);color:var(--accent);}}
  .ybtn.active{{background:var(--accent);border-color:var(--accent);color:#fff;}}

  .sbtn{{background:var(--surface2);border:1px solid var(--border);color:var(--muted);
         padding:5px 12px;border-radius:16px;cursor:pointer;font-size:12px;font-weight:500;transition:all .2s;}}
  .sbtn:hover{{border-color:var(--accent);color:var(--accent);}}
  .sbtn.active{{background:var(--accent);border-color:var(--accent);color:#fff;}}

  .tw{{overflow-x:auto;}}
  table{{width:100%;border-collapse:collapse;font-size:13px;}}
  thead tr{{border-bottom:2px solid var(--border);}}
  th{{color:var(--muted);font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:.6px;padding:9px 10px;text-align:left;white-space:nowrap;}}
  td{{padding:10px 10px;border-bottom:1px solid var(--border);vertical-align:middle;}}
  tr:last-child td{{border-bottom:none;}}
  tr:hover td{{background:var(--surface2);}}
  .dot{{width:9px;height:9px;border-radius:50%;display:inline-block;margin-right:7px;flex-shrink:0;}}
  .pill{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;}}

  .sgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(290px,1fr));gap:12px;margin-bottom:24px;}}
  .scard{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px;}}
  .scard h4{{font-size:12px;font-weight:700;color:var(--accent);margin-bottom:7px;}}
  .scard p{{font-size:12px;color:var(--muted);line-height:1.6;}}
  .scard code{{background:var(--surface2);padding:1px 5px;border-radius:3px;font-family:monospace;font-size:11px;color:var(--accent3);}}
  .fbox{{background:var(--surface2);border-left:3px solid var(--accent3);padding:10px 14px;border-radius:0 6px 6px 0;margin-top:9px;font-family:monospace;font-size:12px;color:var(--accent3);}}
  .wbox{{background:rgba(239,68,68,.07);border:1px solid rgba(239,68,68,.2);border-radius:8px;padding:12px 16px;margin-bottom:16px;font-size:12px;color:#fca5a5;}}
  .ibox{{background:rgba(59,130,246,.07);border:1px solid rgba(59,130,246,.2);border-radius:8px;padding:12px 16px;margin-bottom:16px;font-size:12px;color:#93c5fd;}}

  .pb{{background:var(--surface2);border-radius:3px;height:5px;overflow:hidden;margin-top:3px;}}
  .pf{{height:100%;border-radius:3px;}}

  input[type=text]{{background:var(--surface2);border:1px solid var(--border);border-radius:6px;
    color:var(--text);padding:6px 12px;font-size:12px;outline:none;width:220px;}}
  input[type=text]::placeholder{{color:var(--muted);}}
  input[type=text]:focus{{border-color:var(--accent);}}

  .nb.total-tab{{
    background:linear-gradient(135deg,rgba(99,102,241,.18) 0%,rgba(16,185,129,.12) 100%);
    border-bottom:2px solid transparent;
    color:#a5b4fc;font-weight:700;
  }}
  .nb.total-tab.active{{
    color:#a5b4fc;border-bottom-color:#a5b4fc;
    background:linear-gradient(135deg,rgba(99,102,241,.28) 0%,rgba(16,185,129,.20) 100%);
  }}
  .total-hero{{
    background:linear-gradient(135deg,#1e2d5a 0%,#162340 40%,#0f1117 100%);
    border:1px solid rgba(99,102,241,.35);
    border-radius:16px;padding:28px 32px;margin-bottom:24px;
    position:relative;overflow:hidden;
  }}
  .total-hero::before{{
    content:'';position:absolute;top:-60px;right:-60px;width:220px;height:220px;
    background:radial-gradient(circle,rgba(99,102,241,.18) 0%,transparent 70%);
    border-radius:50%;
  }}
  .total-hero h2{{font-size:20px;font-weight:800;color:#c7d2fe;margin-bottom:6px;}}
  .total-hero p{{font-size:13px;color:#94a3b8;max-width:780px;line-height:1.6;}}
  .total-kgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:24px;}}
  .total-kcard{{
    background:linear-gradient(135deg,var(--surface) 0%,var(--surface2) 100%);
    border:1px solid rgba(99,102,241,.25);border-radius:var(--r);padding:20px;
    position:relative;overflow:hidden;
  }}
  .total-kcard::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;
    background:linear-gradient(90deg,#6366f1,#10b981);}}
  .total-kcard .klabel{{font-size:10px;color:#94a3b8;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;}}
  .total-kcard .kvalue{{font-size:28px;font-weight:800;line-height:1;
    background:linear-gradient(135deg,#c7d2fe,#6ee7b7);-webkit-background-clip:text;-webkit-text-fill-color:transparent;}}
  .total-kcard .ksub{{font-size:11px;color:var(--muted);margin-top:6px;}}
  .comp-legend{{display:flex;gap:20px;flex-wrap:wrap;margin-bottom:12px;align-items:center;}}
  .comp-dot{{width:12px;height:12px;border-radius:3px;display:inline-block;margin-right:6px;flex-shrink:0;}}

  @media(max-width:768px){{
    main{{padding:14px;}}
    .cgrid.c2,.cgrid.c3{{grid-template-columns:1fr;}}
    nav{{padding:0 14px;}}
    .header{{padding:18px 14px;}}
  }}
</style>
</head>
<body>

<div class="header">
  <h1>🌱 Huella Social — Cooperativas de Ahorro y Crédito</h1>
  <div class="header-sub">Cuenta satélite del sector CAC chileno · Manual ONU-TSE 2018 · SCN 2025 · CIRIEC 2006</div>
  <div class="authors">Autores: <span>Ignacio Ureta</span> &amp; <span>Antonio Ruiz Tagle</span> · Supervisor: <span>Sebastián Cea</span> · Universidad de los Andes, 2026</div>
  <div class="badges">
    <div class="badge">CMF <b>7 CAC supervisadas</b></div>
    <div class="badge">DAES <b>≤37 CAC</b></div>
    <div class="badge">Período <b>2013–2025</b></div>
    <div class="badge">Ref. PIB <b>BCCh 2018</b></div>
    <div class="badge">α = <b>0,3776</b> (MIP Chile 2018 · Sector 94)</div>
  </div>
</div>

<nav>
  <button class="nb active" onclick="show('overview',this)">📊 Resumen CMF</button>
  <button class="nb" onclick="show('entidades',this)">🏦 Entidades CMF</button>
  <button class="nb" onclick="show('daes',this)">📁 Segmento DAES</button>
  <button class="nb" onclick="show('daes_panel',this)">📋 Panel DAES</button>
  <button class="nb" onclick="show('region',this)">🗺️ Distribución Regional</button>
  <button class="nb" onclick="show('historico',this)">📈 Serie Histórica CMF</button>
  <button class="nb total-tab" onclick="show('total',this)">🌐 Agregado Total</button>
  <button class="nb" onclick="show('comparativo',this)">⚖️ CMF vs DAES</button>
  <button class="nb" onclick="show('supuestos',this)">📌 Supuestos</button>
</nav>

<main>

<!-- OVERVIEW CMF -->
<div class="section active" id="overview">
  <div class="stitle">Segmento CMF — Cuenta de Producción Agregada</div>
  <div class="ssub">7 cooperativas supervisadas · 2013–2025 · Cifras en MM$ corrientes · PIB ref. 2018 (BCCh)</div>
  <div class="yr-row"><span class="yr-lbl">Año para KPIs:</span><span id="ov-yr-btns"></span></div>
  <div class="kgrid" id="kpi-cmf"></div>
  <div class="cgrid c2">
    <div class="ccard"><div class="ctitle">B1g (VAB bruto) y Aporte al PIB — CMF</div>
      <div class="csub">N encima de cada barra = cooperativas con dato ese año</div>
      <div id="ch-cmf-b1g" style="height:270px"></div></div>
    <div class="ccard"><div class="ctitle">Descomposición P1 → P2 → B1g</div>
      <div class="csub">Consumo intermedio (α×P1) y valor agregado bruto</div>
      <div id="ch-cmf-pila" style="height:270px"></div></div>
  </div>
  <div class="cgrid c2" style="margin-top:16px">
    <div class="ccard"><div class="ctitle">D1 vs B2g — Generación del ingreso</div>
      <div class="csub">Remuneraciones vs excedente de explotación bruto</div>
      <div id="ch-cmf-d1" style="height:250px"></div></div>
    <div class="ccard"><div class="ctitle">Composición B1g por entidad (año seleccionado)</div>
      <div class="csub">Participación de cada cooperativa en el VAB sectorial CMF</div>
      <div id="ch-cmf-pie" style="height:250px"></div></div>
  </div>
</div>

<!-- ENTIDADES CMF -->
<div class="section" id="entidades">
  <div class="stitle">Entidades CMF — Detalle por Cooperativa</div>
  <div class="ssub">Serie histórica individual · selecciona cooperativas para comparar</div>
  <div class="yr-row" id="ent-selector"></div>
  <div class="cgrid c2">
    <div class="ccard"><div class="ctitle">B1g por entidad</div><div class="csub">MM$ corrientes</div>
      <div id="ch-ent-b1g" style="height:290px"></div></div>
    <div class="ccard"><div class="ctitle">Aporte al PIB por entidad</div><div class="csub">%</div>
      <div id="ch-ent-aporte" style="height:290px"></div></div>
  </div>
  <div class="cgrid c2" style="margin-top:16px">
    <div class="ccard"><div class="ctitle">Activos totales</div><div class="csub">MM$</div>
      <div id="ch-ent-activos" style="height:270px"></div></div>
    <div class="ccard"><div class="ctitle">Ingresos operacionales (P1)</div><div class="csub">MM$</div>
      <div id="ch-ent-p1" style="height:270px"></div></div>
  </div>
  <div class="ccard" style="margin-top:16px">
    <div class="ctitle">Tabla comparativa CMF</div>
    <div class="yr-row" style="margin-top:8px"><span class="yr-lbl">Año:</span><span id="cmf-tbl-yrs"></span></div>
    <div class="tw" id="cmf-table-wrap"></div>
  </div>
</div>

<!-- DAES AGRUPADO -->
<div class="section" id="daes">
  <div class="stitle">Segmento DAES — Agregados sectoriales</div>
  <div class="ssub">CAC no supervisadas por CMF · N varía por año según EEFF disponibles</div>
  <div class="wbox">⚠️ <b>Panel no balanceado:</b> el conjunto de entidades varía por año. Las comparaciones interanuales deben interpretarse controlando por N. No se agregan segmentos DAES y CMF.</div>
  <div class="yr-row"><span class="yr-lbl">Año para KPIs:</span><span id="daes-yr-btns"></span></div>
  <div class="kgrid" id="kpi-daes"></div>
  <div class="cgrid c2">
    <div class="ccard"><div class="ctitle">B1g sectorial y N de cooperativas</div>
      <div class="csub">Etiquetas "N=X" sobre cada barra</div>
      <div id="ch-daes-b1g" style="height:270px"></div></div>
    <div class="ccard"><div class="ctitle">Top 10 entidades DAES — B1g acumulado</div>
      <div class="csub">Suma del período</div>
      <div id="ch-daes-top" style="height:270px"></div></div>
  </div>
  <div class="cgrid c2" style="margin-top:16px">
    <div class="ccard"><div class="ctitle">Aporte al PIB — DAES</div><div class="csub">%</div>
      <div id="ch-daes-aporte" style="height:250px"></div></div>
    <div class="ccard"><div class="ctitle">Ratio D1/B1g — Intensidad laboral</div>
      <div class="csub">Proporción del VAB destinada a remuneraciones · fuente: Excel</div>
      <div id="ch-daes-d1b1g" style="height:250px"></div></div>
  </div>
</div>

<!-- PANEL DAES COMPLETO -->
<div class="section" id="daes_panel">
  <div class="stitle">Panel DAES — Cooperativas individuales por año</div>
  <div class="ssub">Una fila por cooperativa × año · todos los valores en MM$</div>
  <div class="yr-row">
    <span class="yr-lbl">Año:</span><span id="panel-yr-btns"></span>
    <input type="text" id="panel-search" placeholder="Buscar cooperativa…" oninput="filterPanel()" style="margin-left:12px">
  </div>
  <div class="tw" id="panel-daes-wrap"></div>
</div>

<!-- REGIÓN -->
<div class="section" id="region">
  <div class="stitle">Distribución Geográfica del Sector CAC</div>
  <div class="ssub">39 cooperativas vigentes · DAES + CMF · cruzado con cuenta satélite</div>
  <div class="cgrid c2">
    <div class="ccard"><div class="ctitle">CAC vigentes por región</div>
      <div id="ch-reg-n" style="height:330px"></div></div>
    <div class="ccard"><div class="ctitle">Socios totales por región</div>
      <div id="ch-reg-socios" style="height:330px"></div></div>
  </div>
  <div class="ccard" style="margin-top:16px">
    <div class="ctitle">Grandes CAC supervisadas CMF — Presencia geográfica y cuenta satélite</div>
    <div class="yr-row" style="margin-top:8px"><span class="yr-lbl">Año:</span><span id="reg-tbl-yrs"></span></div>
    <div class="tw" id="region-table-wrap"></div>
  </div>
  <div class="ibox" style="margin-top:14px">💡 La R. Metropolitana concentra la mayor parte de las CAC y del VAB sectorial. Oriencoop (Maule) y Coonfía (Valparaíso) son las únicas grandes CAC con sede fuera de la RM.</div>
</div>

<!-- SERIE HISTÓRICA CMF -->
<div class="section" id="historico">
  <div class="stitle">Serie Histórica CMF — Evolución del VAB por Cooperativa</div>
  <div class="ssub">7 CAC supervisadas · B1g en MM$ corrientes · Panel 2013–2025</div>
  <div class="cgrid c2" style="margin-bottom:16px">
    <div class="ccard" style="grid-column:1/-1">
      <div class="ctitle">B1g (VAB bruto) histórico — Todas las entidades CMF</div>
      <div class="csub">Líneas por cooperativa · evolución completa del período</div>
      <div id="ch-hist-b1g" style="height:340px"></div>
    </div>
  </div>
  <div class="cgrid c2">
    <div class="ccard">
      <div class="ctitle">Participación en el VAB sectorial CMF</div>
      <div class="csub">Share de cada cooperativa sobre el total del segmento CMF ese año</div>
      <div class="yr-row" style="margin-top:8px"><span class="yr-lbl">Año:</span><span id="hist-share-yrs"></span></div>
      <div id="ch-hist-share" style="height:310px"></div>
    </div>
    <div class="ccard">
      <div class="ctitle">Aporte al PIB — Serie histórica por entidad</div>
      <div class="csub">% sobre PIB Chile ref. 2018 · cada cooperativa</div>
      <div id="ch-hist-pib" style="height:310px"></div>
    </div>
  </div>
</div>

<!-- AGREGADO TOTAL CMF + DAES -->
<div class="section" id="total">
  <div class="total-hero">
    <h2>🌐 Cuenta Satélite — Sector CAC Completo (CMF + DAES)</h2>
    <p>Estimación agregada del aporte al PIB de las Cooperativas de Ahorro y Crédito chilenas, combinando el segmento CMF (7 entidades supervisadas) y el segmento DAES (hasta 37 entidades no supervisadas). Período 2013–2025. Cifras en MM$ corrientes · PIB ref. 2018 (BCCh). Metodología ONU-TSE 2018.</p>
  </div>
  <div class="yr-row"><span class="yr-lbl">Año para KPIs:</span><span id="tot-yr-btns"></span></div>
  <div class="total-kgrid" id="kpi-total"></div>
  <div class="cgrid c2">
    <div class="ccard" style="grid-column:1/-1">
      <div class="ctitle">B1g Total (VAB bruto) y Aporte al PIB — CMF + DAES</div>
      <div class="csub">N = cooperativas con dato ese año · Eje derecho: % aporte al PIB</div>
      <div id="ch-tot-b1g" style="height:310px"></div>
    </div>
  </div>
  <div class="cgrid c2" style="margin-top:16px">
    <div class="ccard">
      <div class="ctitle">Descomposición P1 → P2 → B1g</div>
      <div class="csub">Consumo intermedio estimado (α=0,3776) y valor agregado bruto</div>
      <div id="ch-tot-pila" style="height:280px"></div>
    </div>
    <div class="ccard">
      <div class="ctitle">D1 vs B2g — Generación del ingreso</div>
      <div class="csub">Remuneraciones vs excedente de explotación bruto · sector completo</div>
      <div id="ch-tot-d1" style="height:280px"></div>
    </div>
  </div>
  <div class="cgrid c2" style="margin-top:16px">
    <div class="ccard">
      <div class="ctitle">Aporte al PIB — Serie histórica</div>
      <div class="csub">% sobre PIB Chile ref. 2018 · eje derecho: N cooperativas</div>
      <div id="ch-tot-aporte" style="height:260px"></div>
    </div>
    <div class="ccard">
      <div class="ctitle">Intensidad laboral D1/B1g</div>
      <div class="csub">Proporción del VAB destinada a remuneraciones · sector completo</div>
      <div id="ch-tot-d1b1g" style="height:260px"></div>
    </div>
  </div>
  <div class="ccard" style="margin-top:16px">
    <div class="ctitle">Tabla resumen por año</div>
    <div class="tw" id="tot-table-wrap" style="margin-top:12px"></div>
  </div>
  <div class="ibox" style="margin-top:14px">💡 <b>Nota metodológica:</b> El agregado total suma ambos segmentos para obtener una estimación del sector CAC completo. Las CAC CMF contribuyen con la mayor parte del VAB dada su escala. El N varía por año: el panel no es balanceado en el segmento DAES.</div>
</div>

<!-- COMPARATIVO CMF vs DAES -->
<div class="section" id="comparativo">
  <div class="stitle">⚖️ CMF vs DAES — Análisis Comparativo</div>
  <div class="ssub">Comparación de variables clave entre el segmento supervisado CMF y el segmento DAES · cifras en MM$ corrientes</div>
  <div class="wbox">⚠️ <b>Interpretación cuidadosa:</b> ambos segmentos tienen universos y períodos distintos. El segmento CMF tiene N estable (7); el DAES tiene N variable (2–37). Las diferencias de nivel reflejan también diferencias de cobertura.</div>
  <div class="cgrid c2">
    <div class="ccard" style="grid-column:1/-1">
      <div class="ctitle">B1g (VAB bruto) — CMF vs DAES</div>
      <div class="csub">Comparación directa de valores agregados por segmento</div>
      <div class="comp-legend">
        <span><span class="comp-dot" style="background:#3b82f6"></span>Segmento CMF</span>
        <span><span class="comp-dot" style="background:#8b5cf6"></span>Segmento DAES</span>
      </div>
      <div id="ch-comp-b1g" style="height:300px"></div>
    </div>
  </div>
  <div class="cgrid c2" style="margin-top:16px">
    <div class="ccard">
      <div class="ctitle">Aporte al PIB — CMF vs DAES</div>
      <div class="csub">% sobre PIB Chile ref. 2018</div>
      <div id="ch-comp-aporte" style="height:260px"></div>
    </div>
    <div class="ccard">
      <div class="ctitle">P1 Ingresos operacionales — CMF vs DAES</div>
      <div class="csub">Output bruto estimado · MM$</div>
      <div id="ch-comp-p1" style="height:260px"></div>
    </div>
  </div>
  <div class="cgrid c2" style="margin-top:16px">
    <div class="ccard">
      <div class="ctitle">D1 Remuneraciones — CMF vs DAES</div>
      <div class="csub">Generación del ingreso · MM$</div>
      <div id="ch-comp-d1" style="height:260px"></div>
    </div>
    <div class="ccard">
      <div class="ctitle">Intensidad laboral D1/B1g — CMF vs DAES</div>
      <div class="csub">Proporción del VAB destinada a remuneraciones</div>
      <div id="ch-comp-d1b1g" style="height:260px"></div>
    </div>
  </div>
  <div class="ccard" style="margin-top:16px">
    <div class="ctitle">B1g per cápita (por cooperativa) — CMF vs DAES</div>
    <div class="csub">B1g ÷ N · VAB promedio por cooperativa · permite comparación ajustada por tamaño de muestra</div>
    <div id="ch-comp-percap" style="height:260px"></div>
  </div>
</div>

<!-- SUPUESTOS -->
<div class="section" id="supuestos">
  <div class="stitle">Supuestos Metodológicos</div>
  <div class="ssub">Criterios adoptados para la estimación · Proyecto Huella Social 2026</div>
  <div class="sgrid">
    <div class="scard"><h4>P1 — Producción (Output)</h4>
      <p>Se usa <code>Total_Ingresos_Operación</code> como proxy del output bruto (P1), siguiendo el Manual ONU-TSE 2018 para entidades financieras cooperativas.</p>
      <div class="fbox">P1 = Total_Ingresos_Operación</div></div>
    <div class="scard"><h4>P2 — Consumo Intermedio (α = 0,3776)</h4>
      <p>El consumo intermedio <b>no se extrae directamente de los EEFF</b> sino que se estima como <code>α × P1</code>, donde α = 0,3776 es el coeficiente técnico del sector 94 de la MIP Chile 2018. Garantiza <b>comparabilidad entre segmentos</b>.</p>
      <div class="fbox">P2 = α × P1  |  α = 0,3776 (MIP Chile 2018, Sector 94)</div></div>
    <div class="scard"><h4>B1g — VAB bruto (variable principal)</h4>
      <p>Se reporta B1g (VAB <i>bruto</i>) como variable central, siguiendo a Polonia (Statistics Poland 2021) y Portugal (Pedroso et al. 2023). La depreciación (P51d) no se incluye por falta de desglose consistente.</p>
      <div class="fbox">B1g = P1 − P2</div></div>
    <div class="scard"><h4>D1 — Remuneraciones</h4>
      <p>Se usa el valor absoluto de <code>Remuneraciones_y_Gastos_del_Personal</code>. En algunos EEFF aparece como negativo (convención de gasto), por eso se aplica |D1|.</p>
      <div class="fbox">D1 = |Remuneraciones_y_Gastos_del_Personal|</div></div>
    <div class="scard"><h4>No-agregación DAES / CMF</h4>
      <p>Los segmentos <b>no se suman</b>. Las CAC CMF tienen EEFF IFRS con fiscalización intensa y universo estable (7). Las DAES tienen universo variable. Agregar introduciría problemas de comparabilidad y posible doble cobertura en años de transición (2019).</p></div>
    <div class="scard"><h4>Panel no balanceado (DAES)</h4>
      <p>La cobertura varía por año (N=2 en 2014; N=23 en 2023). Las comparaciones interanuales requieren controlar por N. Consistente con Portugal y Polonia, que presentaron paneles no balanceados en sus primeras iteraciones.</p></div>
    <div class="scard"><h4>D2 y D3 = 0</h4>
      <p>Las cooperativas están exentas de impuestos sobre la renta (DFL N°5/2003, art. 78), por lo que D2 (impuestos sobre producción) y D3 (subsidios) se asumen iguales a cero.</p></div>
    <div class="scard"><h4>PIB de referencia</h4>
      <p>Precios corrientes en MM$ con año base 2018, según Cuentas Nacionales del Banco Central (si3.bcentral.cl). El dato 2025 es provisorio. Para el avance 2026 se usa el PIB 2025 como denominador referencial.</p></div>
  </div>
  <div class="ccard">
    <div class="ctitle">Flujo metodológico — De EEFF a Cuenta Satélite</div>
    <div id="ch-sankey" style="height:290px"></div>
  </div>
</div>

</main>

<script>
const D = JSON.parse(`{DATA_JSON_SAFE}`);
const CFG = {{displayModeBar:true,responsive:true,displaylogo:false}};
const DK = {{
  paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',
  font:{{color:'#94a3b8',size:11}},
  xaxis:{{gridcolor:'#2d3347',linecolor:'#2d3347',zerolinecolor:'#2d3347'}},
  yaxis:{{gridcolor:'#2d3347',linecolor:'#2d3347',zerolinecolor:'#2d3347'}},
  legend:{{bgcolor:'rgba(0,0,0,0)',font:{{color:'#e2e8f0',size:11}}}},
  margin:{{t:10,b:40,l:60,r:20}},
}};

function fmm(v){{ return v===null||v===undefined?'—':'$'+Math.round(v).toLocaleString('es-CL')+' MM'; }}
function fpct(v){{ return v===null||v===undefined?'—':v.toFixed(4)+'%'; }}
function fn(v){{ return v===null||v===undefined?'—':Number(v).toLocaleString('es-CL'); }}

// ── NAV ──────────────────────────────────────────────────────────────────────
const renderedSections = {{}};
function show(id, btn) {{
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nb').forEach(b=>b.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
  if(!renderedSections[id]) {{ renderers[id](); renderedSections[id]=true; }}
}}

function buildYrBtns(containerId, years, activeFn, onChange) {{
  const c = document.getElementById(containerId);
  c.innerHTML = years.map(y=>`<button class="ybtn" onclick="(${{onChange}})(this,'${{y}}')" data-yr="${{y}}">${{y}}</button>`).join('');
  const def = c.querySelector('[data-yr="'+activeFn()+'"]');
  if(def) def.classList.add('active');
}}
function setActive(container, val) {{
  document.querySelectorAll('#'+container+' .ybtn').forEach(b=>{{
    b.classList.toggle('active', b.dataset.yr==val);
  }});
}}

// ═══════════════════════════════════════════════════════════════════════════
// 1. OVERVIEW CMF
// ═══════════════════════════════════════════════════════════════════════════
let ovYear = null;
function renderOverview() {{
  const agg = D.cmf_agg;
  const años = agg['Año']||[];
  const b1g  = agg['B1g (MM$)']||[];
  const p1   = agg['P1 (MM$)']||[];
  const p2   = agg['P2 (MM$)']||[];
  const d1   = agg['D1 (MM$)']||[];
  const b2g  = agg['B2g (MM$)']||[];
  const ap   = agg['Aporte (%)']||[];
  const n    = agg['N']||[];

  const validYrs = años.filter((_,i)=>b1g[i]!==null);
  if(!ovYear) ovYear = validYrs[validYrs.length-1];

  buildYrBtns('ov-yr-btns', años, ()=>ovYear, 'updateOvKpis');
  renderOvKpis();

  Plotly.newPlot('ch-cmf-b1g',[
    {{x:años,y:b1g,type:'bar',name:'B1g (MM$)',marker:{{color:'#3b82f6',opacity:.85}},yaxis:'y',
      text:n.map(v=>'N='+v),textposition:'outside',textfont:{{color:'#94a3b8',size:10}}}},
    {{x:años,y:ap,type:'scatter',mode:'lines+markers',name:'Aporte PIB (%)',
      marker:{{color:'#10b981',size:7}},line:{{color:'#10b981',width:2}},yaxis:'y2'}}
  ],{{...DK,yaxis:{{...DK.yaxis,title:'MM$'}},
      yaxis2:{{title:'% PIB',overlaying:'y',side:'right',gridcolor:'transparent',tickformat:'.3f'}},
      legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}},CFG);

  Plotly.newPlot('ch-cmf-pila',[
    {{x:años,y:p2,type:'bar',name:'P2 Cons. Intermedio',marker:{{color:'#ef4444',opacity:.8}}}},
    {{x:años,y:b1g,type:'bar',name:'B1g VAB bruto',marker:{{color:'#3b82f6',opacity:.85}}}},
  ],{{...DK,barmode:'stack',yaxis:{{...DK.yaxis,title:'MM$'}},legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}},CFG);

  Plotly.newPlot('ch-cmf-d1',[
    {{x:años,y:d1,type:'bar',name:'D1 Remuneraciones',marker:{{color:'#f59e0b',opacity:.85}}}},
    {{x:años,y:b2g,type:'bar',name:'B2g Excedente bruto',marker:{{color:'#10b981',opacity:.8}}}},
  ],{{...DK,barmode:'group',yaxis:{{...DK.yaxis,title:'MM$'}},legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}},CFG);

  renderOvPie();
}}
window.updateOvKpis = function(btn, yr) {{
  ovYear = parseInt(yr);
  setActive('ov-yr-btns', yr);
  renderOvKpis();
  renderOvPie();
}};
function renderOvKpis() {{
  const agg=D.cmf_agg, años=agg['Año']||[], b1g=agg['B1g (MM$)']||[];
  const ap=agg['Aporte (%)']||[], p1=agg['P1 (MM$)']||[], d1=agg['D1 (MM$)']||[], n=agg['N']||[];
  const li = años.indexOf(ovYear);
  const kpis=[
    {{l:'B1g — VAB bruto '+ovYear, v:fmm(b1g[li]), s:'Valor Agregado Bruto · CMF', c:'blue'}},
    {{l:'Aporte al PIB '+ovYear,  v:fpct(ap[li]),  s:'PIB ref. 2018 · BCCh',       c:'green'}},
    {{l:'P1 Ingresos Oper. '+ovYear, v:fmm(p1[li]),s:'Output bruto',                c:'amber'}},
    {{l:'D1 Remuneraciones '+ovYear, v:fmm(d1[li]),s:'Generación del ingreso',      c:'red'}},
    {{l:'N Cooperativas '+ovYear,    v:n[li]??'—', s:'Con dato ese año',             c:'blue'}},
    {{l:'B1g acumulado 2013–2025',   v:fmm(b1g.reduce((a,v)=>a+(v||0),0)), s:'Suma del período', c:'green'}},
  ];
  document.getElementById('kpi-cmf').innerHTML = kpis.map(k=>`
    <div class="kcard ${{k.c}}"><div class="klabel">${{k.l}}</div>
    <div class="kvalue">${{k.v}}</div><div class="ksub">${{k.s}}</div></div>`).join('');
}}
function renderOvPie() {{
  const labels=[], vals=[], colors=[];
  for(const [ent,ed] of Object.entries(D.cmf_entities)) {{
    const i = ed.años.indexOf(ovYear);
    if(i>=0 && ed.B1g[i]!==null && ed.B1g[i]>0) {{ labels.push(ent); vals.push(ed.B1g[i]); colors.push(ed.color); }}
  }}
  Plotly.react('ch-cmf-pie',[
    {{labels,values:vals,type:'pie',hole:.4,marker:{{colors}},textfont:{{size:10}},textinfo:'label+percent'}}
  ],{{...DK,margin:{{t:10,b:10,l:10,r:10}}}},CFG);
}}

// ═══════════════════════════════════════════════════════════════════════════
// 2. ENTIDADES CMF
// ═══════════════════════════════════════════════════════════════════════════
let entSel = Object.keys(D.cmf_entities);
let cmfTblYear = null;
function renderEntidades() {{
  const div = document.getElementById('ent-selector');
  div.innerHTML = '<span class="yr-lbl">Cooperativas:</span>' +
    Object.keys(D.cmf_entities).map(e=>
      `<button class="sbtn active" onclick="toggleEnt('${{e}}',this)">${{e}}</button>`).join('') +
    '<button class="sbtn" onclick="selAll()">Todas</button><button class="sbtn" onclick="selNone()">Ninguna</button>';
  renderEntCharts();

  const allYrs = [...new Set(D.cmf_table_all.map(r=>r.año).filter(y=>y!==null))].sort();
  const validYrs = allYrs.filter(y=>D.cmf_table_all.some(r=>r.año===y && r.B1g!==null));
  if(!cmfTblYear) cmfTblYear = validYrs[validYrs.length-1];
  const tblC = document.getElementById('cmf-tbl-yrs');
  tblC.innerHTML = validYrs.map(y=>`<button class="ybtn${{y===cmfTblYear?' active':''}}" onclick="setCmfTblYear(this,${{y}})">${{y}}</button>`).join('');
  renderCmfTable();
}}
function toggleEnt(e,btn){{
  entSel=entSel.includes(e)?entSel.filter(x=>x!==e):[...entSel,e];
  btn.classList.toggle('active');
  renderEntCharts();
}}
function selAll(){{ entSel=Object.keys(D.cmf_entities); document.querySelectorAll('#ent-selector .sbtn').forEach(b=>b.classList.add('active')); renderEntCharts(); }}
function selNone(){{ entSel=[]; document.querySelectorAll('#ent-selector .sbtn').forEach(b=>b.classList.remove('active')); renderEntCharts(); }}
function renderEntCharts(){{
  const tb=[], ta=[], tact=[], tp1=[];
  for(const ent of entSel){{
    const ed=D.cmf_entities[ent]; if(!ed) continue;
    const base={{x:ed.años,mode:'lines+markers',name:ent,line:{{color:ed.color,width:2}},marker:{{size:6,color:ed.color}}}};
    tb.push({{...base,y:ed.B1g}});
    ta.push({{...base,y:ed.Aporte}});
    tact.push({{...base,y:ed.Activos}});
    tp1.push({{...base,y:ed.P1}});
  }}
  const lay=(t)=>{{return{{...DK,yaxis:{{...DK.yaxis,title:t}},legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}}}};
  Plotly.react('ch-ent-b1g',tb.length?tb:[{{x:[],y:[],type:'scatter'}}],lay('MM$'),CFG);
  Plotly.react('ch-ent-aporte',ta.length?ta:[{{x:[],y:[],type:'scatter'}}],{{...lay('%'),yaxis:{{...DK.yaxis,title:'%',tickformat:'.4f'}}}},CFG);
  Plotly.react('ch-ent-activos',tact.length?tact:[{{x:[],y:[],type:'scatter'}}],lay('MM$'),CFG);
  Plotly.react('ch-ent-p1',tp1.length?tp1:[{{x:[],y:[],type:'scatter'}}],lay('MM$'),CFG);
}}
window.setCmfTblYear = function(btn,yr){{
  cmfTblYear=yr;
  document.querySelectorAll('#cmf-tbl-yrs .ybtn').forEach(b=>b.classList.toggle('active',parseInt(b.textContent)===yr));
  renderCmfTable();
}};
function renderCmfTable(){{
  const rows = D.cmf_table_all.filter(r=>r.año===cmfTblYear).sort((a,b)=>(b.B1g||0)-(a.B1g||0));
  if(!rows.length){{ document.getElementById('cmf-table-wrap').innerHTML='<p style="color:#94a3b8;font-size:12px;padding:12px">Sin datos para este año.</p>'; return; }}
  const maxAct = Math.max(...rows.map(r=>r.Activos||0));
  let h=`<table><thead><tr><th>Cooperativa</th><th>Región</th><th>Socios</th><th>Empleados</th>
    <th>Oficinas</th><th>B1g (MM$)</th><th>Aporte PIB</th><th>Activos (MM$)</th></tr></thead><tbody>`;
  for(const r of rows){{
    const pib = r.pib||310681000;
    const ap = r.B1g!==null ? (r.B1g/pib*100).toFixed(4)+'%' : '—';
    const pct = r.Activos&&maxAct?(r.Activos/maxAct*100).toFixed(0):0;
    h+=`<tr><td><span class="dot" style="background:${{r.color}}"></span>${{r.entidad}}</td>
      <td>${{r.region}}</td><td>${{fn(r.socios)}}</td><td>${{fn(r.empleados)}}</td><td>${{fn(r.oficinas)}}</td>
      <td>${{fmm(r.B1g)}}</td><td>${{ap}}</td>
      <td>${{fmm(r.Activos)}}<div class="pb"><div class="pf" style="width:${{pct}}%;background:${{r.color}}"></div></div></td></tr>`;
  }}
  document.getElementById('cmf-table-wrap').innerHTML = h+'</tbody></table>';
}}

// ═══════════════════════════════════════════════════════════════════════════
// 3. DAES AGREGADOS
// ═══════════════════════════════════════════════════════════════════════════
let daesYear = null;
function renderDaes(){{
  const agg=D.daes_agg;
  const años=agg['Año']||[], b1g=agg['B1g (MM$)']||[], p1=agg['P1 (MM$)']||[];
  const d1=agg['D1 (MM$)']||[], ap=agg['Aporte (%)']||[], n=agg['N']||[];
  const d1b1g=agg['D1/B1g']||[];

  const validYrs=años.filter((_,i)=>b1g[i]!==null);
  if(!daesYear) daesYear=validYrs[validYrs.length-1];

  buildYrBtns('daes-yr-btns', años, ()=>daesYear, 'updateDaesKpis');
  renderDaesKpis();

  Plotly.newPlot('ch-daes-b1g',[
    {{x:años,y:b1g,type:'bar',name:'B1g (MM$)',marker:{{color:'#8b5cf6',opacity:.85}},yaxis:'y',
      text:n.map(v=>'N='+v),textposition:'outside',textfont:{{color:'#f59e0b',size:11}}}},
    {{x:años,y:n,type:'scatter',mode:'lines+markers',name:'N cooperativas',
      marker:{{color:'#f59e0b',size:8}},line:{{color:'#f59e0b',width:2,dash:'dot'}},yaxis:'y2'}}
  ],{{...DK,yaxis:{{...DK.yaxis,title:'MM$'}},
      yaxis2:{{title:'N cooperativas',overlaying:'y',side:'right',gridcolor:'transparent',dtick:2}},
      legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}},CFG);

  const topEnts=Object.entries(D.daes_entities);
  const topAcum=topEnts.map(([e,ed])=>[e,ed.B1g.reduce((a,v)=>a+(v||0),0)]).sort((a,b)=>b[1]-a[1]);
  Plotly.newPlot('ch-daes-top',[
    {{x:topAcum.map(([_,v])=>v),y:topAcum.map(([e])=>e),type:'bar',orientation:'h',
      marker:{{color:topAcum.map(([e])=>D.daes_entities[e]?.color||'#555')}},
      text:topAcum.map(([_,v])=>'$'+Math.round(v).toLocaleString('es-CL')),textposition:'outside'}}
  ],{{...DK,margin:{{t:10,b:40,l:240,r:100}},xaxis:{{...DK.xaxis,title:'B1g acumulado MM$'}}}},CFG);

  Plotly.newPlot('ch-daes-aporte',[
    {{x:años,y:ap,type:'scatter',mode:'lines+markers',name:'Aporte PIB (%)',
      fill:'tozeroy',fillcolor:'rgba(139,92,246,.15)',line:{{color:'#8b5cf6',width:2}},marker:{{size:7,color:'#8b5cf6'}}}}
  ],{{...DK,yaxis:{{...DK.yaxis,title:'%',tickformat:'.4f'}}}},CFG);

  Plotly.newPlot('ch-daes-d1b1g',[
    {{x:años,y:d1b1g,type:'bar',name:'D1/B1g',
      marker:{{color:d1b1g.map(v=>v&&v>0.7?'#ef4444':v&&v>0.5?'#f59e0b':'#10b981')}},
      text:d1b1g.map(v=>v!=null?v.toFixed(2):'—'),textposition:'outside'}},
    {{x:[años[0],años[años.length-1]],y:[0.5,0.5],type:'scatter',mode:'lines',name:'Ref 0.5',
      line:{{color:'#94a3b8',dash:'dot',width:1}}}}
  ],{{...DK,yaxis:{{...DK.yaxis,title:'Ratio D1/B1g',tickformat:'.2f'}},
      legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}},CFG);
}}
window.updateDaesKpis = function(btn,yr){{
  daesYear=parseInt(yr);
  setActive('daes-yr-btns',yr);
  renderDaesKpis();
}};
function renderDaesKpis(){{
  const agg=D.daes_agg,años=agg['Año']||[],b1g=agg['B1g (MM$)']||[];
  const ap=agg['Aporte (%)']||[],p1=agg['P1 (MM$)']||[],n=agg['N']||[],d1b1g=agg['D1/B1g']||[];
  const li=años.indexOf(daesYear);
  const kpis=[
    {{l:'B1g '+daesYear,        v:fmm(b1g[li]),     s:'VAB bruto DAES',         c:'blue'}},
    {{l:'Aporte PIB '+daesYear, v:fpct(ap[li]),     s:'PIB ref. 2018',           c:'green'}},
    {{l:'N Cooperativas '+daesYear, v:n[li]??'—',   s:'Con dato ese año',        c:'amber'}},
    {{l:'P1 Ingresos '+daesYear,v:fmm(p1[li]),     s:'Ingresos operacionales',  c:'red'}},
    {{l:'D1/B1g '+daesYear,     v:d1b1g[li]!=null?d1b1g[li].toFixed(2):'—', s:'Intensidad laboral', c:'blue'}},
  ];
  document.getElementById('kpi-daes').innerHTML=kpis.map(k=>`
    <div class="kcard ${{k.c}}"><div class="klabel">${{k.l}}</div>
    <div class="kvalue">${{k.v}}</div><div class="ksub">${{k.s}}</div></div>`).join('');
}}

// ═══════════════════════════════════════════════════════════════════════════
// 4. PANEL DAES COMPLETO
// ═══════════════════════════════════════════════════════════════════════════
let panelYear = null;
let panelSearch = '';
function renderDaesPanel(){{
  const years = [...new Set(D.daes_panel.map(r=>r.Año).filter(y=>y!==null))].sort();
  if(!panelYear) panelYear = years[years.length-1];
  const c=document.getElementById('panel-yr-btns');
  c.innerHTML=years.map(y=>`<button class="ybtn${{y===panelYear?' active':''}}" onclick="setPanelYear(this,${{y}})">${{y}}</button>`).join('');
  renderPanelTable();
}}
window.setPanelYear=function(btn,yr){{
  panelYear=yr;
  document.querySelectorAll('#panel-yr-btns .ybtn').forEach(b=>b.classList.toggle('active',parseInt(b.textContent)===yr));
  renderPanelTable();
}};
window.filterPanel=function(){{
  panelSearch=document.getElementById('panel-search').value.toLowerCase();
  renderPanelTable();
}};
function renderPanelTable(){{
  let rows=D.daes_panel.filter(r=>r.Año===panelYear);
  if(panelSearch) rows=rows.filter(r=>r.Entidad.toLowerCase().includes(panelSearch));
  rows=rows.sort((a,b)=>(b.B1g||0)-(a.B1g||0));
  if(!rows.length){{
    document.getElementById('panel-daes-wrap').innerHTML=`<p style="color:#94a3b8;font-size:12px;padding:12px">Sin datos para ${{panelYear}}.</p>`;
    return;
  }}
  const maxB1g=Math.max(...rows.map(r=>r.B1g||0));
  let h=`<p style="font-size:12px;color:#94a3b8;margin-bottom:10px;">${{rows.length}} cooperativas · Año ${{panelYear}}</p>
  <table><thead><tr>
    <th>#</th><th>Cooperativa</th><th>RUT</th><th>P1 (MM$)</th><th>P2 (MM$)</th>
    <th>B1g (MM$)</th><th>D1 (MM$)</th><th>B2g (MM$)</th><th>Activos (MM$)</th><th>Patrimonio (MM$)</th><th>Aporte (%)</th>
  </tr></thead><tbody>`;
  rows.forEach((r,i)=>{{
    const pct=maxB1g&&r.B1g?(r.B1g||0)/maxB1g*100:0;
    h+=`<tr>
      <td style="color:#94a3b8">${{i+1}}</td>
      <td style="font-weight:500">${{r.Entidad}}</td>
      <td style="font-size:11px;color:#94a3b8">${{r.RUT}}</td>
      <td>${{r.P1!==null?'$'+r.P1.toFixed(1):'—'}}</td>
      <td>${{r.P2!==null?'$'+r.P2.toFixed(1):'—'}}</td>
      <td>
        ${{r.B1g!==null?'$'+r.B1g.toFixed(1):'—'}}
        <div class="pb"><div class="pf" style="width:${{pct.toFixed(0)}}%;background:#8b5cf6"></div></div>
      </td>
      <td>${{r.D1!==null?'$'+r.D1.toFixed(1):'—'}}</td>
      <td>${{r.B2g!==null?'$'+r.B2g.toFixed(1):'—'}}</td>
      <td>${{r.Activos!==null?'$'+r.Activos.toFixed(1):'—'}}</td>
      <td>${{r.Patrimonio!==null?'$'+r.Patrimonio.toFixed(1):'—'}}</td>
      <td>${{r.Aporte!==null?r.Aporte.toFixed(6)+'%':'—'}}</td>
    </tr>`;
  }});
  document.getElementById('panel-daes-wrap').innerHTML=h+'</tbody></table>';
}}

// ═══════════════════════════════════════════════════════════════════════════
// 5. REGIÓN
// ═══════════════════════════════════════════════════════════════════════════
let regTblYear = null;
function renderRegion(){{
  const rd=D.region;
  Plotly.newPlot('ch-reg-n',[
    {{x:rd.n_cac,y:rd.regiones,type:'bar',orientation:'h',
      marker:{{color:rd.regiones.map(r=>r.includes('Metropolitana')?'#3b82f6':'#8b5cf6')}},
      text:rd.n_cac,textposition:'outside'}}
  ],{{...DK,margin:{{t:10,b:40,l:175,r:50}},xaxis:{{...DK.xaxis,title:'N° CAC'}}}},CFG);

  Plotly.newPlot('ch-reg-socios',[
    {{x:rd.socios,y:rd.regiones,type:'bar',orientation:'h',
      marker:{{color:rd.regiones.map(r=>r.includes('Metropolitana')?'#10b981':'#06b6d4')}},
      text:rd.socios.map(v=>v?Number(v).toLocaleString('es-CL'):'0'),textposition:'outside'}}
  ],{{...DK,margin:{{t:10,b:40,l:175,r:100}},xaxis:{{...DK.xaxis,title:'Socios'}}}},CFG);

  const allYrs=[...new Set(D.cmf_table_all.map(r=>r.año).filter(y=>y!==null))].sort();
  const validYrs=allYrs.filter(y=>D.cmf_table_all.some(r=>r.año===y&&r.B1g!==null));
  if(!regTblYear) regTblYear=validYrs[validYrs.length-1];
  const rc=document.getElementById('reg-tbl-yrs');
  rc.innerHTML=validYrs.map(y=>`<button class="ybtn${{y===regTblYear?' active':''}}" onclick="setRegYear(this,${{y}})">${{y}}</button>`).join('');
  renderRegTable();
}}
window.setRegYear=function(btn,yr){{
  regTblYear=yr;
  document.querySelectorAll('#reg-tbl-yrs .ybtn').forEach(b=>b.classList.toggle('active',parseInt(b.textContent)===yr));
  renderRegTable();
}};
function renderRegTable(){{
  const rows=D.cmf_table_all.filter(r=>r.año===regTblYear).sort((a,b)=>(b.B1g||0)-(a.B1g||0));
  let h=`<table><thead><tr><th>Cooperativa</th><th>Región sede</th><th>Socios</th>
    <th>Empleados</th><th>Oficinas</th><th>B1g (MM$)</th><th>Aporte PIB</th></tr></thead><tbody>`;
  for(const r of rows){{
    const pib=r.pib||310681000;
    const ap=r.B1g!==null?(r.B1g/pib*100).toFixed(4)+'%':'—';
    h+=`<tr><td><span class="dot" style="background:${{r.color}}"></span>${{r.entidad}}</td>
      <td>${{r.region}}</td><td>${{fn(r.socios)}}</td><td>${{fn(r.empleados)}}</td><td>${{fn(r.oficinas)}}</td>
      <td>${{fmm(r.B1g)}}</td><td>${{ap}}</td></tr>`;
  }}
  document.getElementById('region-table-wrap').innerHTML=h+'</tbody></table>';
}}

// ═══════════════════════════════════════════════════════════════════════════
// 6. SERIE HISTÓRICA CMF
// ═══════════════════════════════════════════════════════════════════════════
let histShareYear = null;
function renderHistorico() {{
  const ents = D.cmf_entities;
  const entKeys = Object.keys(ents);

  // ── Gráfico 1: líneas históricas B1g por entidad ──────────────────────
  const traces_b1g = entKeys.map(e => {{
    const ed = ents[e];
    return {{
      x: ed.años, y: ed.B1g,
      mode: 'lines+markers', name: e,
      line: {{ color: ed.color, width: 2.2 }},
      marker: {{ size: 6, color: ed.color }},
      connectgaps: false,
    }};
  }});
  Plotly.newPlot('ch-hist-b1g', traces_b1g, {{
    ...DK,
    yaxis: {{ ...DK.yaxis, title: 'B1g (MM$)' }},
    legend: {{ ...DK.legend, orientation: 'h', x: 0, y: 1.14 }},
    margin: {{ t: 30, b: 50, l: 75, r: 20 }},
    shapes: [
      // Línea vertical: inicio supervisión CMF (2019)
      {{ type: 'line', x0: 2019, x1: 2019, y0: 0, y1: 1, yref: 'paper',
         line: {{ color: '#f59e0b', width: 1.2, dash: 'dot' }} }},
    ],
    annotations: [
      {{ x: 2019, y: 1, xref: 'x', yref: 'paper', text: 'Transición<br>CMF', showarrow: false,
         font: {{ size: 10, color: '#f59e0b' }}, xanchor: 'left', yanchor: 'top', bgcolor: 'rgba(0,0,0,0)' }},
    ],
  }}, CFG);

  // ── Gráfico 2: share por entidad (año seleccionable) ──────────────────
  const allYrs = [...new Set(
    entKeys.flatMap(e => ents[e].años.filter((_,i) => ents[e].B1g[i] !== null))
  )].sort();
  if (!histShareYear) histShareYear = allYrs[allYrs.length - 1];
  const sc = document.getElementById('hist-share-yrs');
  sc.innerHTML = allYrs.map(y =>
    `<button class="ybtn${{y === histShareYear ? ' active' : ''}}" onclick="setHistShareYear(this,${{y}})">${{y}}</button>`
  ).join('');
  renderHistShare();

  // ── Gráfico 3: aporte PIB por entidad (líneas) ────────────────────────
  const traces_pib = entKeys.map(e => {{
    const ed = ents[e];
    return {{
      x: ed.años, y: ed.Aporte,
      mode: 'lines+markers', name: e,
      line: {{ color: ed.color, width: 2 }},
      marker: {{ size: 5, color: ed.color }},
      connectgaps: false,
    }};
  }});
  Plotly.newPlot('ch-hist-pib', traces_pib, {{
    ...DK,
    yaxis: {{ ...DK.yaxis, title: '% PIB', tickformat: '.4f' }},
    legend: {{ ...DK.legend, orientation: 'h', x: 0, y: 1.14 }},
    margin: {{ t: 30, b: 50, l: 75, r: 20 }},
  }}, CFG);
}}

window.setHistShareYear = function(btn, yr) {{
  histShareYear = yr;
  document.querySelectorAll('#hist-share-yrs .ybtn').forEach(b =>
    b.classList.toggle('active', parseInt(b.textContent) === yr));
  renderHistShare();
}};

function renderHistShare() {{
  const ents = D.cmf_entities;
  const entKeys = Object.keys(ents);
  // recoger B1g para el año seleccionado
  const vals = entKeys.map(e => {{
    const ed = ents[e];
    const i = ed.años.indexOf(histShareYear);
    return i >= 0 ? (ed.B1g[i] || 0) : 0;
  }});
  const total = vals.reduce((a, v) => a + v, 0);
  const shares = vals.map(v => total > 0 ? (v / total * 100) : 0);
  // ordenar descendente
  const idx = shares.map((_, i) => i).sort((a, b) => shares[b] - shares[a]);
  const labelsOrd = idx.map(i => entKeys[i]);
  const sharesOrd = idx.map(i => shares[i]);
  const colorsOrd = idx.map(i => D.cmf_entities[entKeys[i]].color);

  Plotly.react('ch-hist-share', [{{
    x: sharesOrd,
    y: labelsOrd,
    type: 'bar',
    orientation: 'h',
    marker: {{ color: colorsOrd, opacity: 0.88 }},
    text: sharesOrd.map(v => v.toFixed(1) + '%'),
    textposition: 'outside',
    textfont: {{ color: '#e2e8f0', size: 11 }},
  }}], {{
    ...DK,
    xaxis: {{ ...DK.xaxis, title: 'Share B1g sectorial (%)', ticksuffix: '%' }},
    margin: {{ t: 10, b: 50, l: 160, r: 70 }},
  }}, CFG);
}}

// ═══════════════════════════════════════════════════════════════════════════
// 7. SUPUESTOS
// ═══════════════════════════════════════════════════════════════════════════
function renderSupuestos(){{
  Plotly.newPlot('ch-sankey',[{{type:'sankey',orientation:'h',
    node:{{pad:18,thickness:22,
      label:['EEFF DAES','EEFF CMF (IFRS)','P1 = Ingr. Operac.','P2 = α×P1 (α=0,3776)','B1g = P1−P2','D1 = |Remuner.|','B2g = B1g−D1','% Aporte PIB'],
      color:['#3b82f6','#10b981','#f59e0b','#ef4444','#6366f1','#ec4899','#06b6d4','#8b5cf6']}},
    link:{{source:[0,1,2,2,3,4,4,5],target:[2,2,3,4,4,5,6,7],
      value:[50,200,60,190,60,90,130,130],
      color:['rgba(59,130,246,.3)','rgba(16,185,129,.3)','rgba(239,68,68,.3)','rgba(99,102,241,.3)',
             'rgba(99,102,241,.3)','rgba(236,72,153,.3)','rgba(6,182,212,.3)','rgba(139,92,246,.3)']}}
  }}],{{...DK,margin:{{t:20,b:20,l:20,r:20}}}},CFG);
}}

// ═══════════════════════════════════════════════════════════════════════════
// 8. AGREGADO TOTAL
// ═══════════════════════════════════════════════════════════════════════════
let totYear = null;
function renderTotal() {{
  const agg = D.total_agg;
  const años = agg['Año']||[], b1g = agg['B1g (MM$)']||[], p1 = agg['P1 (MM$)']||[];
  const p2 = agg['P2 (MM$)']||[], d1 = agg['D1 (MM$)']||[], b2g = agg['B2g (MM$)']||[];
  const ap = agg['Aporte (%)']||[], n = agg['N']||[], d1b1g = agg['D1/B1g']||[];

  const validYrs = años.filter((_,i) => b1g[i] !== null);
  if (!totYear) totYear = validYrs[validYrs.length - 1];
  buildYrBtns('tot-yr-btns', años, () => totYear, 'updateTotKpis');
  renderTotKpis();

  // Gráfico principal: B1g barras + aporte línea
  Plotly.newPlot('ch-tot-b1g', [
    {{x:años, y:b1g, type:'bar', name:'B1g Total (MM$)',
      marker:{{color:'#6366f1', opacity:.88}}, yaxis:'y',
      text:n.map(v=>'N='+v), textposition:'outside', textfont:{{color:'#94a3b8',size:10}}}},
    {{x:años, y:ap, type:'scatter', mode:'lines+markers', name:'Aporte PIB (%)',
      marker:{{color:'#10b981', size:8, symbol:'diamond'}}, line:{{color:'#10b981', width:2.5}}, yaxis:'y2'}},
  ], {{...DK,
    yaxis:{{...DK.yaxis, title:'B1g (MM$)'}},
    yaxis2:{{title:'% PIB', overlaying:'y', side:'right', gridcolor:'transparent', tickformat:'.3f'}},
    legend:{{...DK.legend, orientation:'h', x:0, y:1.12}},
    shapes:[{{type:'line',x0:2019,x1:2019,y0:0,y1:1,yref:'paper',line:{{color:'#f59e0b',width:1.2,dash:'dot'}}}}],
    annotations:[{{x:2019,y:0.98,xref:'x',yref:'paper',text:'Inicio<br>CMF',showarrow:false,
      font:{{size:9,color:'#f59e0b'}},xanchor:'left',bgcolor:'rgba(0,0,0,0)'}}],
    margin:{{t:30,b:40,l:70,r:70}},
  }}, CFG);

  // Pila P2 + B1g
  Plotly.newPlot('ch-tot-pila', [
    {{x:años, y:p2, type:'bar', name:'P2 Cons. Intermedio', marker:{{color:'#ef4444', opacity:.8}}}},
    {{x:años, y:b1g, type:'bar', name:'B1g VAB bruto', marker:{{color:'#6366f1', opacity:.88}}}},
  ], {{...DK, barmode:'stack', yaxis:{{...DK.yaxis, title:'MM$'}},
    legend:{{...DK.legend, orientation:'h', x:0, y:1.18}}}}, CFG);

  // D1 vs B2g
  Plotly.newPlot('ch-tot-d1', [
    {{x:años, y:d1, type:'bar', name:'D1 Remuneraciones', marker:{{color:'#f59e0b', opacity:.88}}}},
    {{x:años, y:b2g, type:'bar', name:'B2g Excedente bruto', marker:{{color:'#10b981', opacity:.8}}}},
  ], {{...DK, barmode:'group', yaxis:{{...DK.yaxis, title:'MM$'}},
    legend:{{...DK.legend, orientation:'h', x:0, y:1.18}}}}, CFG);

  // Aporte PIB con N eje derecho
  Plotly.newPlot('ch-tot-aporte', [
    {{x:años, y:ap, type:'scatter', mode:'lines+markers', name:'Aporte PIB (%)',
      fill:'tozeroy', fillcolor:'rgba(99,102,241,.12)', line:{{color:'#6366f1', width:2.5}},
      marker:{{size:8, color:'#6366f1', symbol:'diamond'}}, yaxis:'y'}},
    {{x:años, y:n, type:'scatter', mode:'lines+markers', name:'N cooperativas',
      line:{{color:'#f59e0b', width:1.5, dash:'dot'}}, marker:{{size:6, color:'#f59e0b'}}, yaxis:'y2'}},
  ], {{...DK,
    yaxis:{{...DK.yaxis, title:'% Aporte PIB', tickformat:'.4f'}},
    yaxis2:{{title:'N', overlaying:'y', side:'right', gridcolor:'transparent', dtick:5}},
    legend:{{...DK.legend, orientation:'h', x:0, y:1.18}},
    margin:{{t:30,b:40,l:70,r:60}},
  }}, CFG);

  // D1/B1g ratio
  Plotly.newPlot('ch-tot-d1b1g', [
    {{x:años, y:d1b1g, type:'bar', name:'D1/B1g',
      marker:{{color:d1b1g.map(v=>v&&v>0.7?'#ef4444':v&&v>0.5?'#f59e0b':'#10b981')}},
      text:d1b1g.map(v=>v!=null?v.toFixed(3):'—'), textposition:'outside'}},
    {{x:[años[0],años[años.length-1]], y:[0.5,0.5], type:'scatter', mode:'lines',
      name:'Ref. 0.5', line:{{color:'#94a3b8', dash:'dot', width:1}}}},
  ], {{...DK, yaxis:{{...DK.yaxis, title:'Ratio D1/B1g', tickformat:'.2f'}},
    legend:{{...DK.legend, orientation:'h', x:0, y:1.18}}}}, CFG);

  // Tabla resumen
  let h = `<table><thead><tr>
    <th>Año</th><th>N</th><th>P1 (MM$)</th><th>P2 (MM$)</th><th>B1g (MM$)</th>
    <th>D1 (MM$)</th><th>B2g (MM$)</th><th>D1/B1g</th><th>PIB Chile (MM$)</th><th>Aporte (%)</th>
  </tr></thead><tbody>`;
  for (let i = 0; i < años.length; i++) {{
    const isSelected = años[i] === totYear;
    h += `<tr style="${{isSelected?'background:rgba(99,102,241,.12);':''}}">`+
      `<td style="font-weight:${{isSelected?700:400}};color:${{isSelected?'#c7d2fe':'inherit'}}">${{años[i]}}</td>`+
      `<td>${{n[i]??'—'}}</td>`+
      `<td>${{agg['P1 (MM$)'][i]!==null?'$'+Math.round(agg['P1 (MM$)'][i]).toLocaleString('es-CL'):'—'}}</td>`+
      `<td>${{agg['P2 (MM$)'][i]!==null?'$'+Math.round(agg['P2 (MM$)'][i]).toLocaleString('es-CL'):'—'}}</td>`+
      `<td style="font-weight:600">${{b1g[i]!==null?'$'+Math.round(b1g[i]).toLocaleString('es-CL'):'—'}}</td>`+
      `<td>${{d1[i]!==null?'$'+Math.round(d1[i]).toLocaleString('es-CL'):'—'}}</td>`+
      `<td>${{b2g[i]!==null?'$'+Math.round(b2g[i]).toLocaleString('es-CL'):'—'}}</td>`+
      `<td>${{d1b1g[i]!=null?d1b1g[i].toFixed(4):'—'}}</td>`+
      `<td>${{agg['PIB Chile (MM$)'][i]!==null?'$'+Math.round(agg['PIB Chile (MM$)'][i]).toLocaleString('es-CL'):'—'}}</td>`+
      `<td style="font-weight:600;color:#10b981">${{ap[i]!=null?ap[i].toFixed(6)+'%':'—'}}</td>`+
      `</tr>`;
  }}
  document.getElementById('tot-table-wrap').innerHTML = h + '</tbody></table>';
}}
window.updateTotKpis = function(btn, yr) {{
  totYear = parseInt(yr);
  setActive('tot-yr-btns', yr);
  renderTotKpis();
}};
function renderTotKpis() {{
  const agg=D.total_agg, años=agg['Año']||[], b1g=agg['B1g (MM$)']||[];
  const ap=agg['Aporte (%)']||[], p1=agg['P1 (MM$)']||[], d1=agg['D1 (MM$)']||[], n=agg['N']||[];
  const li = años.indexOf(totYear);
  const b1gAcum = b1g.reduce((a,v)=>a+(v||0),0);
  const kpis = [
    {{l:'B1g Total '+totYear,    v:fmm(b1g[li]),           s:'VAB bruto CMF + DAES',    c:'indigo'}},
    {{l:'Aporte al PIB '+totYear,v:fpct(ap[li]),            s:'PIB ref. 2018 · BCCh',    c:'green'}},
    {{l:'P1 Ingresos '+totYear,  v:fmm(p1[li]),             s:'Output bruto total',      c:'amber'}},
    {{l:'D1 Remuner. '+totYear,  v:fmm(d1[li]),             s:'Masa salarial sectorial', c:'red'}},
    {{l:'N Cooperativas '+totYear,v:n[li]??'—',             s:'CMF + DAES con dato',     c:'blue'}},
    {{l:'B1g acumul. 2013–2025', v:fmm(b1gAcum),           s:'Suma del período completo',c:'green'}},
  ];
  document.getElementById('kpi-total').innerHTML = kpis.map(k=>`
    <div class="total-kcard">
      <div class="klabel">${{k.l}}</div>
      <div class="kvalue">${{k.v}}</div>
      <div class="ksub">${{k.s}}</div>
    </div>`).join('');
}}

// ═══════════════════════════════════════════════════════════════════════════
// 9. COMPARATIVO CMF vs DAES
// ═══════════════════════════════════════════════════════════════════════════
function renderComparativo() {{
  const cmf  = D.cmf_agg;
  const daes = D.daes_agg;
  const cmfAños  = cmf['Año']||[];
  const daesAños = daes['Año']||[];

  // Helper: alinear dos series por año
  function seriesFor(agg, col) {{
    return (agg['Año']||[]).map((_,i) => agg[col][i]);
  }}

  const cmfB1g   = seriesFor(cmf,  'B1g (MM$)');
  const daesB1g  = seriesFor(daes, 'B1g (MM$)');
  const cmfAp    = seriesFor(cmf,  'Aporte (%)');
  const daesAp   = seriesFor(daes, 'Aporte (%)');
  const cmfP1    = seriesFor(cmf,  'P1 (MM$)');
  const daesP1   = seriesFor(daes, 'P1 (MM$)');
  const cmfD1    = seriesFor(cmf,  'D1 (MM$)');
  const daesD1   = seriesFor(daes, 'D1 (MM$)');
  const cmfRatio = seriesFor(cmf,  'Aporte (%)').map((_,i)=>
    cmf['D1 (MM$)'][i]!=null && cmf['B1g (MM$)'][i]!=null && cmf['B1g (MM$)'][i]>0
    ? cmf['D1 (MM$)'][i]/cmf['B1g (MM$)'][i] : null);
  const daesRatio = seriesFor(daes,'Aporte (%)').map((_,i)=>
    daes['D1 (MM$)'][i]!=null && daes['B1g (MM$)'][i]!=null && daes['B1g (MM$)'][i]>0
    ? daes['D1 (MM$)'][i]/daes['B1g (MM$)'][i] : null);

  const cmfN  = cmf['N']||[];
  const daesN = daes['N']||[];
  const cmfPerCap  = cmfB1g.map((v,i)  => v!=null && cmfN[i]  ? v/cmfN[i]  : null);
  const daesPerCap = daesB1g.map((v,i) => v!=null && daesN[i] ? v/daesN[i] : null);

  const baseCmf  = {{mode:'lines+markers', line:{{color:'#3b82f6',width:2.2}}, marker:{{size:7,color:'#3b82f6'}}}};
  const baseD    = {{mode:'lines+markers', line:{{color:'#8b5cf6',width:2.2}}, marker:{{size:7,color:'#8b5cf6'}}}};

  // B1g comparativo (barras agrupadas)
  Plotly.newPlot('ch-comp-b1g', [
    {{x:cmfAños,  y:cmfB1g,  type:'bar', name:'CMF',  marker:{{color:'#3b82f6',opacity:.85}}}},
    {{x:daesAños, y:daesB1g, type:'bar', name:'DAES', marker:{{color:'#8b5cf6',opacity:.85}}}},
  ], {{...DK, barmode:'group', yaxis:{{...DK.yaxis,title:'B1g (MM$)'}},
    legend:{{...DK.legend,orientation:'h',x:0,y:1.1}},margin:{{t:30,b:40,l:70,r:20}}}}, CFG);

  // Aporte PIB (líneas)
  Plotly.newPlot('ch-comp-aporte', [
    {{x:cmfAños,  y:cmfAp,  ...baseCmf, name:'CMF',  fill:'tozeroy', fillcolor:'rgba(59,130,246,.10)'}},
    {{x:daesAños, y:daesAp, ...baseD,   name:'DAES', fill:'tozeroy', fillcolor:'rgba(139,92,246,.10)'}},
  ], {{...DK, yaxis:{{...DK.yaxis,title:'% PIB',tickformat:'.4f'}},
    legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}}, CFG);

  // P1
  Plotly.newPlot('ch-comp-p1', [
    {{x:cmfAños,  y:cmfP1,  ...baseCmf, name:'CMF'}},
    {{x:daesAños, y:daesP1, ...baseD,   name:'DAES'}},
  ], {{...DK, yaxis:{{...DK.yaxis,title:'P1 (MM$)'}},
    legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}}, CFG);

  // D1
  Plotly.newPlot('ch-comp-d1', [
    {{x:cmfAños,  y:cmfD1,  type:'bar', name:'CMF',  marker:{{color:'#3b82f6',opacity:.85}}}},
    {{x:daesAños, y:daesD1, type:'bar', name:'DAES', marker:{{color:'#8b5cf6',opacity:.85}}}},
  ], {{...DK, barmode:'group', yaxis:{{...DK.yaxis,title:'D1 (MM$)'}},
    legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}}, CFG);

  // D1/B1g ratio
  const daesD1B1g = (daes['D1/B1g']||[]);
  Plotly.newPlot('ch-comp-d1b1g', [
    {{x:cmfAños,  y:cmfRatio,  ...baseCmf, name:'CMF'}},
    {{x:daesAños, y:daesD1B1g, ...baseD,   name:'DAES'}},
    {{x:[cmfAños[0],cmfAños[cmfAños.length-1]], y:[0.5,0.5], type:'scatter',
      mode:'lines', name:'Ref. 0.5', line:{{color:'#94a3b8',dash:'dot',width:1}}}},
  ], {{...DK, yaxis:{{...DK.yaxis,title:'D1/B1g',tickformat:'.3f'}},
    legend:{{...DK.legend,orientation:'h',x:0,y:1.18}}}}, CFG);

  // Per cápita
  Plotly.newPlot('ch-comp-percap', [
    {{x:cmfAños,  y:cmfPerCap,  ...baseCmf, name:'CMF (B1g÷N)'}},
    {{x:daesAños, y:daesPerCap, ...baseD,   name:'DAES (B1g÷N)'}},
  ], {{...DK, yaxis:{{...DK.yaxis,title:'MM$ por cooperativa'}},
    legend:{{...DK.legend,orientation:'h',x:0,y:1.18}},margin:{{t:30,b:40,l:80,r:20}}}}, CFG);
}}

// ─── INIT ─────────────────────────────────────────────────────────────────────
const renderers = {{
  overview:    renderOverview,
  entidades:   renderEntidades,
  daes:        renderDaes,
  daes_panel:  renderDaesPanel,
  region:      renderRegion,
  historico:   renderHistorico,
  total:       renderTotal,
  comparativo: renderComparativo,
  supuestos:   renderSupuestos,
}};
renderOverview();
renderedSections['overview'] = true;
</script>
</body>
</html>"""

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write(HTML)
print(f"✅ Dashboard v3.2 generado: {OUTPUT}")
import webbrowser
webbrowser.open(OUTPUT)