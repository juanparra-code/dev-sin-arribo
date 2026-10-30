import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Guías Devolución Sin Arribo", page_icon="📦", layout="wide")

st.title("📦 Guías Devolución Sin Arribo")
st.markdown("Sube los archivos y el reporte se genera automáticamente.")

# --- SIDEBAR ---
with st.sidebar:
    st.header("📁 Archivos requeridos")
    bd_file = st.file_uploader("BD Devoluciones (.xlsx)", type=["xlsx"], key="bd")
    st.caption("Se usará la hoja 'Devoluciones'")
    ordenes_files = st.file_uploader("Archivos de Órdenes (.xlsx)", type=["xlsx"],
                                      accept_multiple_files=True, key="ordenes")
    st.caption("Puedes subir uno o varios archivos de órdenes")
    run = st.button("🚀 Generar reporte", type="primary", disabled=(bd_file is None or len(ordenes_files) == 0))

if not run:
    st.info("👈 Sube los archivos en el panel izquierdo y haz clic en **Generar reporte**.")
    st.stop()

# --- PROCESAMIENTO ---
with st.spinner("Cargando BD Devoluciones..."):
    dev = pd.read_excel(bd_file, sheet_name="Devoluciones", engine="calamine")
    def norm(g): return str(g).strip().lstrip("0")
    all_dev = set(norm(g) for g in dev["GUIA DEVOLUCION "].dropna()) | \
              set(norm(g) for g in dev["GUIA INICIAL"].dropna())
    st.success(f"BD Devoluciones: {len(all_dev):,} guías cargadas")

dfs = []
for f in ordenes_files:
    with st.spinner(f"Cargando {f.name}..."):
        df = pd.read_excel(f, engine="calamine")
        mask = df["ESTATUS"].astype(str).str.upper().isin(["DEVOLUCION", "DEVUELTA"])
        dfs.append(df[mask])
        st.success(f"{f.name}: {mask.sum():,} devoluciones")

with st.spinner("Analizando..."):
    ordenes = pd.concat(dfs, ignore_index=True)

    # Omitir guías de la bodega WIILOG BUCARAMANGA
    bodega_norm = ordenes["BODEGA"].astype(str).str.upper().str.split().str.join(" ")
    ordenes = ordenes[bodega_norm != "WIILOG BUCARAMANGA"].copy()

    ordenes["GUIA_NORM"] = ordenes["NÚMERO GUIA"].astype(str).str.strip().str.lstrip("0")
    ordenes["EN_BD"] = ordenes["GUIA_NORM"].isin(all_dev)
    ordenes["RESULTADO"] = ordenes["EN_BD"].map({True: "OK", False: "SIN ARRIBO"})

    es_wiilog_bodega = ordenes["BODEGA"].astype(str).str.upper().str.contains("WIILOG")
    wiilog = ordenes[es_wiilog_bodega].copy()
    otras  = ordenes[~es_wiilog_bodega].copy()

    sin_arribo = wiilog[
        (wiilog["RESULTADO"] == "SIN ARRIBO") &
        (~wiilog["TRANSPORTADORA"].astype(str).str.upper().str.contains("WIILOG"))
    ].copy()

    sin_arribo["FECHA"] = pd.to_datetime(sin_arribo["FECHA"], errors="coerce", dayfirst=True)
    sin_arribo["MES"] = sin_arribo["FECHA"].dt.to_period("M").astype(str)

    total_valor = sin_arribo["VALOR DE COMPRA EN PRODUCTOS"].sum()

# --- MÉTRICAS ---
st.divider()
col1, col2, col3 = st.columns(3)
col1.metric("Total órdenes devolucion", f"{len(ordenes):,}")
col2.metric("Guías SIN ARRIBO", f"{len(sin_arribo):,}", delta=f"-{len(wiilog)-len(sin_arribo):,} OK", delta_color="inverse")
col3.metric("Valor total SIN ARRIBO", f"${total_valor:,.0f}")

st.divider()

# --- TABLAS RESUMEN ---
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("Por Mes")
    resumen_mes = sin_arribo.groupby("MES").agg(
        Guías=("NÚMERO GUIA","count"),
        Valor=("VALOR DE COMPRA EN PRODUCTOS","sum")
    ).reset_index().sort_values("MES")
    resumen_mes["Valor"] = resumen_mes["Valor"].apply(lambda x: f"${x:,.0f}")
    st.dataframe(resumen_mes, use_container_width=True, hide_index=True)

with col_b:
    st.subheader("Por Transportadora")
    resumen_trans = sin_arribo.groupby("TRANSPORTADORA").agg(
        Guías=("NÚMERO GUIA","count"),
        Valor=("VALOR DE COMPRA EN PRODUCTOS","sum")
    ).reset_index().sort_values("Valor", ascending=False)
    resumen_trans["Valor"] = resumen_trans["Valor"].apply(lambda x: f"${x:,.0f}")
    st.dataframe(resumen_trans, use_container_width=True, hide_index=True)

# --- DETALLE ---
st.subheader("Detalle — Guías SIN ARRIBO")
cols_show = ["MES","NÚMERO GUIA","TRANSPORTADORA","BODEGA","NOMBRE CLIENTE",
             "DEPARTAMENTO DESTINO","CIUDAD DESTINO","VALOR DE COMPRA EN PRODUCTOS","FECHA"]
st.dataframe(sin_arribo[[c for c in cols_show if c in sin_arribo.columns]].sort_values(["MES","TRANSPORTADORA"]),
             use_container_width=True, hide_index=True)

# --- EXPORT ---
st.divider()

def generar_excel():
    buf = BytesIO()
    resumen_mes_raw = sin_arribo.groupby("MES").agg(
        GUIAS_SIN_ARRIBO=("NÚMERO GUIA","count"),
        VALOR_TOTAL=("VALOR DE COMPRA EN PRODUCTOS","sum")
    ).reset_index()
    resumen_trans_raw = sin_arribo.groupby("TRANSPORTADORA").agg(
        CANTIDAD=("NÚMERO GUIA","count"),
        VALOR_TOTAL=("VALOR DE COMPRA EN PRODUCTOS","sum")
    ).reset_index().sort_values("VALOR_TOTAL", ascending=False)
    otras_res = otras.groupby("BODEGA").agg(
        CANTIDAD=("NÚMERO GUIA","count"),
        VALOR_TOTAL=("VALOR DE COMPRA EN PRODUCTOS","sum"),
        SIN_ARRIBO=("RESULTADO", lambda x: (x=="SIN ARRIBO").sum())
    ).reset_index().sort_values("CANTIDAD", ascending=False)
    resumen_general = pd.DataFrame({
        "CONCEPTO": ["Total órdenes devolucion/devuelta","Bodegas WIILOG","Otras bodegas (excluidas)",
                     "Guías WIILOG OK","Guías SIN ARRIBO","Valor total SIN ARRIBO ($)"],
        "VALOR": [len(ordenes), len(wiilog), len(otras), len(wiilog[wiilog["RESULTADO"]=="OK"]),
                  len(sin_arribo), f"${total_valor:,.0f}"]
    })
    cols = ["MES","ID","NÚMERO GUIA","ESTATUS","TRANSPORTADORA","BODEGA","NOMBRE CLIENTE",
            "DEPARTAMENTO DESTINO","CIUDAD DESTINO","VALOR DE COMPRA EN PRODUCTOS","FECHA","RESULTADO"]
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        resumen_general.to_excel(w, sheet_name="RESUMEN", index=False)
        resumen_mes_raw.to_excel(w, sheet_name="POR MES", index=False)
        resumen_trans_raw.to_excel(w, sheet_name="POR TRANSPORTADORA", index=False)
        sin_arribo[[c for c in cols if c in sin_arribo.columns]].sort_values(["MES","TRANSPORTADORA"]).to_excel(w, sheet_name="SIN ARRIBO", index=False)
        otras_res.to_excel(w, sheet_name="OTRAS BODEGAS", index=False)
    buf.seek(0)
    return buf

st.download_button(
    label="⬇️ Descargar Excel completo",
    data=generar_excel(),
    file_name="GUIAS DEVOLUCION SIN ARRIBO.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
