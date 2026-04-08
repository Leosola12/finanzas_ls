import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Finanzas Personales V9", layout="wide")

# =============================
# Helpers
# =============================

def normalize_colname(c):
    c = c.strip().lower()
    c = c.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
    return c


def find_column(df, options):
    cols = {normalize_colname(c): c for c in df.columns}
    for opt in options:
        opt_norm = normalize_colname(opt)
        for c_norm, original in cols.items():
            if opt_norm in c_norm:
                return original
    return None


def normalize_series(s):
    return (
        s.astype(str)
        .str.strip()
        .str.lower()
        .str.replace("á","a")
        .str.replace("é","e")
        .str.replace("í","i")
        .str.replace("ó","o")
        .str.replace("ú","u")
    )


def format_ars(x):
    return f"${x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# =============================
# UI
# =============================

st.title("📊 Finanzas Personales V9")

file = st.file_uploader("Subí tu Excel", type=["xlsx"])
mapping_file = st.sidebar.file_uploader("Subir mapeo_gastos.csv", type=["csv"])
mapping_ing_file = st.sidebar.file_uploader("Subir mapeo_ingresos.csv", type=["csv"])

if file:
    xls = pd.ExcelFile(file)
    gastos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    ingresos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

    # COLUMNAS
    subcat_col = find_column(gastos_df, ["subcategoria", "subcategoría"])
    monto_col = find_column(gastos_df, ["monto", "importe"])
    fecha_col = find_column(gastos_df, ["fecha"])

    ingresos_monto_col = find_column(ingresos_df, ["monto","importe"])
    ingresos_fecha_col = find_column(ingresos_df, ["fecha"])
    tipo_ing_col = find_column(ingresos_df, ["tipo de ingreso", "tipo ingreso"])

    # FECHAS
    gastos_df["Fecha"] = pd.to_datetime(gastos_df[fecha_col], errors="coerce")
    ingresos_df["Fecha"] = pd.to_datetime(ingresos_df[ingresos_fecha_col], errors="coerce")

    gastos_df["Mes"] = gastos_df["Fecha"].dt.to_period("M").astype(str)
    ingresos_df["Mes"] = ingresos_df["Fecha"].dt.to_period("M").astype(str)

    meses_disponibles = sorted(gastos_df["Mes"].dropna().unique())

    # =============================
    # FILTROS
    # =============================

    st.sidebar.header("📅 Filtros")

    meses_sel = st.sidebar.multiselect("Seleccionar meses", meses_disponibles, default=meses_disponibles)
    modo = st.sidebar.radio("Modo", ["Total", "Promedio mensual"])

    gastos_df = gastos_df[gastos_df["Mes"].isin(meses_sel)]
    ingresos_df = ingresos_df[ingresos_df["Mes"].isin(meses_sel)]

    n_meses = max(len(meses_sel), 1)

    # =============================
    # MAPEO (REVIVE)
    # =============================

    if mapping_file:
        map_gastos = pd.read_csv(mapping_file)
        map_gastos["subcat_norm"] = normalize_series(map_gastos["Subcategoria"])
        gastos_df["subcat_norm"] = normalize_series(gastos_df[subcat_col])
        gastos_df = gastos_df.merge(map_gastos, on="subcat_norm", how="left")
    else:
        gastos_df["Naturaleza"] = "SIN_MAPEAR"

    # =============================
    # COSTO DE VIDA (FIJO + NEC)
    # =============================

    if "Naturaleza" in gastos_df.columns:
        costo_vida_df = gastos_df[gastos_df["Naturaleza"].isin(["FIJO", "NEC"])]
        costo_vida = costo_vida_df[monto_col].sum() / n_meses
    else:
        costo_vida = 0

    # =============================
    # AGRUPACIONES
    # =============================

    gasto_group = gastos_df.groupby(subcat_col)[monto_col].sum()
    ing_group = ingresos_df.groupby(tipo_ing_col)[ingresos_monto_col].sum()

    if modo == "Promedio mensual":
        gasto_group = gasto_group / n_meses
        ing_group = ing_group / n_meses

    gasto_group = gasto_group.sort_values(ascending=False)
    ing_group = ing_group.sort_values(ascending=False)

    # AGRUPAR OTROS
    total_gastos = gasto_group.sum()
    threshold = total_gastos * 0.015

    gastos_main = gasto_group[gasto_group >= threshold]
    otros = gasto_group[gasto_group < threshold].sum()
    if otros > 0:
        gastos_main["Otros"] = otros

    # =============================
    # KPIs
    # =============================

    total_ingresos = ing_group.sum()
    total_gastos = gasto_group.sum()
    ahorro = total_ingresos - total_gastos

    # =============================
    # TABS
    # =============================

    tab1, tab2, tab3 = st.tabs(["📊 Resumen", "📅 Evolución", "🔍 Drill-down"])

    with tab1:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Ingresos", format_ars(total_ingresos))
        col2.metric("Gastos", format_ars(total_gastos))
        col3.metric("Ahorro", format_ars(ahorro))
        col4.metric("Costo de vida mensual", format_ars(costo_vida))

        st.subheader("💰 Ingresos")
        st.plotly_chart(px.pie(values=ing_group.values, names=ing_group.index), use_container_width=True)

        st.subheader("📊 Gastos")
        st.plotly_chart(px.pie(values=gastos_main.values, names=gastos_main.index), use_container_width=True)

    with tab2:
        gastos_mensual = gastos_df.groupby("Mes")[monto_col].sum()
        ingresos_mensual = ingresos_df.groupby("Mes")[ingresos_monto_col].sum()

        evo = pd.DataFrame({"Ingresos": ingresos_mensual, "Gastos": gastos_mensual}).fillna(0)
        evo["Ahorro"] = evo["Ingresos"] - evo["Gastos"]

        st.subheader("📅 Evolución")
        st.line_chart(evo)

        if len(evo) > 1:
            ultimo = evo.iloc[-1]
            promedio = evo.mean()

            st.subheader("📊 Comparación último mes vs promedio")
            comp = pd.DataFrame({
                "Último mes": ultimo,
                "Promedio": promedio
            })
            st.bar_chart(comp)

    with tab3:
        st.subheader("🔍 Drill-down por subcategoría")

        subcat_sel = st.selectbox("Elegir subcategoría", gasto_group.index)

        detalle = gastos_df[gastos_df[subcat_col] == subcat_sel]
        detalle_mensual = detalle.groupby("Mes")[monto_col].sum()

        st.line_chart(detalle_mensual)

else:
    st.info("Subí un Excel para comenzar")
