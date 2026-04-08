import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Finanzas Personales V8", layout="wide")

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

st.title("📊 Finanzas Personales V8")

file = st.file_uploader("Subí tu Excel", type=["xlsx"])

if file:
    xls = pd.ExcelFile(file)
    gastos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    ingresos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

    # =============================
    # COLUMNAS
    # =============================

    subcat_col = find_column(gastos_df, ["subcategoria", "subcategoría"])
    monto_col = find_column(gastos_df, ["monto", "importe"])
    fecha_col = find_column(gastos_df, ["fecha"])

    ingresos_monto_col = find_column(ingresos_df, ["monto","importe"])
    ingresos_fecha_col = find_column(ingresos_df, ["fecha"])
    tipo_ing_col = find_column(ingresos_df, ["tipo de ingreso", "tipo ingreso"])

    # =============================
    # FECHAS
    # =============================

    gastos_df["Fecha"] = pd.to_datetime(gastos_df[fecha_col], errors="coerce")
    ingresos_df["Fecha"] = pd.to_datetime(ingresos_df[ingresos_fecha_col], errors="coerce")

    gastos_df["Mes"] = gastos_df["Fecha"].dt.to_period("M").astype(str)
    ingresos_df["Mes"] = ingresos_df["Fecha"].dt.to_period("M").astype(str)

    meses_disponibles = sorted(gastos_df["Mes"].dropna().unique())

    # =============================
    # FILTROS
    # =============================

    st.sidebar.header("📅 Filtros")

    meses_sel = st.sidebar.multiselect(
        "Seleccionar meses",
        options=meses_disponibles,
        default=meses_disponibles
    )

    modo = st.sidebar.radio("Modo de análisis", ["Total", "Promedio mensual"])

    # Filtrado
    gastos_df = gastos_df[gastos_df["Mes"].isin(meses_sel)]
    ingresos_df = ingresos_df[ingresos_df["Mes"].isin(meses_sel)]

    n_meses = max(len(meses_sel), 1)

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

    # =============================
    # AGRUPAR "OTROS" (<1.5%)
    # =============================

    total_gastos = gasto_group.sum()
    threshold = total_gastos * 0.015

    gastos_main = gasto_group[gasto_group >= threshold]
    gastos_otros = gasto_group[gasto_group < threshold].sum()

    if gastos_otros > 0:
        gastos_main["Otros"] = gastos_otros

    # =============================
    # KPIs
    # =============================

    total_ingresos = ing_group.sum()
    total_gastos = gasto_group.sum()
    ahorro = total_ingresos - total_gastos

    # =============================
    # TABS
    # =============================

    tab1, tab2 = st.tabs(["📊 Resumen", "📅 Evolución"])

    with tab1:
        col1, col2, col3 = st.columns(3)
        col1.metric("Ingresos", format_ars(total_ingresos))
        col2.metric("Gastos", format_ars(total_gastos))
        col3.metric("Ahorro", format_ars(ahorro))

        # INGRESOS
        st.subheader("💰 Ingresos por tipo")
        fig_ing = px.pie(
            values=ing_group.values,
            names=ing_group.index,
            title="Distribución de ingresos"
        )
        st.plotly_chart(fig_ing, use_container_width=True)

        # GASTOS
        st.subheader("📊 Gastos por subcategoría")
        fig_gasto = px.pie(
            values=gastos_main.values,
            names=gastos_main.index,
            title="Distribución de gastos"
        )
        st.plotly_chart(fig_gasto, use_container_width=True)

    with tab2:
        st.subheader("📅 Evolución")

        gastos_mensual = gastos_df.groupby("Mes")[monto_col].sum()
        ingresos_mensual = ingresos_df.groupby("Mes")[ingresos_monto_col].sum()

        evo = pd.DataFrame({
            "Ingresos": ingresos_mensual,
            "Gastos": gastos_mensual
        }).fillna(0)

        evo["Ahorro"] = evo["Ingresos"] - evo["Gastos"]

        st.line_chart(evo)

else:
    st.info("Subí un Excel para comenzar")
