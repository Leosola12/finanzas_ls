import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Finanzas Personales V7", layout="wide")

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

st.title("📊 Finanzas Personales V7")

file = st.file_uploader("Subí tu Excel", type=["xlsx"])
mapping_file = st.sidebar.file_uploader("Subir mapeo.csv", type=["csv"])
mapping_ing_file = st.sidebar.file_uploader("Subir mapeo_ingresos.csv", type=["csv"])

if file:
    xls = pd.ExcelFile(file)
    gastos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    ingresos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

    # =============================
    # GASTOS
    # =============================

    subcat_col = find_column(gastos_df, ["subcategoria", "subcategoría"])
    monto_col = find_column(gastos_df, ["monto", "importe"])
    fecha_col = find_column(gastos_df, ["fecha"])

    gastos_df["subcat_norm"] = normalize_series(gastos_df[subcat_col])

    if fecha_col:
        gastos_df["Fecha"] = pd.to_datetime(gastos_df[fecha_col], errors="coerce")
        gastos_df["Mes"] = gastos_df["Fecha"].dt.to_period("M").astype(str)
    else:
        gastos_df["Mes"] = "Sin fecha"

    # =============================
    # INGRESOS
    # =============================

    ingresos_monto_col = find_column(ingresos_df, ["monto","importe"])
    ingresos_fecha_col = find_column(ingresos_df, ["fecha"])
    tipo_ing_col = find_column(ingresos_df, ["tipo de ingreso", "tipo ingreso"])

    ingresos_df["subcat_norm"] = normalize_series(ingresos_df[tipo_ing_col])

    if ingresos_fecha_col:
        ingresos_df["Fecha"] = pd.to_datetime(ingresos_df[ingresos_fecha_col], errors="coerce")
        ingresos_df["Mes"] = ingresos_df["Fecha"].dt.to_period("M").astype(str)
    else:
        ingresos_df["Mes"] = "Sin fecha"

    # =============================
    # KPIs
    # =============================

    total_ingresos = ingresos_df[ingresos_monto_col].sum()
    total_gastos = gastos_df[monto_col].sum()
    ahorro_teorico = total_ingresos - total_gastos

    # =============================
    # TABS
    # =============================

    tab1, tab2, tab3 = st.tabs(["📊 Resumen", "📅 Evolución", "🔥 Fugas"])

    with tab1:
        st.subheader("📊 Resumen general")

        col1, col2, col3 = st.columns(3)
        col1.metric("Ingresos", format_ars(total_ingresos))
        col2.metric("Gastos", format_ars(total_gastos))
        col3.metric("Ahorro", format_ars(ahorro_teorico))

        # INGRESOS PIE POR SUBCATEGORIA
        st.subheader("💰 ¿De dónde viene tu ingreso?")
        ing_group = ingresos_df.groupby(tipo_ing_col)[ingresos_monto_col].sum().reset_index()

        fig_ing = px.pie(
            ing_group,
            names=tipo_ing_col,
            values=ingresos_monto_col,
            title="Ingresos por tipo"
        )

        st.plotly_chart(fig_ing, use_container_width=True)

        # GASTOS PIE POR SUBCATEGORIA
        st.subheader("📊 ¿En qué se te va la plata?")
        gasto_group = gastos_df.groupby(subcat_col)[monto_col].sum().reset_index()

        fig_gasto = px.pie(
            gasto_group,
            names=subcat_col,
            values=monto_col,
            title="Gastos por subcategoría"
        )

        st.plotly_chart(fig_gasto, use_container_width=True)

    with tab2:
        st.subheader("📅 Evolución financiera completa")

        gastos_mensual = gastos_df.groupby("Mes")[monto_col].sum()
        ingresos_mensual = ingresos_df.groupby("Mes")[ingresos_monto_col].sum()

        evo = pd.DataFrame({
            "Ingresos": ingresos_mensual,
            "Gastos": gastos_mensual
        }).fillna(0)

        evo["Ahorro"] = evo["Ingresos"] - evo["Gastos"]

        st.line_chart(evo)

    with tab3:
        st.subheader("🔥 Principales oportunidades de mejora")
        fugas = gastos_df.groupby(subcat_col)[monto_col].sum().sort_values(ascending=False)
        st.dataframe(fugas.head(10))

else:
    st.info("Subí un Excel para comenzar")
