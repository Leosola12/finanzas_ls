import streamlit as st
import pandas as pd

st.set_page_config(page_title="Finanzas Personales V4", layout="wide")

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

map_naturaleza = {
    "FIJO": "🧱 Gastos fijos",
    "NEC": "🧠 Necesarios",
    "DISC": "🎯 Discrecionales",
    "SIN_MAPEAR": "❓ Sin clasificar"
}

# =============================
# UI
# =============================

st.title("📊 Finanzas Personales V4")

file = st.file_uploader("Subí tu Excel", type=["xlsx"])
mapping_file = st.sidebar.file_uploader("Subir mapeo.csv", type=["csv"])

if file:
    xls = pd.ExcelFile(file)
    gastos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    ingresos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

    subcat_col = find_column(gastos_df, ["subcategoria", "subcategoría"])
    monto_col = find_column(gastos_df, ["monto", "importe"])
    fecha_col = find_column(gastos_df, ["fecha"])

    gastos_df["subcat_norm"] = normalize_series(gastos_df[subcat_col])

    if fecha_col:
        gastos_df["Fecha"] = pd.to_datetime(gastos_df[fecha_col], errors="coerce")
        gastos_df["Mes"] = gastos_df["Fecha"].dt.to_period("M").astype(str)
    else:
        gastos_df["Mes"] = "Sin fecha"

    unique_subcats = gastos_df[["subcat_norm", subcat_col]].drop_duplicates()

    if mapping_file:
        mapping_df = pd.read_csv(mapping_file)
        mapping_df["subcat_norm"] = normalize_series(mapping_df["Subcategoria"])
    else:
        mapping_df = pd.DataFrame({
            "Subcategoria": unique_subcats[subcat_col],
            "Area": "SIN_MAPEAR",
            "Naturaleza": "SIN_MAPEAR",
            "Controlable": "Si"
        })
        mapping_df["subcat_norm"] = normalize_series(mapping_df["Subcategoria"])

    df = gastos_df.merge(mapping_df, on="subcat_norm", how="left")
    df["Naturaleza_label"] = df["Naturaleza"].map(map_naturaleza)

    ingresos_col = find_column(ingresos_df, ["monto","importe"])

    total_ingresos = ingresos_df[ingresos_col].sum() if ingresos_col else 0
    total_gastos = gastos_df[monto_col].sum()
    ahorro_teorico = total_ingresos - total_gastos

    # =============================
    # TABS
    # =============================

    tab1, tab2, tab3 = st.tabs(["📊 Resumen", "📅 Evolución", "🔥 Fugas"])

    # =============================
    # TAB 1
    # =============================

    with tab1:
        st.subheader("📊 Resumen general")

        col1, col2, col3 = st.columns(3)
        col1.metric("Ingresos", format_ars(total_ingresos))
        col2.metric("Gastos", format_ars(total_gastos))
        col3.metric("Ahorro", format_ars(ahorro_teorico))

        st.subheader("📊 ¿En qué se te va la plata?")

        pie_data = df.groupby("Naturaleza_label")[monto_col].sum()
        st.pyplot(pie_data.plot.pie(autopct='%1.1f%%').figure)

        st.subheader("📊 Distribución por subcategoría (Top 10)")
        st.bar_chart(df.groupby(subcat_col)[monto_col].sum().sort_values(ascending=False).head(10))

    # =============================
    # TAB 2
    # =============================

    with tab2:
        st.subheader("📅 Evolución de gastos en el tiempo")

        gasto_mensual = df.groupby("Mes")[monto_col].sum()
        st.line_chart(gasto_mensual)

        st.subheader("📊 Evolución por tipo de gasto")
        evo_nat = df.groupby(["Mes","Naturaleza_label"])[monto_col].sum().unstack()
        st.line_chart(evo_nat)

    # =============================
    # TAB 3
    # =============================

    with tab3:
        st.subheader("🔥 Principales fugas de dinero")

        fugas = df[df["Controlable"] == "Si"]
        st.dataframe(fugas.groupby(subcat_col)[monto_col].sum().sort_values(ascending=False).head(10))

else:
    st.info("Subí un Excel para comenzar")
