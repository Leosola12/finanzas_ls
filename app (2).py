import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Finanzas Personales V6", layout="wide")

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

st.title("📊 Finanzas Personales V6")

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

    # =============================
    # INGRESOS (TU MODELO REAL)
    # =============================

    ingresos_monto_col = find_column(ingresos_df, ["monto","importe"])
    ingresos_fecha_col = find_column(ingresos_df, ["fecha"])
    tipo_ing_col = find_column(ingresos_df, ["tipo de ingreso", "tipo ingreso"])

    # usamos "Tipo de Ingreso" como subcategoria
    ingresos_df["subcat_norm"] = normalize_series(ingresos_df[tipo_ing_col])

    if ingresos_fecha_col:
        ingresos_df["Fecha"] = pd.to_datetime(ingresos_df[ingresos_fecha_col], errors="coerce")
        ingresos_df["Mes"] = ingresos_df["Fecha"].dt.to_period("M").astype(str)
    else:
        ingresos_df["Mes"] = "Sin fecha"

    unique_ing = ingresos_df[["subcat_norm", tipo_ing_col]].drop_duplicates()

    if mapping_ing_file:
        ing_map = pd.read_csv(mapping_ing_file)
        ing_map["subcat_norm"] = normalize_series(ing_map["Subcategoria"])
    else:
        ing_map = pd.DataFrame({
            "Subcategoria": unique_ing[tipo_ing_col],
            "Origen": "SIN_MAPEAR",
            "Estabilidad": "SIN_MAPEAR"
        })
        ing_map["subcat_norm"] = normalize_series(ing_map["Subcategoria"])

    ingresos_df = ingresos_df.merge(ing_map, on="subcat_norm", how="left")

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

        st.subheader("💰 ¿De dónde viene tu ingreso?")
        pie_ing = ingresos_df.groupby("Origen")[ingresos_monto_col].sum()
        fig1, ax1 = plt.subplots()
pie_ing.plot.pie(autopct='%1.1f%%', ax=ax1)
ax1.set_ylabel("")
st.pyplot(fig1)

        st.subheader("📊 ¿En qué se te va la plata?")
        pie_data = df.groupby("Naturaleza_label")[monto_col].sum()
        fig2, ax2 = plt.subplots()
pie_data.plot.pie(autopct='%1.1f%%', ax=ax2)
ax2.set_ylabel("")
st.pyplot(fig2)

    with tab2:
        st.subheader("📅 Evolución financiera completa")

        gastos_mensual = df.groupby("Mes")[monto_col].sum()
        ingresos_mensual = ingresos_df.groupby("Mes")[ingresos_monto_col].sum()

        evo = pd.DataFrame({
            "Ingresos": ingresos_mensual,
            "Gastos": gastos_mensual
        }).fillna(0)

        evo["Ahorro"] = evo["Ingresos"] - evo["Gastos"]

        st.line_chart(evo)

    with tab3:
        st.subheader("🔥 Principales oportunidades de mejora")
        fugas = df[df["Controlable"] == "Si"]
        st.dataframe(fugas.groupby(subcat_col)[monto_col].sum().sort_values(ascending=False).head(10))

else:
    st.info("Subí un Excel para comenzar")
