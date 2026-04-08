import streamlit as st
import pandas as pd

st.set_page_config(page_title="Finanzas Personales V3", layout="wide")

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

st.title("📊 Finanzas Personales V3")

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

    st.sidebar.subheader("🧠 Editor de Mapeo")

    areas = ["ALIM","VIV","TRANS","OCIO","CONS","SALUD","FIN","LAB","SERV","SIN_MAPEAR"]
    naturalezas = ["FIJO","NEC","DISC","SIN_MAPEAR"]
    controlables = ["Si","No"]

    edited_rows = []

    for i, row in mapping_df.iterrows():
        col1, col2, col3 = st.sidebar.columns(3)

        subcat = row["Subcategoria"]

        area = col1.selectbox("Area", areas, index=areas.index(row["Area"]) if row["Area"] in areas else 0, key=f"a{i}")
        nat = col2.selectbox("Nat", naturalezas, index=naturalezas.index(row["Naturaleza"]) if row["Naturaleza"] in naturalezas else 0, key=f"n{i}")
        ctrl = col3.selectbox("Ctrl", controlables, index=controlables.index(row["Controlable"]) if row["Controlable"] in controlables else 0, key=f"c{i}")

        edited_rows.append([subcat, area, nat, ctrl])

    updated_mapping = pd.DataFrame(edited_rows, columns=["Subcategoria","Area","Naturaleza","Controlable"])
    updated_mapping["subcat_norm"] = normalize_series(updated_mapping["Subcategoria"])

    st.sidebar.download_button("💾 Descargar mapeo", updated_mapping.to_csv(index=False), "mapeo.csv")

    df = gastos_df.merge(updated_mapping, on="subcat_norm", how="left")
    df["Naturaleza_label"] = df["Naturaleza"].map(map_naturaleza)

    ingresos_col = find_column(ingresos_df, ["monto","importe"])

    total_ingresos = ingresos_df[ingresos_col].sum() if ingresos_col else 0
    total_gastos = gastos_df[monto_col].sum()
    ahorro_teorico = total_ingresos - total_gastos

    # =============================
    # INPUTS POR MES
    # =============================

    st.sidebar.subheader("💰 Ajustes reales por mes")

    meses = sorted(df["Mes"].unique())

    ajustes = []

    for mes in meses:
        col1, col2 = st.sidebar.columns(2)
        ahorro = col1.number_input(f"Ahorro {mes}", value=0, key=f"ah_{mes}")
        deuda = col2.number_input(f"Deuda {mes}", value=0, key=f"de_{mes}")
        ajustes.append([mes, ahorro, deuda])

    ajustes_df = pd.DataFrame(ajustes, columns=["Mes","Ahorro_real","Deuda"])

    # =============================
    # TABS
    # =============================

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "📅 Evolución", "🔥 Fugas", "🧠 Insights"])

    with tab1:
        col1, col2, col3 = st.columns(3)
        col1.metric("Ingresos", format_ars(total_ingresos))
        col2.metric("Gastos", format_ars(total_gastos))
        col3.metric("Ahorro teórico", format_ars(ahorro_teorico))

        st.subheader("Gasto por tipo")
        st.bar_chart(df.groupby("Naturaleza_label")[monto_col].sum())

    with tab2:
        gasto_mensual = df.groupby("Mes")[monto_col].sum()
        st.line_chart(gasto_mensual)

        st.subheader("Comparación mensual")
        comp = gasto_mensual.reset_index().merge(ajustes_df, on="Mes", how="left")
        st.dataframe(comp)

    with tab3:
        fugas = df[df["Controlable"] == "Si"]
        st.dataframe(fugas.groupby(subcat_col)[monto_col].sum().sort_values(ascending=False).head(10))

    with tab4:
        st.subheader("Insights")

        disc = df[df["Naturaleza"] == "DISC"][monto_col].sum()
        total = df[monto_col].sum()

        if total > 0:
            ratio = disc / total * 100
            st.write(f"🎯 Discrecional: {ratio:.1f}%")

            if ratio > 40:
                st.warning("⚠️ Alto gasto discrecional")
            else:
                st.success("Buen nivel de gasto")

else:
    st.info("Subí un Excel para comenzar")
