import streamlit as st
import pandas as pd

st.set_page_config(page_title="Finanzas Personales", layout="wide")

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

# =============================
# UI
# =============================

st.title("📊 Finanzas Personales")

file = st.file_uploader("Subí tu Excel", type=["xlsx"])
mapping_file = st.sidebar.file_uploader("Subir mapeo.csv", type=["csv"])

if file:
    xls = pd.ExcelFile(file)

    # Detectar hojas
    gastos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    ingresos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

    # Detectar columnas dinámicamente
    subcat_col = find_column(gastos_df, ["subcategoria", "subcategoría"])
    monto_col = find_column(gastos_df, ["monto", "importe"])

    if not subcat_col or not monto_col:
        st.error("No se pudieron detectar columnas clave (Subcategoria / Monto)")
        st.stop()

    # Normalizar
    gastos_df["subcat_norm"] = normalize_series(gastos_df[subcat_col])

    # =============================
    # Mapeo
    # =============================

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

    # Merge
    df = gastos_df.merge(mapping_df, on="subcat_norm", how="left")

    # =============================
    # Editor
    # =============================

    st.sidebar.subheader("🧠 Editor de Mapeo")

    areas = ["ALIM","VIV","TRANS","OCIO","CONS","SALUD","FIN","LAB","SIN_MAPEAR"]
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

    st.sidebar.download_button(
        "💾 Descargar mapeo",
        updated_mapping.to_csv(index=False),
        file_name="mapeo.csv"
    )

    # Re-merge con mapping actualizado
    updated_mapping["subcat_norm"] = normalize_series(updated_mapping["Subcategoria"])
    df = gastos_df.merge(updated_mapping, on="subcat_norm", how="left")

    # =============================
    # KPIs
    # =============================

    ingresos_col = find_column(ingresos_df, ["monto","importe"])

    total_ingresos = ingresos_df[ingresos_col].sum() if ingresos_col else 0
    total_gastos = gastos_df[monto_col].sum()

    ahorro_teorico = total_ingresos - total_gastos

    col1, col2, col3 = st.columns(3)

    col1.metric("Ingresos", f"${total_ingresos:,.0f}")
    col2.metric("Gastos", f"${total_gastos:,.0f}")
    col3.metric("Ahorro teórico", f"${ahorro_teorico:,.0f}")

    # =============================
    # Ajustes
    # =============================

    st.subheader("🧾 Ajustes de realidad")

    ahorro_real = st.number_input("Ahorro real del mes", value=0)
    deuda = st.number_input("Pagos de deuda", value=0)

    st.write(f"Diferencia vs teórico: ${ahorro_real - ahorro_teorico:,.0f}")

    # =============================
    # Dashboard
    # =============================

    st.subheader("📊 Gasto por naturaleza")
    st.bar_chart(df.groupby("Naturaleza")[monto_col].sum())

    st.subheader("🎯 Controlable vs No")
    st.bar_chart(df.groupby("Controlable")[monto_col].sum())

    st.subheader("🔥 Top gastos")
    st.dataframe(df.groupby(subcat_col)[monto_col].sum().sort_values(ascending=False).head(10))

    st.subheader("💣 Fugas (controlables)")
    fugas = df[df["Controlable"] == "Si"]
    st.dataframe(fugas.groupby(subcat_col)[monto_col].sum().sort_values(ascending=False).head(10))

else:
    st.info("Subí un Excel para comenzar")
