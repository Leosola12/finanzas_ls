import streamlit as st
import pandas as pd

st.set_page_config(page_title="Finanzas Personales MVP", layout="wide")

st.title("📊 Finanzas Personales - MVP (mapeo persistente + dashboard)")

# --- Upload principal ---
file = st.file_uploader("Subí tu Excel", type=["xlsx"])

# --- Upload mapeo opcional ---
mapeo_file = st.sidebar.file_uploader("Cargar mapeo (.csv)", type=["csv"])

# --- Estado mapeo ---
if "mapeo" not in st.session_state:
    st.session_state.mapeo = {}

# --- Cargar mapeo desde CSV ---
if mapeo_file is not None:
    df_map = pd.read_csv(mapeo_file)
    for _, row in df_map.iterrows():
        st.session_state.mapeo[row["Subcategoria"]] = {
            "Area": row["Area"],
            "Naturaleza": row["Naturaleza"],
            "Controlable": row["Controlable"]
        }

if file:
    xls = pd.ExcelFile(file)

    gastos = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    ingresos = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

    gastos.columns = [c.strip() for c in gastos.columns]
    ingresos.columns = [c.strip() for c in ingresos.columns]

    monto_col_gastos = [c for c in gastos.columns if "monto" in c.lower() or "importe" in c.lower()]
    monto_col_ingresos = [c for c in ingresos.columns if "monto" in c.lower() or "importe" in c.lower()]

    if monto_col_gastos and monto_col_ingresos and "Subcategoria" in gastos.columns:

        subcats = gastos["Subcategoria"].dropna().unique()

        st.sidebar.header("🧠 Editor de mapeo")

        for sub in subcats:
            if sub not in st.session_state.mapeo:
                st.session_state.mapeo[sub] = {
                    "Area": "Sin clasificar",
                    "Naturaleza": "Sin clasificar",
                    "Controlable": "Sin clasificar"
                }

        for sub in subcats:
            with st.sidebar.expander(sub):
                area = st.selectbox("Área", ["ALIM","VIV","TRANS","OCIO","CONS","SALUD","FIN","LAB","Sin clasificar"],
                                    index=["ALIM","VIV","TRANS","OCIO","CONS","SALUD","FIN","LAB","Sin clasificar"].index(st.session_state.mapeo[sub]["Area"]), key=f"area_{sub}")

                naturaleza = st.selectbox("Naturaleza", ["FIJO","NEC","DISC","Sin clasificar"],
                                         index=["FIJO","NEC","DISC","Sin clasificar"].index(st.session_state.mapeo[sub]["Naturaleza"]), key=f"nat_{sub}")

                controlable = st.selectbox("Controlable", ["Sí","No","Sin clasificar"],
                                           index=["Sí","No","Sin clasificar"].index(st.session_state.mapeo[sub]["Controlable"]), key=f"ctrl_{sub}")

                st.session_state.mapeo[sub] = {
                    "Area": area,
                    "Naturaleza": naturaleza,
                    "Controlable": controlable
                }

        # --- Guardar mapeo ---
        if st.sidebar.button("💾 Descargar mapeo"):
            df_save = pd.DataFrame([
                {"Subcategoria": k, **v} for k, v in st.session_state.mapeo.items()
            ])
            st.sidebar.download_button("Descargar CSV", df_save.to_csv(index=False), "mapeo.csv")

        # --- Aplicar mapeo ---
        def mapear(sub, campo):
            return st.session_state.mapeo.get(sub, {}).get(campo, "Sin clasificar")

        gastos["Area"] = gastos["Subcategoria"].apply(lambda x: mapear(x, "Area"))
        gastos["Naturaleza"] = gastos["Subcategoria"].apply(lambda x: mapear(x, "Naturaleza"))
        gastos["Controlable"] = gastos["Subcategoria"].apply(lambda x: mapear(x, "Controlable"))

        # --- Cálculos ---
        gastos_total = gastos[monto_col_gastos[0]].sum()
        ingresos_total = ingresos[monto_col_ingresos[0]].sum()
        ahorro_teorico = ingresos_total - gastos_total

        st.sidebar.header("⚙️ Ajustes manuales")
        ahorro_real = st.sidebar.number_input("Ahorro real", value=0.0)
        pagos_deuda = st.sidebar.number_input("Pago de deudas", value=0.0)

        # --- KPIs ---
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Ingresos", f"${ingresos_total:,.0f}")
        col2.metric("Gastos", f"${gastos_total:,.0f}")
        col3.metric("Ahorro teórico", f"${ahorro_teorico:,.0f}")
        col4.metric("Ahorro real", f"${ahorro_real:,.0f}")

        # --- Dashboard avanzado ---
        st.subheader("📊 Gasto por naturaleza")
        st.bar_chart(gastos.groupby("Naturaleza")[monto_col_gastos[0]].sum())

        st.subheader("🎯 Gasto controlable vs no")
        st.bar_chart(gastos.groupby("Controlable")[monto_col_gastos[0]].sum())

        st.subheader("🔥 Top gastos")
        top = gastos.groupby("Subcategoria")[monto_col_gastos[0]].sum().sort_values(ascending=False).head(10)
        st.bar_chart(top)

        st.subheader("💣 Fugas (solo controlables)")
        fugas = gastos[gastos["Controlable"]=="Sí"].groupby("Subcategoria")[monto_col_gastos[0]].sum().sort_values(ascending=False).head(5)
        st.bar_chart(fugas)

        st.subheader("🧠 Insight")
        if "Sin clasificar" in gastos["Naturaleza"].values:
            st.warning("Tenés subcategorías sin clasificar")

        if ahorro_real < ahorro_teorico:
            st.warning("Tu ahorro real es menor al teórico (posibles deudas)")

    else:
        st.error("Formato no reconocido")
else:
    st.info("Subí un Excel para comenzar")
