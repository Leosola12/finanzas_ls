import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Finanzas Personales V12", layout="wide")

# =============================
# Helpers
# =============================

def normalize_colname(c):
    c = str(c).strip().lower()
    for a, b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u")]:
        c = c.replace(a,b)
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
    s = s.astype(str).str.strip().str.lower()
    for a, b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u")]:
        s = s.str.replace(a,b)
    return s


def format_ars(x):
    try:
        return f"${x:,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except:
        return "$0,00"

# =============================
# UI
# =============================

st.title("📊 Finanzas Personales V12 — Panel completo")

file = st.file_uploader("Subí tu Excel", type=["xlsx"])
mapping_g_file = st.sidebar.file_uploader("Mapeo gastos", type=["csv"])
mapping_i_file = st.sidebar.file_uploader("Mapeo ingresos", type=["csv"])

if file:
    xls = pd.ExcelFile(file)
    gastos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    ingresos_df = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

    # Column detection
    subcat_col = find_column(gastos_df, ["subcategoria", "subcategoría"])
    monto_col = find_column(gastos_df, ["monto", "importe"])
    fecha_col = find_column(gastos_df, ["fecha"])

    ing_monto_col = find_column(ingresos_df, ["monto","importe"])
    ing_fecha_col = find_column(ingresos_df, ["fecha"])
    tipo_ing_col = find_column(ingresos_df, ["tipo ingreso","tipo de ingreso"])

    # Dates
    gastos_df["Fecha"] = pd.to_datetime(gastos_df[fecha_col], errors="coerce")
    ingresos_df["Fecha"] = pd.to_datetime(ingresos_df[ing_fecha_col], errors="coerce")

    gastos_df["Mes"] = gastos_df["Fecha"].dt.to_period("M").astype(str)
    ingresos_df["Mes"] = ingresos_df["Fecha"].dt.to_period("M").astype(str)

    meses = sorted(gastos_df["Mes"].dropna().unique())

    # =============================
    # Sidebar filtros
    # =============================

    st.sidebar.header("📅 Filtros")
    meses_sel = st.sidebar.multiselect("Meses", meses, default=meses)

    modo = st.sidebar.radio("Modo de análisis", ["Total del período", "Promedio mensual"])

    gastos_df = gastos_df[gastos_df["Mes"].isin(meses_sel)]
    ingresos_df = ingresos_df[ingresos_df["Mes"].isin(meses_sel)]

    n_meses = max(len(meses_sel), 1)

    # =============================
    # Mapping gastos
    # =============================

    if mapping_g_file:
        map_g = pd.read_csv(mapping_g_file)
        map_g["subcat_norm"] = normalize_series(map_g["Subcategoria"])
        gastos_df["subcat_norm"] = normalize_series(gastos_df[subcat_col])
        gastos_df = gastos_df.merge(map_g, on="subcat_norm", how="left")
    else:
        gastos_df["Naturaleza"] = "SIN"

    # =============================
    # Mapping ingresos
    # =============================

    if mapping_i_file:
        map_i = pd.read_csv(mapping_i_file)
        map_i["subcat_norm"] = normalize_series(map_i["Subcategoria"])
        ingresos_df["subcat_norm"] = normalize_series(ingresos_df[tipo_ing_col])
        ingresos_df = ingresos_df.merge(map_i, on="subcat_norm", how="left")
    else:
        ingresos_df["Estabilidad"] = "SIN"

    # =============================
    # Aggregations
    # =============================

    gasto_group = gastos_df.groupby(subcat_col)[monto_col].sum()
    ing_group = ingresos_df.groupby(tipo_ing_col)[ing_monto_col].sum()

    # Naturaleza
    nat = gastos_df.groupby("Naturaleza")[monto_col].sum()
    nec = nat.get("NEC", 0)
    disc = nat.get("DISC", 0)
    fijo = nat.get("FIJO", 0)

    # Ingresos
    est = ingresos_df.groupby("Estabilidad")[ing_monto_col].sum()
    fijo_ing = est.get("FIJO", 0)
    var_ing = est.get("VARIABLE", 0)

    # Aplicar promedio mensual
    if modo == "Promedio mensual":
        gasto_group /= n_meses
        ing_group /= n_meses
        nec /= n_meses
        disc /= n_meses
        fijo /= n_meses
        fijo_ing /= n_meses
        var_ing /= n_meses

    total_ing = ing_group.sum()
    total_gas = gasto_group.sum()
    ahorro = total_ing - total_gas

    tasa_ahorro = (ahorro / total_ing) if total_ing != 0 else 0

    # =============================
    # Sidebar UX análisis
    # =============================

    st.sidebar.header("📊 Qué querés analizar")

    perfil = st.sidebar.selectbox("Perfil", ["Personalizado","Control básico","Optimización","Diagnóstico"])

    defaults = {
        "tasa_nec": False,
        "tasa_disc": False,
        "ratio": False,
        "ingresos_var": False,
        "pareto": False
    }

    if perfil == "Control básico":
        defaults.update({"tasa_nec": True, "tasa_disc": True})
    elif perfil == "Optimización":
        defaults.update({"ratio": True, "pareto": True})
    elif perfil == "Diagnóstico":
        defaults.update({"ingresos_var": True, "ratio": True})

    tasa_nec_on = st.sidebar.checkbox("Costo de vida (%)", defaults["tasa_nec"])
    tasa_disc_on = st.sidebar.checkbox("Gastos personales (%)", defaults["tasa_disc"])
    ratio_on = st.sidebar.checkbox("Balance gustos vs necesidades", defaults["ratio"])
    ing_var_on = st.sidebar.checkbox("Estabilidad de ingresos", defaults["ingresos_var"])
    pareto_on = st.sidebar.checkbox("Dónde se va la plata", defaults["pareto"])

    # =============================
    # KPIs
    # =============================

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Ingresos", format_ars(total_ing))
    c2.metric("Gastos", format_ars(total_gas))
    c3.metric("Ahorro", format_ars(ahorro))
    c4.metric("Tasa ahorro", f"{tasa_ahorro:.1%}")
    c5.metric("Costo de vida", format_ars(nec + fijo))

    # =============================
    # Interpretación
    # =============================

    st.subheader("🧠 Lectura rápida")

    if total_ing > 0:
        if tasa_nec_on:
            tasa_nec = nec / total_ing
            st.write(f"💡 Necesidades básicas: {tasa_nec:.1%} de tus ingresos")

        if tasa_disc_on:
            tasa_disc = disc / total_ing
            st.write(f"🎯 Gustos personales: {tasa_disc:.1%} de tus ingresos")

        if ratio_on:
            ratio = disc / nec if nec else 0
            st.metric("Balance gustos / necesidades", f"{ratio:.2f}")

            if ratio > 1:
                st.warning("Estás gastando más en gustos que en necesidades.")
            else:
                st.success("Tus gastos están equilibrados.")

        if ing_var_on:
            var_ratio = var_ing / total_ing if total_ing else 0
            st.metric("Ingresos variables", f"{var_ratio:.1%}")

            if var_ratio > 0.5:
                st.warning("Tus ingresos son poco estables.")
            else:
                st.success("Tus ingresos son relativamente estables.")

    # =============================
    # Pareto
    # =============================

    if pareto_on and not gasto_group.empty:
        st.subheader("📊 Pareto de gastos")
        dfp = gasto_group.sort_values(ascending=False).reset_index()
        dfp["%"] = dfp[monto_col] / dfp[monto_col].sum()
        dfp["% acum"] = dfp["%"].cumsum()
        st.dataframe(dfp.head(10))

    # =============================
    # Gráficos
    # =============================

    st.subheader("📊 Distribución de gastos")
    st.plotly_chart(px.pie(values=gasto_group.values, names=gasto_group.index))

    st.subheader("💰 Distribución de ingresos")
    st.plotly_chart(px.pie(values=ing_group.values, names=ing_group.index))

    st.subheader("📈 Evolución")
    evo_df = pd.DataFrame({
        "Ingresos": ingresos_df.groupby("Mes")[ing_monto_col].sum(),
        "Gastos": gastos_df.groupby("Mes")[monto_col].sum()
    }).fillna(0)

    st.line_chart(evo_df)

else:
    st.info("Subí un archivo para comenzar")
