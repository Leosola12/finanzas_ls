import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Finanzas Personales V10", layout="wide")

# =============================
# Helpers
# =============================

def normalize_colname(c):
    c = c.strip().lower()
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

st.title("📊 Finanzas Personales V10 — Panel configurable")

file = st.file_uploader("Subí tu Excel", type=["xlsx"])
mapping_file = st.sidebar.file_uploader("Mapeo gastos", type=["csv"])

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

    # Sidebar filters
    st.sidebar.header("📅 Filtros")
    meses_sel = st.sidebar.multiselect("Meses", meses, default=meses)
    modo = st.sidebar.radio("Modo", ["Total","Promedio mensual"])

    gastos_df = gastos_df[gastos_df["Mes"].isin(meses_sel)]
    ingresos_df = ingresos_df[ingresos_df["Mes"].isin(meses_sel)]
    n_meses = max(len(meses_sel),1)

    # Mapping
    if mapping_file:
        map_df = pd.read_csv(mapping_file)
        map_df["subcat_norm"] = normalize_series(map_df["Subcategoria"])
        gastos_df["subcat_norm"] = normalize_series(gastos_df[subcat_col])
        gastos_df = gastos_df.merge(map_df, on="subcat_norm", how="left")
    else:
        gastos_df["Naturaleza"] = "SIN"

    # Aggregations
    gasto_group = gastos_df.groupby(subcat_col)[monto_col].sum()
    ing_group = ingresos_df.groupby(tipo_ing_col)[ing_monto_col].sum()

    if modo == "Promedio mensual":
        gasto_group = gasto_group / n_meses
        ing_group = ing_group / n_meses

    total_ing = ing_group.sum()
    total_gas = gasto_group.sum()
    ahorro = total_ing - total_gas

    tasa_ahorro = (ahorro / total_ing) if total_ing != 0 else 0

    # Naturaleza splits
    if "Naturaleza" in gastos_df.columns:
        nat = gastos_df.groupby("Naturaleza")[monto_col].sum()
        if modo == "Promedio mensual":
            nat = nat / n_meses
        nec = nat.get("NEC",0)
        disc = nat.get("DISC",0)
        fijo = nat.get("FIJO",0)
    else:
        nec = disc = fijo = 0

    # =============================
    # SIDEBAR METRICS (UX)
    # =============================

    st.sidebar.header("📊 Métricas")

    perfil = st.sidebar.selectbox("Perfil", ["Personalizado","Control básico","Optimización","Diagnóstico"])

    default_checks = {
        "tasa_nec": False,
        "tasa_disc": False,
        "ratio_disc_nec": False,
        "top_subcat": False,
        "variabilidad": False,
        "pareto": False
    }

    if perfil == "Control básico":
        default_checks.update({"tasa_nec":True,"tasa_disc":True})
    elif perfil == "Optimización":
        default_checks.update({"ratio_disc_nec":True,"pareto":True})
    elif perfil == "Diagnóstico":
        default_checks.update({"variabilidad":True,"top_subcat":True})

    tasa_nec_on = st.sidebar.checkbox("Tasa NEC", value=default_checks["tasa_nec"])
    tasa_disc_on = st.sidebar.checkbox("Tasa DISC", value=default_checks["tasa_disc"])
    ratio_on = st.sidebar.checkbox("Ratio DISC/NEC", value=default_checks["ratio_disc_nec"])
    top_on = st.sidebar.checkbox("Top subcategoría", value=default_checks["top_subcat"])
    var_on = st.sidebar.checkbox("Variabilidad ahorro", value=default_checks["variabilidad"])
    pareto_on = st.sidebar.checkbox("Pareto", value=default_checks["pareto"])

    # =============================
    # KPIs BASE
    # =============================

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Ingresos", format_ars(total_ing))
    c2.metric("Gastos", format_ars(total_gas))
    c3.metric("Ahorro", format_ars(ahorro))
    c4.metric("Tasa ahorro", f"{tasa_ahorro:.1%}")
    c5.metric("Costo vida", format_ars(nec+fijo))

    # =============================
    # MÉTRICAS ACTIVAS
    # =============================

    st.subheader("📊 Análisis activo")

    if tasa_nec_on:
        st.write("Tasa NEC:", f"{(nec/total_ing if total_ing else 0):.1%}")

    if tasa_disc_on:
        st.write("Tasa DISC:", f"{(disc/total_ing if total_ing else 0):.1%}")

    if ratio_on:
        st.write("Ratio DISC/NEC:", (disc/nec if nec else 0))

    if top_on and not gasto_group.empty:
        top = gasto_group.sort_values(ascending=False).iloc[0]
        st.write("Top subcategoría:", format_ars(top))

    if pareto_on and not gasto_group.empty:
        dfp = gasto_group.sort_values(ascending=False).reset_index()
        dfp["%"] = dfp[monto_col]/dfp[monto_col].sum()
        dfp["% acum"] = dfp["%"].cumsum()
        st.dataframe(dfp.head(10))

    if var_on:
        evo = ingresos_df.groupby("Mes")[ing_monto_col].sum() - gastos_df.groupby("Mes")[monto_col].sum()
        var = evo.std()
        st.write("Variabilidad ahorro:", format_ars(var))

    # =============================
    # GRÁFICOS
    # =============================

    st.subheader("📈 Evolución")
    evo_df = pd.DataFrame({
        "Ingresos": ingresos_df.groupby("Mes")[ing_monto_col].sum(),
        "Gastos": gastos_df.groupby("Mes")[monto_col].sum()
    }).fillna(0)
    st.line_chart(evo_df)

else:
    st.info("Subí un archivo para comenzar")
