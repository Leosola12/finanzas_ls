import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import base64
from datetime import datetime

st.set_page_config(page_title="Finanzas Personales V1.14", layout="wide", page_icon="💰")

# ─────────────────────────────────────────
# CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #0f1117; }
    .kpi-card {
        background: linear-gradient(135deg, #1e2130 0%, #262d40 100%);
        border-radius: 12px;
        padding: 16px 20px;
        border-left: 4px solid #4f8ef7;
        margin-bottom: 8px;
    }
    .kpi-card.green  { border-left-color: #2ecc71; }
    .kpi-card.red    { border-left-color: #e74c3c; }
    .kpi-card.yellow { border-left-color: #f39c12; }
    .kpi-label { color: #8892a4; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; }
    .kpi-value { color: #ffffff; font-size: 24px; font-weight: 700; margin-top: 4px; }
    .kpi-sub   { color: #8892a4; font-size: 12px; margin-top: 2px; }
    .kpi-label-row { display: flex; align-items: center; gap: 6px; }
    .tooltip-icon {
        display: inline-flex; align-items: center; justify-content: center;
        width: 15px; height: 15px; border-radius: 50%;
        background: #3a4460; color: #8892a4; font-size: 10px;
        cursor: help; position: relative; flex-shrink: 0;
    }
    .tooltip-icon:hover::after {
        content: attr(data-tip);
        position: absolute; bottom: 22px; left: 50%; transform: translateX(-50%);
        background: #1a2035; color: #c8d0e0; font-size: 11px; line-height: 1.5;
        padding: 10px 14px; border-radius: 8px; width: 260px;
        border: 1px solid #2e3a55; box-shadow: 0 4px 20px rgba(0,0,0,0.4);
        z-index: 9999; pointer-events: none; white-space: normal; font-weight: 400;
        text-transform: none; letter-spacing: 0;
    }
    .section-title {
        font-size: 18px; font-weight: 600; color: #e0e6f0;
        margin: 28px 0 12px 0; border-bottom: 1px solid #2a3040; padding-bottom: 6px;
    }
    .badge {
        display: inline-block; padding: 3px 10px; border-radius: 20px;
        font-size: 11px; font-weight: 600;
    }
    .badge-nec  { background:#1a3a5c; color:#4f9de8; }
    .badge-disc { background:#2d1f4a; color:#9b59b6; }
    .badge-fijo { background:#1a3d2b; color:#2ecc71; }
    [data-testid="stMetricValue"] { font-size: 22px !important; }
    div[data-testid="stTabs"] button { font-size: 15px !important; font-weight: 600 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────
def normalize_colname(c):
    c = str(c).strip().lower()
    for a, b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u")]:
        c = c.replace(a, b)
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
        s = s.str.replace(a, b)
    return s

def fmt(x):
    try:
        return f"${x:,.0f}".replace(",", ".")
    except:
        return "$0"

def pct(x):
    try:
        return f"{x:.1%}"
    except:
        return "0%"

def kpi_card(label, value, sub="", color="blue", tooltip=""):
    tip_html = (
        f'<span class="tooltip-icon" data-tip="{tooltip}">?</span>'
        if tooltip else ""
    )
    st.markdown(f"""
    <div class="kpi-card {color}">
        <div class="kpi-label-row">
            <span class="kpi-label">{label}</span>{tip_html}
        </div>
        <div class="kpi-value">{value}</div>
        {'<div class="kpi-sub">' + sub + '</div>' if sub else ''}
    </div>
    """, unsafe_allow_html=True)

PLOTLY_THEME = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#c0cce0", family="Inter, sans-serif"),
    margin=dict(t=40, b=20, l=20, r=20),
)

COLORS_MAIN = ["#4f8ef7","#2ecc71","#9b59b6","#f39c12","#e74c3c",
               "#1abc9c","#e67e22","#3498db","#e91e63","#00bcd4",
               "#ff9800","#8bc34a","#9c27b0","#795548","#607d8b"]

# ─────────────────────────────────────────
# SIDEBAR — uploads + filters
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Archivos")
    file = st.file_uploader("Excel de datos", type=["xlsx"])
    mapping_g_file = st.file_uploader("Mapeo gastos (.csv)", type=["csv"])
    mapping_i_file = st.file_uploader("Mapeo ingresos (.csv)", type=["csv"])

if not file:
    st.markdown("## 💰 Finanzas Personales V.14")
    st.info("Subí tu Excel en el panel izquierdo para comenzar.")
    st.stop()

# ─────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────
xls = pd.ExcelFile(file)
gastos_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
ingresos_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

subcat_col   = find_column(gastos_raw, ["subcategoria", "subcategoría"])
monto_col    = find_column(gastos_raw, ["monto", "importe"])
fecha_col    = find_column(gastos_raw, ["fecha"])
ing_monto_col = find_column(ingresos_raw, ["monto","importe"])
ing_fecha_col = find_column(ingresos_raw, ["fecha"])
tipo_ing_col  = find_column(ingresos_raw, ["tipo ingreso","tipo de ingreso"])

gastos_raw["Fecha"] = pd.to_datetime(gastos_raw[fecha_col], errors="coerce")
ingresos_raw["Fecha"] = pd.to_datetime(ingresos_raw[ing_fecha_col], errors="coerce")
gastos_raw["Mes"] = gastos_raw["Fecha"].dt.to_period("M").astype(str)
ingresos_raw["Mes"] = ingresos_raw["Fecha"].dt.to_period("M").astype(str)

# ─── Mapping gastos ───
if mapping_g_file:
    map_g = pd.read_csv(mapping_g_file)
    map_g["subcat_norm"] = normalize_series(map_g["Subcategoria"])
    gastos_raw["subcat_norm"] = normalize_series(gastos_raw[subcat_col])
    gastos_raw = gastos_raw.merge(map_g, on="subcat_norm", how="left")
    gastos_raw["Naturaleza"] = gastos_raw["Naturaleza"].fillna("SIN")
    gastos_raw["Area"] = gastos_raw.get("Area", pd.Series("SIN", index=gastos_raw.index)).fillna("SIN")
    gastos_raw["Controlable"] = gastos_raw.get("Controlable", pd.Series("Si", index=gastos_raw.index)).fillna("Si")
else:
    gastos_raw["Naturaleza"] = "SIN"
    gastos_raw["Area"] = "SIN"
    gastos_raw["Controlable"] = "Si"

# ─── Mapping ingresos ───
if mapping_i_file:
    map_i = pd.read_csv(mapping_i_file)
    map_i["subcat_norm"] = normalize_series(map_i["Subcategoria"])
    ingresos_raw["subcat_norm"] = normalize_series(ingresos_raw[tipo_ing_col])
    ingresos_raw = ingresos_raw.merge(map_i, on="subcat_norm", how="left")
    ingresos_raw["Estabilidad"] = ingresos_raw["Estabilidad"].fillna("SIN")
    ingresos_raw["Origen"] = ingresos_raw.get("Origen", pd.Series("SIN", index=ingresos_raw.index)).fillna("SIN")
else:
    ingresos_raw["Estabilidad"] = "SIN"
    ingresos_raw["Origen"] = "SIN"

# ─────────────────────────────────────────
# SIDEBAR — filters
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("## 🎛️ Filtros")
    meses_all = sorted(gastos_raw["Mes"].dropna().unique())
    meses_sel = st.multiselect("Meses", meses_all, default=meses_all)
    modo = st.radio("Modo", ["Total del período", "Promedio mensual"])

gastos_df   = gastos_raw[gastos_raw["Mes"].isin(meses_sel)].copy()
ingresos_df = ingresos_raw[ingresos_raw["Mes"].isin(meses_sel)].copy()
n_meses = max(len(meses_sel), 1)

# ─────────────────────────────────────────
# AGGREGATIONS
# ─────────────────────────────────────────
gasto_group = gastos_df.groupby(subcat_col)[monto_col].sum()
ing_group   = ingresos_df.groupby(tipo_ing_col)[ing_monto_col].sum()

nat  = gastos_df.groupby("Naturaleza")[monto_col].sum()
nec  = nat.get("NEC", 0)
disc = nat.get("DISC", 0)
fijo = nat.get("FIJO", 0)

est      = ingresos_df.groupby("Estabilidad")[ing_monto_col].sum()
fijo_ing = est.get("Fijo", 0)
var_ing  = est.get("Variable", 0)
ocas_ing = est.get("Ocasional", 0)

if modo == "Promedio mensual":
    gasto_group /= n_meses
    ing_group   /= n_meses
    nec  /= n_meses
    disc /= n_meses
    fijo /= n_meses
    fijo_ing /= n_meses
    var_ing  /= n_meses
    ocas_ing /= n_meses

total_ing = ing_group.sum()
total_gas = gasto_group.sum()
ahorro    = total_ing - total_gas
tasa_ahorro = (ahorro / total_ing) if total_ing > 0 else 0

# ─── Health score ───
def health_score(tasa_ahorro, nec, disc, total_ing, fijo_ing):
    score = 0
    if tasa_ahorro >= 0.20: score += 40
    elif tasa_ahorro >= 0.10: score += 25
    elif tasa_ahorro >= 0: score += 10
    ratio_nec = nec / total_ing if total_ing else 1
    if ratio_nec < 0.50: score += 30
    elif ratio_nec < 0.65: score += 18
    elif ratio_nec < 0.80: score += 8
    ratio_disc = disc / total_ing if total_ing else 1
    if ratio_disc < 0.15: score += 15
    elif ratio_disc < 0.25: score += 8
    ratio_fijo_ing = fijo_ing / total_ing if total_ing else 0
    if ratio_fijo_ing >= 0.70: score += 15
    elif ratio_fijo_ing >= 0.40: score += 8
    return min(score, 100)

score = health_score(tasa_ahorro, nec, disc, total_ing, fijo_ing)

def score_label(s):
    if s >= 75: return "Excelente 🟢", "green"
    if s >= 50: return "Saludable 🟡", "yellow"
    if s >= 30: return "Atención 🟠", "yellow"
    return "Crítico 🔴", "red"

score_text, score_color = score_label(score)

# ─── Month-over-month comparison ───
evo_gas = gastos_raw.groupby("Mes")[monto_col].sum()
evo_ing = ingresos_raw.groupby("Mes")[ing_monto_col].sum()
evo_df  = pd.DataFrame({"Ingresos": evo_ing, "Gastos": evo_gas}).fillna(0)
evo_df["Ahorro"] = evo_df["Ingresos"] - evo_df["Gastos"]

cat_mes = gastos_raw.groupby(["Mes", subcat_col])[monto_col].sum().reset_index()

# ─────────────────────────────────────────
# TABS
# ─────────────────────────────────────────
st.markdown("# 💰 Finanzas Personales V1.14")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Resumen", 
    "💸 Gastos", 
    "💵 Ingresos",
    "📤 Exportar"
])

# ═══════════════════════════════════════════
# TAB 1 — RESUMEN
# ═══════════════════════════════════════════
with tab1:
    
    # KPIs
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        kpi_card("Ingresos", fmt(total_ing), color="green")
    with c2:
        kpi_card("Gastos", fmt(total_gas), color="red")
    with c3:
        kpi_card("Ahorro", fmt(ahorro), color="green" if ahorro >= 0 else "red", tooltip="Este valor es la diferencia entre los ingresos y gastos registrados en el período. No necesariamente refleja dinero efectivamente ahorrado: puede haber gastos no registrados, movimientos fuera del tracker o diferencias de timing.")
    with c4:
        kpi_card("Tasa de ahorro", pct(tasa_ahorro), color="green" if tasa_ahorro >= 0.15 else "yellow" if tasa_ahorro >= 0 else "red")
    with c5:
        kpi_card("Salud financiera", f"{score}/100", sub=score_text, color=score_color)

    st.markdown("")

    # Gauge + lectura
    col_gauge, col_lectura = st.columns([1, 1])

    with col_gauge:
        st.markdown('<div class="section-title">🎯 Score de salud financiera</div>', unsafe_allow_html=True)
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=score,
            number={"suffix": "/100", "font": {"size": 36, "color": "#ffffff"}},
            gauge={
                "axis": {"range": [0, 100], "tickcolor": "#8892a4"},
                "bar": {"color": "#4f8ef7", "thickness": 0.25},
                "bgcolor": "#1e2130",
                "borderwidth": 0,
                "steps": [
                    {"range": [0, 30],  "color": "#2d1018"},
                    {"range": [30, 50], "color": "#2d2010"},
                    {"range": [50, 75], "color": "#1e2d10"},
                    {"range": [75, 100],"color": "#102d1e"},
                ],
                "threshold": {"line": {"color": "#4f8ef7", "width": 3}, "thickness": 0.75, "value": score}
            }
        ))
        fig_gauge.update_layout(**PLOTLY_THEME, height=260)
        st.plotly_chart(fig_gauge, use_container_width=True)

    with col_lectura:
        st.markdown('<div class="section-title">🧠 Lectura rápida</div>', unsafe_allow_html=True)
        
        ratio_nec_val = nec / total_ing if total_ing else 0
        ratio_disc_val = disc / total_ing if total_ing else 0
        ratio_fijo_val = (nec + fijo) / total_ing if total_ing else 0
        ratio_disc_vs_nec = disc / nec if nec else 0
        ratio_ing_fijo = fijo_ing / total_ing if total_ing else 0

        insights = []
        
        if tasa_ahorro >= 0.20:
            insights.append(("✅", f"Excelente tasa de ahorro: **{pct(tasa_ahorro)}** de tus ingresos."))
        elif tasa_ahorro >= 0.10:
            insights.append(("🟡", f"Tasa de ahorro moderada: **{pct(tasa_ahorro)}**. Podés mejorar."))
        elif tasa_ahorro >= 0:
            insights.append(("⚠️", f"Ahorrás poco: **{pct(tasa_ahorro)}** de tus ingresos. Margen estrecho."))
        else:
            insights.append(("🔴", f"Déficit: gastás **{fmt(-ahorro)}** más de lo que ingresás."))

        if ratio_nec_val > 0.65:
            insights.append(("⚠️", f"Tu costo de vida básico consume **{pct(ratio_nec_val)}** de tus ingresos."))
        else:
            insights.append(("✅", f"Costo de vida controlado: **{pct(ratio_nec_val)}** de tus ingresos."))

        if ratio_disc_vs_nec > 0.5:
            insights.append(("🟡", f"Tus gustos representan el **{pct(ratio_disc_val)}** de lo que ganás ({ratio_disc_vs_nec:.1f}x vs. necesidades)."))
        else:
            insights.append(("✅", f"Gastos discrecionales equilibrados: **{pct(ratio_disc_val)}** de ingresos."))

        if ratio_ing_fijo >= 0.7:
            insights.append(("✅", f"Ingresos muy estables: **{pct(ratio_ing_fijo)}** son fijos."))
        elif ratio_ing_fijo >= 0.4:
            insights.append(("🟡", f"Ingresos moderadamente estables: **{pct(ratio_ing_fijo)}** son fijos."))
        else:
            insights.append(("⚠️", f"Alta dependencia de ingresos variables (**{pct(1-ratio_ing_fijo)}**)."))

        for icon, text in insights:
            st.markdown(f"{icon} {text}")

    # Distribución NEC/DISC/FIJO
    st.markdown('<div class="section-title">📊 Composición del gasto</div>', unsafe_allow_html=True)
    
    col_nat, col_area = st.columns([1, 1])
    
    with col_nat:
        nat_df = gastos_df.groupby("Naturaleza")[monto_col].sum().reset_index()
        nat_df.columns = ["Naturaleza", "Monto"]
        nat_labels = {"NEC": "Necesidades", "DISC": "Discrecional", "FIJO": "Fijo", "SIN": "Sin mapear"}
        nat_colors = {"NEC": "#4f8ef7", "DISC": "#9b59b6", "FIJO": "#2ecc71", "SIN": "#8892a4"}
        nat_df["Label"] = nat_df["Naturaleza"].map(nat_labels).fillna(nat_df["Naturaleza"])
        nat_df["Color"] = nat_df["Naturaleza"].map(nat_colors).fillna("#8892a4")
        
        fig_nat = px.pie(
            nat_df, values="Monto", names="Label",
            color="Label",
            color_discrete_map={v: nat_colors[k] for k, v in nat_labels.items()},
            hole=0.45,
        )
        fig_nat.update_traces(textposition="outside", textfont_size=12)
        fig_nat.update_layout(**PLOTLY_THEME, height=300, title="Por naturaleza",
                              showlegend=True, legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(fig_nat, use_container_width=True)

    with col_area:
        area_df = gastos_df.groupby("Area")[monto_col].sum().reset_index()
        area_df.columns = ["Area", "Monto"]
        area_df = area_df[area_df["Area"] != "SIN"]
        if not area_df.empty:
            fig_area = px.bar(
                area_df.sort_values("Monto", ascending=True),
                x="Monto", y="Area", orientation="h",
                color="Area", color_discrete_sequence=COLORS_MAIN,
                text="Monto",
            )
            fig_area.update_traces(texttemplate="$%{x:,.0f}", textposition="outside")
            fig_area.update_layout(**PLOTLY_THEME, height=300, title="Por área",
                                   showlegend=False,
                                   xaxis=dict(showticklabels=False, showgrid=False),
                                   yaxis=dict(showgrid=False))
            st.plotly_chart(fig_area, use_container_width=True)

    # Evolución mensual
    st.markdown('<div class="section-title">📈 Evolución mensual</div>', unsafe_allow_html=True)
    
    fig_evo = go.Figure()
    fig_evo.add_trace(go.Bar(
        x=evo_df.index, y=evo_df["Gastos"],
        name="Gastos", marker_color="#e74c3c", opacity=0.85
    ))
    fig_evo.add_trace(go.Scatter(
        x=evo_df.index, y=evo_df["Ingresos"],
        name="Ingresos", mode="lines+markers",
        line=dict(color="#2ecc71", width=3),
        marker=dict(size=8, color="#2ecc71")
    ))
    fig_evo.add_trace(go.Scatter(
        x=evo_df.index, y=evo_df["Ahorro"],
        name="Ahorro", mode="lines+markers",
        line=dict(color="#4f8ef7", width=2, dash="dot"),
        marker=dict(size=6, color="#4f8ef7")
    ))
    fig_evo.update_layout(**PLOTLY_THEME, height=320,
                          legend=dict(orientation="h", y=1.1),
                          xaxis=dict(showgrid=False),
                          yaxis=dict(showgrid=True, gridcolor="#1e2540"))
    st.plotly_chart(fig_evo, use_container_width=True)


# ═══════════════════════════════════════════
# TAB 2 — GASTOS
# ═══════════════════════════════════════════
with tab2:

    st.markdown('<div class="section-title">📊 Distribución por subcategoría</div>', unsafe_allow_html=True)
    
    col_pie, col_bar = st.columns([1, 1])
    
    with col_pie:
        fig_pie = px.pie(
            values=gasto_group.values,
            names=gasto_group.index,
            hole=0.4,
            color_discrete_sequence=COLORS_MAIN,
        )
        fig_pie.update_traces(textposition="inside", textfont_size=11)
        fig_pie.update_layout(**PLOTLY_THEME, height=380, showlegend=True,
                              legend=dict(orientation="v", x=1.01, font=dict(size=11)))
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_bar:
        pareto_df = gasto_group.sort_values(ascending=False).reset_index()
        pareto_df.columns = ["Subcategoría", "Monto"]
        pareto_df["%"] = pareto_df["Monto"] / pareto_df["Monto"].sum()
        pareto_df["% acum"] = pareto_df["%"].cumsum()
        
        fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
        fig_pareto.add_trace(go.Bar(
            x=pareto_df["Subcategoría"], y=pareto_df["Monto"],
            name="Monto", marker_color="#4f8ef7", opacity=0.85
        ), secondary_y=False)
        fig_pareto.add_trace(go.Scatter(
            x=pareto_df["Subcategoría"], y=pareto_df["% acum"],
            name="% acumulado", mode="lines+markers",
            line=dict(color="#f39c12", width=2),
            marker=dict(size=6)
        ), secondary_y=True)
        fig_pareto.update_layout(**PLOTLY_THEME, height=380,
                                  title="Pareto de gastos",
                                  xaxis=dict(tickangle=-40, showgrid=False),
                                  yaxis=dict(showgrid=True, gridcolor="#1e2540"),
                                  legend=dict(orientation="h", y=1.1))
        fig_pareto.update_yaxes(tickformat=".0%", secondary_y=True,
                                 showgrid=False, range=[0, 1.05])
        st.plotly_chart(fig_pareto, use_container_width=True)

    # Tabla Pareto
    st.markdown('<div class="section-title">📋 Top subcategorías</div>', unsafe_allow_html=True)
    pareto_show = pareto_df.copy()
    pareto_show["Monto"] = pareto_show["Monto"].apply(fmt)
    pareto_show["%"] = pareto_show["%"].apply(pct)
    pareto_show["% acum"] = pareto_show["% acum"].apply(pct)
    st.dataframe(pareto_show, use_container_width=True, hide_index=True)

    # ── Comparativa entre meses ──
    st.markdown('<div class="section-title">📅 Comparativa entre meses</div>', unsafe_allow_html=True)

    # Barras agrupadas por subcategoría
    pivot = cat_mes.pivot_table(index=subcat_col, columns="Mes", values=monto_col, fill_value=0)
    pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index]
    top_cats = pivot.head(12).reset_index()
    top_melt = top_cats.melt(id_vars=subcat_col, var_name="Mes", value_name="Monto")

    fig_comp = px.bar(
        top_melt, x=subcat_col, y="Monto", color="Mes",
        barmode="group",
        color_discrete_sequence=COLORS_MAIN,
    )
    fig_comp.update_layout(**PLOTLY_THEME, height=380,
                            xaxis=dict(tickangle=-40, showgrid=False, title=""),
                            yaxis=dict(showgrid=True, gridcolor="#1e2540"),
                            legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_comp, use_container_width=True)

    # Tabla variación % mes a mes
    st.markdown("**Variación % mes a mes**")
    if len(pivot.columns) >= 2:
        meses_ord = sorted(pivot.columns)
        var_rows = []
        for i in range(1, len(meses_ord)):
            m_prev, m_curr = meses_ord[i-1], meses_ord[i]
            row = {"Subcategoría": list(pivot.index)}
            prev_vals = pivot[m_prev]
            curr_vals = pivot[m_curr]
            var = ((curr_vals - prev_vals) / prev_vals.replace(0, float("nan"))).fillna(0)
            df_var = pd.DataFrame({
                "Subcategoría": pivot.index,
                f"{m_prev}": prev_vals.apply(fmt).values,
                f"{m_curr}": curr_vals.apply(fmt).values,
                "Variación %": var.apply(lambda x: f"+{x:.1%}" if x > 0 else f"{x:.1%}").values,
            })
            st.caption(f"{m_prev} → {m_curr}")
            st.dataframe(df_var, use_container_width=True, hide_index=True)
    else:
        st.info("Seleccioná al menos 2 meses para ver la comparativa.")


# ═══════════════════════════════════════════
# TAB 3 — INGRESOS
# ═══════════════════════════════════════════
with tab3:

    col_a, col_b = st.columns(2)

    with col_a:
        kpi_card("Total ingresos", fmt(total_ing), color="green")
    with col_b:
        kpi_card("Ingresos fijos", fmt(fijo_ing), sub=pct(fijo_ing/total_ing) + " del total" if total_ing else "", color="green")

    st.markdown('<div class="section-title">📊 Distribución por tipo</div>', unsafe_allow_html=True)

    col_i1, col_i2 = st.columns([1, 1])

    with col_i1:
        fig_ing_pie = px.pie(
            values=ing_group.values,
            names=ing_group.index,
            hole=0.4,
            color_discrete_sequence=COLORS_MAIN,
        )
        fig_ing_pie.update_layout(**PLOTLY_THEME, height=320, title="Por fuente")
        st.plotly_chart(fig_ing_pie, use_container_width=True)

    with col_i2:
        est_df = ingresos_df.groupby("Estabilidad")[ing_monto_col].sum().reset_index()
        est_df.columns = ["Estabilidad", "Monto"]
        est_colors = {"Fijo": "#2ecc71", "Variable": "#f39c12", "Ocasional": "#e74c3c", "SIN": "#8892a4"}
        fig_est = px.bar(
            est_df, x="Estabilidad", y="Monto",
            color="Estabilidad",
            color_discrete_map=est_colors,
        )
        fig_est.update_traces(texttemplate="$%{y:,.0f}", textposition="outside")
        fig_est.update_layout(**PLOTLY_THEME, height=320, showlegend=False,
                               title="Por estabilidad",
                               xaxis=dict(showgrid=False),
                               yaxis=dict(showgrid=True, gridcolor="#1e2540", showticklabels=False))
        st.plotly_chart(fig_est, use_container_width=True)

    st.markdown('<div class="section-title">🧠 Análisis de estabilidad</div>', unsafe_allow_html=True)

    ratio_fijo = fijo_ing / total_ing if total_ing else 0
    ratio_var  = var_ing  / total_ing if total_ing else 0
    ratio_ocas = ocas_ing / total_ing if total_ing else 0

    col_x, col_y, col_z = st.columns(3)
    with col_x:
        kpi_card("Fijo", pct(ratio_fijo), sub=fmt(fijo_ing), color="green")
    with col_y:
        kpi_card("Variable", pct(ratio_var), sub=fmt(var_ing), color="yellow")
    with col_z:
        kpi_card("Ocasional", pct(ratio_ocas), sub=fmt(ocas_ing), color="yellow" if ratio_ocas > 0 else "blue")

    if ratio_fijo >= 0.7:
        st.success(f"✅ Base de ingresos sólida: {pct(ratio_fijo)} son fijos. Podés planificar con confianza.")
    elif ratio_fijo >= 0.4:
        st.warning(f"🟡 Ingresos mixtos. Los ingresos variables ({pct(ratio_var)}) suman, pero son impredecibles.")
    else:
        st.error(f"🔴 Alta exposición a variabilidad de ingresos. Tus ingresos fijos son solo {pct(ratio_fijo)}.")

    # Detalle de movimientos
    st.markdown('<div class="section-title">📋 Detalle de ingresos</div>', unsafe_allow_html=True)
    ing_show = ingresos_df[[ing_fecha_col, "Concepto" if "Concepto" in ingresos_df.columns else tipo_ing_col,
                             tipo_ing_col, ing_monto_col, "Estabilidad"]].copy()
    ing_show[ing_monto_col] = ing_show[ing_monto_col].apply(fmt)
    ing_show[ing_fecha_col] = ing_show[ing_fecha_col].astype(str)
    st.dataframe(ing_show, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════
# TAB 4 — EXPORTAR
# ═══════════════════════════════════════════
with tab4:

    st.markdown('<div class="section-title">📤 Exportar reportes</div>', unsafe_allow_html=True)
    st.markdown("Generá tu reporte con los datos del período seleccionado.")

    col_pdf, col_xls = st.columns(2)

    # ─── EXCEL EXPORT ───
    with col_xls:
        st.markdown("### 📊 Excel")
        st.markdown("Incluye: KPIs, tabla de gastos, tabla de ingresos, pareto y evolución mensual.")

        if st.button("⬇️ Generar Excel", use_container_width=True):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:

                # Hoja resumen
                resumen_data = {
                    "Métrica": ["Ingresos", "Gastos", "Ahorro", "Tasa de ahorro", "Score salud",
                                "Costo de vida (NEC+FIJO)", "Gastos discrecionales", "Ingresos fijos"],
                    "Valor": [fmt(total_ing), fmt(total_gas), fmt(ahorro), pct(tasa_ahorro),
                              f"{score}/100", fmt(nec+fijo), fmt(disc), fmt(fijo_ing)],
                }
                pd.DataFrame(resumen_data).to_excel(writer, sheet_name="Resumen", index=False)

                # Hoja gastos
                gas_export = gastos_df[[fecha_col, subcat_col, monto_col, "Naturaleza", "Area", "Controlable"]].copy()
                gas_export.to_excel(writer, sheet_name="Gastos", index=False)

                # Hoja ingresos
                ing_export = ingresos_df[[ing_fecha_col, tipo_ing_col, ing_monto_col, "Estabilidad", "Origen"]].copy()
                ing_export.to_excel(writer, sheet_name="Ingresos", index=False)

                # Hoja pareto
                pareto_df[[subcat_col, monto_col, "%", "% acum"]].to_excel(writer, sheet_name="Pareto", index=False)

                # Hoja evolución
                evo_df.reset_index().to_excel(writer, sheet_name="Evolución", index=False)

            output.seek(0)
            st.download_button(
                label="📥 Descargar Excel",
                data=output,
                file_name=f"finanzas_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    # ─── PDF EXPORT ───
    with col_pdf:
        st.markdown("### 📄 PDF")
        st.markdown("Incluye: KPIs, gauge de salud, gráficos de distribución y evolución mensual.")

        if st.button("⬇️ Generar PDF", use_container_width=True):

            # Build all figures for PDF
            # Gauge
            fig_pdf_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=score,
                number={"suffix": "/100", "font": {"size": 48, "color": "#333"}},
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": "#3b82f6", "thickness": 0.25},
                    "steps": [
                        {"range": [0,30],   "color": "#fee2e2"},
                        {"range": [30,50],  "color": "#fef3c7"},
                        {"range": [50,75],  "color": "#d1fae5"},
                        {"range": [75,100], "color": "#a7f3d0"},
                    ],
                    "threshold": {"line": {"color": "#3b82f6","width":3}, "value": score}
                }
            ))
            fig_pdf_gauge.update_layout(height=300, paper_bgcolor="white", font=dict(color="#333"),
                                         margin=dict(t=30,b=10,l=20,r=20))

            # Pie gastos
            fig_pdf_gas = px.pie(
                values=gasto_group.values, names=gasto_group.index,
                title="Distribución de gastos", hole=0.35,
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig_pdf_gas.update_layout(height=380, paper_bgcolor="white", font=dict(color="#333"),
                                       margin=dict(t=50,b=10,l=10,r=10))

            # Pie naturaleza
            nat_pdf = gastos_df.groupby("Naturaleza")[monto_col].sum().reset_index()
            nat_pdf.columns = ["Naturaleza","Monto"]
            fig_pdf_nat = px.pie(
                nat_pdf, values="Monto", names="Naturaleza",
                title="Composición por naturaleza", hole=0.35,
                color_discrete_sequence=["#3b82f6","#8b5cf6","#10b981","#6b7280"],
            )
            fig_pdf_nat.update_layout(height=380, paper_bgcolor="white", font=dict(color="#333"),
                                       margin=dict(t=50,b=10,l=10,r=10))

            # Pie ingresos
            fig_pdf_ing = px.pie(
                values=ing_group.values, names=ing_group.index,
                title="Distribución de ingresos", hole=0.35,
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            fig_pdf_ing.update_layout(height=380, paper_bgcolor="white", font=dict(color="#333"),
                                       margin=dict(t=50,b=10,l=10,r=10))

            # Evolución
            fig_pdf_evo = go.Figure()
            fig_pdf_evo.add_trace(go.Bar(x=evo_df.index, y=evo_df["Gastos"],
                                          name="Gastos", marker_color="#ef4444", opacity=0.8))
            fig_pdf_evo.add_trace(go.Scatter(x=evo_df.index, y=evo_df["Ingresos"],
                                              name="Ingresos", mode="lines+markers",
                                              line=dict(color="#22c55e", width=3)))
            fig_pdf_evo.add_trace(go.Scatter(x=evo_df.index, y=evo_df["Ahorro"],
                                              name="Ahorro", mode="lines+markers",
                                              line=dict(color="#3b82f6", width=2, dash="dot")))
            fig_pdf_evo.update_layout(height=320, paper_bgcolor="white", font=dict(color="#333"),
                                       title="Evolución mensual",
                                       legend=dict(orientation="h", y=1.1),
                                       margin=dict(t=60,b=30,l=40,r=20))

            # Convert to base64 images
            def fig_to_b64(fig):
                img_bytes = fig.to_image(format="png", width=900, scale=2)
                return base64.b64encode(img_bytes).decode()

            try:
                g_gauge = fig_to_b64(fig_pdf_gauge)
                g_gas   = fig_to_b64(fig_pdf_gas)
                g_nat   = fig_to_b64(fig_pdf_nat)
                g_ing   = fig_to_b64(fig_pdf_ing)
                g_evo   = fig_to_b64(fig_pdf_evo)

                periodo_label = f"{min(meses_sel)} — {max(meses_sel)}" if len(meses_sel) > 1 else meses_sel[0]
                modo_label = modo
                fecha_gen = datetime.today().strftime("%d/%m/%Y")

                pareto_rows = ""
                for _, row in pareto_df.head(10).iterrows():
                    pareto_rows += f"""
                    <tr>
                        <td>{row[subcat_col]}</td>
                        <td style="text-align:right">{fmt(row[monto_col])}</td>
                        <td style="text-align:right">{pct(row['%'])}</td>
                        <td style="text-align:right">{pct(row['% acum'])}</td>
                    </tr>"""

                html_pdf = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #fff; color: #1a1a2e; padding: 0; }}
  
  .cover {{
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    color: white; padding: 60px 50px; min-height: 220px;
    display: flex; flex-direction: column; justify-content: center;
  }}
  .cover h1 {{ font-size: 32px; font-weight: 700; margin-bottom: 8px; }}
  .cover .subtitle {{ font-size: 15px; color: #94a3b8; margin-bottom: 20px; }}
  .cover .meta {{ display: flex; gap: 30px; margin-top: 20px; }}
  .cover .meta-item {{ display: flex; flex-direction: column; }}
  .cover .meta-label {{ font-size: 11px; color: #64748b; text-transform: uppercase; letter-spacing: 1px; }}
  .cover .meta-value {{ font-size: 16px; font-weight: 600; color: #e2e8f0; }}

  .section {{ padding: 30px 40px; }}
  .section-title {{
    font-size: 16px; font-weight: 700; color: #0f3460;
    border-left: 4px solid #3b82f6; padding-left: 12px;
    margin-bottom: 20px; text-transform: uppercase; letter-spacing: 0.5px;
  }}

  .kpi-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 8px; }}
  .kpi {{
    background: linear-gradient(135deg, #f8faff 0%, #eef2ff 100%);
    border-radius: 10px; padding: 16px; border-top: 3px solid #3b82f6;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
  }}
  .kpi.green {{ border-top-color: #22c55e; background: linear-gradient(135deg, #f0fdf4, #dcfce7); }}
  .kpi.red   {{ border-top-color: #ef4444; background: linear-gradient(135deg, #fef2f2, #fee2e2); }}
  .kpi.purple{{ border-top-color: #8b5cf6; background: linear-gradient(135deg, #faf5ff, #ede9fe); }}
  .kpi-label {{ font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #64748b; margin-bottom: 4px; }}
  .kpi-value {{ font-size: 22px; font-weight: 700; color: #1e293b; }}

  .insights {{ background: #f8faff; border-radius: 10px; padding: 20px; margin-top: 10px; }}
  .insight {{ display: flex; align-items: flex-start; gap: 10px; margin-bottom: 10px; font-size: 13px; line-height: 1.5; }}

  .charts-2col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
  .chart-box {{ background: #fafafa; border-radius: 10px; padding: 10px; border: 1px solid #e8ecf0; }}
  .chart-full {{ background: #fafafa; border-radius: 10px; padding: 10px; border: 1px solid #e8ecf0; margin-top: 16px; }}

  img {{ max-width: 100%; height: auto; display: block; }}

  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  thead tr {{ background: #1a1a2e; color: white; }}
  thead th {{ padding: 10px 12px; text-align: left; font-weight: 600; }}
  tbody tr:nth-child(even) {{ background: #f1f5f9; }}
  tbody td {{ padding: 9px 12px; border-bottom: 1px solid #e2e8f0; }}

  .footer {{ background: #f1f5f9; padding: 16px 40px; text-align: center; font-size: 11px; color: #94a3b8; margin-top: 10px; }}
  
  .divider {{ height: 1px; background: linear-gradient(90deg, transparent, #cbd5e1, transparent); margin: 0 40px; }}

  .score-section {{ display: flex; align-items: center; gap: 30px; }}
  .score-text {{ flex: 1; }}
  .score-badge {{
    font-size: 48px; font-weight: 800;
    color: {'#22c55e' if score >= 75 else '#f59e0b' if score >= 50 else '#ef4444'};
    line-height: 1;
  }}
  .score-label {{ font-size: 14px; color: #64748b; margin-top: 4px; }}

  @media print {{ body {{ -webkit-print-color-adjust: exact; }} }}
</style>
</head>
<body>

<!-- PORTADA -->
<div class="cover">
  <h1>💰 Reporte de Finanzas Personales</h1>
  <div class="subtitle">Análisis completo del período seleccionado</div>
  <div class="meta">
    <div class="meta-item">
      <span class="meta-label">Período</span>
      <span class="meta-value">{periodo_label}</span>
    </div>
    <div class="meta-item">
      <span class="meta-label">Modo</span>
      <span class="meta-value">{modo_label}</span>
    </div>
    <div class="meta-item">
      <span class="meta-label">Generado</span>
      <span class="meta-value">{fecha_gen}</span>
    </div>
    <div class="meta-item">
      <span class="meta-label">Meses analizados</span>
      <span class="meta-value">{n_meses}</span>
    </div>
  </div>
</div>

<div class="divider"></div>

<!-- KPIs -->
<div class="section">
  <div class="section-title">📊 Métricas principales</div>
  <div class="kpi-grid">
    <div class="kpi green">
      <div class="kpi-label">Ingresos</div>
      <div class="kpi-value">{fmt(total_ing)}</div>
    </div>
    <div class="kpi red">
      <div class="kpi-label">Gastos</div>
      <div class="kpi-value">{fmt(total_gas)}</div>
    </div>
    <div class="kpi {'green' if ahorro >= 0 else 'red'}">
      <div class="kpi-label">Ahorro</div>
      <div class="kpi-value">{fmt(ahorro)}</div>
    </div>
    <div class="kpi purple">
      <div class="kpi-label">Tasa de ahorro</div>
      <div class="kpi-value">{pct(tasa_ahorro)}</div>
    </div>
  </div>
  <div class="kpi-grid" style="grid-template-columns: repeat(3,1fr); margin-top:14px">
    <div class="kpi">
      <div class="kpi-label">Costo de vida (NEC + FIJO)</div>
      <div class="kpi-value">{fmt(nec+fijo)}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Gastos discrecionales</div>
      <div class="kpi-value">{fmt(disc)}</div>
    </div>
    <div class="kpi green">
      <div class="kpi-label">Ingresos fijos</div>
      <div class="kpi-value">{fmt(fijo_ing)}</div>
    </div>
  </div>
</div>

<div class="divider"></div>

<!-- SALUD FINANCIERA -->
<div class="section">
  <div class="section-title">🎯 Salud financiera</div>
  <div class="score-section">
    <div>
      <img src="data:image/png;base64,{g_gauge}" style="width:400px">
    </div>
    <div class="score-text">
      <div class="score-badge">{score}<span style="font-size:24px">/100</span></div>
      <div class="score-label">{score_text}</div>
      <br>
      <div class="insights">
        {''.join(f'<div class="insight">{icon} <span>{text}</span></div>' for icon, text in insights)}
      </div>
    </div>
  </div>
</div>

<div class="divider"></div>

<!-- GRÁFICOS GASTOS -->
<div class="section">
  <div class="section-title">💸 Análisis de gastos</div>
  <div class="charts-2col">
    <div class="chart-box">
      <img src="data:image/png;base64,{g_gas}">
    </div>
    <div class="chart-box">
      <img src="data:image/png;base64,{g_nat}">
    </div>
  </div>
</div>

<div class="divider"></div>

<!-- GRÁFICO INGRESOS -->
<div class="section">
  <div class="section-title">💵 Análisis de ingresos</div>
  <div class="chart-box" style="max-width:500px">
    <img src="data:image/png;base64,{g_ing}">
  </div>
</div>

<div class="divider"></div>

<!-- EVOLUCIÓN -->
<div class="section">
  <div class="section-title">📈 Evolución mensual</div>
  <div class="chart-full">
    <img src="data:image/png;base64,{g_evo}">
  </div>
</div>

<div class="divider"></div>

<!-- PARETO -->
<div class="section">
  <div class="section-title">🔍 Top 10 subcategorías de gasto</div>
  <table>
    <thead>
      <tr>
        <th>Subcategoría</th>
        <th style="text-align:right">Monto</th>
        <th style="text-align:right">%</th>
        <th style="text-align:right">% Acum.</th>
      </tr>
    </thead>
    <tbody>
      {pareto_rows}
    </tbody>
  </table>
</div>

<div class="footer">
  Finanzas Personales V1.14 — Generado el {fecha_gen} · Período: {periodo_label} · {n_meses} mes(es) analizado(s)
</div>

</body>
</html>"""

                b64_html = base64.b64encode(html_pdf.encode()).decode()
                href = f'<a href="data:text/html;base64,{b64_html}" download="finanzas_{datetime.today().strftime("%Y-%m-%d")}.html" style="display:inline-block;background:#3b82f6;color:white;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:600;width:100%;text-align:center">📥 Descargar Reporte PDF/HTML</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.caption("💡 Abrí el archivo en tu navegador y usá Archivo → Imprimir → Guardar como PDF para obtener el PDF final.")

            except Exception as e:
                st.error(f"Error generando el reporte: {e}")
                st.info("Asegurate de tener `kaleido` instalado: `pip install kaleido`")


# ─────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────
st.markdown("---")
st.markdown(
    """
    <div style="text-align:center; padding: 24px 0 12px 0;">
        <p style="font-size: 0.95rem; color: #8892a4; margin-bottom: 12px;">
            Creado por <span style="color:#e0e6f0; font-weight:600;">Leonardo Sola</span>
        </p>
        <a href="https://github.com/LeoSola12" target="_blank">
            <img src="https://cdn-icons-png.flaticon.com/512/733/733553.png" width="34"
                 style="margin:6px; background:#1e2130; border-radius:8px; padding:5px; opacity:0.8; transition:opacity .2s;">
        </a>
        <a href="https://www.instagram.com/leeeeeeeo_/" target="_blank">
            <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" width="34"
                 style="margin:6px; background:#1e2130; border-radius:8px; padding:5px; opacity:0.8;">
        </a>
        <a href="https://x.com/LeoSola7" target="_blank">
            <img src="https://cdn-icons-png.flaticon.com/512/5968/5968830.png" width="34"
                 style="margin:6px; background:#1e2130; border-radius:8px; padding:5px; opacity:0.8;">
        </a>
        <p style="font-size: 0.75rem; color: #3a4460; margin-top: 14px;">
            Finanzas Personales V1.14
        </p>
    </div>
    """,
    unsafe_allow_html=True
)
