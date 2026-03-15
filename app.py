import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import os, time

# ── CONFIG ──────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Procurement Dashboard", layout="wide", page_icon="📊")

DATA_DIR = Path(".")
COLORS = {
    "primary": "#1B3A5C",
    "accent": "#2E86AB",
    "green": "#28A745",
    "red": "#DC3545",
    "amber": "#FFC107",
    "grey": "#6C757D",
    "light_bg": "#F8F9FA",
}
PALETTE = ["#2E86AB", "#1B3A5C", "#A23B72", "#F18F01", "#28A745", "#6C757D",
           "#DC3545", "#17A2B8", "#FFC107", "#6610F2", "#E83E8C", "#20C997",
           "#FD7E14", "#007BFF"]

MESI_ORD = ["2025-11", "2025-12", "2026-01", "2026-02", "2026-03"]
MESI_LABEL = {"2025-11": "Nov 25", "2025-12": "Dic 25", "2026-01": "Gen 26", "2026-02": "Feb 26", "2026-03": "Mar 26"}

# ── STYLES ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main .block-container { padding-top: 1.5rem; max-width: 1400px; }
    .metric-card {
        background: white; border-radius: 12px; padding: 1.2rem 1.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08); border-left: 4px solid #2E86AB;
    }
    .metric-card h4 { margin: 0; font-size: 0.8rem; color: #6C757D; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }
    .metric-card .value { font-size: 1.8rem; font-weight: 700; color: #1B3A5C; margin: 0.3rem 0; }
    .metric-card .delta { font-size: 0.85rem; }
    .section-title {
        font-size: 1.1rem; font-weight: 600; color: #1B3A5C;
        border-bottom: 2px solid #2E86AB; padding-bottom: 0.5rem; margin: 1.5rem 0 1rem;
    }
    .badge-green { background: #D4EDDA; color: #155724; padding: 3px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 600; }
    .badge-red { background: #F8D7DA; color: #721C24; padding: 3px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 600; }
    .badge-amber { background: #FFF3CD; color: #856404; padding: 3px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 600; }
    .badge-blue { background: #D1ECF1; color: #0C5460; padding: 3px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 600; }
    div[data-testid="stSidebar"] { background: #1B3A5C; }
    div[data-testid="stSidebar"] .stMarkdown h1, div[data-testid="stSidebar"] .stMarkdown p,
    div[data-testid="stSidebar"] label { color: white !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 0; }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 24px; font-weight: 500;
        border-bottom: 3px solid transparent;
    }
    .stTabs [aria-selected="true"] { border-bottom-color: #2E86AB !important; color: #2E86AB !important; }
</style>
""", unsafe_allow_html=True)

# ── DATA LOADING ────────────────────────────────────────────────────────────
@st.cache_data(ttl=30)
def load_data():
    files = sorted(DATA_DIR.glob("*.xlsx"))
    if not files:
        return None, None, None
    fpath = files[-1]
    mod_time = os.path.getmtime(fpath)

    xls = pd.ExcelFile(fpath, engine="openpyxl")
    sheets = xls.sheet_names

    ordini = None
    fatture = None
    if "Dati_Ordini" in sheets:
        ordini = pd.read_excel(xls, "Dati_Ordini")
        ordini["Mese_label"] = ordini["Mese"].map(MESI_LABEL)
    if "Dati_Fatture" in sheets:
        fatture = pd.read_excel(xls, "Dati_Fatture")
        fatture["Mese_label"] = fatture["Mese"].map(MESI_LABEL)
        fatture["Tipo_procedencia"] = fatture["Tipo_procedencia"].fillna("N/D").replace({0: "N/D"})
        fatture["TipoDocumento"] = fatture["TipoDocumento"].fillna("N/D").replace({0: "N/D"})
        fatture["Fornitore"] = fatture["Fornitore"].fillna("N/D").replace({0: "N/D"})
        fatture["Fornitore_Norm"] = fatture["Fornitore_Norm"].fillna("N/D").replace({0: "N/D"})

    return ordini, fatture, fpath.name

ordini, fatture, fname = load_data()

if ordini is None or fatture is None:
    st.error(f"⚠️ Nessun file .xlsx trovato nella cartella `{DATA_DIR}`.\n\nCopia il file Excel nella cartella `data/` e ricarica la pagina.")
    st.stop()

# ── FILTERS ─────────────────────────────────────────────────────────────────
fat_filt = fatture[(fatture["TipoDocumento"].isin(["Fattura", "Nota credito"])) & (fatture["Tipo_procedencia"] == "Fornitore")]
ord_ril = ordini[ordini["Stato"] == "Rilasciato"]

with st.sidebar:
    st.markdown("## 📊 Procurement")
    st.caption(f"File: **{fname}**")
    st.markdown("---")
    centri_all = sorted(set(ord_ril["Centro"].unique()) | set(fat_filt["Centro"].unique()))
    sel_centri = st.multiselect("Centri", centri_all, default=centri_all, key="centri")
    sel_mesi = st.multiselect("Mesi", MESI_ORD, default=MESI_ORD, format_func=lambda x: MESI_LABEL.get(x, x), key="mesi")
    st.markdown("---")
    st.caption("I dati si aggiornano automaticamente quando il file nella cartella cambia.")

fat_f = fat_filt[fat_filt["Centro"].isin(sel_centri) & fat_filt["Mese"].isin(sel_mesi)]
ord_f = ord_ril[ord_ril["Centro"].isin(sel_centri) & ord_ril["Mese"].isin(sel_mesi)]

# ── HELPER ──────────────────────────────────────────────────────────────────
def fmt_eur(v):
    if pd.isna(v) or v == 0: return "€ 0"
    return f"€ {v:,.0f}".replace(",", ".")

def metric_card(title, value, delta=None, delta_color="normal"):
    delta_html = ""
    if delta is not None:
        color = "#28A745" if delta_color == "good" else "#DC3545" if delta_color == "bad" else "#6C757D"
        delta_html = f'<div class="delta" style="color:{color}">{delta}</div>'
    return f'<div class="metric-card"><h4>{title}</h4><div class="value">{value}</div>{delta_html}</div>'

def badge(text, kind="blue"):
    return f'<span class="badge-{kind}">{text}</span>'

# ── KPI HEADER ──────────────────────────────────────────────────────────────
st.markdown("# Procurement Dashboard")
st.caption("Periodo: Nov 2025 – Mar 2026 · Filtri: Fattura/NC + Fornitore | Ordini Rilasciati")

tot_fat = fat_f["Importo"].sum()
tot_ord = ord_f["Importo_Netto"].sum()
n_fat = fat_f["Fattura_Unica"].sum() if "Fattura_Unica" in fat_f.columns else len(fat_f)
n_ord = ord_f["Ordine_Unico"].sum() if "Ordine_Unico" in ord_f.columns else len(ord_f)
copertura = tot_ord / tot_fat if tot_fat != 0 else 0
n_fornitori_ord = ord_f["Fornitore_Norm"].nunique()

c1, c2, c3, c4, c5 = st.columns(5)
with c1: st.markdown(metric_card("Spesa Fatturata", fmt_eur(tot_fat)), unsafe_allow_html=True)
with c2: st.markdown(metric_card("Valore Ordinato", fmt_eur(tot_ord)), unsafe_allow_html=True)
with c3: st.markdown(metric_card("Copertura Navision", f"{copertura:.0%}", f"{'✅ Buona' if copertura > 0.8 else '⚠️ Da migliorare'}", "good" if copertura > 0.8 else "bad"), unsafe_allow_html=True)
with c4: st.markdown(metric_card("N. Fatture", f"{int(n_fat):,}".replace(",", ".")), unsafe_allow_html=True)
with c5: st.markdown(metric_card("Fornitori Attivi", f"{n_fornitori_ord}"), unsafe_allow_html=True)

# ── TABS ────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📈 Spending", "🏥 Navision", "🏭 Fornitori", "📋 Categorie"])

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1: SPENDING
# ═══════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">Spesa Mensile — Fatture vs Ordini</div>', unsafe_allow_html=True)

    # Monthly trend
    fat_mese = fat_f.groupby("Mese")["Importo"].sum().reindex(MESI_ORD).fillna(0)
    ord_mese = ord_f.groupby("Mese")["Importo_Netto"].sum().reindex(MESI_ORD).fillna(0)
    trend_df = pd.DataFrame({"Fatture": fat_mese, "Ordini": ord_mese}).reset_index().rename(columns={"index": "Mese"})
    trend_df["Mese_l"] = trend_df["Mese"].map(MESI_LABEL)

    fig_trend = go.Figure()
    fig_trend.add_trace(go.Bar(x=trend_df["Mese_l"], y=trend_df["Fatture"], name="Fatture", marker_color=COLORS["accent"]))
    fig_trend.add_trace(go.Bar(x=trend_df["Mese_l"], y=trend_df["Ordini"], name="Ordini Rilasciati", marker_color=COLORS["primary"]))
    fig_trend.update_layout(barmode="group", height=380, margin=dict(t=30, b=40), legend=dict(orientation="h", y=1.08),
                            yaxis_tickformat="€,.0f", plot_bgcolor="white", font=dict(family="Inter"))
    fig_trend.update_yaxes(gridcolor="#E9ECEF")
    st.plotly_chart(fig_trend, use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<div class="section-title">Spesa per Centro (Fatture)</div>', unsafe_allow_html=True)
        fat_centro = fat_f.groupby("Centro")["Importo"].sum().sort_values(ascending=True)
        fig_c = px.bar(x=fat_centro.values, y=fat_centro.index, orientation="h", color_discrete_sequence=[COLORS["accent"]])
        fig_c.update_layout(height=400, margin=dict(t=10, b=20, l=0), xaxis_tickformat="€,.0f",
                           showlegend=False, plot_bgcolor="white", font=dict(family="Inter"))
        fig_c.update_xaxes(gridcolor="#E9ECEF")
        st.plotly_chart(fig_c, use_container_width=True)

    with col_b:
        st.markdown('<div class="section-title">Spesa per Categoria Contabile (Top 10)</div>', unsafe_allow_html=True)
        fat_cat = fat_f.groupby("Nome_Conto")["Importo"].sum().sort_values(ascending=False).head(10)
        fat_cat_clean = fat_cat.rename(index=lambda x: x.replace("\xa0", " "))
        fig_cat = px.bar(x=fat_cat_clean.values, y=fat_cat_clean.index, orientation="h",
                        color_discrete_sequence=[COLORS["primary"]])
        fig_cat.update_layout(height=400, margin=dict(t=10, b=20, l=0), xaxis_tickformat="€,.0f",
                             showlegend=False, plot_bgcolor="white", yaxis=dict(autorange="reversed"),
                             font=dict(family="Inter"))
        fig_cat.update_xaxes(gridcolor="#E9ECEF")
        st.plotly_chart(fig_cat, use_container_width=True)

    # Spending heatmap: centro x mese
    st.markdown('<div class="section-title">Heatmap Spesa — Centro × Mese</div>', unsafe_allow_html=True)
    pivot_heat = fat_f.pivot_table(index="Centro", columns="Mese", values="Importo", aggfunc="sum").reindex(columns=MESI_ORD).fillna(0)
    pivot_heat.columns = [MESI_LABEL.get(c, c) for c in pivot_heat.columns]
    fig_heat = px.imshow(pivot_heat, text_auto="€,.0f", color_continuous_scale="Blues", aspect="auto")
    fig_heat.update_layout(height=450, margin=dict(t=10, b=20), font=dict(family="Inter"))
    st.plotly_chart(fig_heat, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 2: NAVISION
# ═══════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Copertura Navision — Ordini vs Fatture per Centro</div>', unsafe_allow_html=True)

    nav_fat = fat_f.groupby("Centro")["Importo"].sum().rename("Fatture")
    nav_ord = ord_f.groupby("Centro")["Importo_Netto"].sum().rename("Ordini")
    nav_df = pd.concat([nav_fat, nav_ord], axis=1).fillna(0)
    nav_df["Gap"] = nav_df["Fatture"] - nav_df["Ordini"]
    nav_df["Copertura"] = (nav_df["Ordini"] / nav_df["Fatture"]).where(nav_df["Fatture"] != 0, 0)
    nav_df["Valutazione"] = nav_df["Copertura"].apply(
        lambda x: "Copertura ottimale" if x >= 0.95 else
                  "Ordini > Fatture" if x > 1.05 else
                  "Sulla strada giusta" if x >= 0.7 else
                  "Da migliorare" if x >= 0.4 else
                  "Critico" if x > 0 else "Non usa Navision"
    )
    nav_df = nav_df.sort_values("Copertura", ascending=False).reset_index()

    # Chart
    fig_nav = go.Figure()
    fig_nav.add_trace(go.Bar(x=nav_df["Centro"], y=nav_df["Fatture"], name="Fatture", marker_color=COLORS["accent"]))
    fig_nav.add_trace(go.Bar(x=nav_df["Centro"], y=nav_df["Ordini"], name="Ordini", marker_color=COLORS["primary"]))
    fig_nav.add_trace(go.Scatter(x=nav_df["Centro"], y=nav_df["Copertura"] * nav_df["Fatture"].max(),
                                  name="% Copertura", yaxis="y2", mode="lines+markers+text",
                                  text=[f"{v:.0%}" for v in nav_df["Copertura"]],
                                  textposition="top center", line=dict(color=COLORS["green"], width=2.5),
                                  marker=dict(size=8)))
    fig_nav.update_layout(barmode="group", height=420, margin=dict(t=30, b=60),
                          yaxis=dict(tickformat="€,.0f", gridcolor="#E9ECEF"),
                          yaxis2=dict(overlaying="y", side="right", showgrid=False, range=[0, nav_df["Fatture"].max() * 1.3]),
                          legend=dict(orientation="h", y=1.08), plot_bgcolor="white", font=dict(family="Inter"))
    st.plotly_chart(fig_nav, use_container_width=True)

    # Table
    nav_display = nav_df.copy()
    nav_display["Fatture"] = nav_display["Fatture"].apply(fmt_eur)
    nav_display["Ordini"] = nav_display["Ordini"].apply(fmt_eur)
    nav_display["Gap"] = nav_display["Gap"].apply(fmt_eur)
    nav_display["Copertura"] = nav_display["Copertura"].apply(lambda x: f"{x:.1%}")
    st.dataframe(nav_display[["Centro", "Fatture", "Ordini", "Gap", "Copertura", "Valutazione"]],
                 use_container_width=True, hide_index=True)

    # Evolution
    st.markdown('<div class="section-title">Evoluzione Mensile N° Ordini per Centro</div>', unsafe_allow_html=True)
    ord_evo = ord_f.groupby(["Mese", "Centro"]).size().reset_index(name="N_Ordini")
    ord_evo["Mese_l"] = ord_evo["Mese"].map(MESI_LABEL)
    fig_evo = px.line(ord_evo, x="Mese_l", y="N_Ordini", color="Centro", markers=True,
                      color_discrete_sequence=PALETTE)
    fig_evo.update_layout(height=380, margin=dict(t=10, b=30), plot_bgcolor="white",
                          legend=dict(orientation="h", y=-0.2), font=dict(family="Inter"))
    fig_evo.update_yaxes(gridcolor="#E9ECEF")
    st.plotly_chart(fig_evo, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 3: FORNITORI
# ═══════════════════════════════════════════════════════════════════════════
with tab3:
    col_f1, col_f2 = st.columns(2)

    with col_f1:
        st.markdown('<div class="section-title">Top 15 Fornitori — Ordini Rilasciati</div>', unsafe_allow_html=True)
        top_forn_ord = ord_f.groupby("Fornitore_Norm")["Importo_Netto"].sum().sort_values(ascending=False).head(15)
        fig_fo = px.bar(y=top_forn_ord.index, x=top_forn_ord.values, orientation="h",
                       color_discrete_sequence=[COLORS["primary"]])
        fig_fo.update_layout(height=500, margin=dict(t=10, b=20, l=0), xaxis_tickformat="€,.0f",
                            showlegend=False, plot_bgcolor="white", font=dict(family="Inter"))
        fig_fo.update_xaxes(gridcolor="#E9ECEF")
        st.plotly_chart(fig_fo, use_container_width=True)

    with col_f2:
        st.markdown('<div class="section-title">Top 15 Fornitori — Fatture</div>', unsafe_allow_html=True)
        top_forn_fat = fat_f[fat_f["Fornitore_Norm"] != "N/D"].groupby("Fornitore_Norm")["Importo"].sum().sort_values(ascending=False).head(15)
        fig_ff = px.bar(y=top_forn_fat.index, x=top_forn_fat.values, orientation="h",
                       color_discrete_sequence=[COLORS["accent"]])
        fig_ff.update_layout(height=500, margin=dict(t=10, b=20, l=0), xaxis_tickformat="€,.0f",
                            showlegend=False, plot_bgcolor="white", font=dict(family="Inter"))
        fig_ff.update_xaxes(gridcolor="#E9ECEF")
        st.plotly_chart(fig_ff, use_container_width=True)

    # Concentrazione
    st.markdown('<div class="section-title">Concentrazione Fornitori (Ordini) — Curva di Pareto</div>', unsafe_allow_html=True)
    forn_tot = ord_f.groupby("Fornitore_Norm")["Importo_Netto"].sum().sort_values(ascending=False)
    forn_cum = forn_tot.cumsum() / forn_tot.sum()
    pareto_df = pd.DataFrame({"Fornitore": range(1, len(forn_cum) + 1), "% Cumulata": forn_cum.values})

    fig_pareto = go.Figure()
    fig_pareto.add_trace(go.Scatter(x=pareto_df["Fornitore"], y=pareto_df["% Cumulata"],
                                     fill="tozeroy", fillcolor="rgba(46,134,171,0.15)",
                                     line=dict(color=COLORS["accent"], width=2.5), name="% cumulata"))
    fig_pareto.add_hline(y=0.80, line_dash="dash", line_color=COLORS["red"], annotation_text="80%")
    n80 = (forn_cum <= 0.80).sum()
    fig_pareto.add_vline(x=n80, line_dash="dash", line_color=COLORS["grey"],
                         annotation_text=f"{n80} fornitori = 80% spesa")
    fig_pareto.update_layout(height=350, margin=dict(t=30, b=30), xaxis_title="N° Fornitori (ordinati per spesa)",
                             yaxis_tickformat=".0%", plot_bgcolor="white", font=dict(family="Inter"))
    fig_pareto.update_yaxes(gridcolor="#E9ECEF")
    st.plotly_chart(fig_pareto, use_container_width=True)

    # Cross-centro
    st.markdown('<div class="section-title">Fornitori Cross-Centro (presenti in più cliniche)</div>', unsafe_allow_html=True)
    forn_centri = ord_f.groupby("Fornitore_Norm")["Centro"].nunique().rename("N_Centri")
    forn_spend = ord_f.groupby("Fornitore_Norm")["Importo_Netto"].sum().rename("Spesa")
    cross = pd.concat([forn_centri, forn_spend], axis=1).sort_values("N_Centri", ascending=False)
    cross = cross[cross["N_Centri"] >= 3].head(20).reset_index()
    cross["Spesa_fmt"] = cross["Spesa"].apply(fmt_eur)
    st.dataframe(cross[["Fornitore_Norm", "N_Centri", "Spesa_fmt"]].rename(
        columns={"Fornitore_Norm": "Fornitore", "N_Centri": "N° Centri", "Spesa_fmt": "Spesa Ordini"}),
        use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 4: CATEGORIE
# ═══════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">Spesa per Categoria Contabile e Centro (Fatture)</div>', unsafe_allow_html=True)

    # Top categories treemap
    cat_df = fat_f.groupby("Nome_Conto")["Importo"].sum().sort_values(ascending=False).head(15).reset_index()
    cat_df["Nome_Conto"] = cat_df["Nome_Conto"].str.replace("\xa0", " ")
    fig_tree = px.treemap(cat_df, path=["Nome_Conto"], values="Importo", color="Importo",
                          color_continuous_scale="Blues")
    fig_tree.update_layout(height=400, margin=dict(t=10, b=10), font=dict(family="Inter"))
    fig_tree.update_traces(textinfo="label+value", texttemplate="%{label}<br>€%{value:,.0f}")
    st.plotly_chart(fig_tree, use_container_width=True)

    # Category x Centro heatmap
    st.markdown('<div class="section-title">Heatmap Categoria × Centro</div>', unsafe_allow_html=True)
    top_cats = fat_f.groupby("Nome_Conto")["Importo"].sum().sort_values(ascending=False).head(12).index
    cat_heat = fat_f[fat_f["Nome_Conto"].isin(top_cats)].pivot_table(
        index="Nome_Conto", columns="Centro", values="Importo", aggfunc="sum").fillna(0)
    cat_heat.index = cat_heat.index.map(lambda x: x.replace("\xa0", " "))
    fig_ch = px.imshow(cat_heat, text_auto="€,.0f", color_continuous_scale="YlOrRd", aspect="auto")
    fig_ch.update_layout(height=500, margin=dict(t=10, b=20), font=dict(family="Inter"))
    st.plotly_chart(fig_ch, use_container_width=True)

    # Categoria ordini (Categoria_Reg)
    st.markdown('<div class="section-title">Categorie Registrazione Ordini — Trend Mensile</div>', unsafe_allow_html=True)
    top_cat_ord = ord_f.groupby("Categoria_Reg")["Importo_Netto"].sum().sort_values(ascending=False).head(10).index
    cat_trend = ord_f[ord_f["Categoria_Reg"].isin(top_cat_ord)].groupby(
        ["Mese", "Categoria_Reg"])["Importo_Netto"].sum().reset_index()
    cat_trend["Mese_l"] = cat_trend["Mese"].map(MESI_LABEL)
    fig_ct = px.bar(cat_trend, x="Mese_l", y="Importo_Netto", color="Categoria_Reg",
                    color_discrete_sequence=PALETTE)
    fig_ct.update_layout(height=420, margin=dict(t=10, b=30), barmode="stack",
                         yaxis_tickformat="€,.0f", plot_bgcolor="white",
                         legend=dict(orientation="h", y=-0.25), font=dict(family="Inter"))
    fig_ct.update_yaxes(gridcolor="#E9ECEF")
    st.plotly_chart(fig_ct, use_container_width=True)

# ── FOOTER ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(f"Dashboard generata da `{fname}` · Ultimo aggiornamento: {time.strftime('%d/%m/%Y %H:%M')}")
