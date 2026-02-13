"""
Pharma Launch Forecast – Use Case 2: GLP-1 Brand Competition
=============================================================
Mounjaro (Tirzepatid, Lilly) vs. Ozempic/Wegovy (Semaglutid, Novo Nordisk)

Two perspectives:
  1. LILLY: Mounjaro as the challenger gaining share
  2. NOVO NORDISK: Defending market leadership with Ozempic/Wegovy
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO

from models.brand_competition_engine import (
    MarketParams, BrandParams, CompetitorParams,
    forecast_brand, calculate_kpis_brand,
)

def show():
    """Render the GLP-1 Brand Competition page."""
        # ─── Custom CSS ─────────────────────────────────────────────────────────
    st.markdown("""
    <style>
        .block-container { padding-top: 1rem !important; padding-bottom: 0.5rem !important; }
        .kpi-card, .kpi-card-orange, .kpi-card-blue,
        .kpi-card-green, .kpi-card-red {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 8px; padding: 10px 8px;
            color: #1e293b;
            text-align: center; margin: 2px 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        .kpi-value { font-size: 18px; font-weight: 700; margin: 2px 0; line-height: 1.2; color: #0f172a; }
        .kpi-label { font-size: 10px; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; }
        .kpi-sublabel { font-size: 9px; color: #94a3b8; margin-top: 1px; }
        .perspective-header {
            font-size: 16px; padding: 10px 15px; border-radius: 8px;
            margin-bottom: 20px; font-weight: 600;
        }
        .lilly-header { background-color: #fef3c7; color: #92400e; border-left: 4px solid #f59e0b; }
        .novo-header { background-color: #dbeafe; color: #1e40af; border-left: 4px solid #2563eb; }
        div[data-testid="stSidebar"] {
            background-color: #f8fafc;
            min-width: 280px !important; max-width: 320px !important; width: 300px !important;
        }
        div[data-testid="stSidebar"] [data-testid="stExpander"] { margin-bottom: -8px; }
        div[data-testid="stSidebar"] [data-testid="stExpander"] summary { font-size: 15px; font-weight: 600; }
        div[data-testid="stSidebar"] .stSlider label p { font-size: 14px !important; }
        div[data-testid="stSidebar"] h1 { font-size: 22px !important; margin-bottom: 0 !important; }
    </style>
    """, unsafe_allow_html=True)


    def fmt_eur(v, d=0):
        if abs(v) >= 1e9: return f"\u20ac{v/1e9:,.{d}f} Mrd."
        if abs(v) >= 1e6: return f"\u20ac{v/1e6:,.{d}f} Mio."
        if abs(v) >= 1e3: return f"\u20ac{v/1e3:,.{d}f}K"
        return f"\u20ac{v:,.{d}f}"

    def fmt_num(v):
        if abs(v) >= 1e6: return f"{v/1e6:,.1f}M"
        if abs(v) >= 1e3: return f"{v/1e3:,.0f}K"
        return f"{v:,.0f}"

    def kpi(label, value, typ="default", sub=""):
        cls = {"default": "kpi-card", "orange": "kpi-card-orange", "blue": "kpi-card-blue",
               "green": "kpi-card-green", "red": "kpi-card-red"}.get(typ, "kpi-card")
        s = f'<div class="kpi-sublabel">{sub}</div>' if sub else ""
        return f'<div class="{cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div>{s}</div>'


    # ─── Sidebar ────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("GLP-1 Forecast")
        st.caption("Mounjaro vs. Ozempic/Wegovy")

        perspective = st.radio(
            "Perspektive",
            ["\U0001f7e1 Lilly (Mounjaro)", "\U0001f535 Novo Nordisk (Ozempic)"],
            index=0, horizontal=True,
        )
        is_lilly = "Lilly" in perspective

        scenario = st.selectbox(
            "Szenario", ["Base Case", "Bull (Lilly)", "Bull (Novo)"],
            index=0, label_visibility="collapsed",
        )

        # ─── Markt ──────────────────────────────────────────────────────
        with st.expander("Markt & Wachstum", expanded=False):
            market_growth = st.slider("Marktwachstum p.a. (%)", 10, 40,
                                      {"Base Case": 25, "Bull (Lilly)": 35, "Bull (Novo)": 30}[scenario], 5) / 100
            forecast_years = st.slider("Horizont (Jahre)", 1, 7, 5)
            t2d_penetration = st.slider("GLP-1-Penetration T2D Ziel (%)", 15, 35,
                                        {"Base Case": 22, "Bull (Lilly)": 28, "Bull (Novo)": 20}[scenario], 1)

        # ─── Indikationen ───────────────────────────────────────────────
        with st.expander("Indikationen", expanded=False):
            obesity_gkv = st.checkbox("Adipositas GKV-Erstattung",
                                      value=(scenario == "Bull (Lilly)"),
                                      help="Gesetzesaenderung: GKV zahlt Adipositas-Therapie")
            if obesity_gkv:
                obesity_gkv_year = st.slider("GKV-Start Jahr", 2027, 2030, 2027)
            else:
                obesity_gkv_year = 2030

            has_cv = st.checkbox("CV-Indikation (Challenger)", value=False,
                                 help="Kardiovaskulaere Risikoreduktion")
            has_mash = st.checkbox("MASH/NASH-Indikation", value=False)

        # ─── Brand-spezifisch ───────────────────────────────────────────
        if is_lilly:
            with st.expander("Mounjaro-Parameter", expanded=True):
                my_peak_t2d = st.slider("Peak Share T2D (%)", 10, 40,
                                        {"Base Case": 25, "Bull (Lilly)": 35, "Bull (Novo)": 15}[scenario], 1)
                my_peak_obesity = st.slider("Peak Share Adipositas (%)", 15, 50,
                                            {"Base Case": 35, "Bull (Lilly)": 45, "Bull (Novo)": 20}[scenario], 5)
                my_months_peak = st.slider("Mon. bis Peak", 18, 48,
                                           {"Base Case": 36, "Bull (Lilly)": 24, "Bull (Novo)": 48}[scenario], 6)
                my_price = st.slider("Preis/Monat (\u20ac)", 200, 500, 350, 25)

            with st.expander("Competitor (Novo)", expanded=False):
                comp_peak_t2d = st.slider("Novo Peak Share T2D (%)", 15, 40, 30, 1)
                comp_price = st.slider("Novo Preis/Monat (\u20ac)", 200, 400, 300, 25)
                comp_supply = st.checkbox("Novo Lieferengpass", value=True)
                comp_supply_end = st.slider("Engpass endet (Mon.)", 0, 24, 12, 3) if comp_supply else 0

            brand = BrandParams(
                name="Mounjaro (Tirzepatid)", company="Eli Lilly",
                current_share_trx=0.08, price_per_month_eur=my_price,
                target_peak_share_t2d=my_peak_t2d / 100,
                target_peak_share_obesity=my_peak_obesity / 100,
                months_to_peak=my_months_peak,
                has_t2d=True, has_obesity=True, has_cv_indication=has_cv, has_mash=has_mash,
                efficacy_score=9.0, evidence_score=7.5,
            )
            competitor = CompetitorParams(
                name="Ozempic / Wegovy", company="Novo Nordisk",
                current_share_trx=0.36, price_per_month_eur=comp_price,
                target_peak_share_t2d=comp_peak_t2d / 100,
                supply_constrained=comp_supply,
                supply_normalization_month=comp_supply_end,
            )
        else:
            with st.expander("Novo Nordisk-Parameter", expanded=True):
                my_peak_t2d = st.slider("Peak Share T2D (%)", 15, 40,
                                        {"Base Case": 30, "Bull (Lilly)": 25, "Bull (Novo)": 38}[scenario], 1)
                my_months_peak = st.slider("Mon. bis Peak", 12, 48, 24, 6)
                my_price = st.slider("Preis/Monat (\u20ac)", 200, 400, 300, 25)
                my_supply = st.checkbox("Lieferengpass aktiv", value=True)
                my_supply_end = st.slider("Engpass endet (Mon.)", 0, 24, 12, 3) if my_supply else 0

            with st.expander("Competitor (Lilly)", expanded=False):
                comp_peak_t2d = st.slider("Lilly Peak Share T2D (%)", 10, 40, 25, 1)
                comp_price = st.slider("Lilly Preis/Monat (\u20ac)", 200, 500, 350, 25)

            brand = BrandParams(
                name="Ozempic / Wegovy (Semaglutid)", company="Novo Nordisk",
                current_share_trx=0.36, price_per_month_eur=my_price,
                target_peak_share_t2d=my_peak_t2d / 100,
                target_peak_share_obesity=0.30,
                months_to_peak=my_months_peak,
                has_t2d=True, has_obesity=True, has_cv_indication=has_cv, has_mash=False,
                supply_constrained=my_supply,
                supply_normalization_month=my_supply_end,
                supply_capacity_monthly_trx=300_000,
                efficacy_score=7.5, evidence_score=9.0,
            )
            competitor = CompetitorParams(
                name="Mounjaro (Tirzepatid)", company="Eli Lilly",
                current_share_trx=0.08, price_per_month_eur=comp_price,
                target_peak_share_t2d=comp_peak_t2d / 100,
                supply_constrained=False,
            )

        with st.expander("Kosten", expanded=False):
            cogs_pct = st.slider("COGS (%)", 10, 40, 20, 5)
            sga = st.number_input("SG&A / Monat (\u20ac)", 200_000, 2_000_000, 800_000, 100_000)

        brand.cogs_pct = cogs_pct / 100
        brand.sga_monthly_eur = sga

        mkt = MarketParams(
            market_growth_annual=market_growth,
            glp1_penetration_target=t2d_penetration / 100,
            obesity_gkv_coverage=obesity_gkv,
            obesity_gkv_start_year=obesity_gkv_year,
        )

    # ─── Main ───────────────────────────────────────────────────────────────
    st.title("GLP-1 Markt \u2013 Brand Competition Forecast")
    st.markdown(
        f"**Markt:** GLP-1-Rezeptor-Agonisten Deutschland | "
        f"**{brand.name}** vs. **{competitor.name}**"
    )

    header_cls = "lilly-header" if is_lilly else "novo-header"
    icon = "\U0001f7e1" if is_lilly else "\U0001f535"
    st.markdown(
        f'<div class="perspective-header {header_cls}">'
        f'{icon} Perspektive: {brand.company} \u2013 {brand.name}</div>',
        unsafe_allow_html=True,
    )

    df = forecast_brand(brand, competitor, mkt, forecast_months=forecast_years * 12)
    kpis = calculate_kpis_brand(df)

    # ─── KPI Row 1 ──────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    color = "orange" if is_lilly else "blue"
    with c1:
        st.markdown(kpi("Umsatz Jahr 1", fmt_eur(kpis["year1_revenue"]), color), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi(f"Umsatz {forecast_years}J", fmt_eur(kpis["total_5y_revenue"]), color), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi("Peak Marktanteil", f"{kpis['peak_share']:.1%}", color), unsafe_allow_html=True)
    with c4:
        ot = kpis.get("overtake_month")
        ot_text = f"Monat {ot}" if ot else "\u2013"
        st.markdown(kpi("Market Leader ab", ot_text, "green" if ot else "default"), unsafe_allow_html=True)

    # ─── KPI Row 2 ──────────────────────────────────────────────────────
    st.markdown("")
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.markdown(kpi("Markt Anfang", fmt_num(kpis["market_size_start"]) + " TRx/M",
                         sub="Monatl. GLP-1 Verordnungen"), unsafe_allow_html=True)
    with c6:
        st.markdown(kpi("Markt Ende", fmt_num(kpis["market_size_end"]) + " TRx/M",
                         sub=f"x{kpis['market_growth_total']:.1f} Wachstum"), unsafe_allow_html=True)
    with c7:
        st.markdown(kpi(f"Gewinn {forecast_years}J", fmt_eur(kpis["total_5y_profit"]),
                         "green" if kpis["total_5y_profit"] > 0 else "red"), unsafe_allow_html=True)
    with c8:
        st.markdown(kpi("Durchschn. Preis", f"\u20ac{kpis['avg_price']:,.0f}/Mon."), unsafe_allow_html=True)

    st.markdown("")

    # ─── Charts Row 1: Market Share Race + Revenue ──────────────────────
    col1, col2 = st.columns(2)

    with col1:
        fig = go.Figure()
        my_color = "#f59e0b" if is_lilly else "#2563eb"
        comp_color = "#2563eb" if is_lilly else "#f59e0b"

        fig.add_trace(go.Scatter(
            x=df["date"], y=df["my_share"], mode="lines",
            name=brand.name.split("(")[0].strip(),
            line=dict(color=my_color, width=3),
            fill="tozeroy", fillcolor=f"rgba({','.join(str(int(my_color.lstrip('#')[i:i+2], 16)) for i in (0,2,4))},0.15)",
            hovertemplate="%{y:.1%}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=df["date"], y=df["competitor_share"], mode="lines",
            name=competitor.name.split("(")[0].strip(),
            line=dict(color=comp_color, width=3, dash="dash"),
            hovertemplate="%{y:.1%}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=df["date"], y=df["rest_share"], mode="lines",
            name="Uebrige (Trulicity etc.)",
            line=dict(color="#94a3b8", width=1, dash="dot"),
            hovertemplate="%{y:.1%}<extra></extra>",
        ))
        fig.update_layout(
            title="Marktanteils-Wettlauf", yaxis_title="Marktanteil",
            yaxis_tickformat=".0%", template="plotly_white", height=420,
            hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig, width="stretch")

    with col2:
        fig2 = make_subplots(specs=[[{"secondary_y": True}]])
        fig2.add_trace(go.Bar(
            x=df["date"], y=df["my_revenue"], name="Mein Umsatz",
            marker_color=my_color, hovertemplate="%{y:\u20ac,.0f}<extra></extra>",
        ), secondary_y=False)
        fig2.add_trace(go.Scatter(
            x=df["date"], y=df["cumulative_revenue"], mode="lines",
            name="Kumuliert", line=dict(color="#166534", width=3),
            hovertemplate="%{y:\u20ac,.0f}<extra></extra>",
        ), secondary_y=True)
        fig2.update_layout(
            title="Umsatz-Aufbau", template="plotly_white", height=420,
            hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        fig2.update_yaxes(title_text="Monatlich (EUR)", secondary_y=False)
        fig2.update_yaxes(title_text="Kumuliert (EUR)", secondary_y=True)
        st.plotly_chart(fig2, width="stretch")

    # ─── Charts Row 2: Market Expansion + Revenue Comparison ────────────
    col3, col4 = st.columns(2)

    with col3:
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(
            x=df["date"], y=df["total_market_trx"], mode="lines",
            name="Gesamtmarkt GLP-1", line=dict(color="#6366f1", width=3),
            fill="tozeroy", fillcolor="rgba(99,102,241,0.1)",
            hovertemplate="%{y:,.0f} TRx<extra></extra>",
        ))
        fig3.add_trace(go.Scatter(
            x=df["date"], y=df["my_actual_trx"], mode="lines",
            name=brand.name.split("(")[0].strip(),
            line=dict(color=my_color, width=2),
            hovertemplate="%{y:,.0f} TRx<extra></extra>",
        ))
        fig3.add_trace(go.Scatter(
            x=df["date"], y=df["competitor_trx"], mode="lines",
            name=competitor.name.split("(")[0].strip(),
            line=dict(color=comp_color, width=2, dash="dash"),
            hovertemplate="%{y:,.0f} TRx<extra></extra>",
        ))
        fig3.update_layout(
            title="Marktexpansion (TRx / Monat)", yaxis_title="TRx",
            template="plotly_white", height=420, hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig3, width="stretch")

    with col4:
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(
            x=df["date"], y=df["my_revenue"], name=brand.company,
            marker_color=my_color, hovertemplate="%{y:\u20ac,.0f}<extra></extra>",
        ))
        fig4.add_trace(go.Bar(
            x=df["date"], y=df["competitor_revenue"], name=competitor.company,
            marker_color=comp_color, hovertemplate="%{y:\u20ac,.0f}<extra></extra>",
        ))
        fig4.update_layout(
            title="Umsatzvergleich (monatlich)", barmode="group",
            yaxis_title="Umsatz (EUR)", template="plotly_white", height=420,
            hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig4, width="stretch")

    # ─── Charts Row 3: Indication Multiplier + Profitability ────────────
    col5, col6 = st.columns(2)

    with col5:
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(
            x=df["date"], y=df["indication_multiplier"], mode="lines",
            name="Indikations-Multiplikator", line=dict(color="#6366f1", width=3),
            fill="tozeroy", fillcolor="rgba(99,102,241,0.1)",
            hovertemplate="x%{y:.2f}<extra></extra>",
        ))
        fig5.update_layout(
            title="Indikations-Hebel (T2D + Adipositas + CV + MASH)",
            yaxis_title="Multiplikator", template="plotly_white", height=350,
            hovermode="x unified",
        )
        st.plotly_chart(fig5, width="stretch")

    with col6:
        fig6 = go.Figure()
        fig6.add_trace(go.Bar(
            x=df["date"], y=df["my_revenue"], name="Umsatz", marker_color="#86efac",
        ))
        fig6.add_trace(go.Bar(
            x=df["date"], y=[-c for c in df["my_cogs"]], name="COGS", marker_color="#fca5a5",
        ))
        fig6.add_trace(go.Scatter(
            x=df["date"], y=df["cumulative_profit"], mode="lines",
            name="Kum. Gewinn", line=dict(color="#166534", width=3), yaxis="y2",
        ))
        fig6.update_layout(
            title="Profitabilitaet", barmode="relative",
            yaxis=dict(title="Monatlich (EUR)"),
            yaxis2=dict(title="Kumuliert (EUR)", overlaying="y", side="right"),
            template="plotly_white", height=350, hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig6, width="stretch")

    # ─── Supply Gap (if applicable) ─────────────────────────────────────
    if df["supply_gap_trx"].sum() > 0:
        fig_sg = go.Figure()
        fig_sg.add_trace(go.Bar(
            x=df["date"], y=df["supply_gap_trx"], name="Ungedeckte Nachfrage",
            marker_color="#ef4444",
        ))
        fig_sg.update_layout(
            title="Supply Gap: Nachfrage vs. Lieferkapazitaet",
            yaxis_title="Fehlende TRx", template="plotly_white", height=280,
        )
        st.plotly_chart(fig_sg, width="stretch")

    # ─── Detail Table ───────────────────────────────────────────────────
    with st.expander("Detaildaten"):
        disp = df[["month", "total_market_trx", "my_share", "my_actual_trx",
                   "my_price", "my_revenue", "competitor_share", "competitor_trx",
                   "indication_multiplier", "my_operating_profit",
                   "cumulative_revenue", "cumulative_profit"]].copy()
        disp.columns = [
            "Monat", "Markt TRx", "Mein Anteil", "Meine TRx",
            "Preis/Mon.", "Umsatz", "Competitor Anteil", "Competitor TRx",
            "Ind.-Multipl.", "Oper. Gewinn", "Kum. Umsatz", "Kum. Gewinn",
        ]
        disp["Mein Anteil"] = disp["Mein Anteil"].apply(lambda x: f"{x:.1%}")
        disp["Competitor Anteil"] = disp["Competitor Anteil"].apply(lambda x: f"{x:.1%}")
        disp["Ind.-Multipl."] = disp["Ind.-Multipl."].apply(lambda x: f"x{x:.2f}")
        disp["Preis/Mon."] = disp["Preis/Mon."].apply(lambda x: f"\u20ac{x:,.0f}")
        for c in ["Markt TRx", "Meine TRx", "Competitor TRx"]:
            disp[c] = disp[c].apply(lambda x: f"{x:,.0f}")
        for c in ["Umsatz", "Oper. Gewinn", "Kum. Umsatz", "Kum. Gewinn"]:
            disp[c] = disp[c].apply(lambda x: f"\u20ac{x:,.0f}")
        st.dataframe(disp, width="stretch", hide_index=True)

    # ─── Export ─────────────────────────────────────────────────────────
    st.divider()
    st.subheader("Export")
    col_e1, col_e2, _ = st.columns([1, 1, 2])
    with col_e1:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.drop(columns=["date"]).to_excel(w, sheet_name=brand.company, index=False)
        st.download_button("Excel herunterladen", buf.getvalue(),
                           "glp1_brand_forecast.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_e2:
        st.info(f"Perspektive: {brand.company} | Szenario: {scenario}")

    # ─── Footer ─────────────────────────────────────────────────────────
    st.divider()
    st.caption(
        "Daten sind synthetisch, basierend auf oeffentlichen Quellen: "
        "BARMER Gesundheitswesen aktuell 2025, G-BA Nutzenbewertung Tirzepatid, "
        "IQVIA Pharmamarkt 2024, EMA EPAR Mounjaro. Kein Bezug zu vertraulichen Daten."
    )


if __name__ == "__main__":
    show()
