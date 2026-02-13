"""
Pharma Launch Forecast – Use Case 3: Rx-to-OTC Switch
=====================================================
Omeprazol/Pantoprazol PPI switch as reference case.
Models dual-channel dynamics: Rx decay + OTC ramp-up + market expansion.
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO

from models.rx_otc_engine import (
    RxOtcParams, AdjacentCategory,
    forecast_rx_otc, calculate_kpis_rx_otc,
)


def show():
    """Render the Rx-to-OTC Switch page."""

    # ─── Custom CSS ─────────────────────────────────────────────────
    st.markdown("""
    <style>
        .block-container { padding-top: 1rem !important; padding-bottom: 0.5rem !important; }
        .kpi-card, .kpi-card-teal, .kpi-card-amber,
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
        div[data-testid="stSidebar"] {
            background-color: #f8fafc;
            min-width: 280px !important; max-width: 320px !important; width: 300px !important;
        }
        div[data-testid="stSidebar"] [data-testid="stExpander"] { margin-bottom: -8px; }
        div[data-testid="stSidebar"] [data-testid="stExpander"] summary { font-size: 15px; font-weight: 600; }
        div[data-testid="stSidebar"] .stSlider label p { font-size: 14px !important; }
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
        cls = {"default": "kpi-card", "teal": "kpi-card-teal", "amber": "kpi-card-amber",
               "green": "kpi-card-green", "red": "kpi-card-red"}.get(typ, "kpi-card")
        s = f'<div class="kpi-sublabel">{sub}</div>' if sub else ""
        return f'<div class="{cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div>{s}</div>'

    # ─── Sidebar ────────────────────────────────────────────────────
    with st.sidebar:
        st.title("Rx-to-OTC Switch")
        st.caption("PPI (Omeprazol/Pantoprazol)")

        perspective = st.radio(
            "Perspektive",
            ["\U0001f48a Hersteller (OTC-Launch)", "\U0001f3e5 Gesamtmarkt"],
            index=0, horizontal=True,
        )
        is_manufacturer = "Hersteller" in perspective

        scenario = st.selectbox(
            "Szenario", ["Base Case", "Optimistisch", "Konservativ"],
            index=0, label_visibility="collapsed",
        )

        defaults = {
            "Base Case":     {"otc_peak": 280_000, "otc_price": 7.99, "ramp": 18, "mktg": 300_000, "rx_decline": 0.15, "new_pat": 0.70},
            "Optimistisch":  {"otc_peak": 400_000, "otc_price": 8.99, "ramp": 12, "mktg": 500_000, "rx_decline": 0.20, "new_pat": 0.75},
            "Konservativ":   {"otc_peak": 180_000, "otc_price": 6.99, "ramp": 24, "mktg": 200_000, "rx_decline": 0.10, "new_pat": 0.60},
        }[scenario]

        with st.expander("Rx-Kanal (vor Switch)", expanded=False):
            rx_packs = st.slider("Rx-Packungen/Monat", 100_000, 600_000, 350_000, 25_000)
            rx_price = st.slider("Rx-Preis (Festbetrag, EUR)", 10.0, 30.0, 16.99, 1.0)
            rx_decline = st.slider("Rx-Abwanderung zu OTC (%)", 5, 30, int(defaults["rx_decline"] * 100), 5) / 100

        with st.expander("OTC-Kanal", expanded=True):
            otc_price = st.slider("OTC-Preis/Packung (EUR)", 4.0, 15.0, defaults["otc_price"], 0.50)
            otc_peak = st.slider("OTC Peak Packungen/Mon.", 50_000, 600_000, defaults["otc_peak"], 25_000)
            otc_ramp = st.slider("Monate bis Peak", 6, 36, defaults["ramp"], 3)
            new_patient = st.slider("Neue Patienten (%)", 30, 90, int(defaults["new_pat"] * 100), 5,
                                    help="Anteil OTC-Volumen von Patienten, die nie Rx hatten") / 100

        with st.expander("Marketing & Awareness", expanded=False):
            marketing = st.slider("Marketing/Monat (EUR)", 50_000, 1_000_000, defaults["mktg"], 50_000)
            awareness_peak = st.slider("Peak Awareness (%)", 30, 90, 65, 5) / 100
            awareness_ramp = st.slider("Awareness-Aufbau (Monate)", 6, 24, 12, 3)

        with st.expander("Saisonalitaet", expanded=False):
            use_seasonality = st.checkbox("Saisonalitaet aktivieren", value=True)
            if use_seasonality:
                st.caption("PPI: Herbst/Winter-Peak (GI-Beschwerden)")
            seasonality = [0.90, 0.85, 0.95, 1.05, 1.05, 1.00,
                           0.95, 0.90, 1.00, 1.10, 1.15, 1.10] if use_seasonality else [1.0] * 12

        with st.expander("Kosten & Margen", expanded=False):
            cogs = st.slider("COGS (%)", 10, 40, 25, 5) / 100
            pharmacy_margin = st.slider("Apothekenmarge (%)", 25, 55, 40, 5) / 100
            distribution = st.slider("Distribution (%)", 3, 15, 8, 1) / 100

        forecast_years = st.slider("Horizont (Jahre)", 1, 7, 5)

    # ─── Build params and run ───────────────────────────────────────
    params = RxOtcParams(
        rx_packs_per_month=rx_packs,
        rx_price_per_pack=rx_price,
        rx_decline_rate=rx_decline,
        otc_price_per_pack=otc_price,
        otc_peak_packs_per_month=otc_peak,
        otc_ramp_months=otc_ramp,
        new_patient_share=new_patient,
        marketing_monthly_eur=marketing,
        awareness_peak=awareness_peak,
        awareness_ramp_months=awareness_ramp,
        seasonality=seasonality,
        cogs_pct=cogs,
        pharmacy_margin_pct=pharmacy_margin,
        distribution_cost_pct=distribution,
        forecast_months=forecast_years * 12,
    )
    adjacent = AdjacentCategory()

    df = forecast_rx_otc(params, adjacent, forecast_months=forecast_years * 12)
    kpis = calculate_kpis_rx_otc(df)

    # ─── Main ───────────────────────────────────────────────────────
    st.title("Rx-to-OTC Switch \u2013 PPI Forecast")
    st.markdown(
        "**Produkt:** Omeprazol/Pantoprazol 20mg (14 St.) | "
        "**Markt:** Sodbrennen & Reflux Deutschland"
    )

    # ─── KPI Row 1 ──────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(kpi("OTC Umsatz Jahr 1", fmt_eur(kpis["year1_otc_revenue"]), "teal",
                         sub="Herstelleranteil"), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi(f"Gesamtumsatz {forecast_years}J", fmt_eur(kpis["total_5y_revenue"]), "teal",
                         sub="Rx + OTC"), unsafe_allow_html=True)
    with c3:
        co = kpis.get("crossover_month")
        co_text = f"Monat {co}" if co else "\u2013"
        st.markdown(kpi("OTC > Rx ab", co_text, "amber" if co else "default"), unsafe_allow_html=True)
    with c4:
        st.markdown(kpi(f"Gewinn {forecast_years}J", fmt_eur(kpis["total_5y_profit"]),
                         "green" if kpis["total_5y_profit"] > 0 else "red"), unsafe_allow_html=True)

    # ─── KPI Row 2 ──────────────────────────────────────────────────
    st.markdown("")
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.markdown(kpi("OTC-Anteil M12", f"{kpis['otc_share_tablets_m12']:.1%}",
                         sub="In Tabletten"), unsafe_allow_html=True)
    with c6:
        st.markdown(kpi("Rx-Rueckgang", f"{kpis['rx_decline_total']:.1%}",
                         sub="Packungen ueber Laufzeit"), unsafe_allow_html=True)
    with c7:
        st.markdown(kpi("Peak Awareness", f"{kpis['peak_awareness']:.0%}"), unsafe_allow_html=True)
    with c8:
        st.markdown(kpi("Kannibalisierung", fmt_eur(kpis["total_adjacent_lost"]), "red",
                         sub="Antazida/H2-Blocker"), unsafe_allow_html=True)

    st.markdown("")

    # ─── Chart Row 1: Dual Channel + Revenue ────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df["date"], y=df["rx_packs"], mode="lines",
            name="Rx-Packungen", line=dict(color="#2563eb", width=3),
            fill="tozeroy", fillcolor="rgba(37,99,235,0.1)",
            hovertemplate="%{y:,.0f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=df["date"], y=df["otc_packs"], mode="lines",
            name="OTC-Packungen", line=dict(color="#0d9488", width=3),
            fill="tozeroy", fillcolor="rgba(13,148,136,0.1)",
            hovertemplate="%{y:,.0f}<extra></extra>",
        ))
        fig.update_layout(
            title="Dual-Kanal: Rx-Abwanderung + OTC-Aufbau",
            yaxis_title="Packungen / Monat",
            template="plotly_white", height=420,
            hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig, width="stretch")

    with col2:
        fig2 = make_subplots(specs=[[{"secondary_y": True}]])
        fig2.add_trace(go.Bar(
            x=df["date"], y=df["rx_revenue"], name="Rx Umsatz",
            marker_color="#2563eb", hovertemplate="\u20ac%{y:,.0f}<extra></extra>",
        ), secondary_y=False)
        fig2.add_trace(go.Bar(
            x=df["date"], y=df["otc_manufacturer_revenue"], name="OTC Umsatz",
            marker_color="#0d9488", hovertemplate="\u20ac%{y:,.0f}<extra></extra>",
        ), secondary_y=False)
        fig2.add_trace(go.Scatter(
            x=df["date"], y=df["cumulative_total_revenue"], mode="lines",
            name="Kumuliert", line=dict(color="#166534", width=3),
            hovertemplate="\u20ac%{y:,.0f}<extra></extra>",
        ), secondary_y=True)
        fig2.update_layout(
            title="Umsatz nach Kanal", barmode="stack",
            template="plotly_white", height=420,
            hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        fig2.update_yaxes(title_text="Monatlich (EUR)", secondary_y=False)
        fig2.update_yaxes(title_text="Kumuliert (EUR)", secondary_y=True)
        st.plotly_chart(fig2, width="stretch")

    # ─── Chart Row 2: Volume Decomposition + Awareness ──────────────
    col3, col4 = st.columns(2)

    with col3:
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(
            x=df["date"], y=df["otc_from_new_patients"], mode="lines",
            name="Neue Patienten", line=dict(width=0.5, color="#0d9488"),
            fill="tozeroy", fillcolor="rgba(13,148,136,0.4)",
            stackgroup="vol",
            hovertemplate="%{y:,.0f}<extra></extra>",
        ))
        fig3.add_trace(go.Scatter(
            x=df["date"], y=df["otc_from_rx_migration"], mode="lines",
            name="Rx-Migration", line=dict(width=0.5, color="#2563eb"),
            fill="tonexty", fillcolor="rgba(37,99,235,0.4)",
            stackgroup="vol",
            hovertemplate="%{y:,.0f}<extra></extra>",
        ))
        fig3.update_layout(
            title="OTC-Volumen: Neue Patienten vs. Rx-Migration",
            yaxis_title="OTC Packungen / Monat",
            template="plotly_white", height=420,
            hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig3, width="stretch")

    with col4:
        fig4 = make_subplots(specs=[[{"secondary_y": True}]])
        fig4.add_trace(go.Scatter(
            x=df["date"], y=df["awareness"], mode="lines",
            name="Awareness", line=dict(color="#d97706", width=3),
            hovertemplate="%{y:.1%}<extra></extra>",
        ), secondary_y=False)
        fig4.add_trace(go.Bar(
            x=df["date"], y=df["marketing_spend"], name="Marketing-Spend",
            marker_color="rgba(217,119,6,0.3)",
            hovertemplate="\u20ac%{y:,.0f}<extra></extra>",
        ), secondary_y=True)
        fig4.update_layout(
            title="Awareness-Aufbau & Marketing",
            template="plotly_white", height=420,
            hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        fig4.update_yaxes(title_text="Awareness", tickformat=".0%", secondary_y=False)
        fig4.update_yaxes(title_text="Marketing (EUR)", secondary_y=True)
        st.plotly_chart(fig4, width="stretch")

    # ─── Chart Row 3: Seasonality + Profitability ───────────────────
    col5, col6 = st.columns(2)

    with col5:
        fig5 = go.Figure()
        fig5.add_trace(go.Bar(
            x=df["date"], y=df["season_factor"], name="Saisonalitaet",
            marker_color=["#0d9488" if s >= 1.0 else "#94a3b8" for s in df["season_factor"]],
            hovertemplate="x%{y:.2f}<extra></extra>",
        ))
        fig5.update_layout(
            title="Saisonalitaet (Nachfragefaktor)",
            yaxis_title="Faktor (1.0 = Durchschnitt)", yaxis_range=[0.7, 1.3],
            template="plotly_white", height=350,
            hovermode="x unified",
        )
        st.plotly_chart(fig5, width="stretch")

    with col6:
        fig6 = go.Figure()
        fig6.add_trace(go.Bar(
            x=df["date"], y=df["total_revenue"], name="Umsatz", marker_color="#86efac",
        ))
        fig6.add_trace(go.Bar(
            x=df["date"], y=[-c for c in df["cogs"]], name="COGS", marker_color="#fca5a5",
        ))
        fig6.add_trace(go.Bar(
            x=df["date"], y=[-m for m in df["marketing_spend"]], name="Marketing", marker_color="#fdba74",
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

    # ─── Adjacent Category Impact ───────────────────────────────────
    if df["adjacent_revenue_lost"].sum() > 0:
        fig_adj = go.Figure()
        fig_adj.add_trace(go.Bar(
            x=df["date"], y=df["adjacent_revenue_lost"], name="Antazida/H2-Blocker Verlust",
            marker_color="#ef4444",
        ))
        fig_adj.update_layout(
            title="Kannibalisierung benachbarter OTC-Kategorien",
            yaxis_title="Umsatzverlust (EUR/Mon.)", template="plotly_white", height=280,
        )
        st.plotly_chart(fig_adj, width="stretch")

    # ─── Detail Table ───────────────────────────────────────────────
    with st.expander("Detaildaten"):
        disp = df[["month", "rx_packs", "rx_revenue", "otc_packs", "otc_price",
                    "otc_manufacturer_revenue", "otc_share_tablets", "awareness",
                    "total_revenue", "operating_profit",
                    "cumulative_total_revenue", "cumulative_profit"]].copy()
        disp.columns = [
            "Monat", "Rx Pack.", "Rx Umsatz", "OTC Pack.", "OTC Preis",
            "OTC Umsatz (Herst.)", "OTC-Anteil", "Awareness",
            "Gesamt-Umsatz", "Oper. Gewinn", "Kum. Umsatz", "Kum. Gewinn",
        ]
        disp["OTC-Anteil"] = disp["OTC-Anteil"].apply(lambda x: f"{x:.1%}")
        disp["Awareness"] = disp["Awareness"].apply(lambda x: f"{x:.1%}")
        disp["OTC Preis"] = disp["OTC Preis"].apply(lambda x: f"\u20ac{x:.2f}")
        for c in ["Rx Pack.", "OTC Pack."]:
            disp[c] = disp[c].apply(lambda x: f"{x:,.0f}")
        for c in ["Rx Umsatz", "OTC Umsatz (Herst.)", "Gesamt-Umsatz", "Oper. Gewinn", "Kum. Umsatz", "Kum. Gewinn"]:
            disp[c] = disp[c].apply(lambda x: f"\u20ac{x:,.0f}")
        st.dataframe(disp, width="stretch", hide_index=True)

    # ─── Export ─────────────────────────────────────────────────────
    st.divider()
    st.subheader("Export")
    col_e1, col_e2, _ = st.columns([1, 1, 2])
    with col_e1:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.drop(columns=["date"]).to_excel(w, sheet_name="Rx-to-OTC Forecast", index=False)
        st.download_button("Excel herunterladen", buf.getvalue(),
                           "rx_otc_switch_forecast.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_e2:
        st.info(f"Szenario: {scenario}")

    # ─── Footer ─────────────────────────────────────────────────────
    st.divider()
    st.caption(
        "Daten sind synthetisch, angelehnt an: Omeprazol/Pantoprazol OTC-Switch Juli 2009, "
        "IQVIA OTC-Markt 2024, Pharma Deutschland Selbstmedikationsmarkt, "
        "BfArM Sachverstaendigenausschuss. Kein Bezug zu vertraulichen Daten."
    )


if __name__ == "__main__":
    show()
