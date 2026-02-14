"""
Pharma Launch Forecast – Use Case 4: Sildenafil Rx-to-OTC Switch (Viatris/Viagra)
==================================================================================
Advanced Rx-to-OTC model with pharmacy-only distribution.
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from models.sildenafil_otc_engine import (
    SildenafilOtcParams, ChannelParams,
    forecast_sildenafil_otc, calculate_kpis_sildenafil,
)


BLUE = "#1e3a5f"
TEAL = "#0d9488"
AMBER = "#d97706"
GREEN = "#27ae60"
RED = "#c0392b"
PURPLE = "#7c3aed"
PINK = "#db2777"


def show():
    """Render the Sildenafil Rx-to-OTC Switch page."""

    st.markdown("""
    <style>
        .block-container { padding-top: 1rem !important; padding-bottom: 0.5rem !important; }
        .kpi-card, .kpi-card-teal, .kpi-card-amber,
        .kpi-card-green, .kpi-card-red, .kpi-card-purple {
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
        div[data-testid="stSidebar"] .stMarkdown p { font-size: 13px; }
        div[data-testid="stSidebar"] label { font-size: 13px !important; }
        .plot-container { margin-top: -10px; }

        /* Tab styling */
        .stTabs [data-baseweb="tab-list"] {
            background: #f8fafc;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            padding: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
            gap: 2px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: 500;
            color: #64748b;
        }
        .stTabs [aria-selected="true"] {
            background: #ffffff !important;
            border: 1px solid #e5e7eb !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.06);
            color: #0f172a !important;
            font-weight: 600;
        }
    </style>
    """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SIDEBAR
    # ═══════════════════════════════════════════════════════════════════
    with st.sidebar:
        st.markdown(f"#### Sildenafil Rx-to-OTC Switch")
        st.markdown("*Viatris / Viagra Connect (DE)*")

        scenario = st.selectbox("Szenario", [
            "Base Case", "Optimistisch (BMG erzwingt Switch)",
            "Konservativ (SVA-Auflagen)"
        ], key="sil_scenario")

        # ─── Scenario-driven slider defaults ──────────────────────
        # The scenario sets the DEFAULT values for sliders.  When the user
        # switches scenario the sliders jump to the new defaults.  The user
        # can still fine-tune every slider afterwards.
        # We detect scenario changes and clear cached slider values so the
        # new defaults actually take effect (Streamlit ignores `value=` when
        # a key already exists in session_state).
        if scenario == "Optimistisch (BMG erzwingt Switch)":
            _d = dict(otc_peak=450_000, otc_ramp=12, mktg=750_000,
                      new_patient=70, rx_decline=5, brand_share=50)
        elif scenario == "Konservativ (SVA-Auflagen)":
            _d = dict(otc_peak=200_000, otc_ramp=24, mktg=300_000,
                      new_patient=55, rx_decline=12, brand_share=35)
        else:  # Base Case
            _d = dict(otc_peak=350_000, otc_ramp=18, mktg=500_000,
                      new_patient=63, rx_decline=8, brand_share=45)

        # Reset scenario-dependent slider keys when scenario changes.
        # We pop the cached values and rerun so Streamlit picks up the
        # new defaults on the NEXT pass (pop + widget in same pass = ignored).
        _prev = st.session_state.get("_sil_prev_scenario")
        if _prev is not None and _prev != scenario:
            for _k in ["sil_rx_dec", "sil_otc_peak", "sil_otc_ramp",
                        "sil_np", "sil_brand", "sil_mktg"]:
                st.session_state.pop(_k, None)
            st.session_state["_sil_prev_scenario"] = scenario
            st.rerun()
        st.session_state["_sil_prev_scenario"] = scenario

        st.markdown("---")

        with st.expander("Rx-Markt", expanded=False):
            rx_packs = st.number_input("Sildenafil Rx Pack./Mon.", 50_000, 500_000, 217_000, 5_000, key="sil_rx_packs")
            rx_price_brand = st.number_input("Viagra Preis/Tabl. (EUR)", 5.0, 25.0, 11.19, 0.5, key="sil_rx_brand")
            rx_price_generic = st.number_input("Generika Preis/Tabl. (EUR)", 0.50, 5.0, 1.50, 0.1, key="sil_rx_gen")
            rx_brand_share = st.slider("Viagra Markenanteil Rx (%)", 2, 25, 10, key="sil_rx_bs") / 100
            rx_decline = st.slider("Rx-Rueckgang durch OTC (%)", 0, 30, _d["rx_decline"], key="sil_rx_dec") / 100

        with st.expander("OTC-Kanal", expanded=False):
            otc_price = st.number_input("OTC Preis/Tablette (EUR)", 2.0, 15.0, 5.99, 0.5, key="sil_otc_p")
            otc_peak = st.number_input("OTC Peak Pack./Mon.", 100_000, 800_000, _d["otc_peak"], 10_000, key="sil_otc_peak")
            otc_ramp = st.slider("Monate bis Peak", 6, 36, _d["otc_ramp"], key="sil_otc_ramp")
            new_patient = st.slider("Neue Patienten (% OTC-Vol.)", 30, 80, _d["new_patient"], key="sil_np") / 100

        with st.expander("Apothekenverteilung", expanded=False):
            st.caption("Apothekenpflichtig: nur Apotheken-Kanaele")
            ch_apo = st.slider("Stationaere Apotheke (%)", 20, 80, 55, key="sil_ch_apo")
            ch_online = 100 - ch_apo
            st.markdown(f"*Online-Apotheke: {ch_online}%*")
            online_growth = st.slider("Online-Wachstum p.a. (Pp.)", 0, 10, 3, key="sil_og") / 100

        with st.expander("Marke vs. Generika OTC", expanded=False):
            brand_share = st.slider("Viagra Connect Anteil OTC (%)", 15, 70, _d["brand_share"], key="sil_brand") / 100
            brand_erosion = st.slider("Markenanteil-Erosion p.a. (Pp.)", 0, 10, 3, key="sil_berosion") / 100
            brand_premium = st.slider("Preispremium Marke (x)", 1.0, 3.0, 1.8, 0.1, key="sil_bprem")

        with st.expander("Marketing & Awareness", expanded=False):
            mktg = st.number_input("Marketing/Mon. (EUR)", 100_000, 2_000_000, _d["mktg"], 50_000, key="sil_mktg")
            aw_peak = st.slider("Peak Awareness (%)", 50, 95, 75, key="sil_awp") / 100
            aw_base = st.slider("Baseline Awareness (%)", 20, 70, 60, key="sil_awb") / 100
            aw_ramp = st.slider("Awareness Ramp (Mon.)", 3, 18, 9, key="sil_awr")

        with st.expander("Tadalafil-Migration", expanded=False):
            tada_monthly = st.number_input("Tadalafil Rx Pack./Mon.", 50_000, 300_000, 120_000, 5_000, key="sil_tada")
            tada_switch = st.slider("Migration zu Sildenafil OTC (%)", 0, 30, 12, key="sil_tadas") / 100

    # ─── Build params ──────────────────────────────────────────────
    # Scenario sets slider defaults; all values come from sliders.
    # User can fine-tune after selecting a scenario.

    channels = [
        ChannelParams(name="Stationaere Apotheke", share_of_otc=ch_apo / 100,
                      share_trend_annual=-online_growth, margin_pct=0.42,
                      distribution_cost_pct=0.06, discretion_factor=0.70),
        ChannelParams(name="Online-Apotheke", share_of_otc=ch_online / 100,
                      share_trend_annual=online_growth, margin_pct=0.30,
                      distribution_cost_pct=0.10, discretion_factor=1.0),
    ]

    params = SildenafilOtcParams(
        rx_packs_per_month=rx_packs,
        rx_price_brand=rx_price_brand,
        rx_price_generic=rx_price_generic,
        rx_brand_share=rx_brand_share,
        rx_decline_rate=rx_decline,
        otc_price_per_tablet=otc_price,
        otc_peak_packs_per_month=otc_peak,
        otc_ramp_months=otc_ramp,
        new_patient_share=new_patient,
        brand_otc_share=brand_share,
        brand_otc_share_trend=-brand_erosion,
        brand_price_premium=brand_premium,
        channels=channels,
        marketing_monthly_eur=mktg,
        awareness_peak=aw_peak,
        awareness_baseline=aw_base,
        awareness_ramp_months=aw_ramp,
        tadalafil_rx_monthly=tada_monthly,
        tadalafil_switch_to_sildenafil_otc=tada_switch,
    )

    try:
        df = forecast_sildenafil_otc(params)
        kpis = calculate_kpis_sildenafil(df)
    except Exception as e:
        import traceback
        st.error(f"Engine error: {type(e).__name__}: {e}")
        st.code(traceback.format_exc())
        st.json({
            "channels_count": len(params.channels),
            "channel_names": [c.name for c in params.channels],
            "forecast_months": params.forecast_months,
        })
        st.stop()

    # ═══════════════════════════════════════════════════════════════════
    # HEADER
    # ═══════════════════════════════════════════════════════════════════
    st.markdown(f"## Sildenafil Rx-to-OTC Switch Forecast")
    st.markdown(f"*Viatris / Viagra Connect – Szenario: **{scenario}** | "
                f"Dual-Channel Apotheken-Modell (apothekenpflichtig)*")

    # ═══════════════════════════════════════════════════════════════════
    # KPI CARDS (Row 1)
    # ═══════════════════════════════════════════════════════════════════
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="kpi-card-teal">
            <div class="kpi-label">OTC Umsatz Jahr 1</div>
            <div class="kpi-value">EUR {kpis['year1_otc_revenue']/1e6:.1f}M</div>
            <div class="kpi-sublabel">Herstellerumsatz</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">Gesamtumsatz 5 Jahre</div>
            <div class="kpi-value">EUR {kpis['total_5y_revenue']/1e6:.0f}M</div>
            <div class="kpi-sublabel">Rx + OTC kumuliert</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        co = kpis.get("crossover_month")
        co_text = f"Monat {co}" if co else "–"
        st.markdown(f"""<div class="kpi-card-amber">
            <div class="kpi-label">OTC > Rx ab</div>
            <div class="kpi-value">{co_text}</div>
            <div class="kpi-sublabel">Umsatz-Crossover</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="kpi-card-green">
            <div class="kpi-label">Gewinn 5 Jahre</div>
            <div class="kpi-value">EUR {kpis['total_5y_profit']/1e6:.0f}M</div>
            <div class="kpi-sublabel">nach Marketing & COGS</div>
        </div>""", unsafe_allow_html=True)

    # KPI Row 2
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.markdown(f"""<div class="kpi-card-purple">
            <div class="kpi-label">Online-Anteil M24</div>
            <div class="kpi-value">{kpis['online_share_m24']:.0%}</div>
            <div class="kpi-sublabel">Online-Apotheke</div>
        </div>""", unsafe_allow_html=True)
    with c6:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">Viagra Markenanteil M12</div>
            <div class="kpi-value">{kpis['brand_share_m12']:.0%}</div>
            <div class="kpi-sublabel">vs. Generika OTC</div>
        </div>""", unsafe_allow_html=True)
    with c7:
        st.markdown(f"""<div class="kpi-card-red">
            <div class="kpi-label">Rx-Rueckgang gesamt</div>
            <div class="kpi-value">{kpis['rx_decline_total']:.0%}</div>
            <div class="kpi-sublabel">ueber 5 Jahre</div>
        </div>""", unsafe_allow_html=True)
    with c8:
        st.markdown(f"""<div class="kpi-card-green">
            <div class="kpi-label">Therapiequote Neu</div>
            <div class="kpi-value">{kpis['treatment_rate_final']:.0%}</div>
            <div class="kpi-sublabel">vorher: 33%</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ═══════════════════════════════════════════════════════════════════
    # CHARTS
    # ═══════════════════════════════════════════════════════════════════

    # ─── Chart 1: Dual Channel (Rx vs OTC packs) ──────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Dual Channel", "Apothekenverteilung", "Brand vs. Generika",
        "Treatment Gap", "Profitabilitaet"
    ])

    with tab1:
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(
            x=df["month"], y=df["rx_packs"], name="Rx Packungen",
            line=dict(color=BLUE, width=2.5),
        ))
        fig1.add_trace(go.Scatter(
            x=df["month"], y=df["otc_packs"], name="OTC Packungen",
            line=dict(color=TEAL, width=2.5),
        ))
        if kpis["crossover_month"]:
            fig1.add_vline(x=kpis["crossover_month"], line_dash="dot",
                           line_color=AMBER, annotation_text="OTC > Rx")
        fig1.update_layout(
            title="Dual-Kanal: Rx vs. OTC Packungen/Monat",
            xaxis_title="Monate nach Switch", yaxis_title="Packungen",
            yaxis_tickformat=",", height=420,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig1, width="stretch")

        # Stacked revenue
        fig1b = go.Figure()
        fig1b.add_trace(go.Bar(
            x=df["month"], y=df["rx_revenue"], name="Rx Umsatz",
            marker_color="#93c5fd",
        ))
        fig1b.add_trace(go.Bar(
            x=df["month"], y=df["otc_manufacturer_revenue"], name="OTC Umsatz (Hersteller)",
            marker_color="#5eead4",
        ))
        fig1b.update_layout(
            barmode="stack",
            title="Monatlicher Umsatz nach Kanal",
            xaxis_title="Monate nach Switch", yaxis_title="EUR",
            height=380,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig1b, width="stretch")

    # ─── Chart 2: Omnichannel ─────────────────────────────────────
    with tab2:
        col_a, col_b = st.columns(2)

        with col_a:
            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=df["month"], y=df["ch_apotheke_packs"],
                name="Stationaere Apotheke", stackgroup="one",
                line=dict(color=BLUE),
            ))
            fig2.add_trace(go.Scatter(
                x=df["month"], y=df["ch_online_packs"],
                name="Online-Apotheke", stackgroup="one",
                line=dict(color=TEAL),
            ))
            fig2.update_layout(
                title="OTC-Volumen nach Vertriebskanal",
                xaxis_title="Monate", yaxis_title="Packungen",
                yaxis_tickformat=",", height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig2, width="stretch")

        with col_b:
            fig2b = go.Figure()
            fig2b.add_trace(go.Scatter(
                x=df["month"], y=df["ch_apotheke_share"],
                name="Stationaere Apotheke",
                line=dict(color=BLUE, width=2),
            ))
            fig2b.add_trace(go.Scatter(
                x=df["month"], y=df["ch_online_share"],
                name="Online-Apotheke",
                line=dict(color=TEAL, width=2),
            ))
            fig2b.update_layout(
                title="Kanalanteil-Entwicklung",
                xaxis_title="Monate", yaxis_title="Anteil",
                yaxis_tickformat=".0%", height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig2b, width="stretch")

        # Revenue by channel
        fig2c = go.Figure()
        fig2c.add_trace(go.Bar(
            x=df["month"], y=df["ch_apotheke_revenue"],
            name="Apotheke", marker_color=BLUE,
        ))
        fig2c.add_trace(go.Bar(
            x=df["month"], y=df["ch_online_revenue"],
            name="Online", marker_color=TEAL,
        ))
        fig2c.update_layout(
            barmode="stack",
            title="OTC-Herstellerumsatz nach Kanal",
            xaxis_title="Monate", yaxis_title="EUR",
            height=380,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig2c, width="stretch")

    # ─── Chart 3: Brand vs Generic ────────────────────────────────
    with tab3:
        col_c, col_d = st.columns(2)

        with col_c:
            fig3 = go.Figure()
            fig3.add_trace(go.Scatter(
                x=df["month"], y=df["otc_brand_packs"],
                name="Viagra Connect (Marke)", stackgroup="one",
                line=dict(color="#1e40af"),
            ))
            fig3.add_trace(go.Scatter(
                x=df["month"], y=df["otc_generic_packs"],
                name="Generika OTC", stackgroup="one",
                line=dict(color="#94a3b8"),
            ))
            fig3.update_layout(
                title="OTC-Volumen: Marke vs. Generika",
                xaxis_title="Monate", yaxis_title="Packungen",
                yaxis_tickformat=",", height=400,
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig3, width="stretch")

        with col_d:
            fig3b = go.Figure()
            fig3b.add_trace(go.Scatter(
                x=df["month"], y=df["otc_brand_share"],
                name="Viagra Connect Anteil",
                line=dict(color="#1e40af", width=3),
                fill="tozeroy", fillcolor="rgba(30,64,175,0.1)",
            ))
            fig3b.update_layout(
                title="Viagra Markenanteil-Erosion",
                xaxis_title="Monate", yaxis_title="Anteil",
                yaxis_tickformat=".0%", height=400,
                yaxis_range=[0, 0.7],
            )
            st.plotly_chart(fig3b, width="stretch")

        # Volume decomposition
        fig3c = go.Figure()
        fig3c.add_trace(go.Scatter(
            x=df["month"], y=df["otc_from_new_patients"],
            name="Neue Patienten (nie Arzt)", stackgroup="one",
            line=dict(color=GREEN),
        ))
        fig3c.add_trace(go.Scatter(
            x=df["month"], y=df["otc_from_rx_migration"],
            name="Rx-Migration (Sildenafil)", stackgroup="one",
            line=dict(color=BLUE),
        ))
        fig3c.add_trace(go.Scatter(
            x=df["month"], y=df["otc_from_tadalafil"],
            name="Tadalafil-Migration", stackgroup="one",
            line=dict(color=PURPLE),
        ))
        fig3c.update_layout(
            title="OTC-Volumen Herkunft (Woher kommen die Patienten?)",
            xaxis_title="Monate", yaxis_title="Packungen",
            yaxis_tickformat=",", height=400,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig3c, width="stretch")

    # ─── Chart 4: Treatment Gap ───────────────────────────────────
    with tab4:
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(
            x=df["month"], y=df["treatment_rate_effective"],
            name="Therapiequote (effektiv)",
            line=dict(color=GREEN, width=3),
            fill="tozeroy", fillcolor="rgba(39,174,96,0.1)",
        ))
        fig5.add_hline(y=params.treatment_rate, line_dash="dash",
                       line_color="#999", annotation_text="Vor Switch (33%)")
        fig5.update_layout(
            title="Schliessung der Therapieluecke durch OTC-Zugang",
            xaxis_title="Monate nach Switch", yaxis_title="Therapiequote",
            yaxis_tickformat=".0%", height=400,
        )
        st.plotly_chart(fig5, width="stretch")

        col_g, col_h = st.columns(2)
        with col_g:
            fig5b = go.Figure()
            fig5b.add_trace(go.Scatter(
                x=df["month"], y=df["awareness"],
                name="Brand Awareness",
                line=dict(color=AMBER, width=2.5),
            ))
            fig5b.update_layout(
                title="Consumer Awareness (Viagra OTC)",
                xaxis_title="Monate", yaxis_title="Awareness",
                yaxis_tickformat=".0%", yaxis_range=[0, 1], height=360,
            )
            st.plotly_chart(fig5b, width="stretch")

        with col_h:
            untreated_before = params.ed_prevalence_men * (1 - params.treatment_rate)
            untreated_after = params.ed_prevalence_men * (1 - kpis["treatment_rate_final"])
            fig5c = go.Figure()
            fig5c.add_trace(go.Bar(
                x=["Vor Switch", "Nach 5J OTC"],
                y=[untreated_before, untreated_after],
                marker_color=[RED, GREEN],
                text=[f"{untreated_before/1e6:.1f}M", f"{untreated_after/1e6:.1f}M"],
                textposition="outside",
            ))
            fig5c.update_layout(
                title="Unbehandelte Maenner mit ED (Deutschland)",
                yaxis_title="Anzahl", yaxis_tickformat=",", height=360,
            )
            st.plotly_chart(fig5c, width="stretch")

    # ─── Chart 5: Profitability ───────────────────────────────────
    with tab5:
        fig6 = go.Figure()
        fig6.add_trace(go.Scatter(
            x=df["month"], y=df["cumulative_total_revenue"],
            name="Kum. Umsatz", line=dict(color=BLUE, width=2.5),
        ))
        fig6.add_trace(go.Scatter(
            x=df["month"], y=df["cumulative_profit"],
            name="Kum. Gewinn", line=dict(color=GREEN, width=2.5),
        ))
        fig6.add_trace(go.Scatter(
            x=df["month"], y=df["cumulative_marketing"],
            name="Kum. Marketing", line=dict(color=RED, width=1.5, dash="dot"),
        ))
        fig6.update_layout(
            title="Kumulierter Umsatz, Gewinn & Marketing-Invest",
            xaxis_title="Monate", yaxis_title="EUR",
            height=420,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig6, width="stretch")

        fig6b = go.Figure()
        fig6b.add_trace(go.Bar(
            x=df["month"], y=df["operating_profit"],
            name="Operativer Gewinn",
            marker_color=[GREEN if v >= 0 else RED for v in df["operating_profit"]],
        ))
        fig6b.update_layout(
            title="Monatlicher operativer Gewinn",
            xaxis_title="Monate", yaxis_title="EUR",
            height=380,
        )
        st.plotly_chart(fig6b, width="stretch")

    # ═══════════════════════════════════════════════════════════════════
    # SCENARIO COMPARISON TABLE
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### Szenario-Vergleich")

    scenarios = {
        "Konservativ": {"otc_peak_packs_per_month": 200_000, "otc_ramp_months": 24,
                        "marketing_monthly_eur": 300_000, "new_patient_share": 0.55,
                        "rx_decline_rate": 0.12, "brand_otc_share": 0.35},
        "Base Case": {},
        "Optimistisch": {"otc_peak_packs_per_month": 450_000, "otc_ramp_months": 12,
                         "marketing_monthly_eur": 750_000, "new_patient_share": 0.70,
                         "rx_decline_rate": 0.05, "brand_otc_share": 0.50},
    }

    comp_rows = []
    for sn, overrides in scenarios.items():
        p = SildenafilOtcParams(**overrides)
        k = calculate_kpis_sildenafil(forecast_sildenafil_otc(p))
        comp_rows.append({
            "Szenario": sn,
            "OTC Umsatz J1": f"EUR {k['year1_otc_revenue']/1e6:.0f}M",
            "Gesamt 5J": f"EUR {k['total_5y_revenue']/1e6:.0f}M",
            "Gewinn 5J": f"EUR {k['total_5y_profit']/1e6:.0f}M",
            "Online M24": f"{k['online_share_m24']:.0%}",
            "Marke M12": f"{k['brand_share_m12']:.0%}",
            "Crossover": f"M{k['crossover_month']}" if k['crossover_month'] else "–",
        })

    st.dataframe(pd.DataFrame(comp_rows).set_index("Szenario"), width=900)

    # ═══════════════════════════════════════════════════════════════════
    # METHODOLOGY
    # ═══════════════════════════════════════════════════════════════════
    with st.expander("Methodik & Datenquellen"):
        st.markdown("""
        **Modell-Architektur:** Dual-Channel Rx/OTC mit Apothekenverteilung

        | Komponente | Methodik | Referenz |
        |---|---|---|
        | OTC-Ramp | Logistische S-Kurve | UK Viagra Connect Launch 2018 |
        | Rx-Effekt | Exponentieller Decay (langsam) | UK: Rx stieg sogar post-Switch |
        | Apothekenverteilung | 2 Kanaele (apothekenpflichtig) | Online-Apotheke CAGR 12.6% |
        | Markenanteil | Linearer Erosion + Premium | UK: Generika ab GBP 0.50/Tab |
        | Treatment Gap | Logistische Schliessung | UK: 63% Neupatienten |
        | Awareness | S-Kurve mit hoher Baseline | Viagra Markenbekanntheit >60% |

        **Datenquellen (alle oeffentlich):**
        - IQVIA Pharmamarkt DE, Apotheke Adhoc (PDE5-Marktdaten)
        - Handelsblatt / Citeline (BfArM SVA-Entscheidungen 2022/2023/2025)
        - PMC / Lee et al. 2021 (UK Real-World-Studie, n=1,162)
        - PAGB/Frontier Economics (UK OTC Impact Report)
        - Viatris Investor Relations, Newsroom
        """)


if __name__ == "__main__":
    st.set_page_config(page_title="Sildenafil OTC Switch", page_icon="💊", layout="wide")
    show()
