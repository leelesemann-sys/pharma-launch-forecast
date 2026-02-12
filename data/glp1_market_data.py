"""
Market data for the GLP-1 Brand-vs-Brand scenario: Mounjaro vs. Wegovy/Ozempic.

Based on publicly available market intelligence:
- GLP-1 RA global market: ~$63B (2025), growing ~25-30% p.a.
- Germany: ~880,000 T2D patients on GLP-1 (13.9% of 6.33M, BARMER 2025)
- Ozempic (Semaglutid s.c.): Market leader T2D, ~€80/pen, ~€300/month maintenance
- Mounjaro (Tirzepatid): Dual GIP/GLP-1, EU approval T2D Sep 2022, Adipositas Dec 2023
- Wegovy (Semaglutid 2.4mg): Adipositas-only, NOT covered by GKV (G-BA June 2024)
- Rybelsus (Semaglutid oral): Oral alternative, smaller share
- Trulicity (Dulaglutid): Legacy leader (~43% of GLP-1 in DE), declining
- Key fact: Mounjaro has confidential GKV price since Aug 2025 (first ever in DE)
- Semaglutide patents: EU/US expire ~2031+; emerging markets 2026

Sources:
- BARMER Gesundheitswesen aktuell 2025 (GLP-1 Versorgungsanalyse)
- G-BA Nutzenbewertung Tirzepatid (A23-112, Sep 2024)
- KV Berlin Verordnungs-News 2024/2025
- IQVIA Pharmamarkt Q1/Q2 2024
- EMA Mounjaro EPAR
"""

import pandas as pd
import numpy as np


def generate_glp1_market_data() -> dict:
    """Generate realistic market data for the German GLP-1 RA market."""

    # --- Market Definition ---
    market = {
        "name": "GLP-1-Rezeptor-Agonisten (GLP-1 RA) - Deutschland",
        "indications": [
            "Diabetes mellitus Typ 2 (T2D)",
            "Adipositas / Gewichtsmanagement (BMI >= 30)",
            "Adipositas mit Komorbiditaten (BMI 27-30)",
        ],
        "total_t2d_patients_de": 6_330_000,       # GKV-Versicherte mit T2D
        "glp1_treated_patients_2024": 880_000,     # 13.9% on GLP-1 RA
        "glp1_penetration_t2d_2024": 0.139,
        "potential_obesity_patients_de": 8_000_000, # BMI >= 30 in DE
        "obesity_treated_glp1_2024": 50_000,       # Estimate (mostly self-pay)
        "total_glp1_monthly_trx_de": 950_000,      # ~950K monthly Rx (all GLP-1)
        "total_glp1_revenue_monthly_eur": 230_000_000,  # ~230M EUR/month
        "currency": "EUR",
        "geography": "Deutschland",
        "data_source": "Synthetisch (basierend auf BARMER 2025, IQVIA, G-BA)",
    }

    # --- Products in Market (current state mid-2025) ---
    products = {
        "Ozempic (Semaglutid s.c.)": {
            "company": "Novo Nordisk",
            "molecule": "Semaglutid",
            "mechanism": "GLP-1 RA",
            "indication_primary": "T2D",
            "indication_secondary": None,  # Off-label weight loss exists
            "market_share_trx": 0.32,      # ~32% of GLP-1 Rx
            "market_share_revenue": 0.35,
            "price_per_month_eur": 300.00,  # ~€300/month maintenance (1mg)
            "ddd_price_eur": 10.00,
            "launch_de": "2020-02",
            "gkv_covered": True,
            "supply_constrained": True,     # Lieferengpass since 2022
            "amnog_status": "Erstattungsbetrag vereinbart",
        },
        "Mounjaro (Tirzepatid)": {
            "company": "Eli Lilly",
            "molecule": "Tirzepatid",
            "mechanism": "Dual GIP/GLP-1 RA",
            "indication_primary": "T2D",
            "indication_secondary": "Adipositas (seit Dez 2023)",
            "market_share_trx": 0.08,      # Still ramping up
            "market_share_revenue": 0.10,
            "price_per_month_eur": 350.00,  # ~€206-490/month, avg ~€350
            "ddd_price_eur": 11.50,
            "launch_de": "2023-11",         # Marktstart DE Nov 2023
            "gkv_covered": True,            # T2D: yes; Adipositas: complex
            "supply_constrained": False,
            "amnog_status": "Geheimer Erstattungsbetrag (ab Aug 2025)",
        },
        "Trulicity (Dulaglutid)": {
            "company": "Eli Lilly",
            "molecule": "Dulaglutid",
            "mechanism": "GLP-1 RA",
            "indication_primary": "T2D",
            "indication_secondary": None,
            "market_share_trx": 0.38,      # Legacy leader, 43% declining
            "market_share_revenue": 0.28,   # Lower price
            "price_per_month_eur": 178.00,
            "ddd_price_eur": 5.93,
            "launch_de": "2015-03",
            "gkv_covered": True,
            "supply_constrained": False,
            "amnog_status": "Erstattungsbetrag vereinbart",
        },
        "Wegovy (Semaglutid 2.4mg)": {
            "company": "Novo Nordisk",
            "molecule": "Semaglutid",
            "mechanism": "GLP-1 RA (high-dose)",
            "indication_primary": "Adipositas",
            "indication_secondary": None,
            "market_share_trx": 0.04,      # Small: not GKV-covered
            "market_share_revenue": 0.06,
            "price_per_month_eur": 277.00,  # ~€277/month (2.4mg, after price cut)
            "ddd_price_eur": 9.89,          # After April 2025 price reduction
            "launch_de": "2023-07",
            "gkv_covered": False,           # G-BA: Lifestyle-Arzneimittel
            "supply_constrained": True,
            "amnog_status": "Verordnungsausschluss GKV (G-BA Juni 2024)",
        },
        "Rybelsus (Semaglutid oral)": {
            "company": "Novo Nordisk",
            "molecule": "Semaglutid",
            "mechanism": "GLP-1 RA (oral)",
            "indication_primary": "T2D",
            "indication_secondary": None,
            "market_share_trx": 0.07,
            "market_share_revenue": 0.08,
            "price_per_month_eur": 280.00,
            "ddd_price_eur": 9.33,
            "launch_de": "2020-06",
            "gkv_covered": True,
            "supply_constrained": False,
            "amnog_status": "Erstattungsbetrag vereinbart",
        },
        "Victoza (Liraglutid)": {
            "company": "Novo Nordisk",
            "molecule": "Liraglutid",
            "mechanism": "GLP-1 RA",
            "indication_primary": "T2D",
            "indication_secondary": None,
            "market_share_trx": 0.06,      # Declining legacy
            "market_share_revenue": 0.05,
            "price_per_month_eur": 195.00,
            "ddd_price_eur": 6.50,
            "launch_de": "2010-07",
            "gkv_covered": True,
            "supply_constrained": True,
            "amnog_status": "Erstattungsbetrag vereinbart",
        },
        "Sonstige (Exenatid etc.)": {
            "company": "Diverse",
            "molecule": "Exenatid, Lixisenatid",
            "mechanism": "GLP-1 RA",
            "indication_primary": "T2D",
            "indication_secondary": None,
            "market_share_trx": 0.05,
            "market_share_revenue": 0.08,
            "price_per_month_eur": 150.00,
            "ddd_price_eur": 5.00,
            "launch_de": "2007",
            "gkv_covered": True,
            "supply_constrained": False,
            "amnog_status": "Diverse",
        },
    }

    # --- Historical quarterly data (Q1 2023 - Q4 2025, 12 quarters) ---
    quarters = pd.date_range("2023-01-01", periods=12, freq="QS")
    np.random.seed(99)

    # Market growth trajectory: GLP-1 segment growing ~30% p.a.
    base_monthly_trx = 650_000  # Q1 2023 starting point

    # Share trajectories over 12 quarters
    share_trajectories = {
        "Ozempic (Semaglutid s.c.)":    [0.28, 0.29, 0.30, 0.31, 0.32, 0.33, 0.33, 0.33, 0.32, 0.32, 0.32, 0.32],
        "Mounjaro (Tirzepatid)":         [0.00, 0.00, 0.00, 0.01, 0.02, 0.04, 0.05, 0.06, 0.07, 0.08, 0.08, 0.08],
        "Trulicity (Dulaglutid)":        [0.45, 0.44, 0.43, 0.42, 0.41, 0.39, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38],
        "Wegovy (Semaglutid 2.4mg)":     [0.00, 0.00, 0.01, 0.02, 0.02, 0.03, 0.03, 0.03, 0.04, 0.04, 0.04, 0.04],
        "Rybelsus (Semaglutid oral)":    [0.08, 0.08, 0.07, 0.07, 0.07, 0.07, 0.07, 0.07, 0.07, 0.07, 0.07, 0.07],
        "Victoza (Liraglutid)":          [0.11, 0.10, 0.10, 0.09, 0.08, 0.07, 0.07, 0.06, 0.06, 0.06, 0.06, 0.06],
        "Sonstige (Exenatid etc.)":      [0.08, 0.09, 0.09, 0.08, 0.08, 0.07, 0.07, 0.07, 0.06, 0.05, 0.05, 0.05],
    }

    history_rows = []
    for i, q in enumerate(quarters):
        # GLP-1 market grows ~7% per quarter (~30% annualized)
        quarterly_growth = 1.07 ** i
        total_trx_q = base_monthly_trx * 3 * quarterly_growth

        for product_name, shares in share_trajectories.items():
            share = shares[i] + np.random.normal(0, 0.003)
            share = max(0.0, share)
            trx = int(total_trx_q * share)
            price = products[product_name]["price_per_month_eur"]
            revenue = trx * price

            history_rows.append({
                "quarter": q.strftime("%Y-Q") + str((q.month - 1) // 3 + 1),
                "date": q,
                "product": product_name,
                "company": products[product_name]["company"],
                "trx": trx,
                "revenue_eur": round(revenue, 2),
                "market_share_trx": round(share, 4),
                "price_per_month_eur": price,
            })

    historical_df = pd.DataFrame(history_rows)

    # --- Market Growth Drivers ---
    growth_drivers = {
        "t2d_guideline_shift": {
            "label": "Leitlinien-Shift: GLP-1 frueher in T2D-Therapie",
            "impact": "Erhoehung GLP-1-Penetration von 14% auf 25-30%",
            "timeline": "2025-2030",
            "certainty": "Hoch",
        },
        "obesity_gkv_coverage": {
            "label": "Potenzielle GKV-Erstattung fuer Adipositas",
            "impact": "Marktexplosion: 8M potenzielle Patienten",
            "timeline": "Fruehestens 2027 (Gesetzaenderung noetig)",
            "certainty": "Mittel",
        },
        "cv_risk_reduction": {
            "label": "Kardiovaskulaere Indikationserweiterung",
            "impact": "SELECT-Studie: -20% CV-Events; erweitert Zielgruppe",
            "timeline": "2025-2026",
            "certainty": "Hoch",
        },
        "nash_mash": {
            "label": "MASH/NASH Indikation (Lebererkrankung)",
            "impact": "Neue Patientenpopulation, Lilly fuehrend",
            "timeline": "2026-2028",
            "certainty": "Mittel",
        },
        "oral_semaglutide_25mg": {
            "label": "Orales Semaglutid Hochdosis (25/50mg)",
            "impact": "Convenience-Shift weg von Injektion",
            "timeline": "2025-2026",
            "certainty": "Hoch",
        },
        "supply_normalization": {
            "label": "Behebung Lieferengpaesse (Semaglutid)",
            "impact": "Aufholen unterdrueckter Nachfrage",
            "timeline": "2025",
            "certainty": "Hoch",
        },
    }

    # --- Competitive Differentiation ---
    differentiation = {
        "Mounjaro (Tirzepatid)": {
            "efficacy_hba1c_reduction": -2.4,       # % points (SURPASS trials)
            "efficacy_weight_loss_pct": -22.5,       # SURMOUNT-1 (72 weeks)
            "mechanism_advantage": "Dual GIP/GLP-1: ueberlegene Wirksamkeit",
            "dosing": "1x woechentlich s.c.",
            "gi_tolerability": "Vergleichbar",
            "supply_reliability": "Gut",
            "gkv_t2d": True,
            "gkv_obesity": True,                     # Via Mounjaro (not Zepbound in EU)
        },
        "Ozempic (Semaglutid s.c.)": {
            "efficacy_hba1c_reduction": -1.8,        # SUSTAIN trials
            "efficacy_weight_loss_pct": -15.0,        # STEP trials (Wegovy dose)
            "mechanism_advantage": "Bewaehrter GLP-1 RA, breiteste Evidenzbasis",
            "dosing": "1x woechentlich s.c.",
            "gi_tolerability": "Maessig (Uebelkeit haeufig)",
            "supply_reliability": "Eingeschraenkt (Lieferengpass)",
            "gkv_t2d": True,
            "gkv_obesity": False,                    # Wegovy = nicht GKV
        },
    }

    # --- Scenario Defaults ---
    scenarios = {
        "base": {
            "label": "Base Case",
            "market_growth_annual": 0.25,
            "mounjaro_peak_share_t2d": 0.25,
            "mounjaro_peak_share_obesity": 0.35,
            "obesity_gkv_year": None,          # No GKV coverage
            "t2d_penetration_target": 0.22,
            "mounjaro_months_to_peak": 36,
        },
        "bull_lilly": {
            "label": "Bull Case (Lilly / Mounjaro)",
            "market_growth_annual": 0.35,
            "mounjaro_peak_share_t2d": 0.35,
            "mounjaro_peak_share_obesity": 0.45,
            "obesity_gkv_year": 2027,          # GKV coverage from 2027
            "t2d_penetration_target": 0.28,
            "mounjaro_months_to_peak": 24,
        },
        "bull_novo": {
            "label": "Bull Case (Novo Nordisk)",
            "market_growth_annual": 0.30,
            "mounjaro_peak_share_t2d": 0.15,
            "mounjaro_peak_share_obesity": 0.20,
            "obesity_gkv_year": None,
            "t2d_penetration_target": 0.20,
            "mounjaro_months_to_peak": 48,
        },
    }

    return {
        "market": market,
        "products": products,
        "historical": historical_df,
        "growth_drivers": growth_drivers,
        "differentiation": differentiation,
        "scenarios": scenarios,
    }


# Convenience constants
MOUNJARO_PRICE_PER_MONTH = 350.00
OZEMPIC_PRICE_PER_MONTH = 300.00
WEGOVY_PRICE_PER_MONTH = 277.00
GLP1_MONTHLY_TRX_DE = 950_000
T2D_PATIENTS_DE = 6_330_000
GLP1_PENETRATION_2024 = 0.139
