"""
Forecast engine for Brand-vs-Brand competition in expanding markets.

Use Case: GLP-1 RA market – Mounjaro (Lilly) vs. Ozempic/Wegovy (Novo Nordisk)

Key differences vs. generic entry model:
- Market EXPANDS (not erodes) – the pie grows
- Two (or more) branded originators compete for share
- Indications drive growth (T2D, Obesity, CV risk, MASH)
- AMNOG pricing (not generic price erosion)
- No aut-idem / pharmacy substitution
- Supply constraints limit theoretical uptake

Core methodology:
- Market expansion: S-curve for treatment penetration
- Share competition: Logistic share shift based on differentiation
- Indication layering: T2D base + Obesity upside + CV/MASH optionality
- Supply constraint modeling: min(demand, supply_capacity)
"""

import numpy as np
import pandas as pd
from dataclasses import dataclass, field


# ═══════════════════════════════════════════════════════════════════════
# PARAMETER CLASSES
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class MarketParams:
    """Parameters for the overall GLP-1 market expansion."""
    # T2D segment
    t2d_patients_de: int = 6_330_000
    glp1_penetration_start: float = 0.139       # 13.9% (2024)
    glp1_penetration_target: float = 0.22       # Target penetration
    penetration_years_to_target: int = 5
    t2d_patient_growth_annual: float = 0.015    # Slight increase in T2D prevalence

    # Obesity segment
    obesity_patients_de: int = 8_000_000
    obesity_treated_start: int = 50_000         # Currently ~50K (mostly self-pay)
    obesity_gkv_coverage: bool = False          # GKV pays for obesity?
    obesity_gkv_start_year: int = 2028          # If enabled, when?
    obesity_penetration_if_gkv: float = 0.05    # 5% of 8M = 400K patients
    obesity_penetration_no_gkv: float = 0.01    # 1% = self-pay only

    # Overall market
    total_glp1_monthly_trx_start: int = 950_000
    market_growth_annual: float = 0.25          # 25% market growth


@dataclass
class BrandParams:
    """Parameters for a specific brand competitor."""
    name: str = "Mounjaro (Tirzepatid)"
    company: str = "Eli Lilly"

    # Current position
    current_share_trx: float = 0.08
    price_per_month_eur: float = 350.00

    # Growth trajectory
    target_peak_share_t2d: float = 0.25
    target_peak_share_obesity: float = 0.35
    months_to_peak: int = 36
    share_ramp_speed: float = 1.0               # 1.0 = normal

    # Differentiation scores (0-10)
    efficacy_score: float = 9.0                 # Wirksamkeit
    safety_score: float = 7.0                   # Vertraeglichkeit
    convenience_score: float = 8.0              # Darreichung
    supply_score: float = 8.0                   # Versorgungssicherheit
    evidence_score: float = 7.5                 # Evidenzbasis / Studienlage

    # Indication coverage
    has_t2d: bool = True
    has_obesity: bool = True
    has_cv_indication: bool = False             # Cardiovascular risk
    has_mash: bool = False                      # MASH/NASH

    # Supply constraints
    supply_constrained: bool = False
    supply_capacity_monthly_trx: int = 200_000  # Max capacity
    supply_normalization_month: int = 0         # When constraint lifts

    # Pricing dynamics
    price_trend_annual: float = -0.03           # -3% annual price erosion
    amnog_price_cut_month: int = 0              # Month of AMNOG price cut
    amnog_price_cut_pct: float = 0.0

    # Revenue & costs
    cogs_pct: float = 0.20
    sga_monthly_eur: float = 800_000
    medical_affairs_monthly_eur: float = 200_000


@dataclass
class CompetitorParams:
    """Simplified parameters for the main competitor."""
    name: str = "Ozempic / Wegovy (Semaglutid)"
    company: str = "Novo Nordisk"
    current_share_trx: float = 0.36             # Ozempic + Wegovy combined
    price_per_month_eur: float = 300.00
    target_peak_share_t2d: float = 0.30
    target_peak_share_obesity: float = 0.30
    supply_constrained: bool = True
    supply_normalization_month: int = 12
    efficacy_score: float = 7.5
    evidence_score: float = 9.0                 # Most established evidence
    has_t2d: bool = True
    has_obesity: bool = True                    # Via Wegovy (but not GKV)


# ═══════════════════════════════════════════════════════════════════════
# CURVE FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════

def _market_expansion_curve(
    t: int,
    start_trx: int,
    growth_annual: float,
    max_growth_factor: float = 4.0,
) -> int:
    """Model market expansion with decelerating growth."""
    monthly_rate = (1 + growth_annual) ** (1/12) - 1
    # Logistic dampening: growth slows as market matures
    progress = min(1.0, t / 60)  # 60 months to maturity
    dampened_rate = monthly_rate * (1 - 0.5 * progress)
    raw = start_trx * (1 + dampened_rate) ** t
    cap = start_trx * max_growth_factor
    return int(min(raw, cap))


def _share_shift_curve(
    t: int,
    current_share: float,
    target_share: float,
    months_to_peak: int,
    speed: float = 1.0,
) -> float:
    """S-curve share shift from current to target."""
    if months_to_peak <= 0:
        return target_share
    midpoint = months_to_peak / 2
    steepness = 4.0 / months_to_peak * speed
    sigmoid = 1 / (1 + np.exp(-steepness * (t - midpoint)))
    share = current_share + (target_share - current_share) * sigmoid
    return float(np.clip(share, 0, 1))


def _indication_multiplier(
    t: int,
    brand: BrandParams,
    market: MarketParams,
    forecast_start_year: int = 2026,
) -> tuple[float, dict]:
    """
    Calculate volume multiplier from indication expansions.

    Returns (multiplier, details_dict)
    Base = 1.0 (T2D only). Each indication adds incremental volume.
    """
    multiplier = 1.0
    details = {"t2d": 1.0}

    # Obesity: big upside if GKV coverage
    if brand.has_obesity:
        if market.obesity_gkv_coverage:
            months_since_gkv = (forecast_start_year * 12 + t) - (market.obesity_gkv_start_year * 12)
            if months_since_gkv >= 0:
                # Ramp up over 24 months
                ramp = min(1.0, months_since_gkv / 24)
                obesity_uplift = 0.5 * ramp  # +50% volume at full ramp
                multiplier += obesity_uplift
                details["obesity_gkv"] = obesity_uplift
        else:
            # Self-pay obesity: small but growing
            self_pay_ramp = min(1.0, t / 36)
            obesity_uplift = 0.08 * self_pay_ramp  # +8% volume
            multiplier += obesity_uplift
            details["obesity_self_pay"] = obesity_uplift

    # CV indication
    if brand.has_cv_indication:
        cv_ramp = min(1.0, max(0, (t - 6)) / 18)  # Starts month 6, ramps 18m
        cv_uplift = 0.15 * cv_ramp
        multiplier += cv_uplift
        details["cv_risk"] = cv_uplift

    # MASH/NASH
    if brand.has_mash:
        mash_ramp = min(1.0, max(0, (t - 18)) / 24)  # Later, slower
        mash_uplift = 0.10 * mash_ramp
        multiplier += mash_uplift
        details["mash"] = mash_uplift

    return multiplier, details


def _supply_constraint(
    t: int,
    demand_trx: int,
    brand: BrandParams,
) -> int:
    """Apply supply constraints if applicable."""
    if not brand.supply_constrained:
        return demand_trx
    if t >= brand.supply_normalization_month:
        return demand_trx
    return min(demand_trx, brand.supply_capacity_monthly_trx)


# ═══════════════════════════════════════════════════════════════════════
# MAIN FORECAST FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════

def forecast_brand(
    brand: BrandParams,
    competitor: CompetitorParams,
    market: MarketParams,
    forecast_months: int = 60,
    start_date: str = "2026-01-01",
) -> pd.DataFrame:
    """
    Generate monthly forecast for a brand in an expanding market.

    Perspective: Lilly (Mounjaro) OR Novo Nordisk (Ozempic/Wegovy)
    """
    start = pd.Timestamp(start_date)
    dates = pd.date_range(start, periods=forecast_months, freq="MS")
    start_year = start.year

    rows = []
    cumulative_revenue = 0
    cumulative_profit = 0

    for i, date in enumerate(dates):
        t = i

        # 1. MARKET EXPANSION
        total_market_trx = _market_expansion_curve(
            t, market.total_glp1_monthly_trx_start,
            market.market_growth_annual,
        )

        # 2. SHARE EVOLUTION
        my_share_t2d = _share_shift_curve(
            t, brand.current_share_trx,
            brand.target_peak_share_t2d,
            brand.months_to_peak,
            brand.share_ramp_speed,
        )

        competitor_share = _share_shift_curve(
            t, competitor.current_share_trx,
            competitor.target_peak_share_t2d,
            36,  # Competitor also evolving
        )

        # Normalize: my_share + competitor + rest = 1.0
        rest_share = max(0.1, 1.0 - my_share_t2d - competitor_share)

        # 3. INDICATION MULTIPLIER
        indication_mult, indication_details = _indication_multiplier(
            t, brand, market, start_year,
        )

        # 4. BASE VOLUME
        base_trx = int(total_market_trx * my_share_t2d)

        # Apply indication multiplier
        demand_trx = int(base_trx * indication_mult)

        # 5. SUPPLY CONSTRAINT
        actual_trx = _supply_constraint(t, demand_trx, brand)
        supply_gap = demand_trx - actual_trx

        # 6. PRICING
        price = brand.price_per_month_eur
        # Annual price erosion
        years_in = t / 12
        price *= (1 + brand.price_trend_annual) ** years_in
        # AMNOG cut
        if brand.amnog_price_cut_month > 0 and t >= brand.amnog_price_cut_month:
            price *= (1 - brand.amnog_price_cut_pct)
        price = max(price, 100.0)

        # 7. REVENUE
        revenue = actual_trx * price

        # 8. PROFITABILITY
        cogs = revenue * brand.cogs_pct
        sga = brand.sga_monthly_eur
        medical = brand.medical_affairs_monthly_eur
        gross_profit = revenue - cogs
        operating_profit = gross_profit - sga - medical

        cumulative_revenue += revenue
        cumulative_profit += operating_profit

        # 9. COMPETITOR METRICS
        competitor_trx = int(total_market_trx * competitor_share)
        competitor_revenue = competitor_trx * competitor.price_per_month_eur

        rows.append({
            "date": date,
            "month": date.strftime("%Y-%m"),
            "months_from_start": t,
            # Market
            "total_market_trx": total_market_trx,
            "market_growth_vs_start": total_market_trx / market.total_glp1_monthly_trx_start,
            # My brand
            "my_share": round(my_share_t2d, 4),
            "my_base_trx": base_trx,
            "indication_multiplier": round(indication_mult, 3),
            "my_demand_trx": demand_trx,
            "my_actual_trx": actual_trx,
            "supply_gap_trx": supply_gap,
            "my_price": round(price, 2),
            "my_revenue": round(revenue, 2),
            "my_cogs": round(cogs, 2),
            "my_gross_profit": round(gross_profit, 2),
            "my_operating_profit": round(operating_profit, 2),
            "cumulative_revenue": round(cumulative_revenue, 2),
            "cumulative_profit": round(cumulative_profit, 2),
            # Competitor
            "competitor_share": round(competitor_share, 4),
            "competitor_trx": competitor_trx,
            "competitor_revenue": round(competitor_revenue, 2),
            # Rest of market
            "rest_share": round(rest_share, 4),
            "rest_trx": int(total_market_trx * rest_share),
            # Indication breakdown
            "indication_t2d": indication_details.get("t2d", 1.0),
            "indication_obesity": indication_details.get("obesity_gkv", indication_details.get("obesity_self_pay", 0)),
            "indication_cv": indication_details.get("cv_risk", 0),
            "indication_mash": indication_details.get("mash", 0),
        })

    return pd.DataFrame(rows)


def calculate_kpis_brand(df: pd.DataFrame) -> dict:
    """Calculate summary KPIs for brand forecast."""
    if len(df) == 0:
        return {}

    year1 = df.head(12)
    peak_revenue_month = df["my_revenue"].max()
    peak_share = df["my_share"].max()

    # Time to #1 (if applicable)
    leading = df[df["my_share"] > df["competitor_share"]]
    overtake_month = leading.iloc[0]["months_from_start"] if len(leading) > 0 else None

    return {
        "year1_revenue": year1["my_revenue"].sum(),
        "year1_profit": year1["my_operating_profit"].sum(),
        "total_5y_revenue": df["my_revenue"].sum(),
        "total_5y_profit": df["my_operating_profit"].sum(),
        "peak_monthly_revenue": peak_revenue_month,
        "peak_share": peak_share,
        "share_month_12": df.iloc[11]["my_share"] if len(df) > 11 else None,
        "share_month_36": df.iloc[35]["my_share"] if len(df) > 35 else None,
        "share_month_60": df.iloc[-1]["my_share"],
        "cumulative_revenue": df["cumulative_revenue"].iloc[-1],
        "cumulative_profit": df["cumulative_profit"].iloc[-1],
        "overtake_month": overtake_month,
        "market_size_start": df.iloc[0]["total_market_trx"],
        "market_size_end": df.iloc[-1]["total_market_trx"],
        "market_growth_total": df.iloc[-1]["total_market_trx"] / df.iloc[0]["total_market_trx"],
        "avg_price": df["my_price"].mean(),
        "avg_indication_multiplier": df["indication_multiplier"].mean(),
    }
