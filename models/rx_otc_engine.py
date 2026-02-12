"""
Rx-to-OTC Switch Forecast Engine
=================================
Models the transition of a prescription drug to OTC status.

Key dynamics:
  - Dual channel: Rx volume decays as OTC ramps up
  - Market expansion: OTC reaches new patients who never visited a doctor
  - Cannibalization: OTC eats into adjacent OTC categories (e.g. antacids)
  - Seasonality: OTC products have seasonal demand patterns
  - Price elasticity: consumer pays 100% → price-sensitive demand
  - Marketing-driven awareness: S-curve from launch investment

Example: Omeprazol/Pantoprazol PPI switch (DE, July 2009)
  - Rx: ~4.2M packs/month, GKV-Festbetrag ~€17 (50 St)
  - OTC: 14 St max, ~€8 retail, patient pays 100%
  - Only ~3% of PPI volume moved to OTC (partial switch)
  - 11.9% heartburn market expansion in year 1
  - Cannibalized antacids (-€11M) and H2-antagonists (-46%)
"""

from dataclasses import dataclass, field
from typing import Optional
import numpy as np
import pandas as pd


@dataclass
class RxOtcParams:
    """Parameters for the Rx-to-OTC switch model."""

    # Product
    product_name: str = "Omeprazol 20mg"
    company: str = "Mein Unternehmen"
    category: str = "PPI (Protonenpumpeninhibitoren)"
    switch_year: int = 2026

    # Rx channel (pre-switch baseline)
    rx_packs_per_month: int = 350_000          # monthly Rx packs before switch
    rx_price_per_pack: float = 16.99           # GKV Festbetrag (50 St)
    rx_patient_copay: float = 5.00             # Zuzahlung
    rx_decline_rate: float = 0.15              # % of Rx volume that migrates to OTC over time
    rx_decline_months: int = 24                # months until Rx decline stabilizes

    # OTC channel
    otc_price_per_pack: float = 7.99           # retail price (14 St)
    otc_peak_packs_per_month: int = 280_000    # peak OTC volume (new + migrated)
    otc_ramp_months: int = 18                  # months to reach peak OTC
    otc_pack_size: int = 14                    # tablets per OTC pack
    rx_pack_size: int = 50                     # tablets per Rx pack

    # Market expansion (new patients who never had Rx)
    new_patient_share: float = 0.70            # % of OTC volume from genuinely new patients
    cannibalization_adjacent_pct: float = 0.20 # % revenue eaten from adjacent OTC categories

    # Price elasticity
    price_elasticity: float = -0.8             # 1% price increase → 0.8% volume decrease
    price_trend_annual: float = -0.02          # annual price erosion from generics/competition

    # Seasonality (12 monthly factors, 1.0 = average)
    seasonality: list = field(default_factory=lambda: [
        0.90, 0.85, 0.95, 1.05, 1.05, 1.00,   # Jan-Jun
        0.95, 0.90, 1.00, 1.10, 1.15, 1.10,    # Jul-Dec (autumn/winter peak for GI)
    ])

    # Marketing & awareness
    marketing_monthly_eur: float = 300_000     # monthly marketing spend
    awareness_peak: float = 0.65               # max brand awareness (0-1)
    awareness_ramp_months: int = 12            # months to reach peak awareness

    # Consumer funnel
    trial_rate: float = 0.30                   # % of aware consumers who try
    repeat_rate: float = 0.55                  # % of trialists who become repeat buyers

    # Costs
    cogs_pct: float = 0.25                     # COGS as % of revenue
    pharmacy_margin_pct: float = 0.40          # pharmacy takes ~40% of retail price
    distribution_cost_pct: float = 0.08        # wholesale + logistics

    # Forecast
    forecast_months: int = 60


@dataclass
class AdjacentCategory:
    """An adjacent OTC category that gets cannibalized."""
    name: str = "Antazida (Maaloxan, Rennie etc.)"
    monthly_revenue: float = 7_750_000         # EUR/month baseline
    cannibalization_peak: float = 0.15         # max share lost to switched product
    cannibalization_months: int = 18           # months to reach peak cannibalization


def _awareness_curve(month: int, peak: float, ramp: int) -> float:
    """S-curve for consumer awareness build-up."""
    if ramp <= 0:
        return peak
    midpoint = ramp / 2
    steepness = 6.0 / ramp
    return peak / (1.0 + np.exp(-steepness * (month - midpoint)))


def _otc_ramp(month: int, peak: int, ramp_months: int) -> float:
    """Logistic ramp for OTC volume."""
    if ramp_months <= 0:
        return float(peak)
    midpoint = ramp_months * 0.45
    steepness = 6.0 / ramp_months
    return peak / (1.0 + np.exp(-steepness * (month - midpoint)))


def _rx_decline(month: int, initial: int, decline_rate: float, decline_months: int) -> float:
    """Exponential decay of Rx volume as OTC takes over."""
    if decline_months <= 0 or decline_rate <= 0:
        return float(initial)
    k = -np.log(1 - decline_rate) / decline_months
    floor = initial * (1 - decline_rate)
    return floor + (initial - floor) * np.exp(-k * month)


def forecast_rx_otc(
    params: RxOtcParams,
    adjacent: Optional[AdjacentCategory] = None,
    forecast_months: Optional[int] = None,
) -> pd.DataFrame:
    """Run the Rx-to-OTC switch forecast.

    Returns DataFrame with monthly data for both Rx and OTC channels.
    """
    months = forecast_months or params.forecast_months
    if adjacent is None:
        adjacent = AdjacentCategory()

    rows = []
    cumulative_otc_revenue = 0.0
    cumulative_rx_revenue = 0.0
    cumulative_total_revenue = 0.0
    cumulative_profit = 0.0
    cumulative_marketing = 0.0

    for m in range(1, months + 1):
        year_frac = m / 12.0
        season_idx = (m - 1) % 12
        season_factor = params.seasonality[season_idx]

        # ─── Awareness ────────────────────────────────────────────
        awareness = _awareness_curve(m, params.awareness_peak, params.awareness_ramp_months)

        # ─── Rx Channel ──────────────────────────────────────────
        rx_packs = _rx_decline(m, params.rx_packs_per_month, params.rx_decline_rate, params.rx_decline_months)
        rx_price = params.rx_price_per_pack  # Festbetrag, stable
        rx_revenue = rx_packs * rx_price

        # ─── OTC Channel ─────────────────────────────────────────
        otc_base = _otc_ramp(m, params.otc_peak_packs_per_month, params.otc_ramp_months)

        # Apply seasonality
        otc_seasonal = otc_base * season_factor

        # Apply awareness scaling (early months: low awareness = fewer buyers)
        awareness_factor = max(0.1, awareness / params.awareness_peak) if params.awareness_peak > 0 else 1.0
        otc_demand = otc_seasonal * awareness_factor

        # Price effect (price changes over time reduce/increase demand)
        price_change = params.price_trend_annual * year_frac
        price_volume_effect = 1.0 + (price_change * params.price_elasticity)
        otc_packs = max(0, otc_demand * price_volume_effect)

        # Current OTC price
        otc_price = params.otc_price_per_pack * (1 + params.price_trend_annual) ** year_frac

        # Revenue: manufacturer gets retail minus pharmacy margin minus distribution
        manufacturer_share = 1.0 - params.pharmacy_margin_pct - params.distribution_cost_pct
        otc_retail_revenue = otc_packs * otc_price
        otc_manufacturer_revenue = otc_retail_revenue * manufacturer_share

        # ─── Volume decomposition ────────────────────────────────
        otc_from_rx_migration = otc_packs * (1 - params.new_patient_share)
        otc_from_new_patients = otc_packs * params.new_patient_share

        # Convert to tablets for fair comparison
        rx_tablets = rx_packs * params.rx_pack_size
        otc_tablets = otc_packs * params.otc_pack_size
        total_tablets = rx_tablets + otc_tablets

        # ─── Adjacent category cannibalization ───────────────────
        cannibal_rate = min(
            adjacent.cannibalization_peak,
            adjacent.cannibalization_peak * min(1.0, m / adjacent.cannibalization_months)
        )
        adjacent_revenue_lost = adjacent.monthly_revenue * cannibal_rate

        # ─── Marketing funnel ────────────────────────────────────
        marketing_spend = params.marketing_monthly_eur
        # Diminishing returns after ramp
        if m > params.awareness_ramp_months * 1.5:
            marketing_spend *= 0.6  # reduce to maintenance level

        # ─── Profitability ───────────────────────────────────────
        total_revenue = rx_revenue + otc_manufacturer_revenue
        cogs = total_revenue * params.cogs_pct
        gross_profit = total_revenue - cogs
        operating_profit = gross_profit - marketing_spend

        cumulative_otc_revenue += otc_manufacturer_revenue
        cumulative_rx_revenue += rx_revenue
        cumulative_total_revenue += total_revenue
        cumulative_profit += operating_profit
        cumulative_marketing += marketing_spend

        # ─── Consumer funnel metrics ─────────────────────────────
        # Potential OTC users (simplified: awareness * category patients)
        potential_users = params.otc_peak_packs_per_month * 2  # rough addressable market
        aware_users = potential_users * awareness
        trialists = aware_users * params.trial_rate
        repeaters = trialists * params.repeat_rate

        rows.append({
            "month": m,
            "date": pd.Timestamp(params.switch_year, 1, 1) + pd.DateOffset(months=m - 1),
            "months_from_switch": m,
            # Rx
            "rx_packs": round(rx_packs),
            "rx_price": round(rx_price, 2),
            "rx_revenue": round(rx_revenue),
            "rx_tablets": round(rx_tablets),
            # OTC
            "otc_packs": round(otc_packs),
            "otc_price": round(otc_price, 2),
            "otc_retail_revenue": round(otc_retail_revenue),
            "otc_manufacturer_revenue": round(otc_manufacturer_revenue),
            "otc_tablets": round(otc_tablets),
            "otc_from_rx_migration": round(otc_from_rx_migration),
            "otc_from_new_patients": round(otc_from_new_patients),
            # Combined
            "total_tablets": round(total_tablets),
            "total_revenue": round(total_revenue),
            "otc_share_tablets": otc_tablets / total_tablets if total_tablets > 0 else 0,
            "otc_share_revenue": otc_manufacturer_revenue / total_revenue if total_revenue > 0 else 0,
            # Seasonality & awareness
            "season_factor": season_factor,
            "awareness": awareness,
            # Adjacent
            "adjacent_revenue_lost": round(adjacent_revenue_lost),
            # Profitability
            "cogs": round(cogs),
            "marketing_spend": round(marketing_spend),
            "gross_profit": round(gross_profit),
            "operating_profit": round(operating_profit),
            "cumulative_otc_revenue": round(cumulative_otc_revenue),
            "cumulative_rx_revenue": round(cumulative_rx_revenue),
            "cumulative_total_revenue": round(cumulative_total_revenue),
            "cumulative_profit": round(cumulative_profit),
            "cumulative_marketing": round(cumulative_marketing),
            # Funnel
            "potential_users": round(potential_users),
            "aware_users": round(aware_users),
            "trialists": round(trialists),
            "repeaters": round(repeaters),
        })

    return pd.DataFrame(rows)


def calculate_kpis_rx_otc(df: pd.DataFrame) -> dict:
    """Calculate KPIs from Rx-to-OTC forecast."""
    last = df.iloc[-1]
    y1 = df[df["months_from_switch"] <= 12]
    y1_otc = y1["otc_manufacturer_revenue"].sum()
    y1_rx = y1["rx_revenue"].sum()

    # Find crossover month (OTC revenue > Rx revenue)
    crossover = df[df["otc_manufacturer_revenue"] > df["rx_revenue"]]
    crossover_month = int(crossover.iloc[0]["months_from_switch"]) if len(crossover) > 0 else None

    # Peak OTC month
    peak_otc_idx = df["otc_packs"].idxmax()
    peak_otc_month = int(df.loc[peak_otc_idx, "months_from_switch"])
    peak_otc_packs = int(df.loc[peak_otc_idx, "otc_packs"])

    # Rx decline
    rx_start = df.iloc[0]["rx_packs"]
    rx_end = last["rx_packs"]
    rx_decline_pct = (rx_start - rx_end) / rx_start if rx_start > 0 else 0

    # OTC share at various points
    m12 = df[df["months_from_switch"] == 12]
    m24 = df[df["months_from_switch"] == 24]
    m36 = df[df["months_from_switch"] == 36]

    return {
        "year1_otc_revenue": y1_otc,
        "year1_rx_revenue": y1_rx,
        "year1_total_revenue": y1_otc + y1_rx,
        "total_5y_revenue": last["cumulative_total_revenue"],
        "total_5y_otc_revenue": last["cumulative_otc_revenue"],
        "total_5y_rx_revenue": last["cumulative_rx_revenue"],
        "total_5y_profit": last["cumulative_profit"],
        "total_5y_marketing": last["cumulative_marketing"],
        "crossover_month": crossover_month,
        "peak_otc_month": peak_otc_month,
        "peak_otc_packs": peak_otc_packs,
        "rx_decline_total": rx_decline_pct,
        "otc_share_tablets_m12": float(m12["otc_share_tablets"].iloc[0]) if len(m12) > 0 else 0,
        "otc_share_tablets_m24": float(m24["otc_share_tablets"].iloc[0]) if len(m24) > 0 else 0,
        "otc_share_tablets_m36": float(m36["otc_share_tablets"].iloc[0]) if len(m36) > 0 else 0,
        "peak_awareness": float(df["awareness"].max()),
        "total_adjacent_lost": df["adjacent_revenue_lost"].sum(),
        "otc_packs_month_12": int(m12["otc_packs"].iloc[0]) if len(m12) > 0 else 0,
        "rx_packs_month_12": int(m12["rx_packs"].iloc[0]) if len(m12) > 0 else 0,
        "avg_otc_price": df["otc_price"].mean(),
    }
