"""
Sildenafil Rx-to-OTC Switch Forecast Engine (Viatris / Viagra)
================================================================
Advanced Rx-to-OTC model with omnichannel distribution modeling.

Key differences vs. generic PPI switch model (Use Case 3):
  - Patient already pays 100% Rx (ED not covered by GKV) → low price friction
  - Massive treatment gap: ~67% of ED patients never see a doctor (shame barrier)
  - OTC removes stigma barrier → huge market expansion (UK ref: 63% new patients)
  - Omnichannel: stationaere Apotheke vs. Online-Apotheke vs. Telemedizin
  - Telemedizin gets cannibalized (Zava, GoSpring lose Rx raison d'être)
  - Brand premium: Viagra vs. generic sildenafil OTC pricing power
  - Competitor dynamics: Tadalafil (Cialis) remains Rx-only → competitive moat

Reference case: UK Viagra Connect (launched March 2018)
  - 50mg OTC, pharmacy-only, GBP 5/tablet
  - Year 1: ~GBP 16M retail, 63% new-to-therapy patients
  - NHS Rx prescriptions ALSO increased (no cannibalization of Rx channel)

Market data (Germany):
  - 4-6M men with ED, only ~33% treated
  - ~2.6M PDE5 packs/year; Sildenafil 55-65% share
  - Viagra brand: ~10% of sildenafil packs (rest generics)
  - Generic Rx: EUR 0.95-2.30/tablet; Viagra brand: EUR 8-14/tablet
  - BfArM SVA rejected 3x (2022, 2023, 2025) but BMG wants switch
"""

from dataclasses import dataclass, field
from typing import Optional
import numpy as np
import pandas as pd


# ═══════════════════════════════════════════════════════════════════
# CHANNEL DEFINITIONS
# ═══════════════════════════════════════════════════════════════════

@dataclass
class ChannelParams:
    """Parameters for a single distribution channel."""
    name: str = "Stationaere Apotheke"
    share_of_otc: float = 0.55           # % of OTC volume through this channel
    share_trend_annual: float = -0.02    # annual shift (neg = losing share)
    margin_pct: float = 0.42            # pharmacy/platform margin
    distribution_cost_pct: float = 0.06  # logistics
    avg_basket_multiplier: float = 1.0   # impulse/cross-sell effect
    discretion_factor: float = 0.8       # 1.0 = max discretion, affects stigma-driven demand


@dataclass
class TelemedChannel:
    """Telemedizin channel that gets disrupted by OTC switch."""
    name: str = "Telemedizin (Zava, GoSpring etc.)"
    monthly_revenue_baseline: float = 4_500_000  # EUR/month pre-switch
    monthly_consultations: int = 75_000           # consultations/month
    avg_consultation_fee: float = 29.0            # EUR per consultation
    decline_rate: float = 0.60                    # revenue lost to OTC over time
    decline_months: int = 24                      # months for full disruption
    pivot_retention: float = 0.15                 # % retained via value-add services


# ═══════════════════════════════════════════════════════════════════
# MAIN MODEL PARAMETERS
# ═══════════════════════════════════════════════════════════════════

@dataclass
class SildenafilOtcParams:
    """Parameters for Sildenafil Rx-to-OTC switch forecast."""

    # ─── Product ────────────────────────────────────────────────
    product_name: str = "Viagra Connect 50mg"
    company: str = "Viatris"
    molecule: str = "Sildenafil"
    otc_dosage: str = "50mg"
    switch_year: int = 2027  # hypothetical switch date

    # ─── Epidemiology ───────────────────────────────────────────
    ed_prevalence_men: int = 5_000_000     # men with moderate-to-complete ED (DE)
    treatment_rate: float = 0.33           # currently treated (rest = treatment gap)
    addressable_otc_pct: float = 0.14      # % of untreated who would buy OTC in steady state

    # ─── Rx channel (pre-switch) ────────────────────────────────
    rx_packs_per_month: int = 217_000      # sildenafil Rx packs/month (65% of ~2.6M/yr ÷ 12)
    rx_price_brand: float = 11.19          # Viagra brand per tablet (50mg, 4er)
    rx_price_generic: float = 1.50         # average generic per tablet
    rx_brand_share: float = 0.10           # Viagra brand % of sildenafil packs
    rx_avg_pack_size: int = 4              # typical pack size
    rx_decline_rate: float = 0.08          # surprisingly low: UK showed Rx INCREASED
    rx_decline_months: int = 36            # slow decline (some patients prefer Rx relationship)

    # ─── OTC channel ────────────────────────────────────────────
    otc_price_per_tablet: float = 5.99     # Viagra Connect OTC per tablet
    otc_pack_sizes: list = field(default_factory=lambda: [4, 8, 12])
    otc_avg_pack_size: int = 6             # weighted average
    otc_peak_packs_per_month: int = 350_000  # peak OTC packs/month (all channels)
    otc_ramp_months: int = 18              # months to reach peak

    # ─── Market expansion (the core thesis) ─────────────────────
    new_patient_share: float = 0.63        # UK reference: 63% were new-to-therapy
    stigma_reduction_factor: float = 0.40  # OTC removes stigma → 40% of barrier gone
    # Treatment gap closure rate (how many untreated men start buying OTC)
    treatment_gap_closure_rate: float = 0.08  # 8% of untreated switch to OTC over 5 years

    # ─── Brand vs. Generic OTC ──────────────────────────────────
    brand_otc_share: float = 0.45          # Viagra Connect vs generic OTC sildenafil
    brand_otc_share_trend: float = -0.03   # annual erosion from generic OTC entrants
    brand_price_premium: float = 1.8       # brand is 1.8x generic OTC price

    # ─── Price dynamics ─────────────────────────────────────────
    price_elasticity: float = -0.5         # lower than PPI: ED patients less price-sensitive
    price_trend_annual: float = -0.03      # generic competition drives prices down
    # Note: lower elasticity because patients already pay 100% for Rx

    # ─── Omnichannel distribution ───────────────────────────────
    channels: list = field(default_factory=lambda: [
        ChannelParams(
            name="Stationaere Apotheke",
            share_of_otc=0.50,
            share_trend_annual=-0.03,
            margin_pct=0.42,
            distribution_cost_pct=0.06,
            avg_basket_multiplier=1.0,
            discretion_factor=0.70,     # lower: face-to-face = more stigma
        ),
        ChannelParams(
            name="Online-Apotheke",
            share_of_otc=0.40,
            share_trend_annual=0.04,    # growing channel
            margin_pct=0.30,            # lower margin, competitive pricing
            distribution_cost_pct=0.10, # shipping costs
            avg_basket_multiplier=1.15, # cross-sell via recommendation engines
            discretion_factor=1.0,      # maximum: anonymous ordering
        ),
        ChannelParams(
            name="Drogerie-/Supermarkt (falls freiverkaeuflich)",
            share_of_otc=0.10,
            share_trend_annual=0.01,
            margin_pct=0.35,
            distribution_cost_pct=0.08,
            avg_basket_multiplier=0.9,
            discretion_factor=0.50,     # lowest: public checkout
        ),
    ])

    # ─── Telemedizin disruption ─────────────────────────────────
    telemed: TelemedChannel = field(default_factory=TelemedChannel)

    # ─── Competitor (Tadalafil stays Rx) ────────────────────────
    tadalafil_rx_monthly: int = 120_000    # Tadalafil Rx packs/month
    tadalafil_switch_to_sildenafil_otc: float = 0.12  # % Tadalafil users switching to OTC Sildenafil
    tadalafil_migration_months: int = 24

    # ─── Seasonality (ED: less seasonal than PPI) ───────────────
    seasonality: list = field(default_factory=lambda: [
        0.90, 1.05, 1.00, 1.00, 1.05, 1.10,   # Jan-Jun (Feb=Valentine's, Jun=summer)
        1.05, 1.00, 0.95, 0.95, 0.95, 1.00,    # Jul-Dec
    ])

    # ─── Marketing & awareness ──────────────────────────────────
    marketing_monthly_eur: float = 500_000   # higher than PPI: DTC brand building
    awareness_peak: float = 0.75             # Viagra already high awareness
    awareness_baseline: float = 0.60         # Viagra brand awareness pre-switch (already known)
    awareness_ramp_months: int = 9           # fast: brand already established

    # ─── Consumer funnel ────────────────────────────────────────
    trial_rate: float = 0.25                 # % of aware men with ED who try OTC
    repeat_rate: float = 0.65               # higher than PPI: chronic/recurring condition

    # ─── Costs ──────────────────────────────────────────────────
    cogs_pct: float = 0.12                   # low COGS for established molecule
    # Channel-specific margins are in ChannelParams

    # ─── Forecast ───────────────────────────────────────────────
    forecast_months: int = 60


# ═══════════════════════════════════════════════════════════════════
# MATHEMATICAL PRIMITIVES
# ═══════════════════════════════════════════════════════════════════

def _logistic(t: float, peak: float, midpoint: float, steepness: float) -> float:
    """Generic logistic S-curve."""
    return peak / (1.0 + np.exp(-steepness * (t - midpoint)))


def _awareness_curve(month: int, baseline: float, peak: float, ramp: int) -> float:
    """Awareness with baseline (Viagra is already known)."""
    if ramp <= 0:
        return peak
    delta = peak - baseline
    midpoint = ramp / 2
    steepness = 6.0 / ramp
    return baseline + delta / (1.0 + np.exp(-steepness * (month - midpoint)))


def _otc_ramp(month: int, peak: int, ramp_months: int) -> float:
    """Logistic ramp for OTC volume."""
    if ramp_months <= 0:
        return float(peak)
    midpoint = ramp_months * 0.45
    steepness = 6.0 / ramp_months
    return peak / (1.0 + np.exp(-steepness * (month - midpoint)))


def _rx_effect(month: int, initial: int, decline_rate: float, decline_months: int) -> float:
    """Rx volume with slow decline (UK reference: Rx may even grow initially)."""
    if decline_months <= 0 or decline_rate <= 0:
        return float(initial)
    k = -np.log(1 - decline_rate) / decline_months
    floor = initial * (1 - decline_rate)
    return floor + (initial - floor) * np.exp(-k * month)


def _telemed_disruption(month: int, baseline: float, decline_rate: float,
                         decline_months: int, pivot_retention: float) -> float:
    """Telemedizin revenue decline as OTC removes need for Rx."""
    if decline_months <= 0:
        return baseline * pivot_retention
    k = -np.log(1 - decline_rate) / decline_months
    floor = baseline * pivot_retention
    return floor + (baseline - floor) * np.exp(-k * month)


def _channel_share(base_share: float, trend: float, month: int) -> float:
    """Channel share evolution over time."""
    return max(0.0, min(1.0, base_share + trend * (month / 12.0)))


# ═══════════════════════════════════════════════════════════════════
# FORECAST ENGINE
# ═══════════════════════════════════════════════════════════════════

def forecast_sildenafil_otc(
    params: SildenafilOtcParams,
    forecast_months: Optional[int] = None,
) -> pd.DataFrame:
    """Run the Sildenafil Rx-to-OTC switch forecast with omnichannel modeling.

    Returns DataFrame with monthly data including:
    - Rx channel performance
    - OTC channel by distribution channel (Apotheke, Online, Drogerie)
    - Brand vs. generic OTC split
    - Telemedizin disruption
    - Tadalafil migration
    - Consumer funnel metrics
    - Profitability by channel
    """
    months = forecast_months or params.forecast_months

    rows = []
    cum = {
        "otc_revenue": 0.0, "rx_revenue": 0.0, "total_revenue": 0.0,
        "profit": 0.0, "marketing": 0.0,
    }

    for m in range(1, months + 1):
        year_frac = m / 12.0
        season_idx = (m - 1) % 12
        season_factor = params.seasonality[season_idx]

        # ─── Awareness (starts high for Viagra) ────────────────
        awareness = _awareness_curve(
            m, params.awareness_baseline,
            params.awareness_peak, params.awareness_ramp_months
        )

        # ─── Rx Channel ────────────────────────────────────────
        rx_packs = _rx_effect(
            m, params.rx_packs_per_month,
            params.rx_decline_rate, params.rx_decline_months
        )
        # Weighted average Rx price
        rx_avg_price = (
            params.rx_brand_share * params.rx_price_brand * params.rx_avg_pack_size +
            (1 - params.rx_brand_share) * params.rx_price_generic * params.rx_avg_pack_size
        )
        rx_revenue = rx_packs * rx_avg_price

        # ─── OTC Volume (total) ────────────────────────────────
        otc_base = _otc_ramp(m, params.otc_peak_packs_per_month, params.otc_ramp_months)
        otc_seasonal = otc_base * season_factor

        # Awareness scaling
        aw_ratio = awareness / params.awareness_peak if params.awareness_peak > 0 else 1.0
        otc_demand = otc_seasonal * max(0.3, aw_ratio)  # floor at 30% (Viagra is known)

        # Price effect
        price_change = params.price_trend_annual * year_frac
        price_vol_effect = 1.0 + (price_change * params.price_elasticity)
        otc_total_packs = max(0, otc_demand * price_vol_effect)

        # ─── Tadalafil migration bonus ─────────────────────────
        tada_migration = _logistic(
            m, params.tadalafil_rx_monthly * params.tadalafil_switch_to_sildenafil_otc,
            params.tadalafil_migration_months * 0.5,
            6.0 / params.tadalafil_migration_months
        )
        otc_total_packs += tada_migration

        # ─── Omnichannel distribution ──────────────────────────
        # Calculate share for each channel with time evolution
        raw_shares = []
        for ch in params.channels:
            raw_shares.append(_channel_share(ch.share_of_otc, ch.share_trend_annual, m))
        # Normalize to sum to 1.0
        share_sum = sum(raw_shares)
        norm_shares = [s / share_sum for s in raw_shares] if share_sum > 0 else raw_shares

        channel_data = {}
        total_otc_manufacturer_revenue = 0.0
        total_otc_retail_revenue = 0.0

        for i, ch in enumerate(params.channels):
            ch_share = norm_shares[i]
            ch_packs = otc_total_packs * ch_share

            # Discretion effect on volume: channels with more privacy get bonus
            discretion_bonus = 1.0 + (ch.discretion_factor - 0.7) * 0.15
            ch_packs *= discretion_bonus

            # Current OTC price with erosion
            otc_price_now = params.otc_price_per_tablet * (
                1 + params.price_trend_annual
            ) ** year_frac
            ch_retail_rev = ch_packs * otc_price_now * params.otc_avg_pack_size

            # Manufacturer revenue = retail - channel margin - distribution
            mfr_share = 1.0 - ch.margin_pct - ch.distribution_cost_pct
            ch_mfr_rev = ch_retail_rev * mfr_share

            total_otc_retail_revenue += ch_retail_rev
            total_otc_manufacturer_revenue += ch_mfr_rev

            channel_data[ch.name] = {
                "packs": round(ch_packs),
                "share": ch_share,
                "retail_revenue": round(ch_retail_rev),
                "manufacturer_revenue": round(ch_mfr_rev),
            }

        # Recalculate total OTC packs (after discretion bonus)
        otc_total_packs_adj = sum(cd["packs"] for cd in channel_data.values())

        # ─── Brand vs. Generic OTC ─────────────────────────────
        brand_share = max(0.15, params.brand_otc_share + params.brand_otc_share_trend * year_frac)
        otc_brand_packs = otc_total_packs_adj * brand_share
        otc_generic_packs = otc_total_packs_adj * (1 - brand_share)

        # Brand premium: Viagra Connect sells at a premium vs. generic OTC.
        # Recalculate manufacturer revenue accounting for brand/generic price split.
        otc_price_now = params.otc_price_per_tablet * (
            1 + params.price_trend_annual
        ) ** year_frac
        generic_price = otc_price_now
        brand_price = otc_price_now * params.brand_price_premium
        # Weighted average revenue: brand packs × brand price + generic packs × generic price
        brand_rev_bonus = (
            otc_brand_packs * (brand_price - generic_price) * params.otc_avg_pack_size
        )
        # Apply average manufacturer share across channels
        avg_mfr_pct = total_otc_manufacturer_revenue / total_otc_retail_revenue if total_otc_retail_revenue > 0 else 0.52
        total_otc_manufacturer_revenue += brand_rev_bonus * avg_mfr_pct
        total_otc_retail_revenue += brand_rev_bonus

        # ─── Volume decomposition ──────────────────────────────
        otc_from_new_patients = otc_total_packs_adj * params.new_patient_share
        otc_from_rx_migration = otc_total_packs_adj * (1 - params.new_patient_share) * 0.7
        otc_from_tadalafil = tada_migration
        otc_from_other = otc_total_packs_adj - otc_from_new_patients - otc_from_rx_migration - otc_from_tadalafil

        # ─── Tablets for comparison ─────────────────────────────
        rx_tablets = rx_packs * params.rx_avg_pack_size
        otc_tablets = otc_total_packs_adj * params.otc_avg_pack_size
        total_tablets = rx_tablets + otc_tablets

        # ─── Telemedizin disruption ─────────────────────────────
        telemed_rev = _telemed_disruption(
            m, params.telemed.monthly_revenue_baseline,
            params.telemed.decline_rate, params.telemed.decline_months,
            params.telemed.pivot_retention
        )
        telemed_lost = params.telemed.monthly_revenue_baseline - telemed_rev

        # ─── Marketing ──────────────────────────────────────────
        marketing_spend = params.marketing_monthly_eur
        if m > params.awareness_ramp_months * 2:
            marketing_spend *= 0.5  # maintenance mode

        # ─── Profitability ──────────────────────────────────────
        total_revenue = rx_revenue + total_otc_manufacturer_revenue
        cogs = total_revenue * params.cogs_pct
        gross_profit = total_revenue - cogs
        operating_profit = gross_profit - marketing_spend

        cum["otc_revenue"] += total_otc_manufacturer_revenue
        cum["rx_revenue"] += rx_revenue
        cum["total_revenue"] += total_revenue
        cum["profit"] += operating_profit
        cum["marketing"] += marketing_spend

        # ─── Treatment gap metric ──────────────────────────────
        untreated = params.ed_prevalence_men * (1 - params.treatment_rate)
        newly_treated_cumulative = min(
            untreated * params.treatment_gap_closure_rate,
            untreated * params.treatment_gap_closure_rate * min(1.0, m / 36)
        )
        treatment_rate_new = params.treatment_rate + (
            newly_treated_cumulative / params.ed_prevalence_men
        )

        # ─── OTC price ─────────────────────────────────────────
        otc_price_now = params.otc_price_per_tablet * (
            1 + params.price_trend_annual
        ) ** year_frac

        rows.append({
            "month": m,
            "date": pd.Timestamp(params.switch_year, 1, 1) + pd.DateOffset(months=m - 1),

            # Rx
            "rx_packs": round(rx_packs),
            "rx_revenue": round(rx_revenue),
            "rx_tablets": round(rx_tablets),

            # OTC total
            "otc_packs": round(otc_total_packs_adj),
            "otc_price_per_tablet": round(otc_price_now, 2),
            "otc_retail_revenue": round(total_otc_retail_revenue),
            "otc_manufacturer_revenue": round(total_otc_manufacturer_revenue),
            "otc_tablets": round(otc_tablets),

            # OTC by channel
            "ch_apotheke_packs": channel_data[params.channels[0].name]["packs"],
            "ch_apotheke_revenue": channel_data[params.channels[0].name]["manufacturer_revenue"],
            "ch_apotheke_share": channel_data[params.channels[0].name]["share"],
            "ch_online_packs": channel_data[params.channels[1].name]["packs"],
            "ch_online_revenue": channel_data[params.channels[1].name]["manufacturer_revenue"],
            "ch_online_share": channel_data[params.channels[1].name]["share"],
            "ch_drogerie_packs": channel_data[params.channels[2].name]["packs"],
            "ch_drogerie_revenue": channel_data[params.channels[2].name]["manufacturer_revenue"],
            "ch_drogerie_share": channel_data[params.channels[2].name]["share"],

            # Brand vs Generic OTC
            "otc_brand_packs": round(otc_brand_packs),
            "otc_generic_packs": round(otc_generic_packs),
            "otc_brand_share": brand_share,

            # Volume decomposition
            "otc_from_new_patients": round(otc_from_new_patients),
            "otc_from_rx_migration": round(otc_from_rx_migration),
            "otc_from_tadalafil": round(otc_from_tadalafil),

            # Combined
            "total_tablets": round(total_tablets),
            "total_revenue": round(total_revenue),
            "otc_share_tablets": otc_tablets / total_tablets if total_tablets > 0 else 0,

            # Awareness & seasonality
            "awareness": awareness,
            "season_factor": season_factor,

            # Telemedizin
            "telemed_revenue": round(telemed_rev),
            "telemed_lost": round(telemed_lost),

            # Profitability
            "cogs": round(cogs),
            "marketing_spend": round(marketing_spend),
            "gross_profit": round(gross_profit),
            "operating_profit": round(operating_profit),

            # Cumulative
            "cumulative_total_revenue": round(cum["total_revenue"]),
            "cumulative_otc_revenue": round(cum["otc_revenue"]),
            "cumulative_rx_revenue": round(cum["rx_revenue"]),
            "cumulative_profit": round(cum["profit"]),
            "cumulative_marketing": round(cum["marketing"]),

            # Treatment gap
            "treatment_rate_effective": treatment_rate_new,
            "newly_treated_cumulative": round(newly_treated_cumulative),
        })

    return pd.DataFrame(rows)


def calculate_kpis_sildenafil(df: pd.DataFrame) -> dict:
    """Calculate KPIs from Sildenafil OTC forecast."""
    last = df.iloc[-1]
    y1 = df[df["month"] <= 12]

    # Crossover month (OTC mfr revenue > Rx revenue)
    crossover = df[df["otc_manufacturer_revenue"] > df["rx_revenue"]]
    crossover_month = int(crossover.iloc[0]["month"]) if len(crossover) > 0 else None

    # Peak OTC
    peak_idx = df["otc_packs"].idxmax()
    peak_otc_packs = int(df.loc[peak_idx, "otc_packs"])
    peak_otc_month = int(df.loc[peak_idx, "month"])

    # Rx decline
    rx_start = df.iloc[0]["rx_packs"]
    rx_end = last["rx_packs"]
    rx_decline_pct = (rx_start - rx_end) / rx_start if rx_start > 0 else 0

    # Channel shares at month 12
    m12 = df[df["month"] == 12]
    m24 = df[df["month"] == 24]

    # Online share evolution
    online_share_m12 = float(m12["ch_online_share"].iloc[0]) if len(m12) > 0 else 0
    online_share_m24 = float(m24["ch_online_share"].iloc[0]) if len(m24) > 0 else 0

    return {
        "year1_otc_revenue": y1["otc_manufacturer_revenue"].sum(),
        "year1_rx_revenue": y1["rx_revenue"].sum(),
        "year1_total_revenue": y1["total_revenue"].sum(),
        "total_5y_revenue": last["cumulative_total_revenue"],
        "total_5y_otc_revenue": last["cumulative_otc_revenue"],
        "total_5y_rx_revenue": last["cumulative_rx_revenue"],
        "total_5y_profit": last["cumulative_profit"],
        "total_5y_marketing": last["cumulative_marketing"],
        "crossover_month": crossover_month,
        "peak_otc_packs": peak_otc_packs,
        "peak_otc_month": peak_otc_month,
        "rx_decline_total": rx_decline_pct,
        "otc_share_m12": float(m12["otc_share_tablets"].iloc[0]) if len(m12) > 0 else 0,
        "otc_share_m24": float(m24["otc_share_tablets"].iloc[0]) if len(m24) > 0 else 0,
        "brand_share_m12": float(m12["otc_brand_share"].iloc[0]) if len(m12) > 0 else 0,
        "brand_share_m24": float(m24["otc_brand_share"].iloc[0]) if len(m24) > 0 else 0,
        "online_share_m12": online_share_m12,
        "online_share_m24": online_share_m24,
        "telemed_lost_5y": df["telemed_lost"].sum(),
        "treatment_rate_final": last["treatment_rate_effective"],
        "newly_treated_total": last["newly_treated_cumulative"],
        "peak_awareness": float(df["awareness"].max()),
    }
