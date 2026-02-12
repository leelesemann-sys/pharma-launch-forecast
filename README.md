# Pharma Launch Forecast

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://pharma-launch-forecast.streamlit.app/)

Interactive forecast models for pharmaceutical product launches in Germany. Built with Python, Streamlit, and Plotly.

> **Live Demo:** [pharma-launch-forecast.streamlit.app](https://pharma-launch-forecast.streamlit.app/)

Three fundamentally different forecasting engines demonstrate the breadth of launch scenarios a Strategic Portfolio Manager faces:

| | Use Case 1 | Use Case 2 | Use Case 3 |
|---|---|---|---|
| **Scenario** | Eliquis Generic Entry | GLP-1: Mounjaro vs. Ozempic | Rx-to-OTC Switch (PPI) |
| **Type** | Generic vs. Originator | Brand vs. Brand | Dual Channel (Rx + OTC) |
| **Market** | Mature, eroding | Expanding, high-growth | Channel transition |
| **Key Mechanic** | Aut-idem, Rabattvertraege | Indication layering, supply | Consumer awareness, pricing |
| **LOE** | May 2026 | n/a (patent protected) | n/a (OTC switch) |
| **Pricing** | Festbetrag / generic erosion | AMNOG Erstattungsbetrag | Free OTC pricing, pharmacy margin |

## Screenshots

Run the apps locally to explore the interactive dashboards (see [Getting Started](#getting-started)).

## Architecture

```
pharma-launch-forecast/
├── models/
│   ├── forecast_engine.py           # Generic Entry engine (Eliquis)
│   ├── brand_competition_engine.py  # Brand Competition engine (GLP-1)
│   └── rx_otc_engine.py             # Rx-to-OTC Switch engine (PPI)
├── app/
│   ├── app.py                       # Multi-page entry point (st.navigation)
│   ├── main.py                      # Page: Eliquis Generic Entry
│   ├── glp1.py                      # Page: GLP-1 Brand Competition
│   └── rx_otc.py                    # Page: Rx-to-OTC Switch
├── data/
│   ├── market_data.py               # NOAK/DOAK market data (synthetic)
│   └── glp1_market_data.py          # GLP-1 market data (synthetic)
├── exports/
│   ├── build_excel_model.py                    # Excel generator: Eliquis
│   ├── build_glp1_excel.py                     # Excel generator: GLP-1
│   ├── build_rx_otc_excel.py                   # Excel generator: Rx-to-OTC
│   ├── Eliquis_Launch_Forecast_v2.xlsx         # Pre-built Excel (6 sheets)
│   ├── GLP1_Brand_Competition_Forecast.xlsx    # Pre-built Excel (6 sheets)
│   └── RxToOTC_Switch_Forecast.xlsx            # Pre-built Excel (5 sheets)
└── requirements.txt
```

## Use Case 1: Eliquis Generic Entry (May 2026)

**Question:** What happens when the #1 anticoagulant loses patent protection?

Two perspectives with independent parameter sets:

- **Originator (BMS/Pfizer):** Revenue-at-Risk quantification. Models market share erosion via exponential decay with configurable floor, authorized generic strategy, and price defense.
- **Generic Challenger:** Market opportunity sizing. Models uptake via logistic S-curve with three volume sources:

| Volume Source | Mechanism | Regulatory Basis |
|---|---|---|
| Organic | Physician switching, formulary listing | Standard market access |
| Aut-idem | Pharmacy-level substitution | SGB V Par.129 (mandatory generic substitution) |
| Tender | Krankenkasse exclusive contracts | Par.130a SGB V (Rabattvertraege) |

**Key model features:**
- Aut-idem ramp curve tied to Festbetrag (reference pricing) timeline
- Per-Kasse tender modeling (TK, BARMER, DAK, AOKs) with win probabilities
- Volume decomposition: organic vs. aut-idem vs. tender
- Originator price defense and authorized generic cannibalization

## Use Case 2: GLP-1 Brand Competition

**Question:** Will Mounjaro overtake Ozempic in the German GLP-1 market?

Two perspectives with symmetric model structure:

- **Eli Lilly (Mounjaro/Tirzepatid):** Challenger gaining share with superior efficacy data
- **Novo Nordisk (Ozempic/Semaglutid):** Defending market leadership despite supply constraints

**Key model features:**
- Market expansion with logistic dampening (GLP-1 segment growing 25-36% p.a.)
- Indication multiplier system: T2D (base) + Obesity + CV risk + MASH/NASH
- Supply constraint modeling with capacity normalization timeline
- Scenario presets: Base Case, Bull (Lilly), Bull (Novo)
- AMNOG pricing with confidential reimbursement amounts

**Real-world data points incorporated:**
- 6.33M T2D patients in Germany, 13.9% on GLP-1 (BARMER 2025)
- Wegovy excluded from GKV coverage (G-BA June 2024: Lifestyle-Arzneimittel)
- Mounjaro: first-ever confidential Erstattungsbetrag in Germany (Aug 2025)
- Ozempic supply constraints since 2022

## Use Case 3: Rx-to-OTC Switch (PPI)

**Question:** What happens when a prescription drug becomes available over the counter?

Reference case: Omeprazol/Pantoprazol 20mg PPI switch in Germany (July 2009).

Two perspectives:

- **Manufacturer:** Dual-channel revenue modeling. Rx volume decays as OTC ramps up, but OTC reaches entirely new patients.
- **Market:** Category disruption. OTC launch cannibalized antacids and collapsed H2-antagonists.

**Key model features:**
- Dual-channel dynamics: Rx exponential decay + OTC logistic S-curve ramp
- Market expansion: 70% of OTC volume from genuinely new patients (never had Rx)
- Adjacent category cannibalization (antacids, H2-antagonists)
- Consumer awareness S-curve tied to marketing spend
- PPI-typical seasonality (GI complaints peak in autumn/winter)
- Price elasticity: patient pays 100% OTC (vs. GKV copay for Rx)
- Pharmacy margin structure: free OTC pricing vs. regulated Rx (AMPreisV)
- Three scenarios: Konservativ / Base Case / Optimistisch

**Real-world data points incorporated:**
- OTC market Germany 2024: EUR 10.15B (+7.1%, IQVIA)
- PPI OTC switch 2009: only ~3% of PPI volume moved to OTC (partial switch)
- 11.9% heartburn market expansion in year 1 post-switch
- Antacid cannibalization: -EUR 11M; H2-antagonists: -46%
- OTC PPI max 14 tablets/20mg per BfArM

## Excel Models

Pre-built Excel workbooks are included in [`exports/`](exports/) for direct download. Each has professionally formatted sheets:

1. **INPUTS** - All parameters in editable yellow cells
2. **Marktdaten** - Competitive landscape and market context
3. **Forecast** - Monthly forecast with full data table + embedded charts
4. **Dashboard** - Executive summary with scenario comparison
5. **Methodik** - Transparency matrix classifying every data point as FAKT / ANNAHME / MODELL / LUECKE

| File | Use Case | Sheets |
|---|---|---|
| [`Eliquis_Launch_Forecast_v2.xlsx`](exports/Eliquis_Launch_Forecast_v2.xlsx) | Generic Entry | 6 |
| [`GLP1_Brand_Competition_Forecast.xlsx`](exports/GLP1_Brand_Competition_Forecast.xlsx) | Brand Competition | 6 |
| [`RxToOTC_Switch_Forecast.xlsx`](exports/RxToOTC_Switch_Forecast.xlsx) | Rx-to-OTC Switch | 5 |

Regenerate them after modifying parameters:

```bash
python exports/build_excel_model.py      # Eliquis
python exports/build_glp1_excel.py       # GLP-1
python exports/build_rx_otc_excel.py     # Rx-to-OTC
```

## Getting Started

```bash
# Clone
git clone https://github.com/leelesemann-sys/pharma-launch-forecast.git
cd pharma-launch-forecast

# Install dependencies
pip install -r requirements.txt

# Run multi-page app (all 3 use cases)
streamlit run app/app.py
```

## Data Disclaimer

All data is synthetic and based on publicly available sources:

- BARMER Gesundheitswesen aktuell 2025
- G-BA Nutzenbewertung (Tirzepatid, Semaglutid)
- IQVIA Pharmamarkt Deutschland (public summaries)
- EMA European Public Assessment Reports
- GKV-Arzneimittelindex, Lauer-Taxe (public pricing)

No proprietary IQVIA data, no confidential company data. The models are designed to demonstrate forecasting methodology, not to provide investment advice.

## Tech Stack

- **Python 3.10+**
- **Streamlit** - Interactive dashboards
- **Plotly** - Charts (dual-axis, stacked areas, grouped bars)
- **XlsxWriter** - Professional Excel generation with formatting and embedded charts
- **Pandas / NumPy** - Data processing and mathematical models

## License

MIT
