"""
Build a professional Excel forecast model for the Rx-to-OTC Switch scenario.

Sheets:
1. INPUTS          – All user-configurable parameters (yellow cells)
2. Marktdaten      – OTC market context, PPI landscape, adjacent categories
3. Forecast        – Monthly dual-channel forecast (Rx + OTC)
4. Dashboard       – Executive summary with scenario comparison
5. Methodik        – Transparency matrix: Facts vs. Assumptions vs. Models
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import xlsxwriter
import numpy as np
import pandas as pd
from models.rx_otc_engine import (
    RxOtcParams, AdjacentCategory,
    forecast_rx_otc, calculate_kpis_rx_otc,
)

OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "RxToOTC_Switch_Forecast.xlsx")


def build_model():
    wb = xlsxwriter.Workbook(OUTPUT_PATH)

    # ─── Format Definitions ─────────────────────────────────────────
    fmt = {}
    fmt["title"] = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1e3a5f", "bottom": 2, "bottom_color": "#1e3a5f"})
    fmt["section"] = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1e3a5f", "bg_color": "#e8f0fe", "border": 1, "border_color": "#b0c4de"})
    fmt["subsection"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "#374151", "bottom": 1, "bottom_color": "#d1d5db"})
    fmt["input"] = wb.add_format({"bg_color": "#FFF9C4", "border": 1, "border_color": "#E0B800", "num_format": "#,##0", "font_color": "#1a237e", "bold": True})
    fmt["input_pct"] = wb.add_format({"bg_color": "#FFF9C4", "border": 1, "border_color": "#E0B800", "num_format": "0.0%", "font_color": "#1a237e", "bold": True})
    fmt["input_eur"] = wb.add_format({"bg_color": "#FFF9C4", "border": 1, "border_color": "#E0B800", "num_format": "€#,##0.00", "font_color": "#1a237e", "bold": True})
    fmt["input_text"] = wb.add_format({"bg_color": "#FFF9C4", "border": 1, "border_color": "#E0B800", "font_color": "#1a237e", "bold": True})
    fmt["input_bool"] = wb.add_format({"bg_color": "#FFF9C4", "border": 1, "border_color": "#E0B800", "font_color": "#1a237e", "bold": True, "align": "center"})
    fmt["label"] = wb.add_format({"font_size": 10, "font_color": "#374151", "text_wrap": True, "valign": "vcenter"})
    fmt["label_bold"] = wb.add_format({"font_size": 10, "font_color": "#374151", "bold": True, "valign": "vcenter"})
    fmt["hint"] = wb.add_format({"font_size": 9, "font_color": "#9ca3af", "italic": True, "text_wrap": True, "valign": "vcenter"})
    fmt["number"] = wb.add_format({"num_format": "#,##0"})
    fmt["eur"] = wb.add_format({"num_format": "€#,##0"})
    fmt["eur_detail"] = wb.add_format({"num_format": "€#,##0.00"})
    fmt["pct"] = wb.add_format({"num_format": "0.0%"})
    fmt["mult"] = wb.add_format({"num_format": "0.00"})
    fmt["th"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "white", "bg_color": "#1e3a5f", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"})
    fmt["th_teal"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "white", "bg_color": "#134e4a", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"})
    fmt["th_blue"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "white", "bg_color": "#1e40af", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"})
    fmt["kpi_label"] = wb.add_format({"bold": True, "font_size": 11, "font_color": "#1e3a5f", "bg_color": "#f0f4ff", "border": 1})
    fmt["kpi_value"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1e3a5f", "bg_color": "#f0f4ff", "border": 1, "num_format": "€#,##0", "align": "center"})
    fmt["kpi_value_pct"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#134e4a", "bg_color": "#f0fdf4", "border": 1, "num_format": "0.0%", "align": "center"})
    fmt["kpi_value_teal"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#134e4a", "bg_color": "#ccfbf1", "border": 1, "num_format": "€#,##0", "align": "center"})
    fmt["kpi_value_plain"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1e3a5f", "bg_color": "#f0f4ff", "border": 1, "align": "center"})
    fmt["row_alt"] = wb.add_format({"bg_color": "#f8fafc"})
    fmt["row_alt_eur"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "€#,##0"})
    fmt["row_alt_pct"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "0.0%"})
    fmt["row_alt_num"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "#,##0"})
    fmt["row_alt_mult"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "0.00"})
    fmt["fact"] = wb.add_format({"bg_color": "#dcfce7", "border": 1, "text_wrap": True, "valign": "vcenter", "font_size": 10})
    fmt["assumption"] = wb.add_format({"bg_color": "#FFF9C4", "border": 1, "text_wrap": True, "valign": "vcenter", "font_size": 10})
    fmt["model"] = wb.add_format({"bg_color": "#dbeafe", "border": 1, "text_wrap": True, "valign": "vcenter", "font_size": 10})
    fmt["gap"] = wb.add_format({"bg_color": "#fef2f2", "border": 1, "text_wrap": True, "valign": "vcenter", "font_size": 10})

    def write_input_row(ws, row, label, value, fmt_key, hint):
        ws.write(row, 1, label, fmt["label_bold"])
        ws.write(row, 2, value, fmt[fmt_key])
        ws.write(row, 3, hint, fmt["hint"])

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 1: INPUTS
    # ═══════════════════════════════════════════════════════════════════
    ws = wb.add_worksheet("INPUTS")
    ws.hide_gridlines(2)
    ws.set_tab_color("#E0B800")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 44)
    ws.set_column("C:C", 18)
    ws.set_column("D:D", 55)

    row = 1
    ws.merge_range(row, 1, row, 3, "Rx-to-OTC Switch - Forecast Modell", fmt["title"])
    row += 1
    ws.merge_range(row, 1, row, 3, "Referenzfall: Omeprazol/Pantoprazol 20mg PPI Switch (DE 2009)", wb.add_format({"font_size": 10, "font_color": "#6b7280", "italic": True}))

    row += 2
    ws.merge_range(row, 1, row, 3, "1. PRODUKT & MARKT", fmt["section"])
    for label, value, fk, hint in [
        ("Produktname", "Omeprazol 20mg (14 St.)", "input_text", "OTC-Konfiguration: niedrige Dosis, kleine Packung"),
        ("Kategorie", "PPI (Protonenpumpeninhibitoren)", "input_text", "ATC-Klasse"),
        ("Indikation", "Sodbrennen & Reflux", "input_text", "OTC-Indikation (eingeschraenkt vs. Rx)"),
        ("Geographie", "Deutschland", "input_text", "Zielmarkt"),
        ("Switch-Datum", "2026", "input_text", "Erstverkauf OTC"),
        ("Regulatorik", "Apothekenpflichtig (nicht freiverkaeuflich)", "input_text", "BfArM Sachverstaendigenausschuss"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    row += 2
    ws.merge_range(row, 1, row, 3, "2. RX-KANAL (vor Switch)", fmt["section"])
    for label, value, fk, hint in [
        ("Rx-Packungen/Monat", 350_000, "input", "Monatliches Rx-Volumen (50 St. Packungen)"),
        ("Rx-Preis/Packung (Festbetrag, EUR)", 16.99, "input_eur", "GKV-Erstattung (50 Tabletten)"),
        ("Patient Zuzahlung (EUR)", 5.00, "input_eur", "Patienten-Eigenanteil bei Rx"),
        ("Rx-Abwanderung zu OTC (%)", 0.15, "input_pct", "Anteil Rx-Volumen, das zu OTC migriert"),
        ("Abwanderungszeitraum (Monate)", 24, "input", "Monate bis Rx-Rueckgang stabilisiert"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    row += 2
    ws.merge_range(row, 1, row, 3, "3. OTC-KANAL", fmt["section"])
    for label, value, fk, hint in [
        ("OTC-Preis/Packung (EUR)", 7.99, "input_eur", "Endverbraucherpreis (14 Tabletten)"),
        ("OTC Peak Packungen/Monat", 280_000, "input", "Maximal erreichbares OTC-Volumen"),
        ("Monate bis Peak", 18, "input", "Ramp-up-Dauer"),
        ("OTC-Packungsgroesse (Tabletten)", 14, "input", "Max. 14 St. fuer OTC PPI"),
        ("Neue Patienten (% des OTC-Vol.)", 0.70, "input_pct", "Patienten, die nie Rx hatten (Marktexpansion)"),
        ("Preiselastizitaet", -0.80, "input_eur", "1% Preiserhoehung = 0.8% Volumenrueckgang"),
        ("Preistrend p.a.", -0.02, "input_pct", "Jaehrl. Preisveraenderung (Generika-Druck)"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    row += 2
    ws.merge_range(row, 1, row, 3, "4. MARKETING & AWARENESS", fmt["section"])
    for label, value, fk, hint in [
        ("Marketing-Budget/Monat (EUR)", 300_000, "input", "TV, Digital, POS, Apotheken-Kooperationen"),
        ("Peak Awareness (%)", 0.65, "input_pct", "Maximale Markenbekanntheit"),
        ("Awareness-Aufbau (Monate)", 12, "input", "S-Kurve bis Peak Awareness"),
        ("Trial Rate (%)", 0.30, "input_pct", "% der bewussten Konsumenten, die kaufen"),
        ("Repeat Rate (%)", 0.55, "input_pct", "% der Kaeufer, die wiederholt kaufen"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    row += 2
    ws.merge_range(row, 1, row, 3, "5. KOSTEN & MARGEN", fmt["section"])
    for label, value, fk, hint in [
        ("COGS (% vom Umsatz)", 0.25, "input_pct", "Herstellkosten"),
        ("Apothekenmarge (%)", 0.40, "input_pct", "OTC: freie Preisgestaltung, ~40% Marge"),
        ("Distribution (%)", 0.08, "input_pct", "Grosshandel + Logistik"),
        ("Herstelleranteil", 0.52, "input_pct", "= 1 - Apotheke - Distribution (berechnet)"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    row += 2
    ws.merge_range(row, 1, row, 3, "6. SAISONALITAET (PPI-typisch)", fmt["section"])
    row += 1
    months_de = ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
    seasonal = [0.90, 0.85, 0.95, 1.05, 1.05, 1.00, 0.95, 0.90, 1.00, 1.10, 1.15, 1.10]
    for c, m in enumerate(months_de):
        ws.write(row, 1 + c, m, fmt["th"])
    row += 1
    for c, s in enumerate(seasonal):
        ws.write(row, 1 + c, s, fmt["input"])

    row += 2
    ws.merge_range(row, 1, row, 4, "7. SZENARIO-DEFINITIONEN", fmt["section"])
    row += 1
    for c, h in enumerate(["Parameter", "Konservativ", "Base Case", "Optimistisch"]):
        ws.write(row, 1 + c, h, fmt["th"])
    for param, cons, base, opti in [
        ("OTC Peak Packungen/Mon.", 180_000, 280_000, 400_000),
        ("OTC Preis/Packung (EUR)", 6.99, 7.99, 8.99),
        ("Monate bis Peak", 24, 18, 12),
        ("Marketing/Monat (EUR)", 200_000, 300_000, 500_000),
        ("Rx-Abwanderung", 0.10, 0.15, 0.20),
        ("Neue Patienten (%)", 0.60, 0.70, 0.75),
    ]:
        row += 1
        ws.write(row, 1, param, fmt["label_bold"])
        for c, val in enumerate([cons, base, opti]):
            if isinstance(val, float) and val < 1:
                ws.write(row, 2 + c, val, fmt["input_pct"])
            else:
                ws.write(row, 2 + c, val, fmt["input"])

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 2: MARKTDATEN
    # ═══════════════════════════════════════════════════════════════════
    ws2 = wb.add_worksheet("Marktdaten")
    ws2.hide_gridlines(2)
    ws2.set_tab_color("#1e3a5f")
    ws2.set_column("A:A", 3)
    ws2.set_column("B:B", 34)
    ws2.set_column("C:G", 18)

    row = 1
    ws2.merge_range(row, 1, row, 5, "OTC-Markt Deutschland & PPI-Landschaft", fmt["title"])

    row += 2
    ws2.merge_range(row, 1, row, 5, "OTC-Markt Deutschland (2024)", fmt["section"])
    for label, value in [
        ("OTC-Apothekenmarkt Umsatz", "10.15 Mrd. EUR (+7.1%)"),
        ("Selbstmedikationsmarkt gesamt", "13.8 Mrd. EUR (+7%)"),
        ("Packungen gesamt", "1.7 Mrd. (+3%)"),
        ("Groesster OTC-Markt in Europa", "Ja"),
        ("Versandhandel-Anteil", "~20% der Packungen, ~16% des Umsatzes"),
        ("Klassifizierung", "Verschreibungspflichtig > Apothekenpflichtig > Freiverkaeuflich"),
    ]:
        row += 1
        ws2.write(row, 1, label, fmt["label_bold"])
        ws2.merge_range(row, 2, row, 5, value, fmt["label"])

    row += 2
    ws2.merge_range(row, 1, row, 5, "PPI-Wettbewerbslandschaft (OTC)", fmt["section"])
    row += 1
    for c, h in enumerate(["Produkt", "Hersteller", "OTC-Preis\n(14 St.)", "Status", "Marktanteil\n(OTC)"]):
        ws2.write(row, 1 + c, h, fmt["th"])
    ws2.set_row(row, 30)
    products = [
        ("Omeprazol-ratiopharm SK", "Ratiopharm/Teva", 7.99, "Generikum", 0.35),
        ("Omep Hexal", "Hexal/Sandoz", 8.49, "Generikum", 0.15),
        ("Pantozol Control", "Nycomed/Takeda", 12.99, "Erstanbieter", 0.12),
        ("Pantoprazol-Generika", "Diverse", 3.50, "Generikum", 0.25),
        ("Omeprazol AL", "Aliud/STADA", 6.49, "Generikum", 0.08),
        ("Sonstige", "Diverse", 5.00, "Generikum", 0.05),
    ]
    for pn, co, pr, st, sh in products:
        row += 1
        ws2.write(row, 1, pn, fmt["label_bold"])
        ws2.write(row, 2, co, fmt["label"])
        ws2.write(row, 3, pr, fmt["eur_detail"])
        ws2.write(row, 4, st, fmt["label"])
        ws2.write(row, 5, sh, fmt["pct"])

    row += 2
    ws2.merge_range(row, 1, row, 5, "Kannibalisierte Kategorien", fmt["section"])
    row += 1
    for c, h in enumerate(["Kategorie", "Umsatz/Mon. (EUR)", "Kannibalis.\nPeak", "Beispielprodukte"]):
        ws2.write(row, 1 + c, h, fmt["th"])
    categories = [
        ("Antazida", 7_750_000, 0.15, "Maaloxan, Rennie, Talcid"),
        ("H2-Rezeptor-Antagonisten", 120_000, 0.46, "Ranitic, Zantic (auslaufend)"),
        ("Pflanzliche Mittel", 3_200_000, 0.05, "Iberogast, Gaviscon"),
    ]
    for cat, rev, cann, ex in categories:
        row += 1
        ws2.write(row, 1, cat, fmt["label_bold"])
        ws2.write(row, 2, rev, fmt["eur"])
        ws2.write(row, 3, cann, fmt["pct"])
        ws2.write(row, 4, ex, fmt["label"])

    row += 2
    ws2.merge_range(row, 1, row, 5, "Apotheken-Margenstruktur OTC", fmt["section"])
    for label, value in [
        ("Rx-Marge (reguliert)", "3% + EUR 8.35 Fixzuschlag/Packung (AMPreisV)"),
        ("OTC-Marge (FREI)", "Freie Preisgestaltung, typisch 40-60% Aufschlag"),
        ("Online-Apotheken", "20-40% guenstiger als stationaer"),
        ("Preisspanne Cetirizin 100 St.", "EUR 3.49 (Generikum) bis EUR 49.13 (Zyrtec)"),
    ]:
        row += 1
        ws2.write(row, 1, label, fmt["label_bold"])
        ws2.merge_range(row, 2, row, 5, value, fmt["label"])

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 3: FORECAST
    # ═══════════════════════════════════════════════════════════════════
    ws3 = wb.add_worksheet("Forecast")
    ws3.hide_gridlines(2)
    ws3.set_tab_color("#0d9488")
    ws3.set_column("A:A", 3)
    ws3.set_column("B:B", 12)
    ws3.set_column("C:R", 14)

    params = RxOtcParams()
    adj = AdjacentCategory()
    df = forecast_rx_otc(params, adj, forecast_months=60)
    kpis = calculate_kpis_rx_otc(df)

    row = 1
    ws3.merge_range(row, 1, row, 12, "Rx-to-OTC Switch Forecast - Dual Channel (Base Case)", fmt["title"])

    row += 2
    co = kpis.get("crossover_month")
    co_text = f"Monat {co}" if co else "–"
    kpi_data = [
        ("OTC Umsatz Jahr 1", kpis["year1_otc_revenue"], "kpi_value_teal"),
        ("Gesamtumsatz 5J", kpis["total_5y_revenue"], "kpi_value"),
        ("Gewinn 5J", kpis["total_5y_profit"], "kpi_value"),
        ("OTC-Anteil M12", kpis["otc_share_tablets_m12"], "kpi_value_pct"),
        ("OTC > Rx ab", co_text, "kpi_value_plain"),
    ]
    for c, (label, value, vfmt) in enumerate(kpi_data):
        col = 1 + c * 2
        ws3.merge_range(row, col, row, col + 1, label, fmt["kpi_label"])
        ws3.merge_range(row + 1, col, row + 1, col + 1, value, fmt[vfmt])

    row += 4
    headers = [
        "Monat", "Mon.\nvon Switch", "Rx\nPack.", "Rx\nUmsatz (EUR)",
        "OTC\nPack.", "OTC Preis\n(EUR)", "OTC Retail\nUmsatz", "OTC Herst.\nUmsatz",
        "OTC von\nNeukunden", "OTC von\nRx-Migr.",
        "Awareness", "Saison", "Marketing\nSpend",
        "Gesamt\nUmsatz", "Oper.\nGewinn",
        "Kum.\nUmsatz", "Kum.\nGewinn",
    ]
    for c, h in enumerate(headers):
        ws3.write(row, 1 + c, h, fmt["th_teal"])
    ws3.set_row(row, 35)
    data_start = row + 1

    for i, (_, r) in enumerate(df.iterrows()):
        row += 1
        is_alt = i % 2 == 1
        ws3.write(row, 1, r["month"], fmt["row_alt"] if is_alt else None)
        ws3.write(row, 2, r["months_from_switch"], fmt["row_alt_num"] if is_alt else fmt["number"])
        ws3.write(row, 3, r["rx_packs"], fmt["row_alt_num"] if is_alt else fmt["number"])
        ws3.write(row, 4, r["rx_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        ws3.write(row, 5, r["otc_packs"], fmt["row_alt_num"] if is_alt else fmt["number"])
        ws3.write(row, 6, r["otc_price"], fmt["row_alt_eur"] if is_alt else fmt["eur_detail"])
        ws3.write(row, 7, r["otc_retail_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        ws3.write(row, 8, r["otc_manufacturer_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        ws3.write(row, 9, r["otc_from_new_patients"], fmt["row_alt_num"] if is_alt else fmt["number"])
        ws3.write(row, 10, r["otc_from_rx_migration"], fmt["row_alt_num"] if is_alt else fmt["number"])
        ws3.write(row, 11, r["awareness"], fmt["row_alt_pct"] if is_alt else fmt["pct"])
        ws3.write(row, 12, r["season_factor"], fmt["row_alt_mult"] if is_alt else fmt["mult"])
        ws3.write(row, 13, r["marketing_spend"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        ws3.write(row, 14, r["total_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        ws3.write(row, 15, r["operating_profit"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        ws3.write(row, 16, r["cumulative_total_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        ws3.write(row, 17, r["cumulative_profit"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
    data_end = row

    # Chart 1: Dual Channel (Rx vs OTC packs)
    chart1 = wb.add_chart({"type": "line"})
    chart1.add_series({"name": "Rx Packungen", "categories": ["Forecast", data_start, 1, data_end, 1],
        "values": ["Forecast", data_start, 3, data_end, 3], "line": {"color": "#2563eb", "width": 2.5}})
    chart1.add_series({"name": "OTC Packungen", "categories": ["Forecast", data_start, 1, data_end, 1],
        "values": ["Forecast", data_start, 5, data_end, 5], "line": {"color": "#0d9488", "width": 2.5}})
    chart1.set_title({"name": "Dual-Kanal: Rx vs. OTC Packungen"})
    chart1.set_y_axis({"name": "Packungen / Monat", "num_format": "#,##0"})
    chart1.set_size({"width": 800, "height": 380})
    chart1.set_legend({"position": "bottom"})
    ws3.insert_chart("B" + str(row + 3), chart1)

    # Chart 2: Revenue stacked (Rx + OTC)
    chart2 = wb.add_chart({"type": "column", "subtype": "stacked"})
    chart2.add_series({"name": "Rx Umsatz", "categories": ["Forecast", data_start, 1, data_end, 1],
        "values": ["Forecast", data_start, 4, data_end, 4], "fill": {"color": "#93c5fd"}})
    chart2.add_series({"name": "OTC Umsatz (Hersteller)", "categories": ["Forecast", data_start, 1, data_end, 1],
        "values": ["Forecast", data_start, 8, data_end, 8], "fill": {"color": "#5eead4"}})
    chart2.set_title({"name": "Umsatz nach Kanal (monatlich)"})
    chart2.set_y_axis({"name": "EUR"})
    chart2.set_size({"width": 800, "height": 380})
    chart2.set_legend({"position": "bottom"})
    ws3.insert_chart("J" + str(row + 3), chart2)

    # Chart 3: Awareness curve
    chart3 = wb.add_chart({"type": "area"})
    chart3.add_series({"name": "Awareness", "categories": ["Forecast", data_start, 1, data_end, 1],
        "values": ["Forecast", data_start, 11, data_end, 11], "fill": {"color": "#fef3c7"}, "line": {"color": "#d97706", "width": 2}})
    chart3.set_title({"name": "Consumer Awareness Aufbau"})
    chart3.set_y_axis({"name": "Awareness", "num_format": "0%", "max": 1.0})
    chart3.set_size({"width": 800, "height": 380})
    ws3.insert_chart("B" + str(row + 23), chart3)

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 4: DASHBOARD
    # ═══════════════════════════════════════════════════════════════════
    ws4 = wb.add_worksheet("Dashboard")
    ws4.hide_gridlines(2)
    ws4.set_tab_color("#6366f1")
    ws4.set_column("A:A", 3)
    ws4.set_column("B:B", 38)
    ws4.set_column("C:E", 22)

    row = 1
    ws4.merge_range(row, 1, row, 4, "Executive Summary - Rx-to-OTC Switch Forecast", fmt["title"])

    row += 2
    ws4.merge_range(row, 1, row, 4, "Szenario-Vergleich", fmt["section"])

    scenario_results = {}
    for sn, sp in [
        ("Konservativ", {"otc_peak_packs_per_month": 180_000, "otc_price_per_pack": 6.99,
                         "otc_ramp_months": 24, "marketing_monthly_eur": 200_000,
                         "rx_decline_rate": 0.10, "new_patient_share": 0.60}),
        ("Base Case", {}),
        ("Optimistisch", {"otc_peak_packs_per_month": 400_000, "otc_price_per_pack": 8.99,
                          "otc_ramp_months": 12, "marketing_monthly_eur": 500_000,
                          "rx_decline_rate": 0.20, "new_patient_share": 0.75}),
    ]:
        p = RxOtcParams(**sp)
        scenario_results[sn] = calculate_kpis_rx_otc(forecast_rx_otc(p))

    row += 1
    ws4.write(row, 1, "")
    ws4.write(row, 2, "Konservativ", fmt["th"])
    ws4.write(row, 3, "Base Case", fmt["th_teal"])
    ws4.write(row, 4, "Optimistisch", fmt["th"])

    for label, key, vtype in [
        ("OTC Umsatz Jahr 1", "year1_otc_revenue", "eur"),
        ("Gesamtumsatz 5J (Rx+OTC)", "total_5y_revenue", "eur"),
        ("OTC Umsatz 5J", "total_5y_otc_revenue", "eur"),
        ("Rx Umsatz 5J", "total_5y_rx_revenue", "eur"),
        ("Gewinn 5J", "total_5y_profit", "eur"),
        ("OTC-Anteil Monat 12 (Tabletten)", "otc_share_tablets_m12", "pct"),
        ("OTC-Anteil Monat 24 (Tabletten)", "otc_share_tablets_m24", "pct"),
        ("Rx-Rueckgang gesamt", "rx_decline_total", "pct"),
        ("Peak Awareness", "peak_awareness", "pct"),
        ("Marketing-Invest 5J", "total_5y_marketing", "eur"),
        ("Kannibalisierung (Antazida etc.)", "total_adjacent_lost", "eur"),
    ]:
        row += 1
        ws4.write(row, 1, label, fmt["label_bold"])
        for c, sn in enumerate(["Konservativ", "Base Case", "Optimistisch"]):
            val = scenario_results[sn].get(key, 0) or 0
            if vtype == "eur":
                ws4.write(row, 2 + c, val, fmt["eur"])
            elif vtype == "pct":
                ws4.write(row, 2 + c, val, fmt["pct"])

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 5: METHODIK
    # ═══════════════════════════════════════════════════════════════════
    ws5 = wb.add_worksheet("Methodik")
    ws5.hide_gridlines(2)
    ws5.set_tab_color("#6366f1")
    ws5.set_column("A:A", 3)
    ws5.set_column("B:B", 34)
    ws5.set_column("C:C", 24)
    ws5.set_column("D:D", 52)
    ws5.set_column("E:E", 42)

    row = 1
    ws5.merge_range(row, 1, row, 4, "Methodik & Transparenz-Matrix", fmt["title"])
    row += 1
    ws5.merge_range(row, 1, row, 4, "Jeder Datenpunkt ist klassifiziert: FAKT / ANNAHME / MODELL / LUECKE",
        wb.add_format({"font_size": 10, "font_color": "#6b7280", "italic": True}))

    row += 2
    ws5.write(row, 1, "Legende:", fmt["label_bold"])
    ws5.write(row, 2, "FAKT", fmt["fact"])
    ws5.write(row, 3, "Aus oeffentlichen Quellen ableitbar", fmt["label"])
    row += 1
    ws5.write(row, 2, "ANNAHME", fmt["assumption"])
    ws5.write(row, 3, "Vom User steuerbar (gelbe Zellen)", fmt["label"])
    row += 1
    ws5.write(row, 2, "MODELL", fmt["model"])
    ws5.write(row, 3, "Mathematisches Modell", fmt["label"])
    row += 1
    ws5.write(row, 2, "LUECKE", fmt["gap"])
    ws5.write(row, 3, "Fuer Produktiveinsatz benoetigt", fmt["label"])

    row += 2
    ws5.merge_range(row, 1, row, 4, "FAKTEN", fmt["section"])
    row += 1
    for c, h in enumerate(["Datenpunkt", "Wert", "Quelle", "Kategorie"]):
        ws5.write(row, 1 + c, h, fmt["th"])
    facts = [
        ("OTC-Markt DE 2024", "10.15 Mrd. EUR", "IQVIA / Pharma Deutschland", "FAKT"),
        ("PPI OTC-Switch Datum", "Juli 2009 (Pantoprazol), Aug 2009 (Omeprazol)", "BfArM Sachverstaendigenausschuss", "FAKT"),
        ("OTC-PPI Jahr 1 Volumen", "3.3 Mio. Packungen, 29 Mio. EUR", "IQVIA / Apotheke Adhoc", "FAKT"),
        ("OTC-PPI Anteil am PPI-Gesamtmarkt", "~3% (Partial Switch)", "IQVIA / PMC-Studie", "FAKT"),
        ("Antazida-Kannibalisierung", "-11 Mio. EUR im ersten Jahr", "IQVIA / Apotheke Adhoc", "FAKT"),
        ("H2-Antagonisten-Kollaps", "-46% im ersten Jahr nach PPI OTC", "IQVIA / Apotheke Adhoc", "FAKT"),
        ("OTC-PPI Packungsgroesse", "Max. 14 Tabletten, 20mg", "AMPreisV / BfArM", "FAKT"),
        ("Apothekenmarge OTC", "Freie Preisgestaltung (nicht reguliert)", "AMPreisV", "FAKT"),
        ("BfArM Switch-Prozess", "2x jaehrlich, Sachverstaendigenausschuss", "BfArM", "FAKT"),
    ]
    for dp, val, source, cat in facts:
        row += 1
        ws5.write(row, 1, dp, fmt["label_bold"])
        ws5.write(row, 2, val, fmt["fact"])
        ws5.write(row, 3, source, fmt["label"])
        ws5.write(row, 4, cat, fmt["fact"])

    row += 2
    ws5.merge_range(row, 1, row, 4, "ANNAHMEN", fmt["section"])
    row += 1
    for c, h in enumerate(["Parameter", "Default", "Begruendung", "Kategorie"]):
        ws5.write(row, 1 + c, h, fmt["th"])
    assumptions = [
        ("OTC Peak Volumen", "280K Pack./Mon.", "Hochrechnung aus hist. PPI-Daten (3.3M/Jahr = 275K/Mon.)", "ANNAHME"),
        ("Rx-Abwanderung", "15%", "Partial Switch: OTC nur 14 St./20mg, Rx fuer chronisch", "ANNAHME"),
        ("Neue Patienten", "70%", "Referenz: 11.9% Marktexpansion Sodbrennen nach PPI OTC", "ANNAHME"),
        ("OTC-Preis", "EUR 7.99", "Marktdurchschnitt Omeprazol-Generika OTC (14 St.)", "ANNAHME"),
        ("Apothekenmarge", "40%", "Branchenueblich OTC, Spanne 35-60%", "ANNAHME"),
        ("Marketing-Budget", "EUR 300K/Mon.", "Firmenspezifisch, Benchmark: Orlistat EUR 12.5M/Jahr", "ANNAHME"),
        ("Peak Awareness", "65%", "Gutes OTC-Marketing; Benchmark: Voltaren ~85%, Orlistat ~45%", "ANNAHME"),
        ("Saisonalitaet", "PPI: Herbst/Winter-Peak", "GI-Beschwerden haeufiger in kalter Jahreszeit", "ANNAHME"),
    ]
    for param, val, reason, cat in assumptions:
        row += 1
        ws5.write(row, 1, param, fmt["label_bold"])
        ws5.write(row, 2, val, fmt["assumption"])
        ws5.write(row, 3, reason, fmt["label"])
        ws5.write(row, 4, cat, fmt["assumption"])

    row += 2
    ws5.merge_range(row, 1, row, 4, "MODELLE", fmt["section"])
    row += 1
    for c, h in enumerate(["Modell", "Formel / Logik", "Referenz", "Kategorie"]):
        ws5.write(row, 1 + c, h, fmt["th"])
    models = [
        ("OTC-Ramp (S-Kurve)", "Logistic: Peak / (1 + e^(-s*(t-midpoint)))", "Standard Launch-Curve; PPI stagnierte nach 3 Monaten", "MODELL"),
        ("Rx-Decline", "Exponential Decay: Floor + (Init-Floor) * e^(-k*t)", "Partial Switch: 97% blieb auf Rx bei PPI", "MODELL"),
        ("Awareness-Kurve", "Logistic S-Curve: 0 -> Peak ueber Ramp-Monate", "Marketing-Funnel: Awareness -> Trial -> Repeat", "MODELL"),
        ("Preiselastizitaet", "Vol(t) = Base * (1 + PriceChange * Elasticity)", "OTC: Patient zahlt 100%, preissensitiv", "MODELL"),
        ("Saisonalitaet", "Monatsfaktoren (0.85-1.15) auf OTC-Nachfrage", "PPI: GI-Beschwerden saisonal", "MODELL"),
        ("Kannibalisierung", "Adjacent_Rev * Cannibalization_Rate, linear ramp", "Antazida -11M EUR, H2-Blocker -46% nach PPI Switch", "MODELL"),
    ]
    for name, formula, ref, cat in models:
        row += 1
        ws5.set_row(row, 35)
        ws5.write(row, 1, name, fmt["label_bold"])
        ws5.write(row, 2, formula, fmt["model"])
        ws5.write(row, 3, ref, fmt["label"])
        ws5.write(row, 4, cat, fmt["model"])

    row += 2
    ws5.merge_range(row, 1, row, 4, "LUECKEN", fmt["section"])
    row += 1
    for c, h in enumerate(["Gap", "Beschreibung", "Datenquelle", "Kategorie"]):
        ws5.write(row, 1 + c, h, fmt["th"])
    gaps = [
        ("Echte IQVIA OTC-Daten", "Sell-out Packungen pro Monat pro PPI-Produkt", "IQVIA OTC Insight", "LUECKE"),
        ("Apotheken-Panel", "Tatsaechliche Empfehlungsrate durch Apotheker", "IQVIA / GfK Panel", "LUECKE"),
        ("Consumer-Tracking", "Awareness, Trial, Repeat auf Produktebene", "GfK / Nielsen Consumer Panel", "LUECKE"),
        ("Marketing-ROI", "Tatsaechliche Wirkung von TV/Digital/POS", "Mediaagenturen / Tracking", "LUECKE"),
        ("Preiselastizitaet (gemessen)", "Tatsaechliche Preissensitivitaet pro Kategorie", "Conjoint-Analyse / Scanner-Daten", "LUECKE"),
        ("Versandhandel-Shift", "OTC-Anteil Online vs. stationaer ueber Zeit", "SEMPORA / Apothekenverband", "LUECKE"),
    ]
    for gap_name, desc, source, cat in gaps:
        row += 1
        ws5.write(row, 1, gap_name, fmt["label_bold"])
        ws5.write(row, 2, desc, fmt["gap"])
        ws5.write(row, 3, source, fmt["label"])
        ws5.write(row, 4, cat, fmt["gap"])

    row += 3
    ws5.merge_range(row, 1, row, 4,
        "Hinweis: Alle Daten sind synthetisch. Angelehnt an Omeprazol/Pantoprazol OTC-Switch 2009, "
        "IQVIA OTC-Markt 2024, Pharma Deutschland. Kein Bezug zu vertraulichen Daten.",
        fmt["hint"])

    wb.close()
    print(f"Excel model saved to: {OUTPUT_PATH}")
    return OUTPUT_PATH


if __name__ == "__main__":
    build_model()
