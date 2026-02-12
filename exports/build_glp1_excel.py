"""
Build a professional Excel forecast model for the GLP-1 Brand Competition scenario.

Sheets:
1. INPUTS          – All user-configurable parameters (yellow cells)
2. Marktdaten      – GLP-1 market context, products, differentiation
3. Lilly Forecast  – Mounjaro perspective (monthly forecast)
4. Novo Forecast   – Ozempic/Wegovy perspective (monthly forecast)
5. Dashboard       – Executive summary, scenario comparison
6. Methodik        – Transparency matrix: Facts vs. Assumptions vs. Models
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import xlsxwriter
import numpy as np
import pandas as pd
from models.brand_competition_engine import (
    MarketParams, BrandParams, CompetitorParams,
    forecast_brand, calculate_kpis_brand,
)

OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "GLP1_Brand_Competition_Forecast.xlsx")


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
    fmt["mult"] = wb.add_format({"num_format": "0.000"})
    fmt["th"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "white", "bg_color": "#1e3a5f", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"})
    fmt["th_orange"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "white", "bg_color": "#9a3412", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"})
    fmt["th_blue"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "white", "bg_color": "#1e40af", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"})
    fmt["th_green"] = wb.add_format({"bold": True, "font_size": 10, "font_color": "white", "bg_color": "#166534", "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter"})
    fmt["kpi_label"] = wb.add_format({"bold": True, "font_size": 11, "font_color": "#1e3a5f", "bg_color": "#f0f4ff", "border": 1})
    fmt["kpi_value"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1e3a5f", "bg_color": "#f0f4ff", "border": 1, "num_format": "€#,##0", "align": "center"})
    fmt["kpi_value_pct"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1e3a5f", "bg_color": "#f0f4ff", "border": 1, "num_format": "0.0%", "align": "center"})
    fmt["kpi_value_orange"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#9a3412", "bg_color": "#fef3c7", "border": 1, "num_format": "€#,##0", "align": "center"})
    fmt["kpi_value_orange_pct"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#9a3412", "bg_color": "#fef3c7", "border": 1, "num_format": "0.0%", "align": "center"})
    fmt["kpi_value_blue"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1e40af", "bg_color": "#dbeafe", "border": 1, "num_format": "€#,##0", "align": "center"})
    fmt["kpi_value_blue_pct"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1e40af", "bg_color": "#dbeafe", "border": 1, "num_format": "0.0%", "align": "center"})
    fmt["kpi_value_green"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#166534", "bg_color": "#f0fdf4", "border": 1, "num_format": "€#,##0", "align": "center"})
    fmt["kpi_value_plain"] = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1e3a5f", "bg_color": "#f0f4ff", "border": 1, "align": "center"})
    fmt["row_alt"] = wb.add_format({"bg_color": "#f8fafc"})
    fmt["row_alt_eur"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "€#,##0"})
    fmt["row_alt_pct"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "0.0%"})
    fmt["row_alt_num"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "#,##0"})
    fmt["row_alt_mult"] = wb.add_format({"bg_color": "#f8fafc", "num_format": "0.000"})
    # Methodik formats
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
    ws.merge_range(row, 1, row, 3, "GLP-1 Brand Competition - Forecast Modell", fmt["title"])
    row += 1
    ws.merge_range(row, 1, row, 3, "Mounjaro (Tirzepatid, Lilly) vs. Ozempic/Wegovy (Semaglutid, Novo Nordisk)", wb.add_format({"font_size": 10, "font_color": "#6b7280", "italic": True}))

    # 1. Marktdefinition
    row += 2
    ws.merge_range(row, 1, row, 3, "1. MARKTDEFINITION", fmt["section"])
    for label, value, fk, hint in [
        ("Markt", "GLP-1-Rezeptor-Agonisten (GLP-1 RA)", "input_text", "ATC-Klasse / Therapiegebiet"),
        ("Geographie", "Deutschland", "input_text", "Zielmarkt"),
        ("T2D-Patienten in DE", 6_330_000, "input", "GKV-Versicherte mit Diabetes mellitus Typ 2"),
        ("GLP-1-Penetration T2D (2024)", 0.139, "input_pct", "13.9% der T2D-Patienten auf GLP-1 (BARMER 2025)"),
        ("GLP-1-Penetration Ziel", 0.22, "input_pct", "Erwartete Penetration durch Leitlinien-Shift"),
        ("Potenzielle Adipositas-Patienten", 8_000_000, "input", "BMI >= 30 in Deutschland"),
        ("GLP-1 TRx/Monat (gesamt)", 950_000, "input", "Monatliche GLP-1-Verordnungen (alle Produkte)"),
        ("Marktwachstum p.a.", 0.25, "input_pct", "GLP-1-Segment waechst ~25% p.a. (IQVIA 2024)"),
        ("Forecast-Horizont (Jahre)", 5, "input", "Prognosezeitraum"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    # 2. Mounjaro (Lilly)
    row += 2
    ws.merge_range(row, 1, row, 3, "2. MOUNJARO (Tirzepatid) - ELI LILLY", fmt["section"])
    for label, value, fk, hint in [
        ("Wirkstoff / Mechanismus", "Tirzepatid (Dual GIP/GLP-1 RA)", "input_text", "Erster dualer Agonist"),
        ("Aktueller Marktanteil (TRx)", 0.08, "input_pct", "~8% und steigend (Marktstart DE Nov 2023)"),
        ("Preis/Monat (EUR)", 350.00, "input_eur", "~EUR 206-490/Monat, Durchschn. ~EUR 350"),
        ("Peak Share T2D", 0.25, "input_pct", "Angestrebter Spitzenanteil im T2D-Segment"),
        ("Peak Share Adipositas", 0.35, "input_pct", "Angestrebter Spitzenanteil im Adipositas-Segment"),
        ("Monate bis Peak", 36, "input", "Monate bis Ziel-Marktanteil"),
        ("Wirksamkeit HbA1c-Reduktion", -2.4, "input_eur", "%-Punkte (SURPASS-Studien)"),
        ("Wirksamkeit Gewichtsverlust", -0.225, "input_pct", "-22.5% Koerpergewicht (SURMOUNT-1, 72 Wochen)"),
        ("Indikation T2D", "JA", "input_bool", "EU-Zulassung seit Sep 2022"),
        ("Indikation Adipositas", "JA", "input_bool", "EU-Zulassung seit Dez 2023"),
        ("Indikation CV-Risiko", "NEIN", "input_bool", "Noch nicht zugelassen (Studien laufen)"),
        ("Indikation MASH/NASH", "NEIN", "input_bool", "Studien laufen (Lilly fuehrend)"),
        ("AMNOG-Status", "Geheimer Erstattungsbetrag (Aug 2025)", "input_text", "Erster vertraulicher EB in DE"),
        ("Lieferengpass", "NEIN", "input_bool", "Gute Versorgungslage"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    # 3. Ozempic/Wegovy (Novo)
    row += 2
    ws.merge_range(row, 1, row, 3, "3. OZEMPIC / WEGOVY (Semaglutid) - NOVO NORDISK", fmt["section"])
    for label, value, fk, hint in [
        ("Wirkstoff / Mechanismus", "Semaglutid (GLP-1 RA)", "input_text", "Mono-Agonist, breiteste Evidenzbasis"),
        ("Aktueller Marktanteil (TRx)", 0.36, "input_pct", "Ozempic + Wegovy kombiniert ~36%"),
        ("Preis/Monat Ozempic (EUR)", 300.00, "input_eur", "~EUR 300/Monat (1mg Erhaltungsdosis)"),
        ("Peak Share T2D", 0.30, "input_pct", "Verteidigbare Position im T2D-Segment"),
        ("Indikation T2D (Ozempic)", "JA", "input_bool", "Marktfuehrer T2D seit 2020"),
        ("Indikation Adipositas (Wegovy)", "JA", "input_bool", "Adipositas-Zulassung, ABER:"),
        ("Wegovy GKV-Erstattung", "NEIN", "input_bool", "G-BA Juni 2024: Lifestyle-Arzneimittel, kein GKV"),
        ("Lieferengpass Ozempic", "JA", "input_bool", "Seit 2022, Normalisierung erwartet 2025/2026"),
        ("Engpass endet (Monat)", 12, "input", "Monate bis volle Versorgung"),
        ("Wirksamkeit HbA1c-Reduktion", -1.8, "input_eur", "%-Punkte (SUSTAIN-Studien)"),
        ("Wirksamkeit Gewichtsverlust", -0.15, "input_pct", "-15% Koerpergewicht (STEP-Studien, Wegovy)"),
        ("Evidenz-Score", "9.0 / 10", "input_text", "Breiteste Evidenzbasis aller GLP-1 RAs"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    # 4. Indikationserweiterung
    row += 2
    ws.merge_range(row, 1, row, 3, "4. INDIKATIONSERWEITERUNGEN & MARKTEXPANSION", fmt["section"])
    for label, value, fk, hint in [
        ("Adipositas GKV-Erstattung?", "NEIN", "input_bool", "Gesetzesaenderung fuer Erstattung noetig"),
        ("Falls ja: GKV-Start Jahr", 2028, "input", "Fruehestens 2027 (optimistisch)"),
        ("GKV-Penetration Adipositas", 0.05, "input_pct", "5% der 8M Adipositas-Patienten = 400K"),
        ("CV-Indikation (Mounjaro)", "NEIN", "input_bool", "Kardiovaskulaere Risikoreduktion"),
        ("MASH/NASH-Indikation (Mounjaro)", "NEIN", "input_bool", "Lebererkrankung, Lilly fuehrend"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    # 5. Kosten
    row += 2
    ws.merge_range(row, 1, row, 3, "5. KOSTEN & PROFITABILITAET", fmt["section"])
    for label, value, fk, hint in [
        ("COGS (% vom Umsatz)", 0.20, "input_pct", "Biologika/Peptide: typisch 15-25%"),
        ("SG&A / Monat (EUR)", 800_000, "input", "Vertrieb, KAM, Marketing"),
        ("Medical Affairs / Monat (EUR)", 200_000, "input", "MSL, Studien, Kongresse"),
        ("Preistrend p.a.", -0.03, "input_pct", "Jaehrliche Preiserosion (-3%)"),
    ]:
        row += 1
        write_input_row(ws, row, label, value, fk, hint)

    # 6. Szenarien
    row += 2
    ws.merge_range(row, 1, row, 4, "6. SZENARIO-DEFINITIONEN", fmt["section"])
    row += 1
    for c, h in enumerate(["Parameter", "Base Case", "Bull (Lilly)", "Bull (Novo)"]):
        ws.write(row, 1 + c, h, fmt["th"])
    for param, base, bull_l, bull_n in [
        ("Marktwachstum p.a.", 0.25, 0.35, 0.30),
        ("Mounjaro Peak Share T2D", 0.25, 0.35, 0.15),
        ("Mounjaro Peak Share Adipositas", 0.35, 0.45, 0.20),
        ("Mounjaro Monate bis Peak", 36, 24, 48),
        ("Novo Peak Share T2D", 0.30, 0.25, 0.38),
        ("Adipositas GKV-Erstattung", "Nein", "Ja (2027)", "Nein"),
        ("GLP-1-Penetration T2D Ziel", 0.22, 0.28, 0.20),
    ]:
        row += 1
        ws.write(row, 1, param, fmt["label_bold"])
        for c, val in enumerate([base, bull_l, bull_n]):
            if isinstance(val, float):
                ws.write(row, 2 + c, val, fmt["input_pct"])
            elif isinstance(val, int):
                ws.write(row, 2 + c, val, fmt["input"])
            else:
                ws.write(row, 2 + c, val, fmt["input_text"])

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 2: MARKTDATEN
    # ═══════════════════════════════════════════════════════════════════
    ws2 = wb.add_worksheet("Marktdaten")
    ws2.hide_gridlines(2)
    ws2.set_tab_color("#1e3a5f")
    ws2.set_column("A:A", 3)
    ws2.set_column("B:B", 32)
    ws2.set_column("C:H", 18)

    row = 1
    ws2.merge_range(row, 1, row, 7, "GLP-1 Wettbewerbslandschaft - Deutschland (Mid-2025)", fmt["title"])
    row += 2
    for c, h in enumerate(["Produkt", "Hersteller", "Marktanteil\n(TRx)", "Preis/Mon.\n(EUR)", "Indikation", "GKV", "Lieferung"]):
        ws2.write(row, 1 + c, h, fmt["th"])
    ws2.set_row(row, 30)

    products = [
        ("Ozempic (Semaglutid s.c.)", "Novo Nordisk", 0.32, 300.00, "T2D", "Ja", "Engpass"),
        ("Trulicity (Dulaglutid)", "Eli Lilly", 0.38, 178.00, "T2D", "Ja", "OK"),
        ("Mounjaro (Tirzepatid)", "Eli Lilly", 0.08, 350.00, "T2D + Adipositas", "Ja (T2D)", "OK"),
        ("Rybelsus (Semaglutid oral)", "Novo Nordisk", 0.07, 280.00, "T2D", "Ja", "OK"),
        ("Victoza (Liraglutid)", "Novo Nordisk", 0.06, 195.00, "T2D", "Ja", "Engpass"),
        ("Wegovy (Semaglutid 2.4mg)", "Novo Nordisk", 0.04, 277.00, "Adipositas", "NEIN", "Engpass"),
        ("Sonstige (Exenatid etc.)", "Diverse", 0.05, 150.00, "T2D", "Ja", "OK"),
    ]
    for pn, co, sh, pr, ind, gkv, supply in products:
        row += 1
        ws2.write(row, 1, pn, fmt["label_bold"])
        ws2.write(row, 2, co, fmt["label"])
        ws2.write(row, 3, sh, fmt["pct"])
        ws2.write(row, 4, pr, fmt["eur_detail"])
        ws2.write(row, 5, ind, fmt["label"])
        ws2.write(row, 6, gkv, fmt["label"])
        ws2.write(row, 7, supply, fmt["label"])

    # Differentiation
    row += 3
    ws2.merge_range(row, 1, row, 7, "Wettbewerbsdifferenzierung: Mounjaro vs. Ozempic", fmt["section"])
    row += 1
    for c, h in enumerate(["Dimension", "Mounjaro (Tirzepatid)", "Ozempic (Semaglutid)", "Vorteil"]):
        ws2.write(row, 1 + c, h, fmt["th"])
    diffs = [
        ("Mechanismus", "Dual GIP/GLP-1 RA", "Mono GLP-1 RA", "Mounjaro"),
        ("HbA1c-Reduktion", "-2.4 %-Punkte", "-1.8 %-Punkte", "Mounjaro"),
        ("Gewichtsverlust", "-22.5%", "-15.0%", "Mounjaro"),
        ("Evidenzbasis", "Wachsend (SURPASS/SURMOUNT)", "Am breitesten (SUSTAIN/STEP/SELECT)", "Ozempic"),
        ("GKV Adipositas", "Ja (via Mounjaro)", "Nein (Wegovy ausgeschlossen)", "Mounjaro"),
        ("Versorgungslage", "Gut", "Lieferengpass seit 2022", "Mounjaro"),
        ("Marktpraesenz", "Seit Nov 2023 (kurz)", "Seit Feb 2020 (etabliert)", "Ozempic"),
        ("AMNOG-Preis", "Vertraulich (ab Aug 2025)", "Vereinbart (oeffentlich)", "Neutral"),
    ]
    for dim, mj, oz, adv in diffs:
        row += 1
        ws2.write(row, 1, dim, fmt["label_bold"])
        ws2.write(row, 2, mj, fmt["label"])
        ws2.write(row, 3, oz, fmt["label"])
        color = fmt["fact"] if adv == "Mounjaro" else (fmt["model"] if adv == "Ozempic" else fmt["assumption"])
        ws2.write(row, 4, adv, color)

    # Growth drivers
    row += 3
    ws2.merge_range(row, 1, row, 5, "Marktwachstumstreiber", fmt["section"])
    row += 1
    for c, h in enumerate(["Treiber", "Auswirkung", "Zeithorizont", "Sicherheit"]):
        ws2.write(row, 1 + c, h, fmt["th"])
    drivers = [
        ("Leitlinien-Shift T2D", "GLP-1 frueher in Therapie: 14% -> 25-30%", "2025-2030", "Hoch"),
        ("Adipositas GKV-Erstattung", "8M potenzielle Patienten", "Fruehestens 2027", "Mittel"),
        ("CV-Risikoreduktion (SELECT)", "-20% CV-Events -> breitere Zielgruppe", "2025-2026", "Hoch"),
        ("MASH/NASH-Indikation", "Neue Patientenpopulation", "2026-2028", "Mittel"),
        ("Orales Semaglutid 25/50mg", "Convenience-Shift weg von Injektion", "2025-2026", "Hoch"),
        ("Behebung Lieferengpaesse", "Aufholen unterdrueckter Nachfrage", "2025", "Hoch"),
    ]
    for driver, impact, timeline, certainty in drivers:
        row += 1
        ws2.write(row, 1, driver, fmt["label_bold"])
        ws2.write(row, 2, impact, fmt["label"])
        ws2.write(row, 3, timeline, fmt["label"])
        cert_fmt = fmt["fact"] if certainty == "Hoch" else fmt["assumption"]
        ws2.write(row, 4, certainty, cert_fmt)

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 3 & 4: FORECAST (Lilly + Novo)
    # ═══════════════════════════════════════════════════════════════════
    def write_forecast_sheet(ws, brand, competitor, market, sheet_name, color_key, th_fmt_key):
        ws.hide_gridlines(2)
        ws.set_column("A:A", 3)
        ws.set_column("B:B", 12)
        ws.set_column("C:Q", 15)

        df = forecast_brand(brand, competitor, market, forecast_months=60)
        kpis = calculate_kpis_brand(df)

        row = 1
        ws.merge_range(row, 1, row, 12, f"{brand.company} Perspektive - {brand.name} (Base Case)", fmt["title"])

        # KPI row
        row += 2
        ot = kpis.get("overtake_month")
        ot_text = f"Monat {ot}" if ot else "–"
        kpi_data = [
            ("Umsatz Jahr 1", kpis["year1_revenue"], f"kpi_value_{color_key}"),
            ("Umsatz 5J", kpis["total_5y_revenue"], f"kpi_value_{color_key}"),
            ("Peak Marktanteil", kpis["peak_share"], f"kpi_value_{color_key}_pct"),
            ("Market Leader ab", ot_text, "kpi_value_plain"),
            ("Gewinn 5J", kpis["total_5y_profit"], "kpi_value_green"),
        ]
        for c, (label, value, vfmt) in enumerate(kpi_data):
            col = 1 + c * 2
            ws.merge_range(row, col, row, col + 1, label, fmt["kpi_label"])
            ws.merge_range(row + 1, col, row + 1, col + 1, value, fmt[vfmt])

        row += 3
        ws.merge_range(row, 1, row, 6, f"Markt: {kpis['market_size_start']:,.0f} -> {kpis['market_size_end']:,.0f} TRx/Mon. (x{kpis['market_growth_total']:.1f})", fmt["subsection"])

        # Data headers
        row += 2
        headers = [
            "Monat", "Mon.\nvon Start", "Gesamtmarkt\nTRx", "Mein\nAnteil",
            "Meine\nTRx", "Ind.-\nMultipl.", "Supply\nGap", "Preis/Mon.\n(EUR)",
            "Mein\nUmsatz (EUR)", "COGS\n(EUR)", "Oper. Gewinn\n(EUR)",
            "Competitor\nAnteil", "Competitor\nTRx", "Competitor\nUmsatz (EUR)",
            "Kum. Umsatz\n(EUR)", "Kum. Gewinn\n(EUR)",
        ]
        for c, h in enumerate(headers):
            ws.write(row, 1 + c, h, fmt[th_fmt_key])
        ws.set_row(row, 35)
        data_start = row + 1

        for i, (_, r) in enumerate(df.iterrows()):
            row += 1
            is_alt = i % 2 == 1
            ws.write(row, 1, r["month"], fmt["row_alt"] if is_alt else None)
            ws.write(row, 2, r["months_from_start"], fmt["row_alt_num"] if is_alt else fmt["number"])
            ws.write(row, 3, r["total_market_trx"], fmt["row_alt_num"] if is_alt else fmt["number"])
            ws.write(row, 4, r["my_share"], fmt["row_alt_pct"] if is_alt else fmt["pct"])
            ws.write(row, 5, r["my_actual_trx"], fmt["row_alt_num"] if is_alt else fmt["number"])
            ws.write(row, 6, r["indication_multiplier"], fmt["row_alt_mult"] if is_alt else fmt["mult"])
            ws.write(row, 7, r["supply_gap_trx"], fmt["row_alt_num"] if is_alt else fmt["number"])
            ws.write(row, 8, r["my_price"], fmt["row_alt_eur"] if is_alt else fmt["eur_detail"])
            ws.write(row, 9, r["my_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
            ws.write(row, 10, r["my_cogs"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
            ws.write(row, 11, r["my_operating_profit"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
            ws.write(row, 12, r["competitor_share"], fmt["row_alt_pct"] if is_alt else fmt["pct"])
            ws.write(row, 13, r["competitor_trx"], fmt["row_alt_num"] if is_alt else fmt["number"])
            ws.write(row, 14, r["competitor_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
            ws.write(row, 15, r["cumulative_revenue"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
            ws.write(row, 16, r["cumulative_profit"], fmt["row_alt_eur"] if is_alt else fmt["eur"])
        data_end = row

        # Chart 1: Market Share Race
        chart1 = wb.add_chart({"type": "line"})
        chart1.add_series({
            "name": brand.name.split("(")[0].strip(),
            "categories": [sheet_name, data_start, 1, data_end, 1],
            "values": [sheet_name, data_start, 4, data_end, 4],
            "line": {"color": "#f59e0b" if color_key == "orange" else "#2563eb", "width": 2.5},
        })
        chart1.add_series({
            "name": competitor.name.split("(")[0].strip(),
            "categories": [sheet_name, data_start, 1, data_end, 1],
            "values": [sheet_name, data_start, 12, data_end, 12],
            "line": {"color": "#2563eb" if color_key == "orange" else "#f59e0b", "width": 2.5, "dash_type": "dash"},
        })
        chart1.set_title({"name": "Marktanteils-Wettlauf"})
        chart1.set_y_axis({"name": "Marktanteil", "num_format": "0%"})
        chart1.set_size({"width": 800, "height": 380})
        chart1.set_legend({"position": "bottom"})
        ws.insert_chart("B" + str(row + 3), chart1)

        # Chart 2: Revenue Build-up
        chart2 = wb.add_chart({"type": "column"})
        chart2.add_series({
            "name": "Monatl. Umsatz",
            "categories": [sheet_name, data_start, 1, data_end, 1],
            "values": [sheet_name, data_start, 9, data_end, 9],
            "fill": {"color": "#f59e0b" if color_key == "orange" else "#2563eb"},
            "border": {"color": "#92400e" if color_key == "orange" else "#1e40af"},
        })
        chart2_line = wb.add_chart({"type": "line"})
        chart2_line.add_series({
            "name": "Kum. Umsatz",
            "categories": [sheet_name, data_start, 1, data_end, 1],
            "values": [sheet_name, data_start, 15, data_end, 15],
            "line": {"color": "#166534", "width": 2.5},
            "y2_axis": True,
        })
        chart2.combine(chart2_line)
        chart2.set_title({"name": "Umsatz-Aufbau"})
        chart2.set_y_axis({"name": "Monatl. Umsatz (EUR)"})
        chart2.set_y2_axis({"name": "Kumuliert (EUR)"})
        chart2.set_size({"width": 800, "height": 380})
        chart2.set_legend({"position": "bottom"})
        ws.insert_chart("J" + str(row + 3), chart2)

        # Chart 3: Market Expansion
        chart3 = wb.add_chart({"type": "area"})
        chart3.add_series({
            "name": "Gesamtmarkt GLP-1",
            "categories": [sheet_name, data_start, 1, data_end, 1],
            "values": [sheet_name, data_start, 3, data_end, 3],
            "fill": {"color": "#c7d2fe"}, "line": {"color": "#6366f1", "width": 2},
        })
        chart3.set_title({"name": "GLP-1 Marktexpansion (TRx / Monat)"})
        chart3.set_y_axis({"name": "TRx", "num_format": "#,##0"})
        chart3.set_size({"width": 800, "height": 380})
        ws.insert_chart("B" + str(row + 23), chart3)

        return df, kpis

    # Lilly sheet
    ws3 = wb.add_worksheet("Lilly Forecast")
    ws3.set_tab_color("#f59e0b")
    brand_lilly = BrandParams(
        name="Mounjaro (Tirzepatid)", company="Eli Lilly",
        current_share_trx=0.08, price_per_month_eur=350.00,
        target_peak_share_t2d=0.25, target_peak_share_obesity=0.35,
        months_to_peak=36, has_t2d=True, has_obesity=True,
        efficacy_score=9.0, evidence_score=7.5,
    )
    comp_novo = CompetitorParams(
        name="Ozempic / Wegovy", company="Novo Nordisk",
        current_share_trx=0.36, price_per_month_eur=300.00,
        target_peak_share_t2d=0.30,
        supply_constrained=True, supply_normalization_month=12,
    )
    mkt = MarketParams()
    df_lilly, kpis_lilly = write_forecast_sheet(ws3, brand_lilly, comp_novo, mkt, "Lilly Forecast", "orange", "th_orange")

    # Novo sheet
    ws4 = wb.add_worksheet("Novo Forecast")
    ws4.set_tab_color("#2563eb")
    brand_novo = BrandParams(
        name="Ozempic / Wegovy (Semaglutid)", company="Novo Nordisk",
        current_share_trx=0.36, price_per_month_eur=300.00,
        target_peak_share_t2d=0.30, target_peak_share_obesity=0.30,
        months_to_peak=24, has_t2d=True, has_obesity=True,
        supply_constrained=True, supply_normalization_month=12,
        supply_capacity_monthly_trx=300_000,
        efficacy_score=7.5, evidence_score=9.0,
    )
    comp_lilly = CompetitorParams(
        name="Mounjaro (Tirzepatid)", company="Eli Lilly",
        current_share_trx=0.08, price_per_month_eur=350.00,
        target_peak_share_t2d=0.25,
        supply_constrained=False,
    )
    df_novo, kpis_novo = write_forecast_sheet(ws4, brand_novo, comp_lilly, mkt, "Novo Forecast", "blue", "th_blue")

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 5: DASHBOARD
    # ═══════════════════════════════════════════════════════════════════
    ws5 = wb.add_worksheet("Dashboard")
    ws5.hide_gridlines(2)
    ws5.set_tab_color("#6366f1")
    ws5.set_column("A:A", 3)
    ws5.set_column("B:B", 38)
    ws5.set_column("C:E", 22)

    row = 1
    ws5.merge_range(row, 1, row, 4, "Executive Summary - GLP-1 Brand Competition 2026-2030", fmt["title"])

    # Head-to-head comparison
    row += 2
    ws5.merge_range(row, 1, row, 3, "Head-to-Head Vergleich (Base Case)", fmt["section"])
    row += 1
    ws5.write(row, 1, "", fmt["th"])
    ws5.write(row, 2, "Mounjaro (Lilly)", fmt["th_orange"])
    ws5.write(row, 3, "Ozempic (Novo)", fmt["th_blue"])

    comparisons = [
        ("Umsatz Jahr 1", kpis_lilly["year1_revenue"], kpis_novo["year1_revenue"], "eur"),
        ("Umsatz 5J kumuliert", kpis_lilly["total_5y_revenue"], kpis_novo["total_5y_revenue"], "eur"),
        ("Gewinn 5J kumuliert", kpis_lilly["total_5y_profit"], kpis_novo["total_5y_profit"], "eur"),
        ("Peak Marktanteil", kpis_lilly["peak_share"], kpis_novo["peak_share"], "pct"),
        ("Share Monat 12", kpis_lilly["share_month_12"], kpis_novo["share_month_12"], "pct"),
        ("Share Monat 36", kpis_lilly["share_month_36"], kpis_novo["share_month_36"], "pct"),
        ("Share Monat 60", kpis_lilly["share_month_60"], kpis_novo["share_month_60"], "pct"),
        ("Durchschn. Preis/Mon.", kpis_lilly["avg_price"], kpis_novo["avg_price"], "eur"),
        ("Ind.-Multiplikator (Durchschn.)", kpis_lilly["avg_indication_multiplier"], kpis_novo["avg_indication_multiplier"], "mult"),
    ]
    for label, val_l, val_n, vtype in comparisons:
        row += 1
        ws5.write(row, 1, label, fmt["label_bold"])
        if vtype == "eur":
            ws5.write(row, 2, val_l, fmt["eur"])
            ws5.write(row, 3, val_n, fmt["eur"])
        elif vtype == "pct":
            ws5.write(row, 2, val_l, fmt["pct"])
            ws5.write(row, 3, val_n, fmt["pct"])
        elif vtype == "mult":
            ws5.write(row, 2, val_l, fmt["mult"])
            ws5.write(row, 3, val_n, fmt["mult"])

    # Scenario comparison for Lilly
    row += 3
    ws5.merge_range(row, 1, row, 4, "Lilly (Mounjaro) - Szenario-Vergleich", fmt["section"])

    scenario_results_lilly = {}
    for sn, bp, cp, mp in [
        ("Base Case", {}, {}, {}),
        ("Bull (Lilly)", {"target_peak_share_t2d": 0.35, "target_peak_share_obesity": 0.45, "months_to_peak": 24, "has_cv_indication": True},
         {}, {"market_growth_annual": 0.35, "obesity_gkv_coverage": True, "obesity_gkv_start_year": 2027}),
        ("Bull (Novo)", {"target_peak_share_t2d": 0.15, "target_peak_share_obesity": 0.20, "months_to_peak": 48},
         {"target_peak_share_t2d": 0.38}, {"market_growth_annual": 0.30}),
    ]:
        b = BrandParams(**{**{"name": "Mounjaro", "company": "Lilly", "current_share_trx": 0.08, "price_per_month_eur": 350.0,
                               "has_t2d": True, "has_obesity": True}, **bp})
        c = CompetitorParams(**{**{"name": "Ozempic", "company": "Novo", "current_share_trx": 0.36, "price_per_month_eur": 300.0,
                                   "supply_constrained": True, "supply_normalization_month": 12}, **cp})
        m = MarketParams(**mp)
        scenario_results_lilly[sn] = calculate_kpis_brand(forecast_brand(b, c, m))

    row += 1
    ws5.write(row, 1, "")
    ws5.write(row, 2, "Base Case", fmt["th"])
    ws5.write(row, 3, "Bull (Lilly)", fmt["th_orange"])
    ws5.write(row, 4, "Bull (Novo)", fmt["th_blue"])

    for label, key, vtype in [
        ("Umsatz Jahr 1", "year1_revenue", "eur"),
        ("Umsatz 5J kumuliert", "total_5y_revenue", "eur"),
        ("Gewinn 5J kumuliert", "total_5y_profit", "eur"),
        ("Peak Marktanteil", "peak_share", "pct"),
        ("Share Monat 12", "share_month_12", "pct"),
        ("Markt-Wachstum Gesamt", "market_growth_total", "mult"),
    ]:
        row += 1
        ws5.write(row, 1, label, fmt["label_bold"])
        for c, sn in enumerate(["Base Case", "Bull (Lilly)", "Bull (Novo)"]):
            val = scenario_results_lilly[sn].get(key, 0) or 0
            if vtype == "eur":
                ws5.write(row, 2 + c, val, fmt["eur"])
            elif vtype == "pct":
                ws5.write(row, 2 + c, val, fmt["pct"])
            elif vtype == "mult":
                ws5.write(row, 2 + c, val, fmt["mult"])

    # ═══════════════════════════════════════════════════════════════════
    # SHEET 6: METHODIK
    # ═══════════════════════════════════════════════════════════════════
    ws6 = wb.add_worksheet("Methodik")
    ws6.hide_gridlines(2)
    ws6.set_tab_color("#6366f1")
    ws6.set_column("A:A", 3)
    ws6.set_column("B:B", 34)
    ws6.set_column("C:C", 24)
    ws6.set_column("D:D", 52)
    ws6.set_column("E:E", 42)

    row = 1
    ws6.merge_range(row, 1, row, 4, "Methodik & Transparenz-Matrix", fmt["title"])
    row += 1
    ws6.merge_range(row, 1, row, 4,
        "Jeder Datenpunkt ist klassifiziert: FAKT / ANNAHME / MODELL / LUECKE",
        wb.add_format({"font_size": 10, "font_color": "#6b7280", "italic": True}))

    # Legend
    row += 2
    ws6.write(row, 1, "Legende:", fmt["label_bold"])
    ws6.write(row, 2, "FAKT", fmt["fact"])
    ws6.write(row, 3, "Aus oeffentlichen Quellen ableitbar", fmt["label"])
    row += 1
    ws6.write(row, 2, "ANNAHME", fmt["assumption"])
    ws6.write(row, 3, "Vom User steuerbar (gelbe Zellen)", fmt["label"])
    row += 1
    ws6.write(row, 2, "MODELL", fmt["model"])
    ws6.write(row, 3, "Mathematisches Modell (branchenvalidiert)", fmt["label"])
    row += 1
    ws6.write(row, 2, "LUECKE", fmt["gap"])
    ws6.write(row, 3, "Fuer Produktiveinsatz benoetigt", fmt["label"])

    # Facts
    row += 2
    ws6.merge_range(row, 1, row, 4, "FAKTEN (aus oeffentlichen Quellen)", fmt["section"])
    row += 1
    for c, h in enumerate(["Datenpunkt", "Wert im Modell", "Quelle", "Kategorie"]):
        ws6.write(row, 1 + c, h, fmt["th"])
    facts = [
        ("T2D-Patienten Deutschland", "6.33M GKV-Versicherte", "BARMER Gesundheitswesen aktuell 2025", "FAKT"),
        ("GLP-1-Penetration T2D (2024)", "13.9%", "BARMER GLP-1-Versorgungsanalyse 2025", "FAKT"),
        ("Mounjaro EU-Zulassung T2D", "Sep 2022", "EMA EPAR Mounjaro", "FAKT"),
        ("Mounjaro EU-Zulassung Adipositas", "Dez 2023", "EMA EPAR Mounjaro", "FAKT"),
        ("Mounjaro Marktstart DE", "Nov 2023", "Lilly Deutschland", "FAKT"),
        ("Mounjaro vertraulicher EB", "Ab Aug 2025", "G-BA Beschluss / Schiedsspruch", "FAKT"),
        ("Wegovy GKV-Ausschluss", "G-BA Juni 2024", "G-BA Beschluss: Lifestyle-Arzneimittel", "FAKT"),
        ("Tirzepatid Wirksamkeit", "-2.4% HbA1c, -22.5% Gewicht", "SURPASS / SURMOUNT Studienprogramm", "FAKT"),
        ("Semaglutid Wirksamkeit", "-1.8% HbA1c, -15% Gewicht", "SUSTAIN / STEP Studienprogramm", "FAKT"),
        ("SELECT-Studie", "-20% CV-Events (Semaglutid)", "NEJM, Lincoff et al. 2023", "FAKT"),
        ("Ozempic Lieferengpass", "Seit 2022, weltweit", "BfArM Lieferengpass-Datenbank", "FAKT"),
        ("Semaglutid-Patent Ablauf", "EU/US: ~2031+", "Espacenet / FDA Orange Book", "FAKT"),
    ]
    for dp, val, source, cat in facts:
        row += 1
        ws6.write(row, 1, dp, fmt["label_bold"])
        ws6.write(row, 2, val, fmt["fact"])
        ws6.write(row, 3, source, fmt["label"])
        ws6.write(row, 4, cat, fmt["fact"])

    # Assumptions
    row += 2
    ws6.merge_range(row, 1, row, 4, "ANNAHMEN (User-steuerbar)", fmt["section"])
    row += 1
    for c, h in enumerate(["Parameter", "Default-Wert", "Begruendung", "Kategorie"]):
        ws6.write(row, 1 + c, h, fmt["th"])
    assumptions = [
        ("Marktwachstum p.a.", "25%", "GLP-1-Segment historisch 25-36% Wachstum (IQVIA)", "ANNAHME"),
        ("Mounjaro Peak Share T2D", "25%", "Abhaengig von KOL-Ueberzeugung, Leitlinien-Update", "ANNAHME"),
        ("Mounjaro Peak Share Adipositas", "35%", "Bei GKV-Erstattung deutlich hoeher", "ANNAHME"),
        ("Monate bis Peak", "36", "S-Kurven-Annahme; Launch Nov 2023 + 36M = Nov 2026", "ANNAHME"),
        ("Novo Peak Share T2D", "30%", "Etablierte Position, breiteste Evidenz", "ANNAHME"),
        ("Adipositas GKV-Erstattung", "Nein (Base)", "Gesetzesaenderung noetig; fruehestens 2027", "ANNAHME"),
        ("GLP-1-Penetration T2D Ziel", "22%", "Leitlinien-Shift treibt Penetration nach oben", "ANNAHME"),
        ("Preistrend", "-3% p.a.", "Erfahrungswert bei etablierten Biologika", "ANNAHME"),
        ("COGS", "20%", "Peptid-Biologika: 15-25% typisch", "ANNAHME"),
        ("SG&A", "EUR 800K/Monat", "Firmenspezifisch (Aussendienst, KAM, Marketing)", "ANNAHME"),
    ]
    for param, val, reason, cat in assumptions:
        row += 1
        ws6.write(row, 1, param, fmt["label_bold"])
        ws6.write(row, 2, val, fmt["assumption"])
        ws6.write(row, 3, reason, fmt["label"])
        ws6.write(row, 4, cat, fmt["assumption"])

    # Models
    row += 2
    ws6.merge_range(row, 1, row, 4, "MATHEMATISCHE MODELLE", fmt["section"])
    row += 1
    for c, h in enumerate(["Modell", "Formel / Logik", "Validierung / Referenz", "Kategorie"]):
        ws6.write(row, 1 + c, h, fmt["th"])
    models = [
        ("Marktexpansion",
         "Decelerating Growth: TRx(t) = Start * (1+r_dampened)^t, cap = 4x Start",
         "Wachstumskurve mit logistischer Daempfung; typisch bei expandierenden Maerkten", "MODELL"),
        ("Share-Shift (S-Kurve)",
         "Logistic: Share(t) = Current + (Target-Current) * sigmoid(t, midpoint, steepness)",
         "Rogers Adoption Curve; Standard bei Brand-Launch-Forecasts", "MODELL"),
        ("Indikations-Multiplikator",
         "Additive Layer: T2D=1.0 + Obesity(0-50%) + CV(0-15%) + MASH(0-10%)",
         "Separate Ramp-Funktionen pro Indikation; Timing abhaengig von Zulassung", "MODELL"),
        ("Supply Constraint",
         "min(Demand, Capacity) bis Normalisierungsmonat",
         "Historisch validierbar durch Ozempic-Engpass 2022-2025", "MODELL"),
        ("Pricing",
         "Price(t) = Base * (1 + annual_trend)^(t/12) * (1 - AMNOG_cut)",
         "AMNOG Erstattungsbetrag-Mechanik; jaehrliche Preisentwicklung", "MODELL"),
    ]
    for name, formula, ref, cat in models:
        row += 1
        ws6.set_row(row, 40)
        ws6.write(row, 1, name, fmt["label_bold"])
        ws6.write(row, 2, formula, fmt["model"])
        ws6.write(row, 3, ref, fmt["label"])
        ws6.write(row, 4, cat, fmt["model"])

    # Gaps
    row += 2
    ws6.merge_range(row, 1, row, 4, "LUECKEN (fuer Produktiveinsatz benoetigt)", fmt["section"])
    row += 1
    for c, h in enumerate(["Gap", "Beschreibung", "Datenquelle", "Kategorie"]):
        ws6.write(row, 1 + c, h, fmt["th"])
    gaps = [
        ("Echte IQVIA-Daten", "TRx, NRx auf Monatsebene pro GLP-1-Produkt", "IQVIA Midas / DPM", "LUECKE"),
        ("Mounjaro vertraulicher Preis", "Tatsaechlicher GKV-Erstattungsbetrag (geheim)", "Lilly intern / GKV-Spitzenverband", "LUECKE"),
        ("Regionale Verschreibungsdaten", "KV-Regionen: unterschiedliche GLP-1-Adoption", "IQVIA Regional / KBV", "LUECKE"),
        ("Arzt-Segmentierung", "Diabetologen vs. Hausaerzte: unterschiedliches Verschreibungsverhalten", "IQVIA / Arztpanel", "LUECKE"),
        ("Supply-Engpass Detail", "Genaue Liefermengen Ozempic pro Monat", "BfArM / Novo Nordisk", "LUECKE"),
        ("Pipeline-Timing", "FDA/EMA-Einreichungen fuer CV/MASH-Indikationen", "ClinicalTrials.gov / Company Filings", "LUECKE"),
        ("Patienten-Switching", "Umstellungsraten zwischen GLP-1 Produkten", "Real World Evidence / IQVIA", "LUECKE"),
    ]
    for name, desc, source, cat in gaps:
        row += 1
        ws6.write(row, 1, name, fmt["label_bold"])
        ws6.write(row, 2, desc, fmt["gap"])
        ws6.write(row, 3, source, fmt["label"])
        ws6.write(row, 4, cat, fmt["gap"])

    # Key difference vs Eliquis model
    row += 3
    ws6.merge_range(row, 1, row, 4, "Strukturelle Unterschiede vs. Generic-Entry-Modell (Eliquis)", fmt["section"])
    row += 1
    for c, h in enumerate(["Dimension", "Generic Entry (Eliquis)", "Brand Competition (GLP-1)"]):
        ws6.write(row, 1 + c, h, fmt["th"])
    comparisons = [
        ("Marktdynamik", "Markt stagniert/schrumpft", "Markt expandiert (S-Kurve)"),
        ("Wettbewerber", "Multiple Generika vs. 1 Originator", "2 Branded Originatoren"),
        ("Preis-Mechanik", "Festbetrag / Generika-Erosion", "AMNOG Erstattungsbetrag"),
        ("Substitution", "Aut-idem (Apotheke)", "Keine; Arzt entscheidet"),
        ("Tender", "Rabattvertraege GKV", "Nicht relevant (Patentschutz)"),
        ("Indikationen", "Single Indication", "Multi-Indication (T2D, Obesity, CV, MASH)"),
        ("Supply", "Normalerweise kein Engpass", "Reale Supply Constraints"),
        ("Zeithorizont", "Post-LOE (Erosion)", "Pre-LOE (Expansion & Competition)"),
    ]
    for dim, eliquis, glp1 in comparisons:
        row += 1
        ws6.write(row, 1, dim, fmt["label_bold"])
        ws6.write(row, 2, eliquis, fmt["label"])
        ws6.write(row, 3, glp1, fmt["label"])

    # Disclaimer
    row += 3
    ws6.merge_range(row, 1, row, 4,
        "Hinweis: Alle Daten sind synthetisch und dienen der Illustration. "
        "Angelehnt an oeffentlich verfuegbare Quellen: BARMER 2025, G-BA, IQVIA, EMA. "
        "Kein Bezug zu vertraulichen Daten.",
        fmt["hint"])

    wb.close()
    print(f"Excel model saved to: {OUTPUT_PATH}")
    return OUTPUT_PATH


if __name__ == "__main__":
    build_model()
