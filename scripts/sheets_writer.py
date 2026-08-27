#!/usr/bin/env python3
"""Writes the revamp report as a Google Sheet instead of an .xlsx file — same
6 tabs, same data as build_report.py, but assembled via the Sheets + Drive
APIs. Text content for the Aide tab is imported from build_report.aide_blocks
so the two output formats never drift apart.
"""

from scripts.build_report import aide_blocks, intent_label, detailed_intent_fr

HEADER_COLOR = {"red": 0.353, "green": 0.322, "blue": 1.0}       # #5A52FF
HEADER_TEXT_COLOR = {"red": 1, "green": 1, "blue": 1}
TITLE_COLOR = {"red": 0.353, "green": 0.322, "blue": 1.0}
DANGER_COLOR = {"red": 1, "green": 0.125, "blue": 0.337}         # #FF2056
SUCCESS_COLOR = {"red": 0.008, "green": 0.773, "blue": 0.682}    # #02C5AE
MUTED_COLOR = {"red": 0.42, "green": 0.45, "blue": 0.5}


def create_spreadsheet(sheets_service, drive_service, title, sheet_titles):
    """Crée le Sheet avec tous les onglets demandés (dans l'ordre), supprime
    l'onglet par défaut, le partage en "anyone with link, can edit" (accord
    explicite de l'utilisateur), retourne (spreadsheet_id, url, {titre: sheetId})."""
    body = {
        "properties": {"title": title},
        "sheets": [{"properties": {"title": t}} for t in sheet_titles],
    }
    resp = sheets_service.spreadsheets().create(body=body, fields="spreadsheetId,spreadsheetUrl,sheets.properties").execute()
    spreadsheet_id = resp["spreadsheetId"]
    sheet_ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in resp["sheets"]}

    drive_service.permissions().create(fileId=spreadsheet_id, body={"type": "anyone", "role": "writer"}).execute()

    return spreadsheet_id, resp["spreadsheetUrl"], sheet_ids


def _values_update(sheets_service, spreadsheet_id, sheet_title, rows):
    if not rows:
        return
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=f"'{sheet_title}'!A1",
        valueInputOption="USER_ENTERED", body={"values": rows},
    ).execute()


def _fmt_header_row(sheet_id, ncols, row=0):
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": row, "endRowIndex": row + 1,
                      "startColumnIndex": 0, "endColumnIndex": ncols},
            "cell": {"userEnteredFormat": {
                "backgroundColor": HEADER_COLOR,
                "textFormat": {"foregroundColor": HEADER_TEXT_COLOR, "bold": True},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }


def _fmt_freeze(sheet_id, rows=1):
    return {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": rows}},
            "fields": "gridProperties.frozenRowCount",
        }
    }


def _fmt_wrap(sheet_id, start_col, end_col, start_row=0, end_row=1000):
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row,
                      "startColumnIndex": start_col, "endColumnIndex": end_col},
            "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP"}},
            "fields": "userEnteredFormat(wrapStrategy,verticalAlignment)",
        }
    }


def _fmt_col_width(sheet_id, col_index, width):
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": col_index, "endIndex": col_index + 1},
            "properties": {"pixelSize": width},
            "fields": "pixelSize",
        }
    }


def _fmt_bold_colored_cell(sheet_id, row, color, size=12):
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": row, "endRowIndex": row + 1,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": size, "foregroundColor": color}}},
            "fields": "userEnteredFormat.textFormat",
        }
    }


def write_aide_sheet(sheets_service, spreadsheet_id, sheet_id):
    rows = []
    bold_rows = []  # (row_index, color, size)
    wrap_rows = []

    def add(cell_text, bold=None, height=None):
        rows.append([cell_text])
        idx = len(rows) - 1
        if bold:
            bold_rows.append((idx, *bold))
        if height:
            wrap_rows.append(idx)

    for block in aide_blocks():
        t = block["type"]
        if t == "heading":
            add(block["text"], bold=(TITLE_COLOR, 16))
            rows.append([""])
        elif t == "title":
            add(block["text"], bold=(TITLE_COLOR, 13))
        elif t == "para":
            add(block["text"])
            wrap_rows.append(len(rows) - 1)
        elif t == "bullet":
            add(f"{block['label']} — {block['text']}")
            wrap_rows.append(len(rows) - 1)
        elif t == "spacer":
            rows.append([""])

    _values_update(sheets_service, spreadsheet_id, "Aide", rows)

    requests = [_fmt_col_width(sheet_id, 0, 900)]
    for idx, color, size in bold_rows:
        requests.append(_fmt_bold_colored_cell(sheet_id, idx, color, size))
    if wrap_rows:
        requests.append(_fmt_wrap(sheet_id, 0, 1, 0, len(rows)))
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def write_serp_sheet(sheets_service, spreadsheet_id, sheet_id, serp_data):
    rows = [[f"SERP — \"{serp_data['keyword']}\" ({serp_data.get('market', 'FR (google.fr)')})"],
            [], ["Position", "Titre", "URL", "Résumé"]]
    for item in serp_data["organic"]:
        rows.append([item["position"], item["title"], item["link"], item["note"]])
    rows.append([])
    rows.append(["Autres questions posées (PAA)"])
    for q in serp_data.get("paa", []):
        rows.append([f"• {q}"])
    rows.append([])
    rows.append(["Recherches associées"])
    for q in serp_data.get("related_searches", []):
        rows.append([f"• {q}"])
    rows.append([])
    rows.append(["Aperçu IA (AI Overview)"])
    rows.append(["Présent" if serp_data.get("ai_overview_present") else "Absent"])
    rows.append([serp_data.get("ai_overview_summary", "")])
    rows.append([])
    rows.append(["Knowledge Graph"])
    rows.append(["Présent" if serp_data.get("knowledge_graph_present") else "Absent"])

    _values_update(sheets_service, spreadsheet_id, "SERP", rows)
    requests = [
        _fmt_header_row(sheet_id, 4, row=2),
        _fmt_freeze(sheet_id, 3),
        _fmt_wrap(sheet_id, 0, 4),
        _fmt_col_width(sheet_id, 1, 260), _fmt_col_width(sheet_id, 2, 320), _fmt_col_width(sheet_id, 3, 400),
    ]
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def write_ahrefs_sheet(sheets_service, spreadsheet_id, sheet_id, ahrefs_rows, brand):
    ahrefs_rows = sorted(ahrefs_rows, key=lambda r: (r.get("volume") or 0), reverse=True)
    rows = [["Ahrefs Keywords Explorer — Matching terms (Terms match, All terms)"], [],
            ["Mot-clé", "Volume/mois", "KD", "CPC ($)", "Potentiel de trafic",
             "Intention (Ahrefs)", "Intention de recherche (description détaillée)", "Parent Topic"]]
    for row in ahrefs_rows:
        rows.append([
            row["keyword"], row.get("volume"), row.get("difficulty"), row.get("cpc"),
            row.get("traffic_potential"), intent_label(row.get("intents")),
            detailed_intent_fr(row["keyword"], brand), row.get("parent_topic"),
        ])

    _values_update(sheets_service, spreadsheet_id, "Ahrefs", rows)
    requests = [
        _fmt_header_row(sheet_id, 8, row=2),
        _fmt_freeze(sheet_id, 3),
        _fmt_wrap(sheet_id, 0, 8),
        _fmt_col_width(sheet_id, 0, 260), _fmt_col_width(sheet_id, 6, 480),
    ]
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def write_gsc_sheet(sheets_service, spreadsheet_id, sheet_id, gsc_rows, page_url):
    gsc_rows = [r for r in gsc_rows if r["clicks"] > 0]
    gsc_rows.sort(key=lambda r: r["clicks"], reverse=True)
    rows = [[f"GSC — {page_url}"], [], ["Requête", "Clics", "Impressions", "CTR", "Position moyenne"]]
    for row in gsc_rows:
        rows.append([row["query"], int(row["clicks"]), int(row["impressions"]),
                     round(row["ctr"] * 100, 2), round(row["position"], 2)])

    _values_update(sheets_service, spreadsheet_id, "GSC", rows)
    requests = [
        _fmt_header_row(sheet_id, 5, row=2),
        _fmt_freeze(sheet_id, 3),
        _fmt_col_width(sheet_id, 0, 320),
    ]
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def write_competitor_sheet(sheets_service, spreadsheet_id, sheet_id, competitors_data):
    rows = []
    bold_rows = []
    for comp in competitors_data:
        rows.append([comp["site"]])
        bold_rows.append(len(rows) - 1)
        rows.append(["URL :", comp["url"]])
        rows.append(["H1 :", comp["h1"]])
        if comp.get("no_dedicated_avis_page"):
            rows.append(["⚠️ Remarque", comp["summary"]])
        for block in comp["structure"]:
            rows.append(["", f"H2 : {block['h2']}"])
            for h3 in block["h3"]:
                rows.append(["", "", f"H3 : {h3}"])
        if not comp.get("no_dedicated_avis_page"):
            rows.append(["Résumé :", comp["summary"]])
        rows.append([])

    _values_update(sheets_service, spreadsheet_id, "Concurrents", rows)
    requests = [_fmt_wrap(sheet_id, 0, 3), _fmt_col_width(sheet_id, 1, 500), _fmt_col_width(sheet_id, 2, 400)]
    for idx in bold_rows:
        requests.append(_fmt_bold_colored_cell(sheet_id, idx, TITLE_COLOR, 12))
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def write_entity_sheet(sheets_service, spreadsheet_id, sheet_id, gap_data, page_url):
    n = gap_data.get("competitors", 0)
    rows = [[f"Écart d'entités — {page_url}"], [f"Comparé à {n} page(s) concurrente(s) via Google Natural Language API"], []]
    rows.append(["⚠️ Manquantes chez nous, présentes chez les concurrents"])
    danger_row = len(rows) - 1
    rows.append(["Entité", "Présente chez", "Reconnue par Google"])
    header_row = len(rows) - 1
    for g in gap_data.get("gaps", []):
        rows.append([g["name"], f"{g['on_competitors']}/{n} concurrent(s)",
                     "Oui" if g.get("mid") else "Vue mais pas identifiée"])
    if not gap_data.get("gaps"):
        rows.append(["(aucun écart détecté)"])
    rows.append([])
    rows.append(["✅ Partagées (nous + concurrents) — le socle attendu"])
    shared_row = len(rows) - 1
    rows.append([", ".join(f"{s['name']} ({s['on_competitors']}/{n})" for s in gap_data.get("shared", [])) or "(aucune)"])
    rows.append([])
    rows.append(["🟢 Uniquement chez nous — différenciant ou hors-sujet à vérifier"])
    unique_row = len(rows) - 1
    rows.append([", ".join(e["name"] for e in gap_data.get("our_unique", [])) or "(aucune)"])

    _values_update(sheets_service, spreadsheet_id, "Entités", rows)
    requests = [
        _fmt_header_row(sheet_id, 3, row=header_row),
        _fmt_freeze(sheet_id, header_row + 1),
        _fmt_wrap(sheet_id, 0, 3),
        _fmt_col_width(sheet_id, 0, 260),
        _fmt_bold_colored_cell(sheet_id, danger_row, DANGER_COLOR, 12),
        _fmt_bold_colored_cell(sheet_id, shared_row, TITLE_COLOR, 12),
        _fmt_bold_colored_cell(sheet_id, unique_row, SUCCESS_COLOR, 12),
    ]
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def build_google_sheet_report(sheets_service, drive_service, keyword, brand, serp_data, ahrefs_rows,
                               gsc_rows, competitors_data, page_url, entity_gap=None):
    """Assemble le rapport complet en Google Sheet, retourne son URL."""
    sheet_titles = ["Aide", "SERP", "Ahrefs", "GSC", "Concurrents"]
    if entity_gap:
        sheet_titles.append("Entités")

    spreadsheet_id, url, sheet_ids = create_spreadsheet(
        sheets_service, drive_service, f"{keyword} — {brand} — Revamp Report", sheet_titles,
    )

    write_aide_sheet(sheets_service, spreadsheet_id, sheet_ids["Aide"])
    write_serp_sheet(sheets_service, spreadsheet_id, sheet_ids["SERP"], serp_data)
    write_ahrefs_sheet(sheets_service, spreadsheet_id, sheet_ids["Ahrefs"], ahrefs_rows, brand)
    write_gsc_sheet(sheets_service, spreadsheet_id, sheet_ids["GSC"], gsc_rows, page_url)
    write_competitor_sheet(sheets_service, spreadsheet_id, sheet_ids["Concurrents"], competitors_data)
    if entity_gap:
        write_entity_sheet(sheets_service, spreadsheet_id, sheet_ids["Entités"], entity_gap, page_url)

    return url
