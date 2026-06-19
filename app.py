"""
DOCX Highlight-Analysator – Flask App für Render.com
=====================================================
Wertet ausschließlich <w:highlight w:val="..."> aus word/document.xml aus.

Explizit IGNORIERT:
  - w:color (Schriftfarben)
  - word/header*.xml, word/footer*.xml
  - Kommentare, Fußnoten, Endnoten

Absicherung: Anfragen müssen den API-Key im Header 'x-api-key' mitsenden.
Den Key als Umgebungsvariable API_KEY in Render setzen.
"""

import os
import json
import zipfile
import io
import base64
from xml.etree import ElementTree as ET
from flask import Flask, request, jsonify

app = Flask(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
SOURCE_FILE = "word/document.xml"

API_KEY = os.environ.get("API_KEY", "")


# ---------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------

def check_api_key():
    """Gibt True zurück wenn kein Key konfiguriert oder Key korrekt ist."""
    if not API_KEY:
        return True
    return request.headers.get("x-api-key", "") == API_KEY


def get_run_highlight(run_elem) -> str | None:
    rpr = run_elem.find(f"{W}rPr")
    if rpr is None:
        return None
    hl = rpr.find(f"{W}highlight")
    if hl is None:
        return None
    val = hl.get(f"{W}val", "")
    return val if val and val != "none" else None


def get_run_text(run_elem) -> str:
    return "".join(t.text or "" for t in run_elem.findall(f".//{W}t"))


def split_words(text: str) -> list:
    return [w for w in text.split() if w]


def extract_runs(root) -> list:
    runs = []
    for run in root.iter(f"{W}r"):
        text = get_run_text(run)
        if not text:
            continue
        runs.append({"text": text, "highlight": get_run_highlight(run)})
    return runs


def build_fundstellen(runs: list) -> tuple:
    fundstellen, unsichere = [], []
    i = 0
    while i < len(runs):
        run = runs[i]
        if run["highlight"] is None:
            i += 1
            continue

        farbe = run["highlight"]
        gruppe = [run]
        j = i + 1
        while j < len(runs):
            nr = runs[j]
            if nr["highlight"] == farbe:
                gruppe.append(nr)
                j += 1
            elif nr["highlight"] is None:
                if nr["text"].strip():
                    break
                gruppe.append(nr)
                j += 1
            else:
                break

        full_text = "".join(r["text"] for r in gruppe)
        woerter = split_words(full_text)
        unsicher_grund = None

        first_text = gruppe[0]["text"]
        last_text = gruppe[-1]["text"]

        if i > 0:
            prev_text = runs[i - 1]["text"]
            if (prev_text and not prev_text[-1].isspace()
                    and first_text and not first_text[0].isspace()):
                unsicher_grund = (
                    f"Erstes Wort möglicherweise nur teilweise markiert "
                    f"(Vorgänger endet auf '{prev_text[-8:].strip()}')"
                )

        if j < len(runs):
            nt = runs[j]["text"]
            if (last_text and not last_text[-1].isspace()
                    and nt and not nt[0].isspace()):
                g2 = (
                    f"Letztes Wort möglicherweise nur teilweise markiert "
                    f"(Nachfolger beginnt mit '{nt[:8].strip()}')"
                )
                unsicher_grund = (unsicher_grund + " | " + g2) if unsicher_grund else g2

        if unsicher_grund:
            unsichere.append({
                "textstelle": full_text[:120] + ("…" if len(full_text) > 120 else ""),
                "farbe": farbe,
                "grund": unsicher_grund,
            })
            woerter_s = woerter[1:-1] if len(woerter) > 2 else []
        else:
            woerter_s = woerter

        fundstellen.append({
            "farbe": farbe,
            "text": full_text[:300] + ("…" if len(full_text) > 300 else ""),
            "woerter": woerter_s if unsicher_grund else woerter,
            "wortanzahl": len(woerter_s if unsicher_grund else woerter),
        })
        i = j

    return fundstellen, unsichere


def parse_docx(file_bytes: bytes) -> dict:
    try:
        zf = zipfile.ZipFile(io.BytesIO(file_bytes))
    except zipfile.BadZipFile:
        raise ValueError("Datei ist kein gültiges DOCX (ZIP-Fehler).")

    if SOURCE_FILE not in zf.namelist():
        zf.close()
        raise ValueError(
            "Technische Prüfung nicht möglich, weil kein Zugriff auf "
            "word/document.xml der echten .docx-Datei besteht."
        )

    with zf.open(SOURCE_FILE) as f:
        try:
            root = ET.parse(f).getroot()
        except ET.ParseError as e:
            zf.close()
            raise ValueError(f"XML-Parsefehler in {SOURCE_FILE}: {e}")
    zf.close()

    alle_highlights = root.findall(f".//{W}highlight")
    highlight_anzahl = len(alle_highlights)
    gefundene_farben = sorted(set(
        hl.get(f"{W}val", "")
        for hl in alle_highlights
        if hl.get(f"{W}val", "") not in ("", "none")
    ))

    runs = extract_runs(root)
    fundstellen, unsichere = build_fundstellen(runs)

    summen: dict = {}
    for fs in fundstellen:
        summen[fs["farbe"]] = summen.get(fs["farbe"], 0) + fs["wortanzahl"]

    return {
        "vorpruefung": {
            "quelle": SOURCE_FILE,
            "highlight_elemente_anzahl": highlight_anzahl,
            "gefundene_farben": gefundene_farben,
            "schriftfarben_ausgewertet": False,
        },
        "fundstellen": fundstellen,
        "zwischensummen": summen,
        "gesamtsumme": sum(summen.values()),
        "unsichere_fundstellen": unsichere,
    }


def format_output(result: dict) -> str:
    lines = []
    vp = result["vorpruefung"]

    lines.append("## 0. Technische Vorprüfung\n")
    lines.append("| Prüfung | Ergebnis |")
    lines.append("|---|---|")
    lines.append(f"| Quelle geprüft | {vp['quelle']} |")
    lines.append(f"| Anzahl w:highlight-Elemente | {vp['highlight_elemente_anzahl']} |")
    farben_str = ", ".join(vp["gefundene_farben"]) if vp["gefundene_farben"] else "keine"
    lines.append(f"| Gefundene Highlight-Farben | {farben_str} |")
    lines.append("| Schriftfarben ausgewertet? | Nein |")

    if vp["highlight_elemente_anzahl"] == 0:
        lines.append("\n**Prüfung abgebrochen:** 0 w:highlight-Elemente gefunden.")
        return "\n".join(lines)

    lines.append("\n## 1. Detailprüfung je Fundstelle\n")
    lines.append("| Nr. | Highlight-Farbe | Exakter Textlaut | Gezählte Einzelwörter | Wortanzahl |")
    lines.append("|---|---|---|---|---|")
    for i, fs in enumerate(result["fundstellen"], 1):
        woerter_str = " / ".join(fs["woerter"]) if fs["woerter"] else "–"
        lines.append(f"| {i} | {fs['farbe']} | {fs['text']} | {woerter_str} | {fs['wortanzahl']} |")

    lines.append("\n## 2. Zwischensummen nach Highlight-Farbe\n")
    lines.append("| Highlight-Farbe | Summe Wörter |")
    lines.append("|---|---|")
    for farbe, summe in sorted(result["zwischensummen"].items()):
        lines.append(f"| {farbe} | {summe} |")

    lines.append("\n## 3. Rechenprüfung\n")
    lines.append("| Prüfung | Rechnung | Ergebnis |")
    lines.append("|---|---|---|")
    for farbe, summe in sorted(result["zwischensummen"].items()):
        nummern = [str(i) for i, fs in enumerate(result["fundstellen"], 1) if fs["farbe"] == farbe]
        rechnung = " + ".join(f"Nr.{n}" for n in nummern)
        lines.append(f"| Summe {farbe} | {rechnung} | {summe} |")
    zs_str = " + ".join(str(v) for v in result["zwischensummen"].values())
    lines.append(f"| Gesamtsumme | {zs_str} | {result['gesamtsumme']} |")

    lines.append("\n## 4. Unsichere Fundstellen\n")
    if result["unsichere_fundstellen"]:
        lines.append("| Nr. | Textstelle | Highlight-Farbe | Grund der Unsicherheit |")
        lines.append("|---|---|---|---|")
        for i, u in enumerate(result["unsichere_fundstellen"], 1):
            lines.append(f"| {i} | {u['textstelle']} | {u['farbe']} | {u['grund']} |")
    else:
        lines.append("Keine unsicheren Fundstellen.")

    lines.append("\n## 5. Ausgeschlossene Formatierungen\n")
    lines.append(
        "Schriftfarben wurden vollständig ignoriert. "
        "Insbesondere #000000, #222B52, #002057 und #231F20 wurden nicht gezählt, "
        "weil sie keine Word-Texthervorhebungen sind."
    )
    return "\n".join(lines)


# ---------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------

@app.route("/health", methods=["GET"])
def health():
    """Render nutzt diesen Endpoint um zu prüfen ob der Service läuft."""
    return jsonify({"status": "ok"}), 200


@app.route("/analyze-docx", methods=["POST"])
def analyze_docx():
    """
    POST /analyze-docx

    Header: x-api-key: <dein-key>

    Body (JSON):  { "file_base64": "<base64-DOCX>" }
    Body (raw):   DOCX-Bytes direkt

    Antwort: JSON mit vorpruefung, fundstellen, zwischensummen,
             gesamtsumme, unsichere_fundstellen, zusammenfassung
    """
    if not check_api_key():
        return jsonify({"error": "Ungültiger oder fehlender API-Key (Header: x-api-key)"}), 401

    try:
        content_type = request.content_type or ""
        if "application/json" in content_type:
            body = request.get_json(force=True)
            if not body or "file_base64" not in body:
                return jsonify({"error": "Feld 'file_base64' fehlt im JSON-Body"}), 400
            file_bytes = base64.b64decode(body["file_base64"])
        else:
            file_bytes = request.get_data()
            if not file_bytes:
                return jsonify({"error": "Kein Dateiinhalt im Request-Body"}), 400

        result = parse_docx(file_bytes)
        result["zusammenfassung"] = format_output(result)
        return jsonify(result), 200

    except ValueError as e:
        msg = str(e)
        return jsonify({"error": msg, "zusammenfassung": msg}), 400
    except Exception as e:
        return jsonify({"error": f"Interner Fehler: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
