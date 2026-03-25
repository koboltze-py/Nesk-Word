# -*- coding: utf-8 -*-
"""
PDF-Beispiele – 6 Designs mit reportlab
PDF bietet vs. Word: echte Gradients, freie Positionierung, Formen, Balken
"""
import os, sys
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.colors import HexColor, Color, white, black
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

W, H = A4   # 595.27 x 841.89 pt

# ── Pfade ─────────────────────────────────────────────────────────────────────
_OD = r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V"
ZIEL = os.path.join(_OD, "Desktop", "bei") if os.path.exists(_OD) else r"C:\Temp\bei"
os.makedirs(ZIEL, exist_ok=True)
LOGO = Path(os.path.dirname(os.path.abspath(__file__))) / "Daten" / "Email" / "Logo.jpg"
EXCEL = (
    _OD + r"\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26"
    r"\04_Tagesdienstpläne\03_März\25.03.2026.xlsx"
)

# ── Farben ────────────────────────────────────────────────────────────────────
def h(hex_): return HexColor(f"#{hex_}")

DRK_ROT   = h("BE0000"); DRK_HELL  = h("E53030")
WEISS     = white;       SCHWARZ   = black
ROT_WARN  = h("FF3333"); GRN_OK    = h("10A050"); ORG_WARN  = h("E07800")

# ── Daten laden ───────────────────────────────────────────────────────────────
def lade():
    try:
        from functions.dienstplan_parser import DienstplanParser
        r = DienstplanParser(EXCEL, alle_anzeigen=True).parse()
        if r.get("success") in (True, "True"): return r
    except Exception as e:
        print(f"  [Parser] {e}")
    return {"betreuer": [], "dispo": [], "kranke": []}

def _personal(d):
    return len([p for p in d.get('betreuer', [])+d.get('dispo', [])
                if p.get('ist_krank') not in (True, 'True')])

def _kranke(d):
    return [p.get('anzeigename', '') for p in d.get('kranke', [])
            if p.get('ist_krank') in (True, 'True')]

def _bulmor_fhr(d):
    alle = d.get('betreuer', []) + d.get('dispo', [])
    return [p.get('anzeigename', '') for p in alle if p.get('ist_bulmorfahrer') in (True, 'True')]

def _gruppen(d, typ='dispo'):
    gru = defaultdict(list)
    for p in d.get(typ, []):
        if p.get('ist_krank') in (True, 'True'): continue
        s = (p.get('start_zeit') or '')[:5]; e = (p.get('end_zeit') or '')[:5]
        gru[f"{s}–{e}"].append(p.get('anzeigename', ''))
    return dict(sorted(gru.items()))

def _bul_status(n):
    if n <= 2:  return ROT_WARN, "KRITISCH"
    elif n == 3: return ORG_WARN, "EINGESCHRÄNKT"
    else:        return GRN_OK,   "VOLLSTÄNDIG"

# ── Zeichen-Hilfsfunktionen ───────────────────────────────────────────────────
def y(c_from_top): return H - c_from_top   # reportlab: y=0 unten

def rect(c, x, yt, w, h_, color, stroke=0):
    c.setFillColor(color); c.setStrokeColor(color)
    c.rect(x, H - yt - h_, w, h_, fill=1, stroke=stroke)

def rect_stroke(c, x, yt, w, h_, fill_col, stroke_col, lw=1):
    c.setFillColor(fill_col); c.setStrokeColor(stroke_col)
    c.setLineWidth(lw); c.rect(x, H - yt - h_, w, h_, fill=1, stroke=1)

def circle(c, x, yt, r_, color):
    c.setFillColor(color); c.circle(x, H - yt, r_, fill=1, stroke=0)

def txt(c, x, yt, text, font="Helvetica", size=10, color=black, align="left"):
    c.setFont(font, size); c.setFillColor(color)
    if align == "center":
        c.drawCentredString(x, H - yt, text)
    elif align == "right":
        c.drawRightString(x, H - yt, text)
    else:
        c.drawString(x, H - yt, text)

def line(c, x1, yt1, x2, yt2, color, lw=1):
    c.setStrokeColor(color); c.setLineWidth(lw)
    c.line(x1, H - yt1, x2, H - yt2)

def logo_draw(c, x, yt, w=60, h_=50):
    if LOGO.exists():
        try:
            img = ImageReader(str(LOGO))
            c.drawImage(img, x, H - yt - h_, w, h_, mask='auto', preserveAspectRatio=True)
        except:
            pass

def progress_bar(cv, x, yt, width, height, wert, maxw, farbe, bg=h("DDDDDD")):
    rect(cv, x, yt, width, height, bg)
    if maxw > 0:
        w_ = width * min(wert, maxw) / maxw
        rect(cv, x, yt, w_, height, farbe)

# ── Verbrauchsmaterial ────────────────────────────────────────────────────────
MAT = [
    ("Einmalhandschuhe Nitril M/L",  "Karton",  "8",  "3"),
    ("Verbandpäckchen groß",         "Stück",   "20", "8"),
    ("Mullbinden 6/8/10 cm",         "Rollen",  "60", "20"),
    ("Wundkompressen steril 10×10",  "Päck.",   "40", "15"),
    ("Pflaster-Sortiment",           "Päck.",   "10", "4"),
    ("Hände-Desinfektion 1L",        "Fl.",     "6",  "2"),
    ("Flächen-Desinfektion 1L",      "Fl.",     "4",  "2"),
    ("Rettungsdecken Gold/Silber",   "Stück",   "20", "8"),
    ("Einmal-Beatmungsmaske CPR",    "Stück",   "10", "4"),
    ("Venenverweilkanüle G18/G20",   "Stück",   "30", "10"),
    ("Einmalspritze 5/10/20 ml",     "Stück",   "50", "20"),
    ("Infusionsset + NaCl 500ml",    "Sets",    "10", "4"),
    ("EKG-Elektroden Einmal",        "Päck.",   "8",  "3"),
    ("Sauerstoffmaske Einmal",       "Stück",   "15", "5"),
    ("FFP2-Atemschutzmaske",         "Stück",   "50", "20"),
    ("AED-Elektroden Einmal",        "Paar",    "4",  "2"),
    ("Blutzucker-Teststreifen",      "Päck.",   "5",  "2"),
    ("Urinbeutel steril",            "Stück",   "12", "4"),
    ("Blasenkatheter Ch 14/16",      "Stück",   "8",  "3"),
    ("CELOX Hämostase-Verband",      "Stück",   "4",  "2"),
]

DATUM = "25.03.2026"
UHRZEIT = "07:45 Uhr"
STATION = "Erste-Hilfe-Station · Flughafen Köln/Bonn"

# ═══════════════════════════════════════════════════════════════════════════════
# P1 – DASHBOARD NOIR (Dunkel + Neon-Akzente, maximale Design-Freiheit)
# ═══════════════════════════════════════════════════════════════════════════════
def pdf1_dashboard_noir(data, bul=4, einz=28, pat=5, pax=42500):
    BG    = h("0A0C10"); BG2   = h("141820"); BG3   = h("1E2530")
    NEON  = h("00E5FF"); NEON2 = h("00FF88"); GOLD  = h("FFD700")
    GRAU  = h("8899AA"); HELL  = h("CCDDEE")

    cv = canvas.Canvas(os.path.join(ZIEL, "P1_DashboardNoir_25032026.pdf"), pagesize=A4)

    # Hintergrund
    rect(cv, 0, 0, W, H, BG)

    # Header-Band
    rect(cv, 0, 0, W, 75, BG2)
    logo_draw(cv, 18, 10, w=55, h_=55)
    txt(cv, W/2, 28, "STÄRKEMELDUNG UND EINSÄTZE", "Helvetica-Bold", 20, NEON, "center")
    txt(cv, W/2, 46, f"DRK KÖLN  ·  {STATION}", "Helvetica", 9, GRAU, "center")
    txt(cv, W/2, 60, f"{DATUM}  ·  {UHRZEIT}", "Helvetica-Bold", 10, GOLD, "center")
    # Trennlinie
    line(cv, 0, 75, W, 75, NEON, 1.5)

    # ── Scorecard (4 Boxen) ───────────────────────────────────────────────────
    top = 95; bh = 70; bw = 120; gap = 10; lx = 28
    kz = [("EINSÄTZE", str(einz), NEON), ("PATIENTEN", str(pat), NEON2),
          ("PAX", f"{pax:,}".replace(",", "."), GOLD), ("PERSONAL", str(_personal(data)), NEON)]
    for i, (label, val, fc) in enumerate(kz):
        bx = lx + i * (bw + gap)
        rect(cv, bx, top, bw, bh, BG3)
        # Rand oben
        rect(cv, bx, top, bw, 3, fc)
        txt(cv, bx + bw/2, top + 18, label, "Helvetica-Bold", 8, GRAU, "center")
        txt(cv, bx + bw/2, top + 50, val, "Helvetica-Bold", 28, fc, "center")
    
    # ── Bulmor-Visualisierung ─────────────────────────────────────────────────
    btop = 185
    txt(cv, 28, btop + 5, "BULMOR – FAHRZEUGSTATUS", "Helvetica-Bold", 11, NEON)
    line(cv, 28, btop + 10, W - 28, btop + 10, NEON, 0.8)
    fc, label = _bul_status(bul)
    bw2 = 88; gap2 = 12; btop2 = btop + 22
    fahr = _bulmor_fhr(data)
    for i in range(5):
        bx2 = 28 + i * (bw2 + gap2)
        aktiv = (i+1) <= bul
        bgc = BG3 if not aktiv else fc
        rect(cv, bx2, btop2, bw2, 48, bgc)
        # LED-Indikator
        circle(cv, bx2 + bw2 - 14, btop2 + 14, 6, fc if aktiv else GRAU)
        txt(cv, bx2 + bw2/2, btop2 + 22, f"B{i+1}", "Helvetica-Bold", 16,
            WEISS if aktiv else GRAU, "center")
        lbl2 = "EINSATZ" if aktiv else "RESERVE"
        txt(cv, bx2 + bw2/2, btop2 + 40, lbl2, "Helvetica-Bold", 7,
            WEISS if aktiv else GRAU, "center")
    # Status-Label
    txt(cv, 28, btop2 + 62, f"  {bul} von 5 Fahrzeugen im Einsatz  —  {label}", "Helvetica-Bold", 11, fc)
    if fahr:
        txt(cv, 28, btop2 + 76, f"  Fahrer: {', '.join(fahr)}", "Helvetica", 8, GRAU)

    # ── Fortschrittsbalken Einsätze z. Tag ────────────────────────────────────
    ptop = 325
    txt(cv, 28, ptop, "EINSATZ-AUSLASTUNG DES TAGES", "Helvetica-Bold", 9, GRAU)
    progress_bar(cv, 28, ptop + 8, W - 56, 14, einz, 50, NEON)
    txt(cv, 28, ptop + 28, f"{einz} Einsätze", "Helvetica-Bold", 8, NEON)
    progress_bar(cv, 28, ptop + 34, W - 56, 14, pat, 20, NEON2)
    txt(cv, 28, ptop + 54, f"{pat} Pat. auf Station", "Helvetica-Bold", 8, NEON2)

    # ── Schichten (2 Spalten) ─────────────────────────────────────────────────
    stop = 390
    line(cv, 28, stop, W - 28, stop, GRAU, 0.4)
    txt(cv, 28, stop + 14, "DISPOSITION", "Helvetica-Bold", 10, NEON)
    txt(cv, W/2 + 10, stop + 14, "BEHINDERTENBETREUER", "Helvetica-Bold", 10, NEON2)
    line(cv, 28, stop + 18, W - 28, stop + 18, GRAU, 0.4)

    cy_d = stop + 30; cy_b = stop + 30
    for zeit, namen in _gruppen(data, 'dispo').items():
        txt(cv, 28, cy_d, zeit, "Helvetica-Bold", 8, GOLD)
        txt(cv, 28, cy_d + 12, ", ".join(namen), "Helvetica", 8, HELL)
        cy_d += 26

    for zeit, namen in _gruppen(data, 'betreuer').items():
        if cy_b > H - 130: break
        txt(cv, W/2 + 10, cy_b, zeit, "Helvetica-Bold", 8, GOLD)
        anzeige = ", ".join(namen)
        if len(anzeige) > 38: anzeige = anzeige[:36] + "…"
        txt(cv, W/2 + 10, cy_b + 12, anzeige, "Helvetica", 8, HELL)
        cy_b += 26

    # Trennlinie Mitte
    lmx = W/2 + 4
    line(cv, lmx, stop + 18, lmx, max(cy_d, cy_b), GRAU, 0.4)

    # ── Krank ─────────────────────────────────────────────────────────────────
    kl = _kranke(data)
    if kl:
        ky = max(cy_d, cy_b) + 8
        line(cv, 28, ky, W - 28, ky, ROT_WARN, 0.5)
        txt(cv, 28, ky + 14, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 9, ROT_WARN)

    # ── Footer ─────────────────────────────────────────────────────────────────
    rect(cv, 0, H - 30, W, 30, BG2)
    line(cv, 0, H - 30, W, H - 30, NEON, 0.8)
    txt(cv, W/2, H - 15, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de  ·  Stationsleitung: L. Peters",
        "Helvetica", 7, GRAU, "center")

    cv.save(); print(f"[OK] P1_DashboardNoir_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P2 – CLEAN WHITE (Professionell-minimalistisch, hoher Kontrast)
# ═══════════════════════════════════════════════════════════════════════════════
def pdf2_clean_white(data, bul=5, einz=28, pat=5, pax=42500):
    BLAU    = h("1A367C"); HELL_B  = h("EEF3FF"); GRAU_H  = h("F7F8FA")
    BLAU_D  = h("0C1F50"); GRAU_M  = h("9999AA"); AKZENT  = h("0060FF")

    cv = canvas.Canvas(os.path.join(ZIEL, "P2_CleanWhite_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, WEISS)

    # Header: volle Breite Blau-Balken oben
    rect(cv, 0, 0, W, 85, BLAU)
    logo_draw(cv, 20, 10, w=60, h_=62)
    txt(cv, W-25, 25, "Deutsches Rotes Kreuz", "Helvetica-Bold", 12, WEISS, "right")
    txt(cv, W-25, 40, "Kreisverband Köln e.V.", "Helvetica", 10, h("AABBDD"), "right")
    txt(cv, W-25, 56, STATION, "Helvetica", 8, h("AABBDD"), "right")
    txt(cv, W-25, 68, f"{DATUM}  |  {UHRZEIT}", "Helvetica-Bold", 9, h("FFD700"), "right")

    # Großer Titel
    txt(cv, 28, 110, "Stärkemeldung und Einsätze", "Helvetica-Bold", 24, BLAU_D)
    line(cv, 28, 120, W - 28, 120, AKZENT, 2.5)

    # ── Metriken-Zeile ─────────────────────────────────────────────────────────
    top = 135; bh = 60; bw = 118; gap = 9
    fc, label = _bul_status(bul)
    kz = [("Einsätze", str(einz), BLAU), ("Patienten", str(pat), h("107E3E")),
          ("Passagiere", f"{pax:,}".replace(",", "."), h("555577")),
          ("Bulmor", f"{bul}/5", fc)]
    for i, (lbl, val, fc2) in enumerate(kz):
        bx = 28 + i * (bw + gap)
        rect(cv, bx, top, bw, bh, GRAU_H)
        # Farbakzent links
        rect(cv, bx, top, 4, bh, fc2)
        txt(cv, bx + 14, top + 20, lbl, "Helvetica", 8, GRAU_M)
        txt(cv, bx + 14, top + 48, val, "Helvetica-Bold", 26, fc2)

    # ── Bulmor ─────────────────────────────────────────────────────────────────
    btop = 215
    txt(cv, 28, btop + 5, "Bulmor – Fahrzeugstatus", "Helvetica-Bold", 12, BLAU)
    line(cv, 28, btop + 10, W - 28, btop + 10, BLAU, 0.8)
    bw2 = 88; gap2 = 12; btop2 = btop + 22
    fahr = _bulmor_fhr(data)
    fc, label = _bul_status(bul)
    for i in range(5):
        bx2 = 28 + i * (bw2 + gap2)
        aktiv = (i+1) <= bul
        rect_stroke(cv, bx2, btop2, bw2, 45, HELL_B if aktiv else GRAU_H, fc if aktiv else h("CCCCCC"), 1.5)
        txt(cv, bx2 + bw2/2, btop2 + 16, f"Bulmor {i+1}", "Helvetica-Bold", 9,
            BLAU if aktiv else GRAU_M, "center")
        sym = "▶  Im Einsatz" if aktiv else "○  Bereit"
        txt(cv, bx2 + bw2/2, btop2 + 33, sym, "Helvetica-Bold", 8, fc if aktiv else GRAU_M, "center")
    txt(cv, 28, btop2 + 56, f"{bul} von 5 Fahrzeugen  —  {label}", "Helvetica-Bold", 10, fc)
    if fahr:
        txt(cv, 28, btop2 + 70, f"Fahrer: {', '.join(fahr)}", "Helvetica", 8, GRAU_M)

    # ── Dispo (links) + Betreuer (rechts) ─────────────────────────────────────
    stop = 370
    line(cv, 28, stop, W - 28, stop, h("CCCCDD"), 0.6)

    # Spaltenköpfe
    rect(cv, 28, stop + 2, (W-56)/2 - 5, 20, BLAU)
    rect(cv, W/2 + 4, stop + 2, (W-56)/2 - 5, 20, BLAU)
    txt(cv, 28 + (W-56)/4 - 5, stop + 16, "Disposition", "Helvetica-Bold", 9, WEISS, "center")
    txt(cv, W/2 + 4 + (W-56)/4 - 5, stop + 16, "Behindertenbetreuer", "Helvetica-Bold", 9, WEISS, "center")

    cy_d = stop + 30; cy_b = stop + 30
    for i, (zeit, namen) in enumerate(_gruppen(data, 'dispo').items()):
        bg = GRAU_H if i % 2 == 0 else WEISS
        rect(cv, 28, cy_d, (W-56)/2 - 5, 20, bg)
        txt(cv, 34, cy_d + 8, zeit, "Helvetica-Bold", 8, BLAU)
        txt(cv, 34, cy_d + 17, ", ".join(namen), "Helvetica", 7.5, h("333355"))
        cy_d += 22

    for i, (zeit, namen) in enumerate(_gruppen(data, 'betreuer').items()):
        if cy_b > H - 130: break
        bg = GRAU_H if i % 2 == 0 else WEISS
        rect(cv, W/2 + 4, cy_b, (W-56)/2 - 5, 20, bg)
        txt(cv, W/2 + 10, cy_b + 8, zeit, "Helvetica-Bold", 8, BLAU)
        anzeige = ", ".join(namen)
        if len(anzeige) > 40: anzeige = anzeige[:38] + "…"
        txt(cv, W/2 + 10, cy_b + 17, anzeige, "Helvetica", 7.5, h("333355"))
        cy_b += 22

    kl = _kranke(data)
    if kl:
        ky = max(cy_d, cy_b) + 6
        rect(cv, 28, ky, W - 56, 18, h("FFF0F0"))
        txt(cv, 34, ky + 12, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 9, h("CC2222"))

    # Footer
    rect(cv, 0, H - 28, W, 28, BLAU)
    txt(cv, W/2, H - 11, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de",
        "Helvetica", 8, h("AABBDD"), "center")

    cv.save(); print(f"[OK] P2_CleanWhite_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P3 – INFOGRAPHIC (Bunte Kacheln, anschaulich, Balken, Kreise)
# ═══════════════════════════════════════════════════════════════════════════════
def pdf3_infographic(data, bul=3, einz=28, pat=5, pax=42500):
    LILA  = h("6C3483"); PINK  = h("D2366A"); CYAN  = h("0097A7"); GELB = h("F5A623")
    GRUEN = h("27AE60"); HELL  = h("F8F9FA"); GRAU  = h("888888"); DUNKEL= h("2C3E50")
    fc, label = _bul_status(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P3_Infographic_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, HELL)

    # Bunter Header
    rect(cv, 0, 0, W/4, 90, PINK)
    rect(cv, W/4, 0, W/4, 90, LILA)
    rect(cv, W/2, 0, W/4, 90, CYAN)
    rect(cv, 3*W/4, 0, W/4, 90, GELB)
    logo_draw(cv, 12, 8, w=50, h_=55)
    txt(cv, W/2, 35, "Stärkemeldung & Einsätze", "Helvetica-Bold", 18, WEISS, "center")
    txt(cv, W/2, 53, f"DRK Köln  ·  {DATUM}  ·  {UHRZEIT}", "Helvetica-Bold", 9, WEISS, "center")
    txt(cv, W/2, 68, STATION, "Helvetica", 8, h("EEEEEE"), "center")

    # ── Runde Zahlen-Kacheln ──────────────────────────────────────────────────
    kacheln = [("Einsätze", str(einz), PINK), ("Patienten", str(pat), GRUEN),
               ("Personal", str(_personal(data)), LILA), ("Bulmor", f"{bul}/5", fc)]
    kw = 105; kh = 80; ktop = 105
    for i, (lbl, val, kc) in enumerate(kacheln):
        kx = 28 + i * (kw + 12)
        rect(cv, kx, ktop, kw, kh, kc)
        # Weißer Kreis
        cv.setFillColor(white); cv.setFillAlpha(0.15)
        cv.circle(kx + kw/2, H - ktop - kh * 0.35, 28, fill=1, stroke=0)
        cv.setFillAlpha(1.0)
        txt(cv, kx + kw/2, ktop + kh * 0.42 - 8, val, "Helvetica-Bold", 22, WEISS, "center")
        txt(cv, kx + kw/2, ktop + 68, lbl, "Helvetica-Bold", 8, WEISS, "center")

    # ── Bulmor – 5 Kreise graphisch ───────────────────────────────────────────
    btop = 202
    txt(cv, 28, btop + 5, "Bulmor-Fahrzeuge", "Helvetica-Bold", 12, DUNKEL)
    line(cv, 28, btop + 12, W - 28, btop + 12, h("DDDDEE"), 1)
    for i in range(5):
        cx_ = 60 + i * 100; cy_ = btop + 55
        aktiv = (i+1) <= bul
        # Großer Kreis  
        cv.setFillColor(fc if aktiv else h("CCCCCC"))
        cv.circle(cx_, H - cy_, 32, fill=1, stroke=0)
        # Innerer weißer Kreis
        cv.setFillColor(HELL if aktiv else WEISS)
        cv.circle(cx_, H - cy_, 22, fill=1, stroke=0)
        txt(cv, cx_, cy_, f"B{i+1}", "Helvetica-Bold", 12, fc if aktiv else h("AAAAAA"), "center")
        txt(cv, cx_, cy_ + 40, "Einsatz" if aktiv else "Bereit", "Helvetica-Bold", 7,
            fc if aktiv else h("AAAAAA"), "center")
    # Status
    rect(cv, 28, btop + 92, W - 56, 22, fc)
    txt(cv, W/2, btop + 108, f"{bul} von 5 Bulmor im Einsatz  —  {label}",
        "Helvetica-Bold", 11, WEISS, "center")
    fahr = _bulmor_fhr(data)
    if fahr:
        txt(cv, 28, btop + 122, f"Fahrer: {', '.join(fahr)}", "Helvetica", 8, GRAU)

    # ── Balken-Chart: Schichten nach Schichtbeginn ────────────────────────────
    ctop = 370; clefte = 28; cbreite = W - 56
    txt(cv, 28, ctop + 5, "Besetzung nach Schichtbeginn", "Helvetica-Bold", 11, DUNKEL)
    line(cv, 28, ctop + 10, W - 28, ctop + 10, h("CCCCDD"), 0.8)
    cy_ = ctop + 24
    alle_g = {**_gruppen(data, 'dispo'), **_gruppen(data, 'betreuer')}
    max_n = max((len(n) for n in alle_g.values()), default=1)
    bar_colors = [PINK, LILA, CYAN, GELB, GRUEN, h("E74C3C"), h("3498DB"), h("1ABC9C"), h("E67E22")]
    for ci, (zeit, namen) in enumerate(alle_g.items()):
        if cy_ > H - 140: break
        bar_w = (cbreite - 90) * len(namen) / max(max_n, 1)
        rect(cv, clefte + 90, cy_, bar_w, 13, bar_colors[ci % len(bar_colors)])
        txt(cv, clefte + 86, cy_ + 10, zeit, "Helvetica-Bold", 7, DUNKEL, "right")
        txt(cv, clefte + 94 + bar_w, cy_ + 10, str(len(namen)), "Helvetica-Bold", 7, DUNKEL)
        cy_ += 17

    kl = _kranke(data)
    if kl:
        kp = max(cy_, ctop + 24) + 6
        rect(cv, 28, kp, W - 56, 16, h("FFE5E5"))
        txt(cv, 34, kp + 11, f"Krank: {', '.join(kl)}", "Helvetica-Bold", 8, h("CC2222"))

    # Footer
    rect(cv, 0, H - 25, W, 25, DUNKEL)
    txt(cv, W/2, H - 9, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de",
        "Helvetica", 7, h("AABBCC"), "center")

    cv.save(); print(f"[OK] P3_Infographic_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P4 – ZEITUNGSSEITE / BROADSHEET (Klare Spalten, Journalisten-Optik)
# ═══════════════════════════════════════════════════════════════════════════════
def pdf4_broadsheet(data, bul=2, einz=28, pat=5, pax=42500):
    ROT_B  = h("AA0000"); DUNKEL = h("1A1A1A"); GRAU_H = h("F5F5F5"); GRAU_M = h("888888")
    LINE_C = h("333333")

    cv = canvas.Canvas(os.path.join(ZIEL, "P4_Broadsheet_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, WEISS)

    # Zeitungskopf
    line(cv, 28, 20, W - 28, 20, DUNKEL, 3)
    line(cv, 28, 22, W - 28, 22, DUNKEL, 0.5)
    txt(cv, W/2, 42, "TAGESBERICHT", "Helvetica-Bold", 32, DUNKEL, "center")
    txt(cv, W/2, 58, "Erste-Hilfe-Station Flughafen Köln/Bonn  ·  DRK Kreisverband Köln e.V.",
        "Helvetica", 8, GRAU_M, "center")
    line(cv, 28, 64, W - 28, 64, DUNKEL, 3)
    # Datum-Zeile
    txt(cv, 28, 76, DATUM, "Helvetica-Bold", 9, GRAU_M)
    txt(cv, W/2, 76, "STÄRKEMELDUNG & EINSÄTZE", "Helvetica-Bold", 9, ROT_B, "center")
    txt(cv, W - 28, 76, UHRZEIT, "Helvetica-Bold", 9, GRAU_M, "right")
    line(cv, 28, 80, W - 28, 80, h("CCCCCC"), 0.5)

    # Schlagzeilen-Kasten (Einsätze groß)
    fc, label = _bul_status(bul)
    rect(cv, 28, 88, W - 56, 65, GRAU_H)
    rect(cv, 28, 88, 4, 65, ROT_B)
    txt(cv, 40, 106, "EINSÄTZE HEUTE:", "Helvetica-Bold", 11, DUNKEL)
    txt(cv, 40, 132, str(einz), "Helvetica-Bold", 36, ROT_B)
    txt(cv, 120, 132, f"|  Patienten auf Station: {pat}", "Helvetica-Bold", 13, DUNKEL)
    txt(cv, 120, 113, f"PAX: {pax:,}".replace(",", ".") + "  |  Personal: " + str(_personal(data)),
        "Helvetica", 9, GRAU_M)
    txt(cv, 120, 148, f"Bulmor: {bul}/5 im Einsatz  —  {label}", "Helvetica-Bold", 10, fc)

    # ── Bulmor-Grafik (5 Rechtecke) ───────────────────────────────────────────
    btop = 165
    txt(cv, 28, btop, "Bulmor Fahrzeugstatus", "Helvetica-Bold", 10, DUNKEL)
    line(cv, 28, btop + 4, W - 28, btop + 4, h("CCCCCC"), 0.4)
    for i in range(5):
        bx = 28 + i * 108; aktiv = (i+1) <= bul
        rect_stroke(cv, bx, btop + 8, 100, 30,
                    fc if aktiv else GRAU_H, fc if aktiv else h("CCCCCC"))
        txt(cv, bx + 50, btop + 21, f"Bulmor {i+1}", "Helvetica-Bold", 9,
            WEISS if aktiv else GRAU_M, "center")
        txt(cv, bx + 50, btop + 32, "EINSATZ" if aktiv else "bereit", "Helvetica-Bold", 7,
            WEISS if aktiv else GRAU_M, "center")
    fahr = _bulmor_fhr(data)
    if fahr:
        txt(cv, 28, btop + 46, f"Fahrer: {', '.join(fahr)}", "Helvetica", 8, GRAU_M)

    # ── 3-Spalten Layout Dispo | Betreuer | Material ───────────────────────────
    col_w = (W - 56 - 20) / 3; col_gap = 10
    c1x = 28; c2x = c1x + col_w + col_gap; c3x = c2x + col_w + col_gap
    stop = 222

    for cx_, titel, gruppe_typ in [(c1x, "DISPOSITION", 'dispo'), (c2x, "BETREUER", 'betreuer')]:
        line(cv, cx_, stop, cx_ + col_w, stop, DUNKEL, 1.5)
        txt(cv, cx_ + col_w/2, stop + 14, titel, "Helvetica-Bold", 9, DUNKEL, "center")
        line(cv, cx_, stop + 17, cx_ + col_w, stop + 17, h("CCCCCC"), 0.4)
        cy_ = stop + 28
        for i, (zeit, namen) in enumerate(_gruppen(data, gruppe_typ).items()):
            if cy_ > H - 60: break
            bg = GRAU_H if i % 2 == 0 else WEISS
            rect(cv, cx_, cy_, col_w, 18, bg)
            txt(cv, cx_ + 4, cy_ + 8, zeit, "Helvetica-Bold", 7, ROT_B)
            anzeige = ", ".join(namen)
            if len(anzeige) > 24: anzeige = anzeige[:22] + "…"
            txt(cv, cx_ + 4, cy_ + 15, anzeige, "Helvetica", 7, DUNKEL)
            cy_ += 19

    # Spalte 3: Material
    line(cv, c3x, stop, c3x + col_w, stop, DUNKEL, 1.5)
    txt(cv, c3x + col_w/2, stop + 14, "MEDIZINPRODUKTE", "Helvetica-Bold", 9, DUNKEL, "center")
    line(cv, c3x, stop + 17, c3x + col_w, stop + 17, h("CCCCCC"), 0.4)
    my_ = stop + 28
    for i, (mat, eh, soll, mind) in enumerate(MAT):
        if my_ > H - 60: break
        bg = GRAU_H if i % 2 == 0 else WEISS
        rect(cv, c3x, my_, col_w, 16, bg)
        txt(cv, c3x + 4, my_ + 11, mat[:22], "Helvetica", 6.5, DUNKEL)
        txt(cv, c3x + col_w - 4, my_ + 11, f"Soll: {soll}", "Helvetica-Bold", 6.5, ROT_B, "right")
        my_ += 17

    kl = _kranke(data)
    if kl:
        line(cv, 28, H - 40, W - 28, H - 40, ROT_B, 1)
        txt(cv, 28, H - 29, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 9, ROT_B)

    # Footer
    line(cv, 28, H - 20, W - 28, H - 20, DUNKEL, 1)
    txt(cv, 28, H - 9, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de", "Helvetica", 7, GRAU_M)
    txt(cv, W - 28, H - 9, "Seite 1", "Helvetica", 7, GRAU_M, "right")

    cv.save(); print(f"[OK] P4_Broadsheet_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P5 – MEDICAL REPORT (Klinisch, Klinikakte-Stil, Blau/Grau/Weiß)
# ═══════════════════════════════════════════════════════════════════════════════
def pdf5_medical_report(data, bul=5, einz=28, pat=5, pax=42500):
    TEAL   = h("006D6D"); TEAL_H = h("E0F5F5"); GRAU_D = h("3D4E5A"); GRAU_H = h("F4F7F9")
    TEAL_M = h("009999"); GRAU_M = h("778899"); LINE_C = h("C0CFD8")
    GRN    = h("168740"); ORG2   = h("C07000"); ROT2   = h("AC1515")

    cv = canvas.Canvas(os.path.join(ZIEL, "P5_MedicalReport_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, GRAU_H)

    # Linker farbiger Rand
    rect(cv, 0, 0, 8, H, TEAL)

    # Header
    rect(cv, 8, 0, W - 8, 80, WEISS)
    logo_draw(cv, 18, 10, w=58, h_=58)
    txt(cv, 90, 25, "STÄRKEMELDUNG UND EINSÄTZE", "Helvetica-Bold", 17, TEAL)
    txt(cv, 90, 43, "Deutsches Rotes Kreuz Kreisverband Köln e.V.", "Helvetica", 9, GRAU_D)
    txt(cv, 90, 57, STATION, "Helvetica", 8.5, GRAU_M)
    # Rechtsbündig Datum
    rect(cv, W - 130, 10, 115, 58, TEAL_H)
    txt(cv, W - 20, 30, DATUM, "Helvetica-Bold", 14, TEAL, "right")
    txt(cv, W - 20, 48, UHRZEIT, "Helvetica-Bold", 10, GRAU_D, "right")
    txt(cv, W - 20, 62, "Mittwoch", "Helvetica", 8.5, GRAU_M, "right")
    line(cv, 8, 80, W, 80, LINE_C, 1)

    # Metriken – klinische Zeile
    rect(cv, 8, 82, W - 8, 50, WEISS)
    fc, label = _bul_status(bul)
    metr = [("Einsätze gesamt", str(einz), TEAL_M), ("Pat. auf Station", str(pat), GRN),
            ("PAX", f"{pax:,}".replace(",", "."), GRAU_M), ("Personal aktiv", str(_personal(data)), TEAL_M),
            ("Bulmor-Status", f"{bul}/5 — {label}", fc)]
    mw = (W - 24) / len(metr)
    for i, (lbl, val, mc) in enumerate(metr):
        mx = 8 + i * mw
        if i > 0: line(cv, mx, 84, mx, 130, LINE_C, 0.6)
        txt(cv, mx + mw/2, 98, lbl, "Helvetica", 7.5, GRAU_M, "center")
        txt(cv, mx + mw/2, 122, val, "Helvetica-Bold", 16, mc, "center")
    line(cv, 8, 132, W, 132, LINE_C, 1)

    # Bulmor-Status-Zeile
    rect(cv, 8, 134, W - 8, 60, WEISS)
    txt(cv, 20, 150, "Fahrzeugstatus Bulmor", "Helvetica-Bold", 10, TEAL)
    bw2 = 82; gap2 = 10; bfahr = _bulmor_fhr(data)
    for i in range(5):
        bx2 = 20 + i * (bw2 + gap2)
        aktiv = (i+1) <= bul
        rect(cv, bx2, 155, bw2, 32, TEAL_H if aktiv else h("EEEEEE"))
        line(cv, bx2, 155, bx2 + bw2, 155, fc if aktiv else LINE_C, 2)
        txt(cv, bx2 + bw2/2, 167, f"Bulmor {i+1}", "Helvetica-Bold", 8.5,
            TEAL if aktiv else GRAU_M, "center")
        txt(cv, bx2 + bw2/2, 180, "Im Einsatz" if aktiv else "Bereit", "Helvetica-Bold", 7.5,
            fc if aktiv else GRAU_M, "center")
    if bfahr:
        txt(cv, 20, 194, f"Bulmor-Fahrer: {', '.join(bfahr)}", "Helvetica", 7.5, GRAU_M)
    line(cv, 8, 198, W, 198, LINE_C, 1)

    # Schraffierte Abschnitte Dispo / Betreuer / Material
    def abschnitt_header(cv_, x_, yt_, breite, titel, farbe):
        rect(cv_, x_, yt_, breite, 18, farbe)
        txt(cv_, x_ + 6, yt_ + 13, titel, "Helvetica-Bold", 8.5, WEISS)

    stop = 200
    # 3 Spalten
    c1w = 175; c2w = 175; c3w = W - 16 - c1w - c2w - 20
    c1x = 10; c2x = c1x + c1w + 10; c3x = c2x + c2w + 10

    abschnitt_header(cv, c1x, stop, c1w, "DISPOSITION", TEAL)
    abschnitt_header(cv, c2x, stop, c2w, "BEHINDERTENBETREUER", TEAL_M)
    abschnitt_header(cv, c3x, stop, c3w, "MEDIZINPRODUKTE", GRAU_D)

    cy1 = stop + 22; cy2 = stop + 22; cy3 = stop + 22
    for i, (zeit, namen) in enumerate(_gruppen(data, 'dispo').items()):
        if cy1 > H - 40: break
        bg = TEAL_H if i % 2 == 0 else WEISS
        rect(cv, c1x, cy1, c1w, 20, bg)
        txt(cv, c1x + 4, cy1 + 9, zeit, "Helvetica-Bold", 7.5, TEAL)
        anzeige = ", ".join(namen)
        if len(anzeige) > 22: anzeige = anzeige[:20] + "…"
        txt(cv, c1x + 4, cy1 + 18, anzeige, "Helvetica", 7, GRAU_D)
        cy1 += 21

    for i, (zeit, namen) in enumerate(_gruppen(data, 'betreuer').items()):
        if cy2 > H - 40: break
        bg = TEAL_H if i % 2 == 0 else WEISS
        rect(cv, c2x, cy2, c2w, 20, bg)
        txt(cv, c2x + 4, cy2 + 9, zeit, "Helvetica-Bold", 7.5, TEAL_M)
        anzeige = ", ".join(namen)
        if len(anzeige) > 22: anzeige = anzeige[:20] + "…"
        txt(cv, c2x + 4, cy2 + 18, anzeige, "Helvetica", 7, GRAU_D)
        cy2 += 21

    for i, (mat, eh, soll, mind) in enumerate(MAT):
        if cy3 > H - 40: break
        bg = h("F0F8F8") if i % 2 == 0 else WEISS
        rect(cv, c3x, cy3, c3w, 18, bg)
        txt(cv, c3x + 4, cy3 + 7, mat[:18], "Helvetica", 6.5, GRAU_D)
        txt(cv, c3x + 4, cy3 + 15, f"Soll: {soll} {eh}", "Helvetica-Bold", 6, TEAL_M)
        txt(cv, c3x + c3w - 4, cy3 + 11, f"Min: {mind}", "Helvetica-Bold", 6.5, ROT2, "right")
        cy3 += 19

    kl = _kranke(data)
    if kl:
        ky = max(cy1, cy2, cy3) + 6
        if ky < H - 30:
            rect(cv, 8, ky, W - 16, 18, h("FFE8E8"))
            txt(cv, 14, ky + 12, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 8, ROT2)

    # Footer
    line(cv, 8, H - 22, W, H - 22, LINE_C, 1)
    rect(cv, 0, H - 22, 8, 22, TEAL)
    txt(cv, 16, H - 9, f"DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de  ·  Stationsleitung: Lars Peters",
        "Helvetica", 7, GRAU_M)

    cv.save(); print(f"[OK] P5_MedicalReport_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P6 – AIRPORT SIGNAGE (Anzeigetafel-Optik: dunkel, weiß, Gate-Stil)
# ═══════════════════════════════════════════════════════════════════════════════
def pdf6_airport_signage(data, bul=4, einz=28, pat=5, pax=42500):
    BG    = h("1C1C1E"); BG2   = h("2C2C2E"); BG3   = h("3A3A3C")
    GRN   = h("30D158"); GELB  = h("FFD60A"); GRAU  = h("8E8E93"); HELL  = h("F2F2F7")
    fc, label = _bul_status(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P6_AirportSignage_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, BG)

    # Top-Banner like flight board
    rect(cv, 0, 0, W, 60, BG2)
    logo_draw(cv, 16, 8, w=45, h_=45)
    txt(cv, W/2, 22, "ERSTE-HILFE-STATION  •  FLUGHAFEN KÖLN/BONN", "Helvetica-Bold", 13, GRN, "center")
    txt(cv, W/2, 40, "DRK KREISVERBAND KÖLN E.V.", "Helvetica", 9, GRAU, "center")
    txt(cv, W - 20, 22, DATUM, "Helvetica-Bold", 12, GELB, "right")
    txt(cv, W - 20, 40, UHRZEIT, "Helvetica-Bold", 10, HELL, "right")
    line(cv, 0, 60, W, 60, GRN, 2)

    # Großer Titeltext
    txt(cv, 28, 92, "STÄRKEMELDUNG", "Helvetica-Bold", 26, HELL)
    txt(cv, 28, 116, "UND EINSÄTZE", "Helvetica-Bold", 26, GRN)
    line(cv, 28, 122, 380, 122, GRAU, 0.5)

    # ── Status-Panel rechts ────────────────────────────────────────────────────
    rect(cv, W - 160, 65, 148, 66, BG3)
    txt(cv, W - 84, 82, "PAX HEUTE", "Helvetica-Bold", 8, GRAU, "center")
    txt(cv, W - 84, 105, f"{pax:,}".replace(",", "."), "Helvetica-Bold", 20, GELB, "center")
    txt(cv, W - 84, 122, f"Einsätze: {einz}  |  Pat.: {pat}", "Helvetica", 8, GRAU, "center")

    # ── Bulmor-Board ──────────────────────────────────────────────────────────
    btop = 132; bw2 = 88; gap2 = 11
    txt(cv, 28, btop + 5, "BULMOR FLOTTENÜBERSICHT", "Helvetica-Bold", 10, GRN)
    line(cv, 28, btop + 10, W - 28, btop + 10, BG3, 1)
    for i in range(5):
        bx = 28 + i * (bw2 + gap2)
        aktiv = (i+1) <= bul
        rect(cv, bx, btop + 14, bw2, 52, BG3)
        # Status-LED
        cv.setFillColor(fc if aktiv else h("555555"))
        cv.circle(bx + bw2 - 12, H - (btop + 14) - 10, 7, fill=1, stroke=0)
        # Gate-Nummer-Stil
        txt(cv, bx + bw2/2, btop + 38, f"B – 0{i+1}", "Helvetica-Bold", 14,
            HELL if aktiv else GRAU, "center")
        status_txt = "IM EINSATZ" if aktiv else "STANDBY"
        rect(cv, bx + 6, btop + 44, bw2 - 12, 16, fc if aktiv else h("555555"))
        txt(cv, bx + bw2/2, btop + 55, status_txt, "Helvetica-Bold", 7,
            WEISS, "center")
    txt(cv, 28, btop + 74, f"Status: {bul}/5  —  {label}",
        "Helvetica-Bold", 10, fc)
    fahr = _bulmor_fhr(data)
    if fahr:
        txt(cv, 28, btop + 88, f"Fahrer: {', '.join(fahr)}", "Helvetica", 8, GRAU)

    # ── Anzeigetafel-Raster Schichten ─────────────────────────────────────────
    stop = 238
    # Spaltenköpfe wie Abflugtafel
    rect(cv, 28, stop, W - 56, 20, BG2)
    line(cv, 28, stop, W - 28, stop, GRN, 1)
    txt(cv, 50, stop + 14, "SCHICHT", "Helvetica-Bold", 8.5, GRN)
    txt(cv, 170, stop + 14, "ROLLE", "Helvetica-Bold", 8.5, GRN)
    txt(cv, 250, stop + 14, "MITARBEITER", "Helvetica-Bold", 8.5, GRN)
    txt(cv, W - 35, stop + 14, "ANZ.", "Helvetica-Bold", 8.5, GRN, "right")

    cy_ = stop + 22; ri = 0
    for gruppe_typ, rolle in [('dispo', 'DISPO'), ('betreuer', 'BETREUER')]:
        for zeit, namen in _gruppen(data, gruppe_typ).items():
            if cy_ > H - 50: break
            bg = BG3 if ri % 2 == 0 else BG2
            rect(cv, 28, cy_, W - 56, 18, bg)
            txt(cv, 50, cy_ + 12, zeit, "Helvetica-Bold", 8, GELB)
            txt(cv, 170, cy_ + 12, rolle, "Helvetica-Bold", 7.5, GRN)
            anzeige = ", ".join(namen)
            if len(anzeige) > 42: anzeige = anzeige[:40] + "…"
            txt(cv, 250, cy_ + 12, anzeige, "Helvetica", 8, HELL)
            txt(cv, W - 35, cy_ + 12, str(len(namen)), "Helvetica-Bold", 9, GELB, "right")
            cy_ += 19; ri += 1

    kl = _kranke(data)
    if kl:
        rect(cv, 28, cy_ + 4, W - 56, 18, h("3D1010"))
        line(cv, 28, cy_ + 4, W - 28, cy_ + 4, ROT_WARN, 1.5)
        txt(cv, 40, cy_ + 17, f"KRANK: {', '.join(kl)}", "Helvetica-Bold", 8.5, ROT_WARN)

    # Material
    mp = max(cy_ + 28, btop + 320); 
    if mp < H - 60:
        rect(cv, 28, mp, W - 56, 18, BG2)
        line(cv, 28, mp, W - 28, mp, GRN, 1)
        txt(cv, 50, mp + 13, "MEDIZINPRODUKTE (Auswahl)", "Helvetica-Bold", 8.5, GRN)
        mcy = mp + 22
        for i, (mat, eh, soll, mind) in enumerate(MAT[:10]):
            if mcy > H - 35: break
            bg = BG3 if i % 2 == 0 else BG2
            rect(cv, 28, mcy, W - 56, 15, bg)
            txt(cv, 50, mcy + 10, mat, "Helvetica", 7.5, HELL)
            txt(cv, W - 50, mcy + 10, f"Soll: {soll} {eh}", "Helvetica-Bold", 7.5, GELB, "right")
            mcy += 16

    # Footer-Bar
    line(cv, 0, H - 28, W, H - 28, GRN, 1.5)
    rect(cv, 0, H - 28, W, 28, BG2)
    txt(cv, W/2, H - 11, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de  ·  Stationsleitung: Lars Peters",
        "Helvetica", 7, GRAU, "center")

    cv.save(); print(f"[OK] P6_AirportSignage_25032026.pdf")


# ── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Lade Dienstplan ...")
    data = lade()
    print(f"  Betreuer: {len(data.get('betreuer',[]))}  |  Dispo: {len(data.get('dispo',[]))}  |  Krank: {len(data.get('kranke',[]))}")
    print(f"\nZielordner: {ZIEL}\n")
    print("Erstelle 6 PDF-Designs ...")
    pdf1_dashboard_noir(data,     bul=4, einz=28, pat=5, pax=42500)
    pdf2_clean_white(data,        bul=5, einz=28, pat=5, pax=42500)
    pdf3_infographic(data,        bul=3, einz=28, pat=5, pax=42500)
    pdf4_broadsheet(data,         bul=2, einz=28, pat=5, pax=42500)
    pdf5_medical_report(data,     bul=5, einz=28, pat=5, pax=42500)
    pdf6_airport_signage(data,    bul=4, einz=28, pat=5, pax=42500)
    print(f"\n✓ Alle 6 PDFs gespeichert in:\n  {ZIEL}")
