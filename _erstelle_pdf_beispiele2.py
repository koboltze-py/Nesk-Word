# -*- coding: utf-8 -*-
"""
PDF-Beispiele Batch 2 – 6 neue Designs
Inspiriert durch Online-Recherche: Warby Parker, Adidas, WWF, Upstack, Geometric, Corporate
"""
import os, sys, math
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.colors import HexColor, Color, white, black
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.graphics.shapes import Drawing, Wedge, Rect, String, Line
from reportlab.graphics import renderPDF

W, H = A4

# ── Pfade ─────────────────────────────────────────────────────────────────────
_OD = r"C:\Users\DRKairport\OneDrive - Deutsches Rotes Kreuz - Kreisverband Köln e.V"
ZIEL = os.path.join(_OD, "Desktop", "bei") if os.path.exists(_OD) else r"C:\Temp\bei"
os.makedirs(ZIEL, exist_ok=True)
LOGO = Path(os.path.dirname(os.path.abspath(__file__))) / "Daten" / "Email" / "Logo.jpg"
EXCEL = (
    _OD + r"\Dateien von Erste-Hilfe-Station-Flughafen - DRK Köln e.V_ - !Gemeinsam.26"
    r"\04_Tagesdienstpläne\03_März\25.03.2026.xlsx"
)
DATUM = "25.03.2026"; UHRZEIT = "07:45 Uhr"
STATION = "Erste-Hilfe-Station · Flughafen Köln/Bonn"

# ── Farb-Helfer ───────────────────────────────────────────────────────────────
def h(hex_): return HexColor(f"#{hex_}")

ROT_WARN = h("FF3333"); GRN_OK = h("10A050"); ORG_WARN = h("E07800")

def _bul_col(n):
    if n <= 2:  return ROT_WARN, "KRITISCH"
    elif n == 3: return ORG_WARN, "EINGESCHRÄNKT"
    else:        return GRN_OK,   "VOLLSTÄNDIG"

# ── Daten ─────────────────────────────────────────────────────────────────────
def lade():
    try:
        from functions.dienstplan_parser import DienstplanParser
        r = DienstplanParser(EXCEL, alle_anzeigen=True).parse()
        if r.get("success") in (True, "True"): return r
    except Exception as e:
        print(f"  [Parser] {e}")
    return {"betreuer": [], "dispo": [], "kranke": []}

def _personal(d): return len([p for p in d.get('betreuer',[])+d.get('dispo',[]) if p.get('ist_krank') not in (True,'True')])
def _kranke(d):   return [p.get('anzeigename','') for p in d.get('kranke',[]) if p.get('ist_krank') in (True,'True')]
def _bulfhr(d):
    alle = d.get('betreuer',[])+d.get('dispo',[])
    return [p.get('anzeigename','') for p in alle if p.get('ist_bulmorfahrer') in (True,'True')]
def _gruppen(d, typ='dispo'):
    gru = defaultdict(list)
    for p in d.get(typ,[]):
        if p.get('ist_krank') in (True,'True'): continue
        s=(p.get('start_zeit') or '')[:5]; e=(p.get('end_zeit') or '')[:5]
        gru[f"{s}–{e}"].append(p.get('anzeigename',''))
    return dict(sorted(gru.items()))

# ── Zeichen-Primitiven ────────────────────────────────────────────────────────
def rect(cv, x, yt, w, ht, color):
    cv.setFillColor(color); cv.rect(x, H-yt-ht, w, ht, fill=1, stroke=0)

def rect_outline(cv, x, yt, w, ht, fill, stroke_c, lw=1):
    cv.setFillColor(fill); cv.setStrokeColor(stroke_c)
    cv.setLineWidth(lw); cv.rect(x, H-yt-ht, w, ht, fill=1, stroke=1)

def line(cv, x1, yt1, x2, yt2, color, lw=1):
    cv.setStrokeColor(color); cv.setLineWidth(lw)
    cv.line(x1, H-yt1, x2, H-yt2)

def t(cv, x, yt, text, font="Helvetica", size=10, color=black, align="left"):
    cv.setFont(font, size); cv.setFillColor(color)
    if align == "center": cv.drawCentredString(x, H-yt, str(text))
    elif align == "right": cv.drawRightString(x, H-yt, str(text))
    else: cv.drawString(x, H-yt, str(text))

def logo(cv, x, yt, w=55, ht=50):
    if LOGO.exists():
        try: cv.drawImage(ImageReader(str(LOGO)), x, H-yt-ht, w, ht, mask='auto', preserveAspectRatio=True)
        except: pass

def vline(cv, x, yt1, yt2, color, lw=0.5):
    cv.setStrokeColor(color); cv.setLineWidth(lw)
    cv.line(x, H-yt1, x, H-yt2)

def donut_chart(cv, cx, cy, outer_r, inner_r, value, max_val, fill_col, bg_col=h("EEEEEE")):
    """Einfaches Donut-Segment"""
    sweepA = 360.0 * min(value, max_val) / max(max_val, 1)
    # Hintergrund-Kreis
    cv.setFillColor(bg_col); cv.circle(cx, cy, outer_r, fill=1, stroke=0)
    # Segment (approximiert mit vielen kleinen Dreiecken)
    cv.setFillColor(fill_col)
    import math
    steps = max(int(sweepA), 1)
    start = 90  # oben
    for i in range(steps):
        angle = math.radians(start - i)
        angle2 = math.radians(start - (i + 1))
        px1 = cx + outer_r * math.cos(angle); py1 = cy + outer_r * math.sin(angle)
        px2 = cx + outer_r * math.cos(angle2); py2 = cy + outer_r * math.sin(angle2)
        cv.setFillColor(fill_col)
        p = cv.beginPath()
        p.moveTo(cx, cy); p.lineTo(px1, py1); p.lineTo(px2, py2)
        p.close(); cv.drawPath(p, fill=1, stroke=0)
    # Ausschnitt (innerer Kreis = Loch)
    cv.setFillColor(white); cv.circle(cx, cy, inner_r, fill=1, stroke=0)

def bar_h(cv, x, yt, total_w, ht, value, max_v, fill_c, bg_c=h("E8E8E8")):
    rect(cv, x, yt, total_w, ht, bg_c)
    if max_v > 0:
        w = total_w * min(value, max_v) / max_v
        rect(cv, x, yt, w, ht, fill_c)

# ── Material ──────────────────────────────────────────────────────────────────
MAT = [
    ("Einmalhandschuhe Nitril M/L","Karton","8","3"),("Verbandpäckchen groß","Stück","20","8"),
    ("Mullbinden 6/8/10 cm","Rollen","60","20"),("Wundkompressen 10×10","Päck.","40","15"),
    ("Pflaster-Sortiment","Päck.","10","4"),("Hände-Desinfektion 1L","Fl.","6","2"),
    ("Flächen-Desinfektion 1L","Fl.","4","2"),("Rettungsdecken","Stück","20","8"),
    ("Einmal-Beatmungsmaske","Stück","10","4"),("Venenverweilkanüle G18/20","Stück","30","10"),
    ("Einmalspritze 5/10/20 ml","Stück","50","20"),("Infusionsset + NaCl 500ml","Sets","10","4"),
    ("EKG-Elektroden Einmal","Päck.","8","3"),("Sauerstoffmaske Einmal","Stück","15","5"),
    ("FFP2-Atemschutzmaske","Stück","50","20"),("AED-Elektroden Einmal","Paar","4","2"),
    ("Blutzucker-Teststreifen","Päck.","5","2"),("Urinbeutel steril","Stück","12","4"),
    ("Blasenkatheter Ch14/16","Stück","8","3"),("CELOX Hämostase-Verband","Stück","4","2"),
]


# ═══════════════════════════════════════════════════════════════════════════════
# P7 – WARBY MAGAZINE (Magazine-Stil: viel Weißraum, Bold Blue, serifenlos)
# Inspiriert: Warby Parker Impact Report – bold blue on white, magazine layout
# ═══════════════════════════════════════════════════════════════════════════════
def P7_warby_magazine(data, bul=5, einz=28, pat=5, pax=42500):
    BLAU   = h("0A3D7C"); BLAU_H = h("E8F0FF"); AKZENT = h("E8003D")
    GRAU_L = h("F9F9F9"); GRAU_M = h("999999"); DUNKEL = h("111111")
    TEXT   = h("333333"); LINT   = h("CCCCCC")
    fc, label = _bul_col(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P7_WarbyMagazin_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, white)

    # Schmaler blauer Top-Streifen
    rect(cv, 0, 0, W, 8, BLAU)
    # Logo + Org-Name
    logo(cv, 28, 16, w=50, ht=48)
    t(cv, 88, 32, "Deutsches Rotes Kreuz", "Helvetica-Bold", 11, BLAU)
    t(cv, 88, 46, "Kreisverband Köln e.V.  ·  " + STATION, "Helvetica", 8, GRAU_M)
    t(cv, W-28, 32, DATUM, "Helvetica-Bold", 11, DUNKEL, "right")
    t(cv, W-28, 46, UHRZEIT, "Helvetica", 9, GRAU_M, "right")

    # Dicke horizontale Bar unter Header
    line(cv, 28, 72, W-28, 72, DUNKEL, 2)
    line(cv, 28, 74, W-28, 74, DUNKEL, 0.4)

    # Mammut-Titel (Magazine-Stil)
    t(cv, 28, 110, "STÄRKEMELDUNG", "Helvetica-Bold", 38, BLAU)
    t(cv, 28, 143, "UND EINSÄTZE", "Helvetica-Bold", 38, DUNKEL)
    # Rote Akzentlinie
    rect(cv, 28, 152, 120, 5, AKZENT)

    # Datum groß rechts (Magazine-Seitenzahl-Optik)
    t(cv, W-30, 110, "25", "Helvetica-Bold", 72, h("EEEEEE"), "right")
    t(cv, W-30, 145, "MRZ", "Helvetica-Bold", 18, GRAU_M, "right")

    # ── 4 Kennzahlen-Boxen ────────────────────────────────────────────────────
    top = 170; bw = 118; bht = 65; gap = 10
    kz = [("EINSÄTZE", str(einz), BLAU), ("PATIENTEN", str(pat), h("0A7A3D")),
          ("PAX", f"{pax:,}".replace(",","."), h("555577")), ("BULMOR", f"{bul}/5", fc)]
    for i, (lbl, val, fc2) in enumerate(kz):
        bx = 28 + i*(bw+gap)
        rect(cv, bx, top, bw, bht, GRAU_L)
        rect(cv, bx, top, bw, 4, fc2)
        t(cv, bx+12, top+22, lbl, "Helvetica-Bold", 7.5, GRAU_M)
        t(cv, bx+12, top+58, val, "Helvetica-Bold", 28, fc2)

    # ── Bulmor – horizontale Fortschrittsbalken (Magazine-Check-Stil) ──────────
    btop = 250
    line(cv, 28, btop, W-28, btop, LINT, 0.6)
    t(cv, 28, btop+16, "BULMOR – FAHRZEUGSTATUS", "Helvetica-Bold", 11, BLAU)
    for i in range(5):
        aktiv = (i+1) <= bul
        by = btop + 28 + i*22
        sym = "●" if aktiv else "○"
        sym_col = fc if aktiv else GRAU_M
        t(cv, 28, by, f"{sym}  Bulmor {i+1}", "Helvetica-Bold", 10, sym_col)
        bar_h(cv, 130, by-10, 340, 12, 1 if aktiv else 0, 1, fc if aktiv else LINT)
        t(cv, 478, by, "Im Einsatz" if aktiv else "Stand-by", "Helvetica-Bold", 8,
          fc if aktiv else GRAU_M)

    sp = btop + 28 + 5*22 + 6
    t(cv, 28, sp, f"  {bul} von 5 Bulmor im Einsatz  —  Status: {label}", "Helvetica-Bold", 10, fc)
    fhr = _bulfhr(data)
    if fhr:
        t(cv, 28, sp+14, f"  Fahrer: {', '.join(fhr)}", "Helvetica", 8.5, GRAU_M)

    # ── 2-Spalten Schichten ────────────────────────────────────────────────────
    stop = sp + 30; lw2 = (W-76)/2
    line(cv, 28, stop, W-28, stop, LINT, 0.6)

    # Spalte 1: Dispo
    t(cv, 28, stop+16, "DISPOSITION", "Helvetica-Bold", 10, BLAU)
    cy1 = stop+28
    for i, (zeit, namen) in enumerate(_gruppen(data,'dispo').items()):
        if cy1 > H-90: break
        t(cv, 28, cy1, zeit, "Helvetica-Bold", 8, AKZENT)
        t(cv, 28, cy1+11, "  "+", ".join(namen), "Helvetica", 8.5, TEXT)
        cy1 += 25

    # Trennlinie
    vline(cv, W/2, stop+8, max(cy1, stop+200), LINT)

    # Spalte 2: Betreuer
    t(cv, W/2+12, stop+16, "BEHINDERTENBETREUER", "Helvetica-Bold", 10, BLAU)
    cy2 = stop+28
    for i, (zeit, namen) in enumerate(_gruppen(data,'betreuer').items()):
        if cy2 > H-90: break
        t(cv, W/2+12, cy2, zeit, "Helvetica-Bold", 8, AKZENT)
        an = ", ".join(namen); an = an[:42]+"…" if len(an)>43 else an
        t(cv, W/2+12, cy2+11, "  "+an, "Helvetica", 8.5, TEXT)
        cy2 += 25

    kl = _kranke(data)
    if kl:
        ky = max(cy1,cy2)+8
        line(cv, 28, ky, W-28, ky, LINT, 0.5)
        t(cv, 28, ky+14, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 9, AKZENT)

    # Footer
    rect(cv, 0, H-22, W, 22, BLAU)
    line(cv, 0, H-22, W, H-22, BLAU, 1)
    t(cv, W/2, H-8, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de  ·  Stationsleitung: L. Peters",
      "Helvetica", 7.5, white, "center")

    cv.save(); print("[OK] P7_WarbyMagazin_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P8 – ADIDAS BOLD (Schwarz + Elektrisch-Gelb, Dashboard, High-Contrast)
# Inspiriert: Adidas Annual Report – bold high-contrast, gradient numbers
# ═══════════════════════════════════════════════════════════════════════════════
def P8_adidas_bold(data, bul=3, einz=28, pat=5, pax=42500):
    BG    = h("0A0A0A"); GELB  = h("FFE500"); WEISS_A = white
    GRAU  = h("1C1C1C"); GRAU2 = h("2A2A2A"); GRAU_M  = h("666666")
    GRAU_H= h("BBBBBB"); HELL  = h("F0F0F0")
    fc, label = _bul_col(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P8_AdidasBold_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, BG)

    # HEADER: 3 gelbe diagonale Balken oben links
    for i in range(3):
        offset = i * 14
        p = cv.beginPath()
        p.moveTo(offset, H); p.lineTo(offset+9, H); p.lineTo(offset+9+50, H-50); p.lineTo(offset+50, H-50)
        p.close(); cv.setFillColor(GELB); cv.drawPath(p, fill=1, stroke=0)
    logo(cv, 58, 10, w=50, ht=48)
    t(cv, 118, 28, "STÄRKEMELDUNG", "Helvetica-Bold", 22, WEISS_A)
    t(cv, 118, 46, "UND EINSÄTZE", "Helvetica-Bold", 22, GELB)
    t(cv, W-28, 28, DATUM, "Helvetica-Bold", 13, GELB, "right")
    t(cv, W-28, 45, UHRZEIT, "Helvetica", 9, GRAU_H, "right")
    t(cv, W-28, 58, STATION, "Helvetica", 7.5, GRAU_M, "right")
    line(cv, 0, 72, W, 72, GELB, 2)

    # ── Kennzahlen – riesige Zahlen, gelb ─────────────────────────────────────
    kz = [("EINSÄTZE", str(einz), GELB), ("PATIENTEN", str(pat), h("00FF88")),
          ("PAX", f"{pax:,}".replace(",","."), GRAU_H), ("PERSONAL", str(_personal(data)), GELB)]
    bw = 118; gap = 10
    for i, (lbl, val, fc2) in enumerate(kz):
        bx = 28 + i*(bw+gap)
        rect(cv, bx, 80, bw, 75, GRAU)
        rect(cv, bx, 80, bw, 3, fc2)
        t(cv, bx+bw/2, 100, lbl, "Helvetica-Bold", 7.5, GRAU_M, "center")
        t(cv, bx+bw/2, 140, val, "Helvetica-Bold", 30, fc2, "center")

    # ── Bulmor: 5 dicke Quadrate ───────────────────────────────────────────────
    btop = 168; bw2 = 88; gap2 = 11
    t(cv, 28, btop, "BULMOR – FLOTTE", "Helvetica-Bold", 11, GELB)
    line(cv, 28, btop+4, W-28, btop+4, GRAU2, 1)
    fhr = _bulfhr(data)
    for i in range(5):
        bx2 = 28 + i*(bw2+gap2)
        aktiv = (i+1) <= bul
        fill_c = fc if aktiv else GRAU2
        rect(cv, bx2, btop+8, bw2, 60, fill_c)
        # Diagonaler Trennstrich (Design-Akzent)
        p2 = cv.beginPath()
        p2.moveTo(bx2+bw2-18, H-(btop+8))
        p2.lineTo(bx2+bw2, H-(btop+8))
        p2.lineTo(bx2+bw2, H-(btop+8+18))
        p2.close()
        cv.setFillColor(BG if aktiv else GRAU_M); cv.drawPath(p2, fill=1, stroke=0)
        t(cv, bx2+bw2/2, btop+35, f"B{i+1}", "Helvetica-Bold", 18, WEISS_A if aktiv else GRAU_M, "center")
        t(cv, bx2+bw2/2, btop+56, "AKTIV" if aktiv else "STANDBY", "Helvetica-Bold", 7,
          BG if aktiv else GRAU_M, "center")
    sp = btop+76; fc2 = fc
    t(cv, 28, sp, f"  Status: {bul}/5 im Einsatz  —  {label}", "Helvetica-Bold", 11, fc2)
    if fhr:
        t(cv, 28, sp+14, f"  Fahrer: {', '.join(fhr)}", "Helvetica", 8.5, GRAU_H)

    # ── Balken-Auslastung ─────────────────────────────────────────────────────
    bp = sp+30
    line(cv, 28, bp, W-28, bp, GRAU2, 0.8)
    t(cv, 28, bp+14, "AUSLASTUNG", "Helvetica-Bold", 9, GELB)
    bar_h(cv, 28, bp+20, W-56, 16, einz, 50, GELB, bg_c=GRAU2)
    t(cv, 28, bp+43, f"Einsätze: {einz}/50 mgl.", "Helvetica-Bold", 8, GRAU_H)
    bar_h(cv, 28, bp+48, W-56, 16, pat, 20, h("00FF88"), bg_c=GRAU2)
    t(cv, 28, bp+71, f"Patienten auf Station: {pat}/20 mgl.", "Helvetica-Bold", 8, GRAU_H)

    # ── Schichten: kompakte Liste ─────────────────────────────────────────────
    stop = bp+84
    line(cv, 28, stop, W-28, stop, GRAU2, 0.8)
    t(cv, 28, stop+14, "DISPOSITION", "Helvetica-Bold", 9, GELB)
    vline(cv, W/2+4, stop+8, H-24, GRAU2)
    t(cv, W/2+14, stop+14, "BEHINDERTENBETREUER", "Helvetica-Bold", 9, GELB)

    cy1 = stop+26
    for i, (zeit, namen) in enumerate(_gruppen(data,'dispo').items()):
        if cy1 > H-40: break
        bg = GRAU if i%2==0 else BG
        rect(cv, 28, cy1, W/2-36, 20, bg)
        t(cv, 34, cy1+9, zeit, "Helvetica-Bold", 8, GELB)
        t(cv, 34, cy1+17, ", ".join(namen), "Helvetica", 7.5, GRAU_H)
        cy1 += 21

    cy2 = stop+26
    for i, (zeit, namen) in enumerate(_gruppen(data,'betreuer').items()):
        if cy2 > H-40: break
        bg = GRAU if i%2==0 else BG
        rect(cv, W/2+8, cy2, W/2-36, 20, bg)
        t(cv, W/2+14, cy2+9, zeit, "Helvetica-Bold", 8, GELB)
        an = ", ".join(namen)[:40]
        t(cv, W/2+14, cy2+17, an, "Helvetica", 7.5, GRAU_H)
        cy2 += 21

    kl = _kranke(data)
    if kl:
        ky = max(cy1, cy2)+4
        if ky < H-28:
            rect(cv, 28, ky, W-56, 16, h("3A0000"))
            t(cv, 34, ky+11, f"KRANK: {', '.join(kl)}", "Helvetica-Bold", 8.5, ROT_WARN)

    # Footer
    rect(cv, 0, H-20, W, 20, GELB)
    t(cv, W/2, H-7, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de",
      "Helvetica-Bold", 7.5, BG, "center")

    cv.save(); print("[OK] P8_AdidasBold_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P9 – WWF TURQUOISE (Türkis/Petrol + weiße Textfelder + Natur-Ästhetik)
# Inspiriert: WWF Living Planet Report – turquoise, square frames, bold white
# ═══════════════════════════════════════════════════════════════════════════════
def P9_wwf_turquoise(data, bul=5, einz=28, pat=5, pax=42500):
    TEAL   = h("009B8D"); TEAL_D = h("006D63"); TEAL_H = h("E0F7F5")
    TEAL_M = h("00B5A5"); WEISS  = white; GRAU_H = h("F5F5F5"); GRAU_M = h("88AAAA")
    TERRA  = h("D45A30"); DUNKEL = h("1A2E2C")
    fc, label = _bul_col(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P9_WWFTurquoise_25032026.pdf"), pagesize=A4)

    # Türkis-Seite
    rect(cv, 0, 0, W, H, TEAL)

    # Weißes Fenster (rechts)
    rect(cv, W*0.42, 0, W*0.58, H, WEISS)

    # ── Linke Seite (Türkis) ──────────────────────────────────────────────────
    logo(cv, 22, 18, w=55, ht=52)
    t(cv, 22, 88, "Stärkemeldung", "Helvetica-Bold", 26, WEISS)
    t(cv, 22, 112, "und Einsätze", "Helvetica-Bold", 26, TEAL_H)
    line(cv, 22, 120, W*0.4, 120, h("FFFFFF"), 2)
    t(cv, 22, 135, DATUM + "  ·  " + UHRZEIT, "Helvetica-Bold", 10, TEAL_H)
    t(cv, 22, 150, STATION, "Helvetica", 8.5, h("AADDDD"))

    # Schmale weiße Quadrat-Kacheln (WWF-Stil)
    kacheln = [("Einsätze", str(einz), TEAL_H), ("Patienten", str(pat), h("AAFF88")),
               ("Personal", str(_personal(data)), TEAL_H)]
    for i, (lbl, val, fc2) in enumerate(kacheln):
        kx = 22; ky = 172 + i*62
        rect(cv, kx, ky, W*0.38, 54, h("00706A"))
        t(cv, kx+10, ky+18, lbl, "Helvetica", 8, TEAL_H)
        t(cv, kx+10, ky+46, val, "Helvetica-Bold", 26, fc2)

    # Bulmor (links)
    b_yt = 370
    t(cv, 22, b_yt, "BULMOR", "Helvetica-Bold", 11, WEISS)
    fc2, lbl2 = _bul_col(bul)
    rect(cv, 22, b_yt+6, W*0.38, 28, h("00706A"))
    t(cv, 22+W*0.19, b_yt+24, f"{bul}/5  —  {lbl2}", "Helvetica-Bold", 13, fc2, "center")
    # 5 Kreise
    for i in range(5):
        aktiv = (i+1) <= bul
        cx_ = 30 + i*36; cy_rl = b_yt+52
        cv.setFillColor(fc2 if aktiv else h("004D47"))
        cv.circle(cx_, H-cy_rl, 12, fill=1, stroke=0)
        t(cv, cx_, cy_rl, f"B{i+1}", "Helvetica-Bold", 7, WEISS if aktiv else GRAU_M, "center")

    # Bulmor Fahrer
    fhr = _bulfhr(data)
    if fhr:
        t(cv, 22, b_yt+74, "Fahrer:", "Helvetica-Bold", 8, TEAL_H)
        for j, n in enumerate(fhr[:4]):
            t(cv, 22, b_yt+86+j*12, f"  {n}", "Helvetica", 8, TEAL_H)

    # Krank (links unten)
    kl = _kranke(data)
    if kl:
        ky_k = H - 60
        t(cv, 22, ky_k, f"Krank: {', '.join(kl)}", "Helvetica-Bold", 8.5, TERRA)

    # ── Rechte Seite (Weiß) ───────────────────────────────────────────────────
    rx = W*0.42 + 18
    rw = W*0.58 - 36

    t(cv, rx, 30, "Disposition", "Helvetica-Bold", 14, TEAL_D)
    line(cv, rx, 36, W-18, 36, TEAL, 1.5)
    cy1 = 48
    for i, (zeit, namen) in enumerate(_gruppen(data,'dispo').items()):
        if cy1 > H/2-20: break
        bg = TEAL_H if i%2==0 else GRAU_H
        rect(cv, rx, cy1, rw, 22, bg)
        t(cv, rx+6, cy1+10, zeit, "Helvetica-Bold", 8, TEAL_D)
        t(cv, rx+6, cy1+19, ", ".join(namen), "Helvetica", 7.5, DUNKEL)
        cy1 += 23

    t(cv, rx, cy1+12, "Behindertenbetreuer", "Helvetica-Bold", 14, TEAL_D)
    line(cv, rx, cy1+18, W-18, cy1+18, TEAL, 1.5)
    cy2 = cy1+30
    for i, (zeit, namen) in enumerate(_gruppen(data,'betreuer').items()):
        if cy2 > H-100: break
        bg = TEAL_H if i%2==0 else GRAU_H
        rect(cv, rx, cy2, rw, 20, bg)
        t(cv, rx+6, cy2+9, zeit, "Helvetica-Bold", 8, TEAL_M)
        an = ", ".join(namen); an = an[:40]+"…" if len(an)>41 else an
        t(cv, rx+6, cy2+17, an, "Helvetica", 7.5, DUNKEL)
        cy2 += 21

    # Verbrauchsmaterial (rechts, kompakt)
    t(cv, rx, cy2+14, "Medizinprodukte (Auswahl)", "Helvetica-Bold", 11, TEAL_D)
    line(cv, rx, cy2+20, W-18, cy2+20, TEAL, 1)
    my = cy2+32
    for i, (mat, eh, soll, mind) in enumerate(MAT[:8]):
        if my > H-12: break
        bg = TEAL_H if i%2==0 else GRAU_H
        rect(cv, rx, my, rw, 17, bg)
        t(cv, rx+6, my+12, mat, "Helvetica", 7.5, DUNKEL)
        t(cv, rx+rw-4, my+12, f"Soll: {soll}", "Helvetica-Bold", 7.5, TEAL_D, "right")
        my += 18

    # Footer
    rect(cv, 0, H-18, W, 18, TEAL_D)
    t(cv, W/2, H-6, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de",
      "Helvetica", 7.5, TEAL_H, "center")

    cv.save(); print("[OK] P9_WWFTurquoise_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P10 – UPSTACK GRADIENT (Lila→Blau Gradient, weiße Karten, moderne Kacheln)
# Inspiriert: Upstack Annual Report – gradient sections, milestone highlights
# ═══════════════════════════════════════════════════════════════════════════════
def P10_upstack_gradient(data, bul=4, einz=28, pat=5, pax=42500):
    LILA   = h("5B2D8E"); BLAU_G = h("1B4FE8"); DUNKELB= h("101840")
    MINT   = h("00C9A7"); ROSA   = h("FF6B9D"); GOLD   = h("FFD166")
    WEISS  = white; GRAU_H = h("F0F0F8"); GRAU_M = h("8888AA")
    fc, label = _bul_col(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P10_UpstackGradient_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, WEISS)

    # Gradient-Header (simuliert: mehrere Rechtecke von Lila→Blau)
    steps = 20
    for i in range(steps):
        t_val = i / steps
        r_ = int(0x5B + (0x1B - 0x5B) * t_val)
        g_ = int(0x2D + (0x4F - 0x2D) * t_val)
        b_ = int(0x8E + (0xE8 - 0x8E) * t_val)
        clr = h(f"{r_:02X}{g_:02X}{b_:02X}")
        bx_ = i * W/steps
        rect(cv, bx_, 0, W/steps+1, 110, clr)

    # Logo + Titel im Header
    logo(cv, 20, 14, w=50, ht=48)
    t(cv, W/2, 38, "STÄRKEMELDUNG UND EINSÄTZE", "Helvetica-Bold", 20, WEISS, "center")
    t(cv, W/2, 56, f"DRK Köln  ·  {STATION}", "Helvetica", 8.5, h("CCDDFF"), "center")
    t(cv, W/2, 72, f"{DATUM}  ·  {UHRZEIT}", "Helvetica-Bold", 10, GOLD, "center")
    # Weiche untere Kante (abgerundet simuliert)
    rect(cv, 0, 108, W, 8, h("F8F8FF"))

    # ── Kacheln mit weichen Schatten (simuliert mit Offset-Rechteck) ───────────
    kz = [("Einsätze", str(einz), LILA), ("Patienten", str(pat), MINT),
          ("PAX", f"{pax:,}".replace(",","."), BLAU_G), ("Personal", str(_personal(data)), ROSA)]
    bw = 115; gap = 10
    for i, (lbl, val, fc2) in enumerate(kz):
        bx = 25 + i*(bw+gap)
        # Schatten
        rect(cv, bx+3, 121, bw, 70, h("E0E0EE"))
        # Karte
        rect(cv, bx, 118, bw, 70, WEISS)
        # Farbiger oberer Rand
        rect(cv, bx, 118, bw, 5, fc2)
        t(cv, bx+bw/2, 136, lbl, "Helvetica", 8, GRAU_M, "center")
        t(cv, bx+bw/2, 172, val, "Helvetica-Bold", 30, fc2, "center")

    # ── Donut-Chart: Bulmor ────────────────────────────────────────────────────
    btop = 200
    t(cv, 28, btop, "Bulmor – Einsatzquote", "Helvetica-Bold", 12, DUNKELB)
    line(cv, 28, btop+6, W-28, btop+6, GRAU_H, 0.8)
    # Donut
    fc2, lbl2 = _bul_col(bul)
    donut_chart(cv, 80, H-(btop+65), 42, 25, bul, 5, fc2, GRAU_H)
    t(cv, 80, btop+55, f"{bul}/5", "Helvetica-Bold", 12, fc2, "center")
    t(cv, 80, btop+70, lbl2, "Helvetica-Bold", 8, fc2, "center")

    # 5 Bulmor-Chips rechts
    for i in range(5):
        aktiv = (i+1) <= bul
        cx_ = 160 + i*78
        rect(cv, cx_, btop+12, 68, 55, fc2 if aktiv else GRAU_H)
        t(cv, cx_+34, btop+35, f"B{i+1}", "Helvetica-Bold", 14,
          WEISS if aktiv else GRAU_M, "center")
        t(cv, cx_+34, btop+56, "AKTIV" if aktiv else "STANDBY", "Helvetica-Bold", 7,
          WEISS if aktiv else GRAU_M, "center")

    fhr = _bulfhr(data)
    if fhr:
        t(cv, 28, btop+82, f"Fahrer: {', '.join(fhr)}", "Helvetica", 8.5, GRAU_M)

    # ── Milestone-Band (Highlight Kacheln – Upstack-Stil) ─────────────────────
    mtp = btop+98
    rect(cv, 0, mtp, W, 38, GRAU_H)
    line(cv, 0, mtp, W, mtp, h("DDDDEE"), 1)
    marks = [("👤 Personal", str(_personal(data))), ("⚕ Krank", str(len(_kranke(data)))),
             ("🚑 Einsätze", str(einz)), ("🏥 Patienten", str(pat)), ("✈ PAX", str(pax))]
    mw = W / len(marks)
    for i, (lbl, val) in enumerate(marks):
        mx = i*mw
        t(cv, mx+mw/2, mtp+14, lbl, "Helvetica-Bold", 8, GRAU_M, "center")
        t(cv, mx+mw/2, mtp+28, val, "Helvetica-Bold", 11, DUNKELB, "center")
        if i > 0: vline(cv, mx, mtp+4, mtp+36, h("CCCCDD"))

    # ── Schichten ─────────────────────────────────────────────────────────────
    stop = mtp+50
    # Dispo (Gradient-Überschrift)
    rect(cv, 28, stop, (W-66)/2, 18, LILA)
    t(cv, 28+(W-66)/4, stop+12, "DISPOSITION", "Helvetica-Bold", 9, WEISS, "center")
    cy1 = stop+20
    for i, (zeit, namen) in enumerate(_gruppen(data,'dispo').items()):
        if cy1 > H-60: break
        bg = h("F4F0FF") if i%2==0 else WEISS
        rect(cv, 28, cy1, (W-66)/2, 20, bg)
        t(cv, 34, cy1+9, zeit, "Helvetica-Bold", 8, LILA)
        t(cv, 34, cy1+17, ", ".join(namen), "Helvetica", 7.5, DUNKELB)
        cy1 += 21

    # Betreuer
    rx2 = 28 + (W-66)/2 + 10
    rect(cv, rx2, stop, (W-66)/2, 18, BLAU_G)
    t(cv, rx2+(W-66)/4, stop+12, "BEHINDERTENBETREUER", "Helvetica-Bold", 9, WEISS, "center")
    cy2 = stop+20
    for i, (zeit, namen) in enumerate(_gruppen(data,'betreuer').items()):
        if cy2 > H-60: break
        bg = h("EEF4FF") if i%2==0 else WEISS
        rect(cv, rx2, cy2, (W-66)/2, 20, bg)
        t(cv, rx2+6, cy2+9, zeit, "Helvetica-Bold", 8, BLAU_G)
        an = ", ".join(namen)[:40]
        t(cv, rx2+6, cy2+17, an, "Helvetica", 7.5, DUNKELB)
        cy2 += 21

    kl = _kranke(data)
    if kl:
        ky = max(cy1,cy2)+6
        if ky < H-24:
            rect(cv, 28, ky, W-56, 16, h("FFE8F0"))
            t(cv, 34, ky+11, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 8.5, ROSA)

    # Footer Gradient
    steps2 = 10
    for i in range(steps2):
        t2 = i/steps2
        r2 = int(0x5B + (0x1B-0x5B)*t2); g2 = int(0x2D + (0x4F-0x2D)*t2); b2 = int(0x8E + (0xE8-0x8E)*t2)
        rect(cv, i*W/steps2, H-18, W/steps2+1, 18, h(f"{r2:02X}{g2:02X}{b2:02X}"))
    t(cv, W/2, H-6, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de",
      "Helvetica", 7.5, WEISS, "center")

    cv.save(); print("[OK] P10_UpstackGradient_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P11 – CORPORATE YELLOW-BLACK (Gelb auf Schwarz-Weiß, Donut-Charts, elegant)
# Inspiriert: Business Annual Report Template – yellow on black/white, data widgets
# ═══════════════════════════════════════════════════════════════════════════════
def P11_corporate_yb(data, bul=2, einz=28, pat=5, pax=42500):
    SCHWARZ = h("111111"); GELB   = h("F5C800"); WEISS  = white
    GRAU_L  = h("F7F7F7"); GRAU_M = h("999999"); GRAU_D = h("4A4A4A")
    ROT_C   = h("E8002A"); GRN_C  = h("28A745")
    fc, label = _bul_col(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P11_CorporateYB_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, WEISS)

    # Linkes schwarzes Band
    rect(cv, 0, 0, 90, H, SCHWARZ)
    # Gelber Akzent-Balken auf schwarzem Band
    rect(cv, 0, 200, 90, 8, GELB)
    # Vertikaler Text links (simuliert mit normaler Schrift)
    for i, ch in enumerate("TAGESBERICHT"):
        t(cv, 45, 225 + i*15, ch, "Helvetica-Bold", 8.5, GELB, "center")

    # Header (weiß)
    logo(cv, 100, 12, w=55, ht=52)
    t(cv, 165, 28, "STÄRKEMELDUNG UND EINSÄTZE", "Helvetica-Bold", 18, SCHWARZ)
    t(cv, 165, 44, f"DRK Köln  ·  {STATION}", "Helvetica", 8.5, GRAU_M)
    t(cv, W-28, 28, DATUM, "Helvetica-Bold", 12, SCHWARZ, "right")
    t(cv, W-28, 44, UHRZEIT, "Helvetica", 9, GRAU_M, "right")
    line(cv, 94, 60, W-18, 60, SCHWARZ, 2)
    line(cv, 94, 63, W-18, 63, GELB, 3)

    # ── Donut-Charts als Kennzahlen ────────────────────────────────────────────
    donuts = [(str(einz), "Einsätze", einz, 50, GELB), (str(pat), "Patienten", pat, 20, GRN_C),
              (f"{bul}/5", "Bulmor", bul, 5, fc), (str(_personal(data)), "Personal", _personal(data), 50, GRAU_D)]
    dw = (W-94-18-30) / 4
    for i, (val, lbl, v, maxv, fc2) in enumerate(donuts):
        dcx = 94 + 15 + i*dw + dw/2; dcy = H - 110
        donut_chart(cv, dcx, dcy, 32, 20, v, maxv, fc2, GRAU_L)
        t(cv, dcx, 97, val, "Helvetica-Bold", 14, fc2, "center")
        t(cv, dcx, 112, lbl, "Helvetica", 7.5, GRAU_M, "center")

    # ── Bulmor-Detailkacheln ──────────────────────────────────────────────────
    btop = 128
    line(cv, 94, btop, W-18, btop, h("EEEEEE"), 0.8)
    t(cv, 100, btop+13, "Bulmor – Fahrzeugstatus", "Helvetica-Bold", 10, SCHWARZ)
    bw2 = 82; gap2 = 8
    fhr = _bulfhr(data)
    for i in range(5):
        bx2 = 100 + i*(bw2+gap2)
        aktiv = (i+1) <= bul
        rect(cv, bx2, btop+18, bw2, 48, fc if aktiv else GRAU_L)
        t(cv, bx2+bw2/2, btop+38, f"B{i+1}", "Helvetica-Bold", 14,
          WEISS if aktiv else GRAU_D, "center")
        t(cv, bx2+bw2/2, btop+54, "EINSATZ" if aktiv else "BEREIT", "Helvetica-Bold", 7.5,
          WEISS if aktiv else GRAU_M, "center")
    t(cv, 100, btop+76, f"{bul}/5 Bulmor im Einsatz  —  {label}", "Helvetica-Bold", 9.5, fc)
    if fhr:
        t(cv, 100, btop+90, f"Fahrer: {', '.join(fhr)}", "Helvetica", 8, GRAU_M)

    # ── Balken (horizontal, Einsatz/Patienten) ─────────────────────────────────
    bap = btop+104
    line(cv, 94, bap, W-18, bap, h("EEEEEE"), 0.8)
    t(cv, 100, bap+14, "Einsatz-Auslastung", "Helvetica-Bold", 9, SCHWARZ)
    bar_h(cv, 100, bap+18, W-130, 15, einz, 50, GELB, h("F0F0F0"))
    t(cv, W-24, bap+28, f"{einz}/50", "Helvetica-Bold", 8.5, SCHWARZ, "right")
    t(cv, 100, bap+42, "Patienten auf Station", "Helvetica-Bold", 9, SCHWARZ)
    bar_h(cv, 100, bap+46, W-130, 15, pat, 20, GRN_C, h("F0F0F0"))
    t(cv, W-24, bap+56, f"{pat}/20", "Helvetica-Bold", 8.5, SCHWARZ, "right")

    # ── Schichten auf weißem Hintergrund ──────────────────────────────────────
    stop = bap+70
    line(cv, 94, stop, W-18, stop, SCHWARZ, 1.5)
    line(cv, 94, stop+2, 94+60, stop+2, GELB, 3)

    # Dispo-Tabelle
    t(cv, 100, stop+16, "Disposition", "Helvetica-Bold", 10, SCHWARZ)
    cy1 = stop+28
    for i, (zeit, namen) in enumerate(_gruppen(data,'dispo').items()):
        if cy1 > H/2+120: break
        bg = GRAU_L if i%2==0 else WEISS
        rect(cv, 100, cy1, (W-130)/2, 20, bg)
        t(cv, 106, cy1+9, zeit, "Helvetica-Bold", 8, h("1A1A1A"))
        t(cv, 106, cy1+17, ", ".join(namen), "Helvetica", 7.5, GRAU_D)
        cy1 += 21

    # Betreuer-Tabelle
    rx3 = 100 + (W-130)/2 + 8
    t(cv, rx3, stop+16, "Behindertenbetreuer", "Helvetica-Bold", 10, SCHWARZ)
    cy2 = stop+28
    for i, (zeit, namen) in enumerate(_gruppen(data,'betreuer').items()):
        if cy2 > H-90: break
        bg = GRAU_L if i%2==0 else WEISS
        rect(cv, rx3, cy2, (W-130)/2, 20, bg)
        t(cv, rx3+6, cy2+9, zeit, "Helvetica-Bold", 8, SCHWARZ)
        an = ", ".join(namen)[:40]
        t(cv, rx3+6, cy2+17, an, "Helvetica", 7.5, GRAU_D)
        cy2 += 21

    # Verbrauchsmaterial
    my_yt = max(cy1, cy2) + 10
    if my_yt < H - 80:
        line(cv, 94, my_yt, W-18, my_yt, SCHWARZ, 1)
        line(cv, 94, my_yt+2, 94+90, my_yt+2, GELB, 3)
        t(cv, 100, my_yt+16, "Verbrauchsmaterial (Auswahl)", "Helvetica-Bold", 9, SCHWARZ)
        my = my_yt+28
        cols = 2; cw2 = (W-130)//cols
        for i, (mat, eh, soll, mind) in enumerate(MAT[:12]):
            if my > H-24: break
            col = i % cols; row = i // cols
            bx_ = 100 + col*cw2; by_ = my + row*16
            if by_ > H-24: break
            if col == 0 and row > 0 and i > 0 and i % cols == 0: my += 16
        # Lineardarstellung
        my2 = my_yt+28
        for i, (mat, eh, soll, mind) in enumerate(MAT[:10]):
            if my2 > H-22: break
            bg = GRAU_L if i%2==0 else WEISS
            rect(cv, 100, my2, W-130, 16, bg)
            t(cv, 106, my2+11, mat, "Helvetica", 7.5, GRAU_D)
            t(cv, W-25, my2+11, f"Soll: {soll} {eh}", "Helvetica-Bold", 7.5, GELB, "right")
            my2 += 17

    kl = _kranke(data)
    if kl:
        t(cv, 100, H-30, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 8.5, ROT_C)

    # Footer
    rect(cv, 0, H-18, W, 18, SCHWARZ)
    t(cv, W/2, H-6, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de",
      "Helvetica", 7.5, GELB, "center")

    cv.save(); print("[OK] P11_CorporateYB_25032026.pdf")


# ═══════════════════════════════════════════════════════════════════════════════
# P12 – GEOMETRIC (Diagonale Farbflächen, Dreiecke, Coral/Dunkelblau/Weiß)
# Inspiriert: Geometric Annual Report Template – eye-catching geometric shapes
# ═══════════════════════════════════════════════════════════════════════════════
def P12_geometric(data, bul=5, einz=28, pat=5, pax=42500):
    CORAL  = h("E8603A"); DUNKELB= h("1A2C5B"); SENF   = h("F5AE29")
    WEISS  = white; GRAU_H = h("F7F7FA"); GRAU_M = h("888888"); HELL_B = h("EEF3FF")
    fc, label = _bul_col(bul)

    cv = canvas.Canvas(os.path.join(ZIEL, "P12_Geometric_25032026.pdf"), pagesize=A4)
    rect(cv, 0, 0, W, H, WEISS)

    # ── Geometrische Header-Komposition ───────────────────────────────────────
    # Großes Dreieck oben rechts (Coral)
    p = cv.beginPath()
    p.moveTo(W, H); p.lineTo(W, H-200); p.lineTo(W-220, H); p.close()
    cv.setFillColor(CORAL); cv.drawPath(p, fill=1, stroke=0)
    # Kleineres Dunkelblau-Dreieck
    p2 = cv.beginPath()
    p2.moveTo(W, H); p2.lineTo(W, H-120); p2.lineTo(W-130, H); p2.close()
    cv.setFillColor(DUNKELB); cv.drawPath(p2, fill=1, stroke=0)
    # Gelbes Dreieck links oben
    p3 = cv.beginPath()
    p3.moveTo(0, H); p3.lineTo(0, H-80); p3.lineTo(80, H); p3.close()
    cv.setFillColor(SENF); cv.drawPath(p3, fill=1, stroke=0)
    # Dunkelblauer oberer Streifen
    rect(cv, 0, 0, W, 85, DUNKELB)
    # Coral-Dreieck im Header
    p4 = cv.beginPath()
    p4.moveTo(0, H-85); p4.lineTo(100, H-85); p4.lineTo(0, H-130); p4.close()
    cv.setFillColor(CORAL); cv.drawPath(p4, fill=1, stroke=0)

    logo(cv, 20, 14, w=55, ht=55)
    t(cv, W/2, 30, "STÄRKEMELDUNG", "Helvetica-Bold", 22, WEISS, "center")
    t(cv, W/2, 50, "UND EINSÄTZE", "Helvetica-Bold", 22, SENF, "center")
    t(cv, W/2, 66, f"DRK Köln  ·  {DATUM}  ·  {UHRZEIT}", "Helvetica", 8.5, h("AABBDD"), "center")

    # ── Kennzahlen-Kacheln mit Dreieck-Akzent ─────────────────────────────────
    kz = [("Einsätze", str(einz), CORAL), ("Patienten", str(pat), h("1E8C50")),
          ("PAX", f"{pax:,}".replace(",","."), DUNKELB), ("Personal", str(_personal(data)), SENF)]
    bw = 115; gap = 10
    for i, (lbl, val, fc2) in enumerate(kz):
        bx = 25 + i*(bw+gap)
        rect(cv, bx, 95, bw, 70, GRAU_H)
        # Diagonales Dreieck oben rechts
        p5 = cv.beginPath()
        p5.moveTo(bx+bw-30, H-95); p5.lineTo(bx+bw, H-95); p5.lineTo(bx+bw, H-95-30)
        p5.close(); cv.setFillColor(fc2); cv.drawPath(p5, fill=1, stroke=0)
        t(cv, bx+12, 115, lbl, "Helvetica", 8, GRAU_M)
        t(cv, bx+12, 153, val, "Helvetica-Bold", 28, fc2)

    # ── Bulmor – Polygon-Stil ─────────────────────────────────────────────────
    btop = 178
    t(cv, 28, btop, "Bulmor – Fahrzeugstatus", "Helvetica-Bold", 12, DUNKELB)
    line(cv, 28, btop+6, W-28, btop+6, CORAL, 2)
    # Hexagon-ähnliche Kacheln (simuliert als Rechtecke mit Schräge)
    bw2 = 88; gap2 = 11; fhr = _bulfhr(data)
    for i in range(5):
        bx2 = 28 + i*(bw2+gap2)
        aktiv = (i+1) <= bul
        # Haupt-Rect
        fill2 = fc if aktiv else GRAU_H
        rect(cv, bx2, btop+10, bw2, 55, fill2)
        # Schräge untere Ecke
        p6 = cv.beginPath()
        p6.moveTo(bx2+bw2-20, H-(btop+10+55))
        p6.lineTo(bx2+bw2, H-(btop+10+55))
        p6.lineTo(bx2+bw2, H-(btop+10+55)+20)
        p6.close(); cv.setFillColor(DUNKELB if aktiv else h("DDDDDD"))
        cv.drawPath(p6, fill=1, stroke=0)
        t(cv, bx2+bw2/2, btop+36, f"B{i+1}", "Helvetica-Bold", 16,
          WEISS if aktiv else GRAU_M, "center")
        t(cv, bx2+bw2/2, btop+53, "EINSATZ" if aktiv else "BEREIT", "Helvetica-Bold", 7.5,
          WEISS if aktiv else GRAU_M, "center")
    t(cv, 28, btop+76, f"{bul}/5 Bulmor im Einsatz  —  {label}", "Helvetica-Bold", 10, fc)
    if fhr:
        t(cv, 28, btop+89, f"Fahrer: {', '.join(fhr)}", "Helvetica", 8.5, GRAU_M)

    # ── Schichten ─────────────────────────────────────────────────────────────
    stop = btop+104
    rect(cv, 28, stop, (W-56)/2-5, 20, DUNKELB)
    rect(cv, W/2+4, stop, (W-56)/2-5, 20, CORAL)
    t(cv, 28+(W-56)/4-5, stop+14, "DISPOSITION", "Helvetica-Bold", 9, WEISS, "center")
    t(cv, W/2+4+(W-56)/4-5, stop+14, "BEHINDERTENBETREUER", "Helvetica-Bold", 9, WEISS, "center")

    cy1 = stop+22; cy2 = stop+22
    for i, (zeit, namen) in enumerate(_gruppen(data,'dispo').items()):
        if cy1 > H-110: break
        bg = HELL_B if i%2==0 else WEISS
        rect(cv, 28, cy1, (W-56)/2-5, 20, bg)
        t(cv, 34, cy1+9, zeit, "Helvetica-Bold", 8, DUNKELB)
        t(cv, 34, cy1+17, ", ".join(namen), "Helvetica", 7.5, h("333333"))
        cy1 += 21

    for i, (zeit, namen) in enumerate(_gruppen(data,'betreuer').items()):
        if cy2 > H-110: break
        bg = h("FFF4EE") if i%2==0 else WEISS
        rect(cv, W/2+4, cy2, (W-56)/2-5, 20, bg)
        t(cv, W/2+10, cy2+9, zeit, "Helvetica-Bold", 8, CORAL)
        an = ", ".join(namen); an = an[:40]+"…" if len(an)>41 else an
        t(cv, W/2+10, cy2+17, an, "Helvetica", 7.5, h("333333"))
        cy2 += 21

    # Verbrauchsmaterial
    my_yt = max(cy1, cy2)+10
    if my_yt < H-80:
        rect(cv, 28, my_yt, W-56, 18, SENF)
        t(cv, W/2, my_yt+13, "MEDIZINPRODUKTE / VERBRAUCHSMATERIAL", "Helvetica-Bold", 9, DUNKELB, "center")
        my = my_yt+22
        for i, (mat, eh, soll, mind) in enumerate(MAT[:8]):
            if my > H-30: break
            bg = GRAU_H if i%2==0 else WEISS
            rect(cv, 28, my, W-56, 16, bg)
            t(cv, 34, my+11, mat, "Helvetica", 7.5, h("333333"))
            t(cv, W-32, my+11, f"Soll: {soll}", "Helvetica-Bold", 7.5, CORAL, "right")
            my += 17

    kl = _kranke(data)
    if kl:
        t(cv, 28, H-34, f"Krankmeldung: {', '.join(kl)}", "Helvetica-Bold", 9, ROT_WARN)

    # Geometrischer Footer
    p7 = cv.beginPath()
    p7.moveTo(0, 0); p7.lineTo(180, 0); p7.lineTo(120, 18); p7.lineTo(0, 18)
    p7.close(); cv.setFillColor(CORAL); cv.drawPath(p7, fill=1, stroke=0)
    p8 = cv.beginPath()
    p8.moveTo(180, 0); p8.lineTo(W, 0); p8.lineTo(W, 18); p8.lineTo(120, 18)
    p8.close(); cv.setFillColor(DUNKELB); cv.drawPath(p8, fill=1, stroke=0)
    t(cv, W/2, 11, "DRK Köln  ·  +49 2203 40-2323  ·  flughafen@drk-koeln.de",
      "Helvetica", 7, WEISS, "center")

    cv.save(); print("[OK] P12_Geometric_25032026.pdf")


# ── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("Lade Dienstplan ...")
    data = lade()
    print(f"  Betreuer: {len(data.get('betreuer',[]))}  |  Dispo: {len(data.get('dispo',[]))}  |  Krank: {len(data.get('kranke',[]))}")
    print(f"\nZielordner: {ZIEL}\n")
    print("Erstelle 6 neue PDF-Designs (Batch 2) ...")
    P7_warby_magazine(data,     bul=5, einz=28, pat=5, pax=42500)
    P8_adidas_bold(data,        bul=3, einz=28, pat=5, pax=42500)
    P9_wwf_turquoise(data,      bul=5, einz=28, pat=5, pax=42500)
    P10_upstack_gradient(data,  bul=4, einz=28, pat=5, pax=42500)
    P11_corporate_yb(data,      bul=2, einz=28, pat=5, pax=42500)
    P12_geometric(data,         bul=5, einz=28, pat=5, pax=42500)
    print(f"\n✓ Alle 6 neuen PDFs gespeichert in:\n  {ZIEL}")
