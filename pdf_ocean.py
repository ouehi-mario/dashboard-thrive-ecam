"""Mise en page des rapports PDF — charte OCEAN (Source & Deliver).

Les chiffres sont calculés à partir des données reçues ; la mise en page ne modifie aucune valeur.
"""
from __future__ import annotations

import io
import os
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.shapes import Drawing, Line, Rect, String
from reportlab.graphics.widgets.markers import makeMarker
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "assets", "ocean_logo_pdf.png")  # copie allégée du logo (mêmes proportions et couleurs)
LOGO_RATIO = 572 / 1434
ADRESSE = "Magasin CFS1, Boulevard de Vridi – 15 BP 360 Abidjan 15 – RCI – cocoasource.ch"

# ---------------------------------------------------------------- charte
BORD = colors.HexColor("#A4204A")       # bordeaux / framboise OCEAN
BORD_10 = colors.HexColor("#F6E9ED")    # teinte 10 % (lignes alternées)
BORD_18 = colors.HexColor("#EFD7DF")    # ligne TOTAL
BORD_40 = colors.HexColor("#D9A3B5")    # barres secondaires
GRIS = colors.HexColor("#6E6768")       # gris du logo
GRIS_L = colors.HexColor("#E6E1E2")     # filets
INK = colors.HexColor("#2B2B2B")        # texte (jamais de noir pur)
VERT = colors.HexColor("#4E8A5B")       # plantations saines
FOND_BOX = colors.HexColor("#FBF6F8")

F, FM, FSB, FB = "Helvetica", "Helvetica", "Helvetica-Bold", "Helvetica-Bold"
try:  # Poppins (sans-serif géométrique proche du logo), sinon Helvetica
    for name, file in (("Poppins", "Regular"), ("Poppins-Medium", "Medium"), ("Poppins-SemiBold", "SemiBold"),
                       ("Poppins-Bold", "Bold")):
        if name not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(name, os.path.join(HERE, "assets", "fonts", f"Poppins-{file}.ttf")))
    F, FM, FSB, FB = "Poppins", "Poppins-Medium", "Poppins-SemiBold", "Poppins-Bold"
except Exception:  # police absente : Helvetica
    pass

S = {
    "title": ParagraphStyle("title", fontName=FB, fontSize=21, leading=26, textColor=BORD),
    "sub": ParagraphStyle("sub", fontName=FM, fontSize=11.5, leading=15, textColor=INK),
    "meta": ParagraphStyle("meta", fontName=F, fontSize=9, leading=12.5, textColor=INK),
    "small": ParagraphStyle("small", fontName=F, fontSize=7.5, leading=10, textColor=GRIS),
    "h": ParagraphStyle("h", fontName=FB, fontSize=13, leading=17, textColor=BORD, spaceBefore=10, spaceAfter=2),
    "body": ParagraphStyle("body", fontName=F, fontSize=9.5, leading=13, textColor=INK),
    "td": ParagraphStyle("td", fontName=F, fontSize=8.5, leading=11, textColor=INK),
    "tdr": ParagraphStyle("tdr", fontName=F, fontSize=8.5, leading=11, textColor=INK, alignment=TA_RIGHT),
    "tdb": ParagraphStyle("tdb", fontName=FB, fontSize=8.5, leading=11, textColor=INK),
    "tdbr": ParagraphStyle("tdbr", fontName=FB, fontSize=8.5, leading=11, textColor=INK, alignment=TA_RIGHT),
    "th": ParagraphStyle("th", fontName=FSB, fontSize=8.5, leading=10.5, textColor=colors.white),
    "thr": ParagraphStyle("thr", fontName=FSB, fontSize=8.5, leading=10.5, textColor=colors.white, alignment=TA_RIGHT),
    "kpi": ParagraphStyle("kpi", fontName=FB, fontSize=20, leading=23, textColor=BORD, alignment=TA_CENTER),
    "kpi_l": ParagraphStyle("kpi_l", fontName=FSB, fontSize=8.5, leading=10.5, textColor=INK, alignment=TA_CENTER),
    "kpi_c": ParagraphStyle("kpi_c", fontName=F, fontSize=7.5, leading=9.5, textColor=GRIS, alignment=TA_CENTER),
}
PAGE = landscape(A4)
MARGE = 1.6 * cm
W = PAGE[0] - 2 * MARGE


# ---------------------------------------------------------------- formats français
def nb(x) -> str:
    """Nombre entier, espace insécable comme séparateur de milliers."""
    return f"{int(round(float(x))):,}".replace(",", " ")


def pct(a, b, dec: int | None = None) -> str:
    """Pourcentage à la française ; une décimale en dessous de 10 % (ex. 3,7 %)."""
    if not b:
        return "–"
    v = a / b * 100
    if dec is None:
        dec = 1 if 0 < v < 10 and round(v, 1) != round(v) else 0
    return f"{v:.{dec}f}".replace(".", ",") + " %"


def P(x, st="td") -> Paragraph:
    if x is None or x is pd.NA or (isinstance(x, float) and np.isnan(x)):
        x = ""
    return Paragraph(escape(str(x)), S[st])


# ---------------------------------------------------------------- briques
def heading(txt: str):
    """Titre de section + filet bordeaux ; toujours gardés avec le bloc qui suit (jamais seuls en bas de page)."""
    p = Paragraph(escape(txt), S["h"])
    p.keepWithNext = 1
    rule = Table([[""]], colWidths=[W], rowHeights=[2], style=[("LINEABOVE", (0, 0), (-1, 0), 1, BORD)])
    rule.keepWithNext = 1
    return [p, rule]


def data_table(head: list[str], rows: list[list], widths: list[float], right: set = frozenset(), total: bool = False):
    """Tableau : en-tête bordeaux, lignes alternées 10 %, chiffres à droite, ligne TOTAL en gras."""
    hdr = [P(h, "thr" if i in right else "th") for i, h in enumerate(head)]
    body = []
    for k, r in enumerate(rows):
        last = total and k == len(rows) - 1
        body.append([c if not isinstance(c, (str, int, float, np.integer, np.floating)) else
                     P(c, ("tdbr" if i in right else "tdb") if last else ("tdr" if i in right else "td"))
                     for i, c in enumerate(r)])
    t = Table([hdr] + body, colWidths=widths, repeatRows=1)
    st = [("BACKGROUND", (0, 0), (-1, 0), BORD), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, BORD_10]),
          ("LINEBELOW", (0, 1), (-1, -1), .3, GRIS_L),
          ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
          ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]
    if total:
        st += [("BACKGROUND", (0, -1), (-1, -1), BORD_18), ("LINEABOVE", (0, -1), (-1, -1), .8, BORD)]
    t.setStyle(TableStyle(st))
    return t


def coach_cell(cid, nom) -> Paragraph:
    nom = "" if nom is None or (isinstance(nom, float) and np.isnan(nom)) else str(nom)
    cid = "" if cid is None or (isinstance(cid, float) and np.isnan(cid)) else str(cid)
    return Paragraph(f"{escape(nom)}&nbsp;&nbsp;<font size='7' color='#6E6768'>{escape(cid)}</font>", S["td"])


def box(text_html: str):
    """Encadré « Pour aller plus loin » : filet gauche bordeaux."""
    p = Paragraph(f"<font name='{FSB}' color='#A4204A'>Pour aller plus loin</font><br/>{text_html}", S["small"])
    t = Table([[p]], colWidths=[W])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), FOND_BOX), ("LINEBEFORE", (0, 0), (0, -1), 3, BORD),
                           ("LEFTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 6),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t


def progress(p: float, width: float) -> Drawing:
    d = Drawing(width, 7)
    d.add(Rect(0, 0, width, 6, rx=3, ry=3, fillColor=GRIS_L, strokeColor=None))
    d.add(Rect(0, 0, max(width * min(max(p, 0), 1), 4), 6, rx=3, ry=3, fillColor=BORD, strokeColor=None))
    return d


def kpi_cards(cards: list[dict]):
    """Cartes d'indicateurs alignées, de même taille (grand chiffre, libellé, ligne de contexte)."""
    gap = .22 * cm
    cw = (W - gap * (len(cards) - 1)) / len(cards)
    cells, widths = [], []
    for i, c in enumerate(cards):
        content = [Paragraph(c["v"], S["kpi"]), Spacer(1, 1), Paragraph(escape(c["lab"]), S["kpi_l"]),
                   Spacer(1, 2), Paragraph(escape(c["ctx"]), S["kpi_c"])]
        if c.get("bar") is not None:
            content += [Spacer(1, 4), progress(c["bar"], cw - 16)]
        cells.append(content)
        widths.append(cw)
        if i < len(cards) - 1:
            cells.append("")
            widths.append(gap)
    t = Table([cells], colWidths=widths, rowHeights=[3.0 * cm])
    st = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 8),
          ("RIGHTPADDING", (0, 0), (-1, -1), 8)]
    for i in range(0, len(cells), 2):
        st += [("BOX", (i, 0), (i, 0), .6, GRIS_L), ("LINEABOVE", (i, 0), (i, 0), 3, BORD),
               ("BACKGROUND", (i, 0), (i, 0), colors.white)]
    t.setStyle(TableStyle(st))
    return t


def side_by_side(left, right, wl: float):
    t = Table([[left, right]], colWidths=[wl, W - wl])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (0, 0), 0), ("LEFTPADDING", (1, 0), (1, 0), 18),
                           ("RIGHTPADDING", (1, 0), (1, 0), 0)]))
    return t


def chart_title(d: Drawing, txt: str):
    d.add(String(0, d.height - 11, txt, fontName=FSB, fontSize=9, fillColor=INK))


def legend(d: Drawing, items: list[tuple], x: float, y: float):
    for lab, col, dashed in items:
        if dashed:
            d.add(Line(x, y + 3, x + 14, y + 3, strokeColor=col, strokeWidth=1.2, strokeDashArray=(3, 2)))
        else:
            d.add(Rect(x, y, 9, 7, fillColor=col, strokeColor=None))
        d.add(String(x + 17, y, lab, fontName=F, fontSize=7.5, fillColor=INK))
        x += 22 + pdfmetrics.stringWidth(lab, F, 7.5) + 10


def hbar_chart(labels: list[str], values: list[float], width: float, title: str, fmt_label,
               colors_: list | None = None, vmax: float | None = None) -> Drawing:
    """Barres horizontales, plus grande valeur en haut, valeurs affichées sur les barres."""
    n = max(len(labels), 1)
    h = 34 + 17 * n
    d = Drawing(width, h)
    chart_title(d, title)
    ch = HorizontalBarChart()
    lab_w = min(max([pdfmetrics.stringWidth(str(x), F, 7.5) for x in labels] + [20]) + 14, width * .45)
    ch.x, ch.y, ch.width, ch.height = lab_w, 6, width - lab_w - 46, h - 30
    ch.data = [list(reversed(values))]
    ch.categoryAxis.categoryNames = list(reversed([str(x) for x in labels]))
    ch.categoryAxis.labels.fontName, ch.categoryAxis.labels.fontSize = F, 7.5
    ch.categoryAxis.labels.fillColor = INK
    ch.categoryAxis.strokeColor = GRIS_L
    ch.valueAxis.visible = False
    ch.valueAxis.valueMin = 0
    ch.valueAxis.valueMax = vmax or (max(values) * 1.05 if values and max(values) > 0 else 1)
    ch.bars[0].fillColor = BORD
    ch.bars.strokeColor = None
    ch.barSpacing, ch.groupSpacing = 0, 5
    if colors_:
        for i, c in enumerate(reversed(colors_)):
            ch.bars[(0, i)].fillColor = c
    vals = list(reversed(values))
    ch.barLabelFormat = lambda v: fmt_label(v)
    ch.barLabels.fontName, ch.barLabels.fontSize, ch.barLabels.fillColor = FSB, 7.5, INK
    ch.barLabels.boxAnchor, ch.barLabels.dx = "w", 4
    ch.barLabels.nudge = 0
    _ = vals
    d.add(ch)
    return d


def stacked_chart(labels: list[str], a: list[int], b: list[int], width: float, title: str,
                  names: tuple[str, str], cols: tuple, mark: set = frozenset()) -> Drawing:
    """Barres horizontales empilées (ex. avec signes / sains par section)."""
    n = max(len(labels), 1)
    h = 50 + 17 * n
    d = Drawing(width, h)
    chart_title(d, title)
    legend(d, [(names[0], cols[0], False), (names[1], cols[1], False)], 0, h - 27)
    cats = [f"{x} •" if x in mark else str(x) for x in labels]
    ch = HorizontalBarChart()
    lab_w = min(max([pdfmetrics.stringWidth(c, F, 7.5) for c in cats] + [20]) + 14, width * .45)
    ch.x, ch.y, ch.width, ch.height = lab_w, 6, width - lab_w - 10, h - 46
    ch.data = [list(reversed(a)), list(reversed(b))]
    ch.categoryAxis.categoryNames = list(reversed(cats))
    ch.categoryAxis.style = "stacked"
    ch.categoryAxis.labels.fontName, ch.categoryAxis.labels.fontSize = F, 7.5
    ch.categoryAxis.labels.fillColor = INK
    ch.categoryAxis.strokeColor = GRIS_L
    ch.valueAxis.visible = False
    ch.valueAxis.valueMin = 0
    tot = [x + y for x, y in zip(a, b)]
    ch.valueAxis.valueMax = max(tot) if tot and max(tot) > 0 else 1
    ch.bars[0].fillColor, ch.bars[1].fillColor = cols
    ch.bars.strokeColor = colors.white
    ch.barSpacing, ch.groupSpacing = 0, 5
    ch.barLabelFormat = lambda v: nb(v) if v else ""
    ch.barLabels.fontName, ch.barLabels.fontSize, ch.barLabels.fillColor = FSB, 7.5, colors.white
    ch.barLabels.boxAnchor = "e"
    ch.barLabels.dx = -3
    d.add(ch)
    return d


def cumul_chart(weeks: list[str], series: list[tuple[str, list[int], object]], objectif: int, width: float,
                title: str) -> Drawing:
    """Courbes du cumul (abscisse : date de soumission), avec la ligne de l'objectif."""
    h = 150
    d = Drawing(width, h)
    chart_title(d, title)
    legend(d, [(n, c, False) for n, _, c in series] + [(f"Objectif ({nb(objectif)})", VERT, True)], 0, h - 27)
    ch = HorizontalLineChart()
    ch.x, ch.y, ch.width, ch.height = 34, 18, width - 34 - 34, h - 52  # marge à droite pour les valeurs
    ch.data = [s[1] for s in series] + [[objectif] * len(weeks)]
    step = max(1, -(-len(weeks) // 9))  # au plus ~9 dates écrites sous l'axe
    last = len(weeks) - 1
    ch.categoryAxis.categoryNames = [w if i == last or (i % step == 0 and last - i >= max(step // 2, 1)) else ""
                                     for i, w in enumerate(weeks)]
    ch.categoryAxis.labels.fontName, ch.categoryAxis.labels.fontSize = F, 7
    ch.categoryAxis.labels.fillColor = GRIS
    ch.categoryAxis.strokeColor = GRIS_L
    ch.valueAxis.valueMin = 0
    ch.valueAxis.valueMax = max([objectif] + [max(s[1]) for s in series if s[1]]) * 1.08
    ch.valueAxis.labels.fontName, ch.valueAxis.labels.fontSize = F, 7
    ch.valueAxis.labels.fillColor = GRIS
    ch.valueAxis.labelTextFormat = lambda v: nb(v)
    ch.valueAxis.strokeColor = GRIS_L
    ch.valueAxis.gridStrokeColor = GRIS_L
    ch.valueAxis.visibleGrid = True
    for i, (_, _, c) in enumerate(series):
        ch.lines[i].strokeColor = c
        ch.lines[i].strokeWidth = 1.8
        ch.lines[i].symbol = makeMarker("FilledCircle", size=4, fillColor=c, strokeColor=c)
    k = len(series)
    ch.lines[k].strokeColor = VERT
    ch.lines[k].strokeDashArray = (3, 2)
    ch.lines[k].strokeWidth = 1
    ch.lineLabelFormat = lambda v: ""  # valeurs posées à la main (sauf la ligne objectif)
    d.add(ch)
    # valeurs : la 1re courbe sur chaque point (si peu de dates), et la dernière valeur de chaque courbe à droite,
    # empilées sans chevauchement
    n = len(weeks)
    if n:
        vmax = ch.valueAxis.valueMax
        ypos = lambda v: ch.y + ch.height * v / vmax  # noqa: E731
        xpos = lambda j: ch.x + ch.width * (j + .5) / n  # noqa: E731
        _, main, cmain = series[0]
        if n <= 12:
            for j, v in enumerate(main[:-1]):
                d.add(String(xpos(j), ypos(v) + 5, nb(v), fontName=FSB, fontSize=7, fillColor=cmain,
                             textAnchor="middle"))
        pts = sorted((ypos(vals[-1]), vals[-1], c) for _, vals, c in series if vals)  # du plus bas au plus haut
        prev = ch.y - 10
        for y, v, c in pts:
            ty = max(y - 3, prev + 9)
            d.add(String(xpos(n - 1) + 7, ty, nb(v), fontName=FSB, fontSize=7, fillColor=c))
            prev = ty
    return d


def daily_cumul(dates: dict[str, pd.Series]) -> pd.DataFrame:
    """Cumul par date de soumission : une ligne par date où au moins une fiche a été soumise."""
    cols = {}
    for name, d in dates.items():
        d = pd.to_datetime(d, errors="coerce").dropna()
        cols[name] = d.dt.normalize().value_counts()
    t = pd.DataFrame(cols).fillna(0).astype(int).sort_index()
    return t.cumsum()


def weekly(dates: dict[str, pd.Series], n_weeks: int = 8) -> pd.DataFrame:
    cols = {}
    for name, d in dates.items():
        d = pd.to_datetime(d, errors="coerce").dropna()
        cols[name] = d.dt.to_period("W-SUN").dt.start_time.value_counts()
    t = pd.DataFrame(cols).fillna(0).astype(int).sort_index()
    if t.empty:
        return t
    t = t.reindex(pd.date_range(t.index.min(), t.index.max(), freq="W-MON"), fill_value=0)
    for name in dates:
        t[f"{name} (cumul)"] = t[name].cumsum()
    return t.tail(n_weeks)


# ---------------------------------------------------------------- page : en-tête, pied de page, « page X / Y »
def build(story: list, titre: str, coop: str) -> bytes:
    def decorate(c, total):
        page = c._pageNumber
        c.saveState()
        if page > 1:  # petit logo + titre dans l'en-tête des pages suivantes
            lh = .85 * cm
            c.drawImage(LOGO, MARGE, PAGE[1] - 1.05 * cm - lh, width=lh * 1 / LOGO_RATIO, height=lh, mask="auto")
            c.setFont(F, 8)
            c.setFillColor(GRIS)
            c.drawRightString(PAGE[0] - MARGE, PAGE[1] - 1.6 * cm, titre)
            c.setStrokeColor(GRIS_L)
            c.setLineWidth(.6)
            c.line(MARGE, PAGE[1] - 2.05 * cm, PAGE[0] - MARGE, PAGE[1] - 2.05 * cm)
        c.setStrokeColor(GRIS_L)
        c.setLineWidth(.6)
        c.line(MARGE, 1.35 * cm, PAGE[0] - MARGE, 1.35 * cm)
        c.setFont(F, 7.5)
        c.setFillColor(GRIS)
        c.drawString(MARGE, .9 * cm, ADRESSE)
        c.drawRightString(PAGE[0] - MARGE, .9 * cm, f"page {page} / {total}")
        c.restoreState()

    class NumberedCanvas(rl_canvas.Canvas):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self._pages = []

        def showPage(self):
            self._pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._pages)
            for st in self._pages:
                self.__dict__.update(st)
                decorate(self, total)
                super().showPage()
            super().save()

    buf = io.BytesIO()
    SimpleDocTemplate(buf, pagesize=PAGE, leftMargin=MARGE, rightMargin=MARGE, topMargin=2.35 * cm,
                      bottomMargin=1.8 * cm, title=titre, author=f"OCEAN — Thrive · {coop}"
                      ).build(story, canvasmaker=NumberedCanvas)
    return buf.getvalue()


def cover(titre: str, perimetre: str, periode: str, source: str, filtres: str) -> list:
    """Bloc d'en-tête de la 1re page : grand logo à gauche, titre et informations à droite."""
    lw = 6.2 * cm
    logo = Image(LOGO, width=lw, height=lw * LOGO_RATIO)
    gen = pd.Timestamp.now(tz="Africa/Abidjan")
    right = [Paragraph(escape(titre), S["title"]),
             Paragraph("Projet Thrive · Suivi CSSVD" + (f" · {escape(perimetre)}" if perimetre else ""), S["sub"]),
             Spacer(1, 4),
             Paragraph(f"<font name='{FSB}'>Période</font> : {escape(periode)} &nbsp;·&nbsp; "
                       f"<font name='{FSB}'>Généré le</font> {gen:%d/%m/%Y à %H:%M}", S["meta"]),
             Paragraph(escape(f"Source : {source}" + (f" · filtres : {filtres}" if filtres and filtres != "toutes les données" else "")),
                       S["small"])]
    t = Table([[logo, right]], colWidths=[lw + .9 * cm, W - lw - .9 * cm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    return [t, Spacer(1, 12)]


def _periode(dates: pd.Series) -> str:
    d = pd.to_datetime(dates, errors="coerce").dropna()
    return f"du {d.min():%d/%m/%Y} au {d.max():%d/%m/%Y}" if len(d) else "—"


# ====================================================================== ECAM
def ecam_report(pp: pd.DataFrame, parc: pd.DataFrame, n_saisies: int, objectif: int, filtres: str, source: str,
                symptomes: dict, perimetre: str = "", par: str = "coach", periode: str | None = None) -> bytes:
    n = len(pp)
    obj = objectif or n or 1
    eng_ = pp["engagement"] == True  # noqa: E712
    ref_ = pp["engagement"] == False  # noqa: E712
    n_eng, n_ref = int(eng_.sum()), int(ref_.sum())
    n_fi, n_pr = int((pp["doc_fiche"] == 1).sum()), int((pp["doc_promesse"] == 1).sum())
    n_ss, n_sain = int((pp["cssvd"] == True).sum()), int((pp["cssvd"] == False).sum())  # noqa: E712
    n_ins = n_ss + n_sain  # producteurs inspectés (réponse Oui / Non sur les signes de CSSVD)
    sym_cols = list(symptomes)
    pl = parc[parc["prod_code"].notna()].copy()
    pl["nb_sym"] = (pl[sym_cols] == 1).sum(axis=1)
    sym = pl[(pl["nb_sym"] > 0) | (pl["cssvd"] == True)]  # noqa: E712
    titre = "Rapport d'avancement — Coopérative ECAM"
    story = cover(titre, perimetre, periode or _periode(pp["date"]), source, filtres)

    # 1) tableau de bord
    story.append(kpi_cards([
        {"v": nb(n), "lab": "Producteurs enquêtés", "ctx": f"{pct(n, obj)} de l'objectif de {nb(obj)}", "bar": n / obj},
        {"v": nb(n_eng), "lab": "Engagements du ménage", "ctx": f"{pct(n_eng, n)} des enquêtés"},
        {"v": nb(n_ref), "lab": "Refus d'engagement", "ctx": f"{pct(n_ref, n)} des enquêtés"},
        {"v": nb(n_fi), "lab": "Inspections CSSVD", "ctx": f"{pct(n_fi, n)} des enquêtés"},
        {"v": nb(n_pr), "lab": "Promesses de taille", "ctx": f"{pct(n_pr, n)} des enquêtés"},
        {"v": nb(n_ss), "lab": "Swollen shoot", "ctx": "producteurs avec signes"},
        {"v": nb(n_sain), "lab": "Sains", "ctx": f"{pct(n_sain, n_ins)} des inspectés"},
    ]))
    story += [Spacer(1, 4), Paragraph(
        f"Calculs par code producteur : les fiches d'un même producteur sont fusionnées "
        f"({nb(n_saisies)} fiches saisies pour {nb(n)} producteurs).", S["small"])]

    # 2) indicateurs par coach (ou par section dans le rapport d'un coach)
    by = "agent_id" if par == "coach" else "section"
    base = pp.assign(section=pp["section"].fillna("Sans section"))
    g = base.groupby(by, dropna=False)
    bc = pd.DataFrame({
        "nom": g["agent_nom"].last() if by == "agent_id" else g["section"].first(),
        "prod": g.size(), "eng": g["engagement"].apply(lambda s: int((s == True).sum())),  # noqa: E712
        "ref": g["engagement"].apply(lambda s: int((s == False).sum())),  # noqa: E712
        "fi": g["doc_fiche"].apply(lambda s: int((s == 1).sum())),
        "pr": g["doc_promesse"].apply(lambda s: int((s == 1).sum())),
        "ss": g["cssvd"].apply(lambda s: int((s == True).sum())),  # noqa: E712
        "sa": g["cssvd"].apply(lambda s: int((s == False).sum()))}).reset_index()  # noqa: E712
    bc = bc.sort_values("prod", ascending=False)
    first = (lambda r: coach_cell(r[by], r["nom"])) if by == "agent_id" else (lambda r: P(r["section"]))
    rows = [[first(r)] + [nb(r[k]) for k in ("prod", "eng", "ref", "fi", "pr", "ss", "sa")]
            for _, r in bc.iterrows()]
    rows.append(["TOTAL", nb(n), nb(n_eng), nb(n_ref), nb(n_fi), nb(n_pr), nb(n_ss), nb(n_sain)])
    story.append(KeepTogether(heading("Indicateurs par coach" if by == "agent_id" else "Indicateurs par section") + [data_table(["Coach" if by == "agent_id" else "Section", "Producteurs", "Engagements", "Refus",
                             "Inspections CSSVD", "Promesses de taille", "Swollen shoot", "Sains"], rows,
                            [W * x for x in (.3, .1, .1, .08, .12, .12, .1, .08)], right=set(range(1, 8)), total=True)]))

    # 3) avancement semaine par semaine
    wk = weekly({"Enquêtés": pp["date"], "Engagés": pp.loc[eng_, "date"]})
    if not wk.empty:
        dc = daily_cumul({"Enquêtés": pp["date"], "Engagés": pp.loc[eng_, "date"]})
        rows = [[f"{i:%d/%m/%Y}", nb(r["Enquêtés"]), nb(r["Engagés"]), nb(r["Enquêtés (cumul)"]),
                 nb(r["Engagés (cumul)"]), pct(r["Enquêtés (cumul)"], obj)] for i, r in wk.iterrows()]
        tbl = data_table(["Semaine du", "Enquêtés", "Engagés", "Cumul enquêtés", "Cumul engagés", "Cumul / objectif"],
                         rows, [W * .55 * x for x in (.2, .14, .14, .18, .17, .17)], right=set(range(1, 6)))
        chart = cumul_chart([f"{i:%d/%m}" for i in dc.index],
                            [("Cumul enquêtés", [int(x) for x in dc["Enquêtés"]], BORD),
                             ("Cumul engagés", [int(x) for x in dc["Engagés"]], GRIS)],
                            obj, W * .45 - 18, "Cumul des enquêtés et engagés, par date de soumission")
        story.append(KeepTogether(heading("Avancement semaine par semaine") + [
            side_by_side(tbl, chart, W * .55),
            Paragraph("Tableau : semaine du lundi au dimanche. Graphique : cumul à chaque date de soumission. "
                      "Date retenue : la dernière fiche soumise de chaque producteur.", S["small"])]))

    # 4) refus d'engagement par coach
    if by == "agent_id":
        h_rf = heading(f"Refus d'engagement par coach — {nb(n_ref)} refus sur {nb(n)} producteurs enquêtés "
                       f"({pct(n_ref, n)})")
        rc = bc[bc["ref"] > 0].copy()
        if len(rc):
            rc["taux"] = rc["ref"] / rc["prod"]
            rc = rc.sort_values(["taux", "ref"], ascending=[False, False])
            refus = base[ref_]
            secs = refus.groupby("agent_id", dropna=False)["section"].agg(
                lambda x: ", ".join(f"{k} ({v})" for k, v in x.value_counts().items()))
            rows = [[coach_cell(r["agent_id"], r["nom"]), nb(r["prod"]), nb(r["ref"]),
                     P(pct(r["ref"], r["prod"]), "tdbr" if r["taux"] >= .4 else "tdr"),
                     P(secs.get(r["agent_id"], ""))] for _, r in rc.iterrows()]
            rows.append(["TOTAL", nb(n), nb(n_ref), pct(n_ref, n), ""])
            tbl = data_table(["Coach", "Enquêtés", "Refus", "Taux de refus", "Sections concernées"], rows,
                             [W * .58 * x for x in (.36, .12, .1, .14, .28)], right={1, 2, 3}, total=True)
            chart = hbar_chart([str(x) for x in rc["agent_id"]], [float(x) * 100 for x in rc["taux"]], W * .42 - 18,
                               "Taux de refus par coach", lambda v: pct(v, 100),
                               colors_=[BORD if t >= .4 else BORD_40 for t in rc["taux"]], vmax=100)
            story.append(KeepTogether(h_rf + [side_by_side(tbl, chart, W * .58),
                                              Paragraph("En bordeaux foncé : taux de refus de 40 % et plus.",
                                                        S["small"])]))
            story.append(Spacer(1, 6))
            story.append(box(f"{len(rc)} coachs sur {len(bc)} ont au moins un refus. Liste nominative des "
                             "producteurs : application Thrive-Track, page <b>Producteurs</b> (export Excel)."))
        else:
            story += h_rf + [Paragraph("Aucun refus d'engagement.", S["body"])]
    else:
        story += heading(f"Refus d'engagement — {nb(n_ref)} refus sur {nb(n)} producteurs enquêtés ({pct(n_ref, n)})")
        story.append(Paragraph("Détail par section : colonne « Refus » du tableau « Indicateurs par section ».",
                               S["body"]))

    # 5) swollen shoot par section
    ins = base[base["cssvd"].notna()]
    h_ss = heading(f"Swollen shoot par section — {nb(n_ss)} producteurs avec signes sur {nb(n_ins)} inspectés")
    if len(ins):
        gs = ins.groupby("section")
        ts = pd.DataFrame({"insp": gs.size(), "sig": gs["cssvd"].apply(lambda x: int((x == True).sum()))})  # noqa: E712
        ts["sain"] = ts["insp"] - ts["sig"]
        ts["pl"] = sym.assign(section=sym["section"].fillna("Sans section")).groupby("section").size() \
            .reindex(ts.index).fillna(0).astype(int)
        ts["taux"] = ts["sig"] / ts["insp"]
        ts = ts.sort_values(["taux", "sig"], ascending=[False, False])
        hi = {s for s, r in ts.iterrows() if r["taux"] >= .75}
        rows = [[P(f"{s} •" if s in hi else s, "tdb" if s in hi else "td"), nb(r["insp"]), nb(r["sig"]), nb(r["sain"]),
                 pct(r["sig"], r["insp"]), nb(r["pl"])] for s, r in ts.iterrows()]
        rows.append(["TOTAL", nb(n_ins), nb(n_ss), nb(n_sain), pct(n_ss, n_ins), nb(len(sym))])
        tbl = data_table(["Section", "Inspectés", "Avec signes", "Sains", "% avec signes", "Plantations avec sympt."],
                         rows, [W * .56 * x for x in (.34, .14, .12, .09, .13, .18)], right=set(range(1, 6)), total=True)
        chart = stacked_chart(list(ts.index), [int(x) for x in ts["sig"]], [int(x) for x in ts["sain"]],
                              W * .44 - 18, "Producteurs inspectés par section", ("Avec signes", "Sains"),
                              (BORD, VERT), mark=hi)
        story.append(KeepTogether(h_ss + [side_by_side(tbl, chart, W * .56)] + (
            [Paragraph("• Section dont le taux de signes atteint 75 % ou plus.", S["small"])] if hi else [])))
    else:
        story += h_ss + [Paragraph("Aucune inspection CSSVD.", S["body"])]

    # 6) fréquence des symptômes
    if len(sym):
        cnt = sorted(((symptomes[k][0], int((pl[k] == 1).sum())) for k in sym_cols), key=lambda x: -x[1])
        tot = len(sym)
        chart = hbar_chart([a for a, _ in cnt], [b for _, b in cnt], W * .62,
                           f"Fréquence des symptômes — sur {nb(tot)} plantations avec symptômes",
                           lambda v: f"{nb(v)} · {pct(v, tot)}")
        story.append(KeepTogether(heading("Fréquence des symptômes") + [chart, Spacer(1, 6), box(
            "Liste des producteurs et plantations avec symptômes : application Thrive-Track, page "
            "<b>Symptômes</b> (export Excel ou Excel par section).")]))
    return build(story, titre + (f" · {perimetre}" if perimetre else ""), "ECAM")


# ====================================================================== UPAS
def upas_report(eng: pd.DataFrame, insp: pd.DataFrame, prom: pd.DataFrame, suivi: pd.DataFrame,
                anomalies: pd.DataFrame, objectif: int, filtres: str, n_lignes: dict, split_counts,
                perimetre: str = "") -> bytes:
    n = len(suivi)
    obj = objectif or n or 1
    n_eng, n_insp, n_prom = eng["code"].nunique(), insp["code"].nunique(), prom["code"].nunique()
    n_comp = int(suivi["complet"].sum()) if n else 0
    n_pi = int((suivi["promesse"] & suivi["inspecte"]).sum()) if n else 0
    n_bad = int((suivi["promesse"] & ~suivi["inspecte"]).sum()) if n else 0
    n_obs = int(insp.loc[insp["signes"].notna(), "code"].nunique())
    n_sig = int(insp.loc[insp["signes"] == True, "code"].nunique())  # noqa: E712
    titre = "Rapport d'avancement — Coopérative UPAS"
    dates = pd.concat([eng["date"], insp["date"], prom["date"]])
    src = " · ".join(f"{lab} : {nb(n_lignes[k])} fiches" for k, lab in
                     (("eng", "engagements"), ("insp", "inspections"), ("prom", "promesses")))
    story = cover(titre, perimetre, _periode(dates), src, filtres)
    story.append(kpi_cards([
        {"v": nb(n), "lab": "Producteurs suivis", "ctx": f"{pct(n, obj)} de l'objectif de {nb(obj)}", "bar": n / obj},
        {"v": nb(n_eng), "lab": "Engagements du ménage", "ctx": f"{pct(n_eng, n)} des suivis"},
        {"v": nb(n_insp), "lab": "Inspections CSSVD", "ctx": f"{pct(n_insp, n_eng)} des engagés"},
        {"v": nb(n_prom), "lab": "Promesses de taille",
         "ctx": f"{pct(n_pi, n_insp)} des inspectés" + (f" · {n_bad} sans inspection" if n_bad else "")},
        {"v": nb(n_comp), "lab": "Parcours complet", "ctx": f"{pct(n_comp, n)} des suivis"},
        {"v": nb(n_sig), "lab": "Swollen shoot", "ctx": "parcelles avec signes"},
        {"v": nb(n_obs - n_sig), "lab": "Sains", "ctx": f"{pct(n_obs - n_sig, n_obs)} des inspectés"},
    ]))
    story += [Spacer(1, 4), Paragraph("Une fiche par code producteur et par fichier (doublons retirés).", S["small"])]

    if n:
        g = suivi.assign(section=suivi["section"].fillna("Sans section")).groupby("section")
        sec = pd.DataFrame({"prod": g.size(), "eng": g["engage"].sum(), "insp": g["inspecte"].sum(),
                            "prom": g["promesse"].sum(), "comp": g["complet"].sum()}).astype(int)
        sec["afi"] = g.apply(lambda x: int((x["engage"] & ~x["inspecte"]).sum()), include_groups=False)
        sec["afp"] = g.apply(lambda x: int((x["inspecte"] & ~x["promesse"]).sum()), include_groups=False)
        sec = sec.sort_values("prod", ascending=False)
        rows = [[P(s), nb(r["prod"]), nb(r["eng"]), nb(r["insp"]), nb(r["prom"]), nb(r["comp"]), pct(r["comp"], r["prod"]),
                 nb(r["afi"]), nb(r["afp"])] for s, r in sec.iterrows()]
        rows.append(["TOTAL", nb(n), nb(sec["eng"].sum()), nb(sec["insp"].sum()), nb(sec["prom"].sum()), nb(n_comp),
                     pct(n_comp, n), nb(sec["afi"].sum()), nb(sec["afp"].sum())])
        story.append(KeepTogether(heading("Avancement par section") + [data_table(["Section", "Producteurs", "Engagés", "Inspectés", "Promesses", "Parcours complet",
                                 "% complet", "Inspection à faire", "Promesse à faire"], rows,
                                [W * x for x in (.2, .1, .09, .09, .1, .12, .09, .11, .1)], right=set(range(1, 9)),
                                total=True)]))
        story += [Spacer(1, 6), box("« Inspection à faire » = engagés pas encore inspectés ; « Promesse à faire » = "
                                    "inspectés sans promesse. Liste nominative : application Thrive-Track, page "
                                    "<b>Suivi producteurs</b> (export Excel par section).")]

    wk = weekly({"Engagements": eng["date"], "Inspections": insp["date"], "Promesses": prom["date"]})
    if not wk.empty:
        sub = lambda t: t["synchro"].fillna(t["date"]) if "synchro" in t else t["date"]  # noqa: E731
        dc = daily_cumul({"Engagements": sub(eng), "Inspections": sub(insp), "Promesses": sub(prom)})
        rows = [[f"{i:%d/%m/%Y}", nb(r["Engagements"]), nb(r["Inspections"]), nb(r["Promesses"]),
                 nb(r["Engagements (cumul)"]), nb(r["Inspections (cumul)"]), nb(r["Promesses (cumul)"])]
                for i, r in wk.iterrows()]
        tbl = data_table(["Semaine du", "Engag.", "Inspect.", "Promesses", "Cumul engag.", "Cumul inspect.",
                          "Cumul promesses"], rows, [W * .55 * x for x in (.16, .11, .12, .15, .15, .15, .16)],
                         right=set(range(1, 7)))
        chart = cumul_chart([f"{i:%d/%m}" for i in dc.index],
                            [("Engagements", [int(x) for x in dc["Engagements"]], BORD),
                             ("Inspections", [int(x) for x in dc["Inspections"]], GRIS),
                             ("Promesses", [int(x) for x in dc["Promesses"]], BORD_40)],
                            obj, W * .45 - 18, "Cumul par date de soumission")
        story.append(KeepTogether(heading("Avancement semaine par semaine") + [
            side_by_side(tbl, chart, W * .55),
            Paragraph("Tableau : semaine du lundi au dimanche, selon la date de chaque fiche. Graphique : cumul à "
                      "chaque date de soumission (synchronisation Kobo).", S["small"])]))

    obs = insp[insp["signes"].notna()]
    h_us = heading(f"Swollen shoot par section — {nb(n_sig)} parcelles avec signes sur {nb(n_obs)} inspectées")
    if len(obs):
        go_ = obs.assign(section=obs["section"].fillna("Sans section")).groupby("section")
        ts = pd.DataFrame({"insp": go_["code"].nunique(),
                           "sig": go_.apply(lambda x: x.loc[x["signes"] == True, "code"].nunique(),  # noqa: E712
                                            include_groups=False)})
        ts["sain"] = ts["insp"] - ts["sig"]
        ts["taux"] = ts["sig"] / ts["insp"]
        ts = ts.sort_values(["taux", "insp"], ascending=[False, False])
        hi = {s for s, r in ts.iterrows() if r["taux"] >= .75}
        rows = [[P(f"{s} •" if s in hi else s, "tdb" if s in hi else "td"), nb(r["insp"]), nb(r["sig"]),
                 nb(r["sain"]), pct(r["sig"], r["insp"])] for s, r in ts.iterrows()]
        rows.append(["TOTAL", nb(n_obs), nb(n_sig), nb(n_obs - n_sig), pct(n_sig, n_obs)])
        tbl = data_table(["Section", "Inspectées", "Avec signes", "Sains", "% avec signes"], rows,
                         [W * .5 * x for x in (.34, .16, .17, .14, .19)], right=set(range(1, 5)), total=True)
        chart = stacked_chart(list(ts.index), [int(x) for x in ts["sig"]], [int(x) for x in ts["sain"]], W * .5 - 18,
                              "Parcelles inspectées par section", ("Avec signes", "Sains"), (BORD, VERT), mark=hi)
        story.append(KeepTogether(h_us + [side_by_side(tbl, chart, W * .5)]))
    else:
        story += h_us + [Paragraph("Aucune inspection CSSVD.", S["body"])]

    etat = split_counts(insp["etat"]) if "etat" in insp else pd.Series(dtype=int)
    if len(etat):
        chart = hbar_chart(list(etat.index), [int(v) for v in etat.values], W * .62,
                           f"État des parcelles inspectées — sur {nb(n_insp)} parcelles",
                           lambda v: f"{nb(v)} · {pct(v, n_insp)}")
        story.append(KeepTogether(heading("État des parcelles inspectées") + [chart]))

    h_cq = heading(f"Contrôle qualité des fichiers — {nb(len(anomalies))} anomalies")
    if len(anomalies):
        ta = anomalies.groupby("Anomalie").agg(nb_=("Code producteur", "size"), np_=("Code producteur", "nunique")) \
            .sort_values("nb_", ascending=False)
        rows = [[P(k), nb(r["nb_"]), nb(r["np_"])] for k, r in ta.iterrows()]
        story.append(KeepTogether(h_cq + [data_table(["Type d'anomalie", "Nombre", "Producteurs concernés"], rows,
                                              [W * .45, W * .15, W * .2], right={1, 2}), Spacer(1, 6),
                                   box("Détail de chaque anomalie : application Thrive-Track, bandeau "
                                       "<b>Contrôle qualité UPAS</b> (export Excel).")]))
    else:
        story += h_cq + [Paragraph("Aucune anomalie : 1 producteur = 1 engagement → 1 inspection → 1 promesse.",
                                   S["body"])]
    return build(story, titre + (f" · {perimetre}" if perimetre else ""), "UPAS")
