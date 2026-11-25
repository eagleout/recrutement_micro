# -*- coding: utf-8 -*-
"""Construit le classeur de recrutement micro-crèches.
Usage : python build.py <sortie.xlsx> [demo]
"""
import sys, math, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule, DataBarRule, CellIsRule
from openpyxl.comments import Comment
from openpyxl.chart import BarChart3D, PieChart3D, RadarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.title import Title
from openpyxl.chart.text import RichText, Text
from openpyxl.drawing.text import (Paragraph, ParagraphProperties, CharacterProperties,
                                   RegularTextRun, Font as DFont)
import content as C

OUT = sys.argv[1]
DEMO = len(sys.argv) > 2 and sys.argv[2] == "demo"

# ---------------------------------------------------------------- palette
NAVY, TEAL, CORAL, SUN, LAV, SKY, GREEN = "1B3A5C", "2A9D8F", "E76F51", "E9C46A", "8E7CC3", "6FA8DC", "3D8B5A"
OCHRE = "B7791F"
BG, CARD, LINE = "F6F4F0", "FFFFFF", "DDD6CC"
INP_FILL, INP_FONT = "EAF2FB", "1F4E9E"
CALC_FILL = "F2F2F2"
TXT, MUTED = "2B2B2B", "7A7A7A"
POSTE_COL = {"EJE": LAV, "AP": TEAL, "CAP AEPE": CORAL}
GO_F, GO_T = "E3F4E8", "1E7B3C"
AP_F, AP_T = "FFF1DC", "B26A00"
NO_F, NO_T = "FDE4E1", "B3261E"

FIRST, N = 7, 150
LAST = FIRST + N - 1          # 156
PR = f"${FIRST}:$"            # helper

def F(size=10, bold=False, color=TXT, italic=False):
    return Font(name="Arial", size=size, bold=bold, color=color, italic=italic)

def FILL(c):
    return PatternFill("solid", start_color=c, end_color=c)

THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOTTOM = Border(bottom=Side(style="thin", color=LINE))
WRAP = Alignment(wrap_text=True, vertical="top")
WRAPC = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFTC = Alignment(horizontal="left", vertical="center", wrap_text=False, indent=1)

def rng(ws, ref):
    return [c for row in ws[ref] for c in row]

def put(ws, ref, value=None, font=None, fill=None, align=None, border=None, fmt=None, merge=True):
    """Écrit dans la cellule haut-gauche de ref (fusionne si plage) et style toute la plage."""
    first = ref.split(":")[0]
    if ":" in ref and merge:
        ws.merge_cells(ref)
    if value is not None:
        ws[first] = value
    cells = rng(ws, ref) if ":" in ref else [ws[ref]]
    for c in cells:
        if font: c.font = font
        if fill: c.fill = fill
        if align: c.alignment = align
        if border: c.border = border
        if fmt: c.number_format = fmt
    return ws[first]

def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v

def bg(ws, ncols, nrows, color=BG):
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            ws.cell(r, c).fill = FILL(color)

def nav(ws, cell="B1", extra=None):
    put(ws, cell, '=HYPERLINK("#\'Accueil\'!A1","⌂ Accueil")', font=F(9, True, TEAL), align=LEFTC)
    if extra:
        for ref, sheet, label in extra:
            put(ws, ref, f'=HYPERLINK("#\'{sheet}\'!A1","{label}")', font=F(9, True, TEAL), align=LEFTC)
    ws.row_dimensions[1].height = 20

def banner(ws, ref, title, sub_ref=None, sub=None, color=NAVY):
    put(ws, ref, title, F(18, True, "FFFFFF"), FILL(color), Alignment(vertical="center", indent=1))
    ws.row_dimensions[int(''.join(ch for ch in ref.split(':')[0] if ch.isdigit()))].height = 36
    if sub_ref:
        put(ws, sub_ref, sub, F(10, False, "FFFFFF", italic=True), FILL(color), Alignment(vertical="center", indent=1))

def section(ws, ref, text, color=NAVY):
    put(ws, ref, text, F(10, True, "FFFFFF"), FILL(color), Alignment(vertical="center", indent=1))

def header_row(ws, row, col1, labels, color=NAVY, height=34, size=9):
    for i, lab in enumerate(labels):
        c = ws.cell(row, col1 + i, lab)
        c.font = F(size, True, "FFFFFF"); c.fill = FILL(color); c.alignment = CENTER; c.border = BOX
    ws.row_dimensions[row].height = height

def inp(c, fmt=None, center=False):
    c.font = F(10, False, INP_FONT); c.fill = FILL(INP_FILL); c.border = BOX
    c.alignment = CENTER if center else Alignment(vertical="center", wrap_text=False)
    if fmt: c.number_format = fmt

def calc(c, fmt=None, center=False, bold=False, color=TXT):
    c.font = F(10, bold, color); c.fill = FILL(CALC_FILL); c.border = BOX
    c.alignment = CENTER if center else Alignment(vertical="center")
    if fmt: c.number_format = fmt

def lines(text, width, size=9):
    cpl = max(8, width * (1.25 if size <= 9 else 1.1))
    return sum(max(1, math.ceil(len(p) / cpl)) for p in str(text).split("\n"))

def ctitle(text, size=1100, color=NAVY):
    cp = CharacterProperties(sz=size, b=True, solidFill=color, latin=DFont(typeface="Arial"))
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)

def finish_chart(ch, axes=True):
    ch.roundedCorners = False
    if ch.legend is not None:
        ch.legend.position = "b"; ch.legend.overlay = False
    if axes:
        ch.x_axis.delete = False; ch.y_axis.delete = False

def bar3d(title, cats, refs, colors, w, h, labels=True, ymax=None, kind="col"):
    ch = BarChart3D(); ch.type = kind; ch.grouping = "clustered"; ch.shape = "box"
    ch.title = ctitle(title)
    for ref in refs:
        ch.add_data(ref, titles_from_data=True)
    ch.set_categories(cats)
    for s, col in zip(ch.series, colors):
        s.graphicalProperties.solidFill = col
        s.graphicalProperties.line.solidFill = col
    ch.view3D.rotX = 12; ch.view3D.rotY = 18; ch.view3D.rAngAx = True
    ch.gapWidth = 70
    ch.z_axis.delete = True
    if labels:
        ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
    if ymax is not None:
        ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = ymax
    if kind == "bar":
        ch.x_axis.scaling.orientation = "maxMin"
    ch.width, ch.height = w, h
    finish_chart(ch)
    if len(refs) == 1:
        ch.legend = None
    return ch

wb = Workbook()
wb._named_styles["Normal"].font = Font(name="Arial", size=10)

def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)

# ================================================================ feuilles
ws_acc = wb.active; ws_acc.title = "Accueil"
ws_dash = wb.create_sheet("Dashboard")
ws_plan = wb.create_sheet("Plan")
ws_pipe = wb.create_sheet("Pipeline")
ws_sc = wb.create_sheet("Scorecard")
ws_gr = wb.create_sheet("Grille_Notes")
ws_cmp = wb.create_sheet("Comparateur")
ws_onb = wb.create_sheet("Onboarding")
ws_fp = wb.create_sheet("Fiches_Poste")
ws_q = wb.create_sheet("Questions")
ws_scr = wb.create_sheet("Scripts_Modeles")
ws_par = wb.create_sheet("Parametres")
ws_calc = wb.create_sheet("Calculs")

tabs = {ws_acc: NAVY, ws_dash: CORAL, ws_plan: OCHRE, ws_pipe: TEAL, ws_sc: LAV, ws_gr: LAV,
        ws_cmp: LAV, ws_onb: GREEN, ws_fp: SKY, ws_q: SKY, ws_scr: SKY, ws_par: MUTED, ws_calc: MUTED}
for w, col in tabs.items():
    w.sheet_properties.tabColor = col
    w.sheet_view.showGridLines = False

# ================================================================ PARAMETRES
p = ws_par
widths(p, {"A": 2, "B": 18, "C": 40, "D": 20, "E": 26, "F": 30, "G": 18, "H": 30, "I": 12, "J": 16, "K": 16})
bg(p, 12, 70)
nav(p)
banner(p, "B2:K2", "PARAMÈTRES", "B3:K3", "Cellules bleues = à personnaliser. Renommez vos crèches AVANT de saisir des candidats.", MUTED)

section(p, "B4:K4", "1 · LES 2 MICRO-CRÈCHES")
header_row(p, 5, 2, ["Code", "Nom de la crèche", "Ville", "Capacité (places)", "Date de reprise",
                     "Prise de poste cible", "Direction (nom)", "ETP direction", "Cat. 1 déjà en poste", "Cat. 2 déjà en poste"])
for i, r in enumerate((6, 7)):
    calc(p.cell(r, 2, f"C{i+1}"), center=True, bold=True)
    vals = [f"Micro-crèche {i+1}", "À compléter", 12, None, dt.date(2027, 2, 1), None, None, 0, 0]
    for j, v in enumerate(vals):
        c = p.cell(r, 3 + j, v)
        inp(c, "dd/mm/yyyy" if j in (3, 4) else None, center=j >= 2)
    p.cell(r, 7).fill = FILL("FFF6CC")
p["G6"].comment = Comment("Hypothèse à ajuster : date de prise de poste visée pour les 4 recrues de la crèche. "
                          "Alimente le Plan et les alertes d'échéance.", "Modèle")
p["I6"].number_format = p["I7"].number_format = "0.0"
p.row_dimensions[5].height = 34

section(p, "B9:K9", "2 · POSTES")
header_row(p, 10, 2, ["Code poste", "Intitulé", "Catégorie (1 / 2)", "Postes par crèche", "Diplôme requis"])
for i, (code, lib, cat, nb, dip) in enumerate(C.POSTES):
    r = 11 + i
    calc(p.cell(r, 2, code), bold=True)
    for j, v in enumerate((lib, cat, nb, dip)):
        c = p.cell(r, 3 + j, v); calc(c, center=j in (1, 2))

section(p, "B15:K15", "3 · ÉTAPES DU WORKFLOW (délai max = alerte de relance)")
header_row(p, 16, 2, ["N°", "Étape", "Délai max (jours)", "Responsable", "Livrable / outil"])
for i, (lib, sla, resp, liv, _) in enumerate(C.ETAPES):
    r = 17 + i
    calc(p.cell(r, 2, i + 1), center=True, bold=True)
    calc(p.cell(r, 3, lib), bold=True)
    inp(p.cell(r, 4, sla), center=True)
    inp(p.cell(r, 5, resp))
    c = p.cell(r, 6, liv); inp(c)
    p.merge_cells(start_row=r, start_column=6, end_row=r, end_column=11)

p.merge_cells("F16:K16")
section(p, "B26:K26", "4 · BLOCS DE COMPÉTENCES & PONDÉRATIONS DE LA SCORECARD (total = 100 % par poste)")
header_row(p, 27, 2, ["Code", "Bloc", "EJE", "AP", "CAP AEPE", "Éliminatoire", "Ce qu'on évalue"])
p["D27"] = "=$B$11"; p["E27"] = "=$B$12"; p["F27"] = "=$B$13"
for i, (code, lib, short, elim, what) in enumerate(C.BLOCS):
    r = 28 + i
    calc(p.cell(r, 2, code), center=True, bold=True)
    calc(p.cell(r, 3, lib), bold=True)
    for j, poste in enumerate(("EJE", "AP", "CAP AEPE")):
        inp(p.cell(r, 4 + j, C.POIDS[poste][i]), "0%", center=True)
    calc(p.cell(r, 7, elim), center=True)
    c = p.cell(r, 8, what); calc(c); c.alignment = WRAPC
    p.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
    p.row_dimensions[r].height = 28
p.merge_cells("H27:K27")
calc(p.cell(34, 3, "TOTAL (doit faire 100 %)"), bold=True)
for j in range(3):
    col = L(4 + j)
    c = p.cell(34, 4 + j, f"=SUM({col}28:{col}33)"); calc(c, "0%", center=True, bold=True)
p.conditional_formatting.add("D34:F34", FormulaRule(formula=["ROUND(D34,4)<>1"], fill=FILL(NO_F), font=F(10, True, NO_T)))

section(p, "B36:K36", "5 · RÈGLES DE DÉCISION")
for i, (lab, v, fmt) in enumerate([("Seuil GO (score ≥)", 75, "0"), ("Seuil « À approfondir » (score ≥)", 60, "0"),
                                    ("Profil cible du radar (/5)", 4, "0.0")]):
    r = 37 + i
    calc(p.cell(r, 2, lab), bold=True); p.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    inp(p.cell(r, 4, v), fmt, center=True)
c = p.cell(40, 2, "Veto automatique : une note de 1 sur une question Sécurité (B1) ou Bientraitance (B2) = non retenu, quel que soit le score.")
c.font = F(9, True, NO_T); p.merge_cells("B40:K40")

section(p, "B42:K42", "6 · ÉCHELLE DE NOTATION (1 à 5)")
header_row(p, 43, 2, ["Note", "Niveau", "Description"], height=22)
for i, (n, niv, desc) in enumerate(C.ECHELLE):
    r = 44 + i
    calc(p.cell(r, 2, n), center=True, bold=True); calc(p.cell(r, 3, niv), bold=True)
    c = p.cell(r, 4, desc); calc(c); c.alignment = WRAPC
    p.merge_cells(start_row=r, start_column=4, end_row=r, end_column=11)
    p.row_dimensions[r].height = 24
p.merge_cells("D43:K43")

section(p, "B50:K50", "7 · LISTES DÉROULANTES (modifiables : remplacez un libellé, ne laissez pas de trou)")
LISTS = [("Statuts", C.STATUTS, "lstStatuts"), ("Sources", C.SOURCES, "lstSources"), ("Diplômes", C.DIPLOMES, "lstDiplomes"),
         ("Contrats", C.CONTRATS, "lstContrats"), ("Oui / Non", C.OUINON, "lstOuiNon"), ("Honorabilité", C.HONOR, "lstHonor"),
         ("Vaccins / aptitude", C.VACCINS, "lstVaccins"), ("Références", C.REFS, "lstRefs"), ("Notes", C.NOTES, "lstNotes"),
         ("Onboarding", C.ONB, "lstOnb")]
header_row(p, 51, 2, [l[0] for l in LISTS], height=22)
for j, (lab, vals, nm) in enumerate(LISTS):
    col = 2 + j
    for i, v in enumerate(vals):
        inp(p.cell(52 + i, col, v))
    name(nm, f"Parametres!${L(col)}$52:${L(col)}${52 + len(vals) - 1}")
p.freeze_panes = "A4"

name("lstCreches", "Parametres!$C$6:$C$7")
name("lstPostes", "Parametres!$B$11:$B$13")
name("lstEtapes", "Parametres!$C$17:$C$24")
name("lstSLA", "Parametres!$D$17:$D$24")
name("SeuilGo", "Parametres!$D$37")
name("SeuilApp", "Parametres!$D$38")
name("CibleRadar", "Parametres!$D$39")
WEIGHTS = "Parametres!$D$28:$F$33"
WHEAD = "Parametres!$D$27:$F$27"

# ================================================================ CALCULS (masqué)
k = ws_calc
k["A1"], k["B1"] = "k", "Libellé candidat (liste déroulante Scorecard)"
for i in range(1, N + 1):
    k.cell(i + 1, 1, i)
    k.cell(i + 1, 2, f'=IFERROR(INDEX(Pipeline!$AH${FIRST}:$AH${LAST},MATCH(A{i+1},Pipeline!$AG${FIRST}:$AG${LAST},0)),"")')
name("lstCandidats", f'OFFSET(Calculs!$B$2,0,0,MAX(1,COUNTIF(Calculs!$B$2:$B${N+1},"?*")),1)')
k["D1"] = "Filtres"
k["E2"] = "Toutes"; k["E3"] = "=Parametres!$C$6"; k["E4"] = "=Parametres!$C$7"
name("lstCrechesToutes", "Calculs!$E$2:$E$4")
k["D6"] = "critCreche (Dashboard)"
k["E6"] = '=IF(Dashboard!$D$5="Toutes","*",Dashboard!$D$5)'
name("critCreche", "Calculs!$E$6")
k["G1"], k["H1"], k["I1"] = "Bloc", "Candidat·e", "Profil cible"
for i, b in enumerate(C.BLOCS):
    r = 2 + i
    k.cell(r, 7, b[2])
    k.cell(r, 8, f'=IF(Scorecard!$K$7="",0,N(INDEX(Grille_Notes!$S${FIRST}:$X${LAST},Scorecard!$K$7,{i+1})))')
    k.cell(r, 9, "=CibleRadar")
ws_calc.sheet_state = "hidden"

# ================================================================ PIPELINE
pp = ws_pipe
PCOLS = [
    # (lettre, en-tête, largeur, type)  type: in / calc / helper
    ("A", "ID", 7, "id"), ("B", "Nom", 16, "in"), ("C", "Prénom", 13, "in"), ("D", "Téléphone", 14, "in"),
    ("E", "E-mail", 22, "in"), ("F", "Crèche ▼", 16, "in"), ("G", "Poste ▼", 11, "in"), ("H", "Source ▼", 18, "in"),
    ("I", "Date candidature", 11, "in"), ("J", "Diplôme ▼", 26, "in"), ("K", "Diplôme / poste", 13, "calc"),
    ("L", "Exp. 0-3 ans (ans)", 8, "in"), ("M", "Disponible le", 11, "in"), ("N", "Prétentions (€ brut/mois)", 11, "in"),
    ("O", "Étape ▼", 28, "in"), ("P", "Statut ▼", 12, "in"), ("Q", "Dernière action (date)", 11, "in"),
    ("R", "Prochaine action", 24, "in"), ("S", "Note tél. /5 ▼", 7, "in"), ("T", "Score entretien /100", 9, "calc"),
    ("U", "Verdict", 30, "calc"), ("V", "Diplôme vérifié ▼", 10, "in"), ("W", "Attestation d'honorabilité ▼", 13, "in"),
    ("X", "Vaccins / aptitude ▼", 11, "in"), ("Y", "Références ▼", 13, "in"), ("Z", "Dossier conformité", 24, "calc"),
    ("AA", "Date d'embauche", 11, "in"), ("AB", "Délai (jours)", 8, "calc"), ("AC", "⚠ Alerte", 34, "calc"),
    ("AD", "Commentaires", 30, "in"),
    ("AE", "Nom complet", 18, "h"), ("AF", "N° étape", 6, "h"), ("AG", "N° actif", 6, "h"), ("AH", "Libellé liste", 30, "h"),
    ("AI", "Clé rang crèche", 22, "h"), ("AJ", "Clé rang toutes", 22, "h"), ("AK", "Clé embauche", 22, "h"), ("AL", "Rang", 6, "h"),
]
widths(pp, {"A": 7, **{c[0]: c[2] for c in PCOLS}})
nav(pp, "B1", [("D1", "Scorecard", "🎯 Scorecard"), ("F1", "Dashboard", "📊 Dashboard")])
banner(pp, "A2:AD2", "PIPELINE CANDIDATS — 1 ligne = 1 candidature", color=TEAL)
put(pp, "A3:AD3", "🟦 Cellules bleues = saisie   ·   ▼ = liste déroulante   ·   gris = calcul automatique (ne pas modifier)   ·   "
    "La note d'entretien se saisit dans « Grille_Notes » et remonte ici automatiquement." + ("" if DEMO else "   ·   Ligne C001 = EXEMPLE à effacer."),
    F(9, False, MUTED, True), None, Alignment(vertical="center", indent=1))
groups = [("A5:E5", "IDENTITÉ", NAVY), ("F5:N5", "CANDIDATURE", TEAL), ("O5:S5", "SUIVI DU WORKFLOW", LAV),
          ("T5:U5", "ÉVALUATION", CORAL), ("V5:Z5", "CONFORMITÉ (obligatoire avant embauche)", GREEN),
          ("AA5:AD5", "EMBAUCHE & ALERTES", OCHRE), ("AE5:AL5", "calculs internes", MUTED)]
for ref, lab, col in groups:
    put(pp, ref, lab, F(9, True, "FFFFFF"), FILL(col), CENTER)
for i, (col, lab, w, t) in enumerate(PCOLS):
    gcol = {"A": NAVY}.get(col)
    c = pp[f"{col}6"]; c.value = lab
    c.font = F(9, True, "FFFFFF"); c.alignment = CENTER; c.border = BOX
    c.fill = FILL(MUTED if t == "h" else (NAVY if t in ("id",) else ("4A6278" if t == "calc" else NAVY)))
pp.row_dimensions[6].height = 44
pp.row_dimensions[5].height = 20

DATEF, EURF = "dd/mm/yyyy", '#,##0 "€"'
for r in range(FIRST, LAST + 1):
    idx = r - FIRST + 1
    a = pp.cell(r, 1, f"C{idx:03d}"); a.font = F(9, True, NAVY); a.fill = FILL("E8EEF4"); a.border = BOX; a.alignment = CENTER
    for col, lab, w, t in PCOLS[1:]:
        c = pp[f"{col}{r}"]
        if t == "in":
            fmt = DATEF if col in ("I", "M", "Q", "AA") else (EURF if col == "N" else None)
            inp(c, fmt, center=col in ("F", "G", "L", "P", "S", "V", "W", "X", "Y", "I", "M", "Q", "AA", "N"))
        elif t == "calc":
            calc(c, center=col in ("K", "T", "AB"))
        else:
            c.font = F(8, False, MUTED); c.border = BOX
    R = r
    pp[f"K{R}"] = (f'=IF(OR(G{R}="",J{R}=""),"",IF(OR(AND(G{R}="EJE",J{R}="DE Éducateur de jeunes enfants"),'
                   f'AND(G{R}="AP",J{R}="DE Auxiliaire de puériculture"),AND(G{R}="CAP AEPE",J{R}<>"Autre / en cours")),'
                   f'"✅ Conforme","⚠️ À vérifier"))')
    pp[f"T{R}"] = f'=IF(B{R}="","",IFERROR(INDEX(Grille_Notes!$Y${FIRST}:$Y${LAST},MATCH(A{R},Grille_Notes!$A${FIRST}:$A${LAST},0)),""))'
    pp[f"T{R}"].number_format = "0.0"
    pp[f"U{R}"] = (f'=IF(B{R}="","",IF(Y{R}="Négatif","❌ Références négatives",'
                   f'IFERROR(INDEX(Grille_Notes!$AB${FIRST}:$AB${LAST},MATCH(A{R},Grille_Notes!$A${FIRST}:$A${LAST},0)),"")))')
    ok = f'((V{R}="Oui")+(W{R}="Reçue")+(X{R}="OK")+OR(Y{R}="OK",Y{R}="Réserve"))'
    pp[f"Z{R}"] = f'=IF(B{R}="","",IF({ok}=4,"✅ Complet","⏳ "&(4-{ok})&" pièce(s) manquante(s)"))'
    pp[f"AB{R}"] = f'=IF(AND(N(AA{R})>0,N(I{R})>0),AA{R}-I{R},"")'
    last = f'MAX(N(I{R}),N(Q{R}))'
    pp[f"AC{R}"] = (f'=IF(B{R}="","",IF(AND(N(AF{R})>=7,Z{R}<>"✅ Complet"),"⚠️ Dossier incomplet — bloquer l\'embauche",'
                    f'IF(AND(P{R}="En cours",OR(ISNUMBER(SEARCH("Non retenu",U{R})),ISNUMBER(SEARCH("négatives",U{R})))),"⚠️ Verdict négatif — clôturer la candidature",'
                    f'IF(AND(P{R}="En cours",N(AF{R})>0,N(AF{R})<8,{last}>0),IF(TODAY()-{last}>INDEX(lstSLA,AF{R}),'
                    f'"🔔 Relancer : "&(TODAY()-{last})&" j sans action",""),""))))')
    pp[f"AE{R}"] = f'=IF(B{R}="","",UPPER(B{R})&" "&C{R})'
    pp[f"AF{R}"] = f'=IF(O{R}="","",IFERROR(MATCH(O{R},lstEtapes,0),""))'
    pp[f"AG{R}"] = f'=IF(B{R}="","",COUNTIF($B${FIRST}:B{R},"?*"))'
    pp[f"AH{R}"] = f'=IF(B{R}="","",A{R}&" · "&AE{R}&" · "&G{R}&" · "&F{R})'
    pp[f"AL{R}"] = (f'=IF(ISNUMBER(T{R}),COUNTIFS($F${FIRST}:$F${LAST},F{R},$G${FIRST}:$G${LAST},G{R},$T${FIRST}:$T${LAST},">"&T{R})'
                    f'+COUNTIFS($F${FIRST}:F{R},F{R},$G${FIRST}:G{R},G{R},$T${FIRST}:T{R},T{R}),"")')
    pp[f"AI{R}"] = f'=IF(AL{R}="","",F{R}&"|"&G{R}&"|"&AL{R})'
    pp[f"AJ{R}"] = (f'=IF(ISNUMBER(T{R}),"Toutes|"&G{R}&"|"&(COUNTIFS($G${FIRST}:$G${LAST},G{R},$T${FIRST}:$T${LAST},">"&T{R})'
                    f'+COUNTIFS($G${FIRST}:G{R},G{R},$T${FIRST}:T{R},T{R})),"")')
    pp[f"AK{R}"] = (f'=IF(N(AF{R})=8,F{R}&"|"&G{R}&"|"&COUNTIFS($F${FIRST}:F{R},F{R},$G${FIRST}:G{R},G{R},$AF${FIRST}:AF{R},8),"")')
    pp.row_dimensions[r].height = 18
for col in ("AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL"):
    pp.column_dimensions[col].hidden = True
pp.freeze_panes = f"D{FIRST}"
pp.auto_filter.ref = f"A6:AD{LAST}"

def dv_list(ws, formula, ref, title="Valeur", msg=None):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True,
                        errorTitle="Valeur non autorisée", error="Choisissez une valeur dans la liste déroulante.")
    if msg:
        dv.promptTitle, dv.prompt, dv.showInputMessage = title, msg, True
    dv.add(ref); ws.add_data_validation(dv)

PR_ = lambda col: f"{col}{FIRST}:{col}{LAST}"
dv_list(pp, "=lstCreches", PR_("F")); dv_list(pp, "=lstPostes", PR_("G")); dv_list(pp, "=lstSources", PR_("H"))
dv_list(pp, "=lstDiplomes", PR_("J")); dv_list(pp, "=lstEtapes", PR_("O")); dv_list(pp, "=lstStatuts", PR_("P"))
dv_list(pp, "=lstNotes", PR_("S")); dv_list(pp, "=lstOuiNon", PR_("V")); dv_list(pp, "=lstHonor", PR_("W"))
dv_list(pp, "=lstVaccins", PR_("X")); dv_list(pp, "=lstRefs", PR_("Y"))
dvd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, showErrorMessage=True,
                     errorTitle="Date attendue", error="Saisissez une date au format jj/mm/aaaa.")
for col in ("I", "M", "Q", "AA"):
    dvd.add(PR_(col))
pp.add_data_validation(dvd)

def verdict_cf(ws, ref, anchor):
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'ISNUMBER(SEARCH("GO",{anchor}))'], fill=FILL(GO_F), font=F(10, True, GO_T)))
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'ISNUMBER(SEARCH("approfondir",{anchor}))'], fill=FILL(AP_F), font=F(10, True, AP_T)))
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("Non retenu",{anchor})),ISNUMBER(SEARCH("négatives",{anchor})))'],
                                                   fill=FILL(NO_F), font=F(10, True, NO_T)))

verdict_cf(pp, PR_("U"), f"U{FIRST}")
pp.conditional_formatting.add(PR_("T"), DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color=CORAL, showValue=True))
pp.conditional_formatting.add(PR_("Z"), FormulaRule(formula=[f'LEFT(Z{FIRST},1)="✅"'], fill=FILL(GO_F), font=F(10, True, GO_T)))
pp.conditional_formatting.add(PR_("Z"), FormulaRule(formula=[f'LEFT(Z{FIRST},1)="⏳"'], fill=FILL(AP_F), font=F(10, False, AP_T)))
pp.conditional_formatting.add(PR_("AC"), FormulaRule(formula=[f'LEN(AC{FIRST})>0'], fill=FILL(NO_F), font=F(10, True, NO_T)))
pp.conditional_formatting.add(PR_("K"), FormulaRule(formula=[f'LEFT(K{FIRST},1)="⚠"'], font=F(10, True, AP_T)))
pp.conditional_formatting.add(PR_("P"), FormulaRule(formula=[f'P{FIRST}="Embauché(e)"'], fill=FILL(GO_F), font=F(10, True, GO_T)))
pp.conditional_formatting.add(PR_("P"), FormulaRule(formula=[f'OR(P{FIRST}="Refusé(e)",P{FIRST}="Désisté(e)")'], font=F(10, False, MUTED, True)))
for poste, col in POSTE_COL.items():
    pp.conditional_formatting.add(PR_("G"), FormulaRule(formula=[f'G{FIRST}="{poste}"'], fill=FILL(col), font=F(10, True, "FFFFFF")))

# ================================================================ GRILLE_NOTES
g = ws_gr
gw = {"A": 7, "B": 24, "C": 16, "D": 10, "E": 16, "F": 11}
for i in range(12): gw[L(7 + i)] = 5.5
for i in range(6): gw[L(19 + i)] = 6.5
gw.update({"Y": 9, "Z": 7, "AA": 7, "AB": 36, "AC": 36, "AD": 36})
for i in range(6): gw[L(31 + i)] = 6
widths(g, gw)
nav(g, "B1", [("C1", "Scorecard", "🎯 Scorecard"), ("E1", "Questions", "❓ Questions")])
banner(g, "A2:AD2", "GRILLE DE NOTATION — saisie des notes d'entretien (1 à 5)", color=LAV)
put(g, "A3:AD3", "1) Ouvrez la Scorecard du candidat (questions + réponses attendues)  ·  2) Saisissez ici une note de 1 à 5 par question  ·  "
    "3) Le score /100 pondéré, le veto et le verdict se calculent seuls. Q1-Q2 = Sécurité, Q3-Q4 = Bientraitance (une note de 1 = VETO).",
    F(9, False, MUTED, True), None, Alignment(vertical="center", indent=1))
gg = [("A5:F5", "CANDIDAT", NAVY), ("G5:H5", "🔒 B1 Sécurité", CORAL), ("I5:J5", "🔒 B2 Bientraitance", CORAL),
      ("K5:L5", "B3 Éveil", LAV), ("M5:N5", "B4 Familles", TEAL), ("O5:P5", "B5 Équipe", SKY), ("Q5:R5", "B6 Exp./Motiv.", OCHRE),
      ("S5:X5", "MOYENNE PAR BLOC /5", "4A6278"), ("Y5:AB5", "RÉSULTAT", NAVY), ("AC5:AD5", "COMMENTAIRES", MUTED),
      ("AE5:AJ5", "poids (calc.)", MUTED)]
for ref, lab, col in gg:
    put(g, ref, lab, F(9, True, "FFFFFF"), FILL(col), CENTER)
gh = ["ID", "Candidat·e", "Crèche", "Poste", "Évaluateur(s)", "Date entretien"] + [f"Q{i}" for i in range(1, 13)] + \
     [b[0] for b in C.BLOCS] + ["Score /100", "Saisie", "Veto", "Verdict", "Points forts", "Points de vigilance"] + [f"w{i}" for i in range(1, 7)]
header_row(g, 6, 1, gh, NAVY, 30)
for r in range(FIRST, LAST + 1):
    idx = r - FIRST + 1
    a = g.cell(r, 1, f"C{idx:03d}"); a.font = F(9, True, NAVY); a.fill = FILL("E8EEF4"); a.border = BOX; a.alignment = CENTER
    g[f"B{r}"] = f'=IFERROR(INDEX(Pipeline!$AE${FIRST}:$AE${LAST},MATCH(A{r},Pipeline!$A${FIRST}:$A${LAST},0)),"")'
    g[f"C{r}"] = f'=IF(B{r}="","",INDEX(Pipeline!$F${FIRST}:$F${LAST},MATCH(A{r},Pipeline!$A${FIRST}:$A${LAST},0))&"")'
    g[f"D{r}"] = f'=IF(B{r}="","",INDEX(Pipeline!$G${FIRST}:$G${LAST},MATCH(A{r},Pipeline!$A${FIRST}:$A${LAST},0))&"")'
    for col in "BCD":
        calc(g[f"{col}{r}"], center=col == "D")
    inp(g[f"E{r}"]); inp(g[f"F{r}"], DATEF, True)
    for i in range(12):
        inp(g.cell(r, 7 + i), center=True)
    pairs = [("G", "H"), ("I", "J"), ("K", "L"), ("M", "N"), ("O", "P"), ("Q", "R")]
    for i, (c1, c2) in enumerate(pairs):
        col = L(19 + i)
        g[f"{col}{r}"] = f'=IF(COUNT({c1}{r}:{c2}{r})=0,"",AVERAGE({c1}{r}:{c2}{r}))'
        calc(g[f"{col}{r}"], "0.0", center=True)
        wc = L(31 + i)
        g[f"{wc}{r}"] = f'=IFERROR(INDEX({WEIGHTS},{i+1},MATCH($D{r},{WHEAD},0)),0)'
        g[f"{wc}{r}"].font = F(8, False, MUTED); g[f"{wc}{r}"].number_format = "0%"
    blocks = [L(19 + i) for i in range(6)]; ws_ = [L(31 + i) for i in range(6)]
    num = "+".join(f"{w}{r}*N({b}{r})" for w, b in zip(ws_, blocks))
    den = "+".join(f"{w}{r}*ISNUMBER({b}{r})" for w, b in zip(ws_, blocks))
    g[f"Y{r}"] = f'=IF(COUNT(S{r}:X{r})=0,"",IFERROR(ROUND(({num})/({den})/5*100,1),""))'
    calc(g[f"Y{r}"], "0.0", center=True, bold=True)
    g[f"Z{r}"] = f'=IF(COUNT(G{r}:R{r})=0,"",COUNT(G{r}:R{r})&"/12")'
    calc(g[f"Z{r}"], center=True)
    g[f"AA{r}"] = f'=IF(COUNTIF(G{r}:J{r},1)>0,"VETO","")'
    calc(g[f"AA{r}"], center=True, bold=True, color=NO_T)
    g[f"AB{r}"] = (f'=IF(Y{r}="","",IF(AA{r}="VETO","⛔ Non retenu — alerte sécurité/bientraitance",'
                   f'IF(Y{r}>=SeuilGo,"✅ GO",IF(Y{r}>=SeuilApp,"🟠 À approfondir","❌ Non retenu")))'
                   f'&IF(COUNT(G{r}:R{r})<12," (évaluation partielle)",""))')
    calc(g[f"AB{r}"])
    inp(g[f"AC{r}"]); inp(g[f"AD{r}"])
    g.row_dimensions[r].height = 18
for i in range(6):
    g.column_dimensions[L(31 + i)].hidden = True
g.freeze_panes = f"G{FIRST}"
g.auto_filter.ref = f"A6:AD{LAST}"
dvn = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True, showErrorMessage=True,
                     errorTitle="Note de 1 à 5", error="Saisissez un entier de 1 (insuffisant) à 5 (excellent).",
                     promptTitle="Note 1 → 5", prompt="1 Insuffisant · 2 Fragile · 3 Conforme · 4 Solide · 5 Excellent", showInputMessage=True)
dvn.add(f"G{FIRST}:R{LAST}"); g.add_data_validation(dvn)
g.add_data_validation(dvd2 := DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True))
dvd2.add(f"F{FIRST}:F{LAST}")
g.conditional_formatting.add(f"G{FIRST}:R{LAST}", ColorScaleRule(start_type="num", start_value=1, start_color="F4A6A0",
                                                                 mid_type="num", mid_value=3, mid_color="FBE7A1",
                                                                 end_type="num", end_value=5, end_color="9ED8AE"))
g.conditional_formatting.add(f"S{FIRST}:X{LAST}", ColorScaleRule(start_type="num", start_value=1, start_color="F4A6A0",
                                                                 mid_type="num", mid_value=3, mid_color="FBE7A1",
                                                                 end_type="num", end_value=5, end_color="9ED8AE"))
g.conditional_formatting.add(f"Y{FIRST}:Y{LAST}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color=CORAL, showValue=True))
verdict_cf(g, f"AB{FIRST}:AB{LAST}", f"AB{FIRST}")

# ================================================================ QUESTIONS
q = ws_q
widths(q, {"A": 12, "B": 10, "C": 5, "D": 22, "E": 15, "F": 52, "G": 52, "H": 40, "I": 34})
bg(q, 10, 45)
nav(q, "B1", [("D1", "Scorecard", "🎯 Scorecard")])
banner(q, "B2:I2", "BANQUE DE QUESTIONS — 12 questions par poste (expérience, mises en situation, motivation)", "B3:I3",
       "2 questions par bloc de compétences. La Scorecard affiche automatiquement le bon jeu de questions selon le poste du candidat.", SKY)
header_row(q, 5, 1, ["Clé", "Poste", "N°", "Bloc", "Type", "Question à poser", "Réponse attendue (indicateurs ✅)",
                     "Signaux d'alerte (🚩)", "Relance possible"], NAVY, 30)
r = 6
for poste in ("EJE", "AP", "CAP AEPE"):
    for n, (b, typ, ques, att, red, rel) in enumerate(C.QUESTIONS[poste], start=1):
        q.cell(r, 1, f'=B{r}&"|"&C{r}').font = F(8, False, MUTED)
        c = q.cell(r, 2, poste); c.font = F(9, True, "FFFFFF"); c.fill = FILL(POSTE_COL[poste]); c.alignment = CENTER; c.border = BOX
        c = q.cell(r, 3, f"Q{n}"); calc(c, center=True, bold=True)
        c = q.cell(r, 4, f"=Parametres!$C${27 + b}"); calc(c); c.alignment = WRAPC
        if b <= 2: c.font = F(10, True, NO_T)
        c = q.cell(r, 5, typ); calc(c, center=True)
        for j, txt in enumerate((ques, att, red, rel)):
            c = q.cell(r, 6 + j, txt); c.font = F(9, j == 0, TXT if j != 2 else NO_T); c.fill = FILL(CARD); c.border = BOX; c.alignment = WRAP
        h = max(lines(ques, 52), lines(att, 52), lines(red, 40), lines(rel, 34)) * 12.5 + 8
        q.row_dimensions[r].height = max(42, h)
        r += 1
q.column_dimensions["A"].hidden = True
q.freeze_panes = "F6"
q.auto_filter.ref = "A5:I41"

# ================================================================ SCRIPTS & MODELES
s = ws_scr
widths(s, {"A": 2, "B": 6, "C": 60, "D": 32, "E": 13, "F": 50})
bg(s, 7, 60)
nav(s)
banner(s, "B2:F2", "SCRIPTS & MODÈLES", "B3:F3", "Pré-qualification téléphonique (15 min) + messages prêts à envoyer", SKY)
section(s, "B5:F5", "A · ENTRETIEN TÉLÉPHONIQUE DE PRÉ-QUALIFICATION (étape 3)", TEAL)
header_row(s, 6, 2, ["N°", "Question", "Objectif", "Éliminatoire", "Ce qu'on veut entendre"], NAVY, 24)
for i, (ques, obj, elim, att) in enumerate(C.SCRIPT_TEL):
    rr = 7 + i
    calc(s.cell(rr, 2, i + 1), center=True, bold=True)
    for j, v in enumerate((ques, obj, elim, att)):
        c = s.cell(rr, 3 + j, v); c.font = F(9, j == 0, NO_T if (j == 2 and v == "Oui") else TXT)
        c.fill = FILL(CARD); c.border = BOX; c.alignment = CENTER if j == 2 else WRAPC
    s.row_dimensions[rr].height = max(30, max(lines(ques, 60), lines(att, 50)) * 12.5 + 8)
put(s, "B16:F16", "➜ Décision : passer à l'entretien si aucun critère éliminatoire n'est bloquant ET note téléphonique ≥ 3/5 "
    "(à saisir dans le Pipeline, colonne « Note tél. »).", F(10, True, NAVY), FILL("E8EEF4"), WRAPC, BOX)
s.row_dimensions[16].height = 30
section(s, "B18:F18", "B · MODÈLES DE MESSAGES (remplacez les [crochets])", TEAL)
rr = 19
for titre, corps in C.MODELES:
    put(s, f"B{rr}:F{rr}", "✉ " + titre, F(10, True, NAVY), FILL("E8EEF4"), Alignment(vertical="center", indent=1), BOX)
    put(s, f"B{rr+1}:F{rr+1}", corps, F(9), FILL(CARD), WRAP, BOX)
    s.row_dimensions[rr + 1].height = lines(corps, 150) * 12.5 + 10
    rr += 3

# ================================================================ FICHES DE POSTE
fp = ws_fp
widths(fp, {"A": 2, "B": 58, "C": 2, "D": 58, "E": 2, "F": 58})
bg(fp, 7, 40)
nav(fp)
banner(fp, "B2:F2", "FICHES DE POSTE — micro-crèche (12 berceaux)", "B3:F3",
       "Synthèse des référentiels métiers (solidarites.gouv.fr, ONISEP, CAF) et du cadre 2026. Rémunérations = repères indicatifs, à ajuster.", SKY)
cols = {"EJE": "B", "AP": "D", "CAP AEPE": "F"}
for poste, col in cols.items():
    put(fp, f"{col}5", C.FICHES[poste]["titre"], F(12, True, "FFFFFF"), FILL(POSTE_COL[poste]), CENTER)
fp.row_dimensions[5].height = 34
r = 6
for sec in C.FICHE_SECTIONS:
    hmax = 1
    for poste, col in cols.items():
        put(fp, f"{col}{r}", sec.upper(), F(8, True, POSTE_COL[poste]), FILL(CARD), Alignment(vertical="bottom", indent=1))
        if sec == "Prérequis à l'embauche":
            val = C.PREREQUIS
        else:
            val = C.FICHES[poste][sec]
            if isinstance(val, list):
                val = "\n".join("• " + x for x in val)
        c = put(fp, f"{col}{r+1}", val, F(9, sec in ("Diplôme requis",), TXT), FILL(CARD), Alignment(wrap_text=True, vertical="top", indent=1))
        if sec == "Annonce prête à publier":
            c.fill = FILL("FFF9E8")
        hmax = max(hmax, lines(val, 54))
    fp.row_dimensions[r].height = 18
    fp.row_dimensions[r + 1].height = hmax * 12.5 + 8
    r += 2

# ================================================================ PLAN
pl = ws_plan
widths(pl, {"A": 7, "B": 16, "C": 11, "D": 5, "E": 8, "F": 13, "G": 8, "H": 12, "I": 12, "J": 10, "K": 9,
            "L": 9, "M": 10, "N": 10, "O": 8, "P": 22, "Q": 22, "R": 10, "S": 18})
bg(pl, 20, 30)
nav(pl, "B1", [("C1", "Dashboard", "📊 Dashboard")])
banner(pl, "A2:S2", "PLAN DE RECRUTEMENT — 8 postes (4 par crèche)", "A3:S3",
       "Statut, recrue et alertes se mettent à jour depuis le Pipeline. Cellules bleues = à ajuster (contrat, horaires, dates).", OCHRE)
header_row(pl, 6, 1, ["ID poste", "Crèche", "Poste", "N°", "Catégorie", "Contrat ▼", "Heures / sem.", "Ouverture recrutement",
                      "Prise de poste cible", "Candidats", "En cours", "Évalués", "Score recrue / meilleur en cours", "Embauchés", "Pourvu",
                      "Statut", "Recrue", "J restants", "Alerte échéance"], NAVY, 40)
slots = [("EJE", 1), ("AP", 1), ("CAP AEPE", 1), ("CAP AEPE", 2)]
r = 7
for ci in range(2):
    for poste, n in slots:
        pid = f"P{r-6:02d}"
        calc(pl.cell(r, 1, pid), center=True, bold=True)
        calc(pl.cell(r, 2, f"=Parametres!$C${6+ci}"), bold=True)
        c = pl.cell(r, 3, poste); c.font = F(9, True, "FFFFFF"); c.fill = FILL(POSTE_COL[poste]); c.alignment = CENTER; c.border = BOX
        calc(pl.cell(r, 4, n), center=True)
        calc(pl.cell(r, 5, f'=INDEX(Parametres!$D$11:$D$13,MATCH(C{r},lstPostes,0))'), center=True)
        inp(pl.cell(r, 6, "CDI"), center=True); inp(pl.cell(r, 7, 35), center=True)
        inp(pl.cell(r, 8, dt.date(2026, 12, 1)), DATEF, True)
        inp(pl.cell(r, 9, f"=Parametres!$G${6+ci}"), DATEF, True)
        crit = f'Pipeline!$F${FIRST}:$F${LAST},B{r},Pipeline!$G${FIRST}:$G${LAST},C{r}'
        pl.cell(r, 10, f'=COUNTIFS({crit},Pipeline!$B${FIRST}:$B${LAST},"?*")')
        pl.cell(r, 11, f'=COUNTIFS({crit},Pipeline!$P${FIRST}:$P${LAST},"En cours")')
        pl.cell(r, 12, f'=COUNTIFS({crit},Pipeline!$T${FIRST}:$T${LAST},">=0")')
        pl.cell(r, 13, f'=IF(O{r}=1,IFERROR(INDEX(Pipeline!$T${FIRST}:$T${LAST},MATCH(B{r}&"|"&C{r}&"|"&D{r},Pipeline!$AK${FIRST}:$AK${LAST},0)),"—"),IF(COUNTIFS({crit},Pipeline!$P${FIRST}:$P${LAST},"En cours",Pipeline!$T${FIRST}:$T${LAST},">=0")=0,"—",_xlfn.MAXIFS(Pipeline!$T${FIRST}:$T${LAST},{crit},Pipeline!$P${FIRST}:$P${LAST},"En cours")))')
        pl.cell(r, 14, f'=COUNTIFS({crit},Pipeline!$AF${FIRST}:$AF${LAST},8)')
        pl.cell(r, 15, f'=IF(N{r}>=D{r},1,0)')
        pl.cell(r, 16, f'=IF(O{r}=1,"✅ Pourvu",IF(K{r}>0,"🔄 En cours",IF(J{r}>0,"🟡 À traiter","⏳ À lancer")))')
        pl.cell(r, 17, f'=IFERROR(INDEX(Pipeline!$AE${FIRST}:$AE${LAST},MATCH(B{r}&"|"&C{r}&"|"&D{r},Pipeline!$AK${FIRST}:$AK${LAST},0)),"")')
        pl.cell(r, 18, f'=IF(N(I{r})=0,"",I{r}-TODAY())')
        pl.cell(r, 19, f'=IF(O{r}=1,"",IF(R{r}="","",IF(R{r}<0,"🔴 Échéance dépassée",IF(R{r}<=30,"🟠 Moins de 30 jours","🟢 Dans les temps"))))')
        for cc in range(10, 20):
            calc(pl.cell(r, cc), "0.0" if cc == 13 else None, center=cc not in (17, 19))
        pl.row_dimensions[r].height = 22
        r += 1
calc(pl.cell(15, 2, "TOTAL"), bold=True)
for cc, f_ in ((10, "=SUM(J7:J14)"), (11, "=SUM(K7:K14)"), (12, "=SUM(L7:L14)"), (14, "=SUM(N7:N14)"), (15, "=SUM(O7:O14)")):
    calc(pl.cell(15, cc, f_), center=True, bold=True)
calc(pl.cell(15, 16, '=O15&" / 8 postes pourvus"'), bold=True)
dv_list(pl, "=lstContrats", "F7:F14")
pl.conditional_formatting.add("P7:P14", FormulaRule(formula=['O7=1'], fill=FILL(GO_F), font=F(10, True, GO_T)))
pl.conditional_formatting.add("P7:P14", FormulaRule(formula=['LEFT(P7,1)="⏳"'], font=F(10, True, MUTED)))
pl.conditional_formatting.add("S7:S14", FormulaRule(formula=['ISNUMBER(SEARCH("dépassée",S7))'], fill=FILL(NO_F), font=F(10, True, NO_T)))
pl.conditional_formatting.add("S7:S14", FormulaRule(formula=['ISNUMBER(SEARCH("30 jours",S7))'], fill=FILL(AP_F), font=F(10, True, AP_T)))
put(pl, "A17:S17", "Catégorie 1 = diplômé·e d'État (EJE, AP…), catégorie 2 = CAP AEPE / qualification. Contrôle du ratio ≥ 40 % de cat. 1 : voir Dashboard › Conformité.",
    F(9, False, MUTED, True), None, Alignment(vertical="center", indent=1))
pl.freeze_panes = "A7"

# ================================================================ SCORECARD
sc = ws_sc
widths(sc, {"A": 2, "B": 6, "C": 18, "D": 15, "E": 46, "F": 42, "G": 34, "H": 9, "I": 14, "J": 2, "K": 10})
bg(sc, 10, 60)
nav(sc, "B1", [("D1", "Grille_Notes", "✏️ Grille de notes"), ("E1", "Comparateur", "⚖️ Comparateur")])
banner(sc, "B2:I2", "FICHE D'ENTRETIEN & SCORECARD INTERACTIVE", "B3:I3",
       "Choisissez un candidat : questions, réponses attendues, notes, score pondéré, verdict et radar s'affichent. Imprimable (A4 paysage).", LAV)
put(sc, "B5:C5", "👤 CANDIDAT ▸", F(10, True, NAVY), None, Alignment(vertical="center", horizontal="right"))
c = put(sc, "D5:F5", "C001 · EXEMPLE CANDIDAT · AP · Micro-crèche 1" if not DEMO else None, F(11, True, INP_FONT), FILL(INP_FILL), Alignment(vertical="center", indent=1), BOX)
dv_list(sc, "=lstCandidats", "D5", "Candidat", "Choisissez le candidat à évaluer (liste issue du Pipeline).")
put(sc, "G5:I5", '=IF($K$7="","",HYPERLINK("#\'Grille_Notes\'!G"&($K$7+' + str(FIRST - 1) + '),"✏️ Saisir / modifier ses notes ➜"))',
    F(10, True, TEAL), None, Alignment(vertical="center", indent=1))
put(sc, "B6:C6", "📄 OU GRILLE VIERGE ▸", F(9, True, MUTED), None, Alignment(vertical="center", horizontal="right"))
inp(sc["D6"], center=True)
dv_list(sc, "=lstPostes", "D6", "Grille vierge", "Sans candidat sélectionné : choisissez un poste pour imprimer un guide d'entretien vide.")
put(sc, "E6:I6", "← pour imprimer un guide d'entretien vide (laissez la case candidat vide)", F(9, False, MUTED, True), None, Alignment(vertical="center"))
# aides (colonne K masquée)
sc["K4"] = "aides"
sc["K5"] = '=IF(D5="","",LEFT(D5,4))'
sc["K6"] = f'=IF(K5="","",IFERROR(MATCH(K5,Pipeline!$A${FIRST}:$A${LAST},0),""))'
sc["K7"] = f'=IF(K5="","",IFERROR(MATCH(K5,Grille_Notes!$A${FIRST}:$A${LAST},0),""))'
sc["K8"] = f'=IF(K6<>"",INDEX(Pipeline!$G${FIRST}:$G${LAST},K6)&"",D6&"")'
sc.column_dimensions["K"].hidden = True

def PV(col):
    return f'IF($K$6="","",INDEX(Pipeline!${col}${FIRST}:${col}${LAST},$K$6)&"")'
def PD(col):
    return f'IF($K$6="","",IF(N(INDEX(Pipeline!${col}${FIRST}:${col}${LAST},$K$6))=0,"",INDEX(Pipeline!${col}${FIRST}:${col}${LAST},$K$6)))'
def GV(col):
    return f'IF($K$7="","",INDEX(Grille_Notes!${col}${FIRST}:${col}${LAST},$K$7)&"")'

ident = [("Candidat·e", f'={PV("AE")}'), ("Crèche · Poste", f'=IF($K$6="","",{PV("F")}&"  ·  "&{PV("G")})'),
         ("Diplôme", f'=IF($K$6="","",{PV("J")}&"   "&{PV("K")})'),
         ("Expérience · Source", f'=IF($K$6="","",{PV("L")}&" an(s) 0-3 ans  ·  "&{PV("H")})'),
         ("Étape · Statut", f'=IF($K$6="","",{PV("O")}&"  ·  "&{PV("P")})')]
for i, (lab, f_) in enumerate(ident):
    rr = 8 + i
    put(sc, f"B{rr}:C{rr}", lab, F(9, True, MUTED), FILL(CARD), Alignment(vertical="center", indent=1), BOTTOM)
    put(sc, f"D{rr}:E{rr}", f_, F(10, i == 0, NAVY if i == 0 else TXT), FILL(CARD), Alignment(vertical="center", indent=1), BOTTOM)
mid = [("Évaluateur(s)", f'={GV("E")}', None), ("Date d'entretien", f'=IF($K$7="","",IF(N(INDEX(Grille_Notes!$F${FIRST}:$F${LAST},$K$7))=0,"",INDEX(Grille_Notes!$F${FIRST}:$F${LAST},$K$7)))', DATEF),
       ("Disponible le", f'={PD("M")}', DATEF), ("Prétentions", f'={PD("N")}', EURF), ("Dossier conformité", f'={PV("Z")}', None)]
for i, (lab, f_, fmt) in enumerate(mid):
    rr = 8 + i
    sc[f"F{rr}"] = lab; sc[f"F{rr}"].font = F(9, True, MUTED); sc[f"F{rr}"].fill = FILL(CARD); sc[f"F{rr}"].border = BOTTOM
    sc[f"F{rr}"].alignment = Alignment(vertical="center", horizontal="right", indent=1)
    c = sc[f"G{rr}"]; c.value = f_; c.font = F(10); c.fill = FILL(CARD); c.border = BOTTOM; c.alignment = Alignment(vertical="center", indent=1)
    if fmt: c.number_format = fmt
put(sc, "H8:I8", "SCORE /100", F(9, True, "FFFFFF"), FILL(NAVY), CENTER)
put(sc, "H9:I10", f'=IF($K$7="","",INDEX(Grille_Notes!$Y${FIRST}:$Y${LAST},$K$7))', F(26, True, NAVY), FILL(CARD), CENTER, BOX, "0.0")
put(sc, "H11:I11", (f'=IF(ISNUMBER(H9),"Rang "&INDEX(Pipeline!$AL${FIRST}:$AL${LAST},$K$6)&" / "&COUNTIFS(Pipeline!$F${FIRST}:$F${LAST},'
                    f'INDEX(Pipeline!$F${FIRST}:$F${LAST},$K$6),Pipeline!$G${FIRST}:$G${LAST},$K$8,Pipeline!$T${FIRST}:$T${LAST},">=0")&" évalués","")'),
    F(9, True, LAV), FILL(CARD), CENTER, BOX)
put(sc, "H12:I12", f'=IF($K$7="","","Notes saisies : "&IF({GV("Z")}="","0/12",{GV("Z")}))', F(9, False, MUTED), FILL(CARD), CENTER, BOX)
for rr in range(8, 13): sc.row_dimensions[rr].height = 21
put(sc, "B13:I13", f'=IF($K$6="","VERDICT : —","VERDICT : "&IF({PV("U")}="","pas encore évalué",{PV("U")}))', F(13, True, NAVY), FILL("E8EEF4"), CENTER, BOX)
sc.row_dimensions[13].height = 30
verdict_cf(sc, "B13:I13", "$B$13")

section(sc, "B15:I15", "GRILLE D'ENTRETIEN — 12 QUESTIONS (posez-les dans l'ordre, notez de 1 à 5)", LAV)
header_row(sc, 16, 2, ["N°", "Bloc", "Type", "Question à poser", "Réponse attendue (indicateurs ✅)", "Signaux d'alerte (🚩)", "Note /5", "Évaluation"], NAVY, 26)
for n in range(1, 13):
    rr = 16 + n
    key = f'$K$8&"|Q{n}"'
    def QL(col):
        return f'IFERROR(INDEX(Questions!${col}$6:${col}$41,MATCH({key},Questions!$A$6:$A$41,0)),"")'
    calc(sc.cell(rr, 2, f"Q{n}"), center=True, bold=True)
    c = sc.cell(rr, 3, f"={QL('D')}"); c.font = F(9, True, NO_T if n <= 4 else NAVY); c.fill = FILL(CARD); c.border = BOX; c.alignment = WRAPC
    c = sc.cell(rr, 4, f"={QL('E')}"); c.font = F(9, False, MUTED); c.fill = FILL(CARD); c.border = BOX; c.alignment = CENTER
    first_q = f'=IF($K$8="","⬆ Sélectionnez un candidat (ou un poste pour une grille vierge).",{QL("F")})' if n == 1 else f"={QL('F')}"
    c = sc.cell(rr, 5, first_q); c.font = F(9, True); c.fill = FILL(CARD); c.border = BOX; c.alignment = WRAP
    c = sc.cell(rr, 6, f"={QL('G')}"); c.font = F(9); c.fill = FILL(CARD); c.border = BOX; c.alignment = WRAP
    c = sc.cell(rr, 7, f"={QL('H')}"); c.font = F(9, False, NO_T); c.fill = FILL(CARD); c.border = BOX; c.alignment = WRAP
    c = sc.cell(rr, 8, f'=IF($K$7="","",IF(INDEX(Grille_Notes!$G${FIRST}:$R${LAST},$K$7,{n})="","",INDEX(Grille_Notes!$G${FIRST}:$R${LAST},$K$7,{n})))')
    c.font = F(16, True, NAVY); c.fill = FILL(CARD); c.border = BOX; c.alignment = CENTER
    c = sc.cell(rr, 9, f'=IF(H{rr}="","☐ ☐ ☐ ☐ ☐",REPT("★",H{rr})&REPT("☆",5-H{rr}))')
    c.font = F(13, False, "E0A800"); c.fill = FILL(CARD); c.border = BOX; c.alignment = CENTER
    sc.row_dimensions[rr].height = 92
sc.conditional_formatting.add("H17:H28", ColorScaleRule(start_type="num", start_value=1, start_color="F4A6A0", mid_type="num", mid_value=3,
                                                        mid_color="FBE7A1", end_type="num", end_value=5, end_color="9ED8AE"))

section(sc, "B30:I30", "SYNTHÈSE PAR BLOC DE COMPÉTENCES (pondérée selon le poste)", LAV)
header_row(sc, 31, 2, ["Code", "Bloc", "Poids", "Score du bloc (moyenne /5 → %)"], NAVY, 24)
for i, b in enumerate(C.BLOCS):
    rr = 32 + i
    calc(sc.cell(rr, 2, b[0]), center=True, bold=True)
    c = sc.cell(rr, 3, ("🔒 " if b[3] == "Oui" else "") + b[2]); calc(c, bold=True)
    c = sc.cell(rr, 4, f'=IFERROR(INDEX({WEIGHTS},{i+1},MATCH($K$8,{WHEAD},0)),"")'); calc(c, "0%", center=True)
    c = sc.cell(rr, 5, f'=IF($K$7="","",IF(INDEX(Grille_Notes!$S${FIRST}:$X${LAST},$K$7,{i+1})="","",INDEX(Grille_Notes!$S${FIRST}:$X${LAST},$K$7,{i+1})/5))')
    c.font = F(10, True, NAVY); c.fill = FILL(CARD); c.border = BOX; c.alignment = Alignment(vertical="center"); c.number_format = "0%"
    sc.row_dimensions[rr].height = 22
sc.conditional_formatting.add("E32:E37", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=TEAL, showValue=True))
calc(sc.cell(38, 3, "Score global"), bold=True)
c = sc.cell(38, 5, "=H9"); calc(c, "0.0", bold=True)
put(sc, "B40:E40", "✅ POINTS FORTS", F(9, True, GO_T), FILL(GO_F), Alignment(vertical="center", indent=1))
put(sc, "B41:E42", f'={GV("AC")}', F(9), FILL(CARD), WRAP, BOX)
put(sc, "B43:E43", "🚩 POINTS DE VIGILANCE", F(9, True, NO_T), FILL(NO_F), Alignment(vertical="center", indent=1))
put(sc, "B44:E45", f'={GV("AD")}', F(9), FILL(CARD), WRAP, BOX)
for rr in (41, 42, 44, 45): sc.row_dimensions[rr].height = 24

radar = RadarChart(); radar.type = "marker"
radar.title = ctitle("Profil du candidat vs profil cible (/5)")
radar.add_data(Reference(ws_calc, min_col=8, min_row=1, max_row=7), titles_from_data=True)
radar.add_data(Reference(ws_calc, min_col=9, min_row=1, max_row=7), titles_from_data=True)
radar.set_categories(Reference(ws_calc, min_col=7, min_row=2, max_row=7))
for s_, col in zip(radar.series, (LAV, CORAL)):
    s_.graphicalProperties.line.solidFill = col; s_.graphicalProperties.line.width = 28000
    s_.marker.symbol = "circle"; s_.marker.size = 6
    s_.marker.graphicalProperties.solidFill = col; s_.marker.graphicalProperties.line.solidFill = col
radar.series[1].graphicalProperties.line.dashStyle = "dash"
radar.y_axis.scaling.min = 0; radar.y_axis.scaling.max = 5; radar.y_axis.majorUnit = 1
radar.width, radar.height = 17.5, 9.2
finish_chart(radar)
sc.add_chart(radar, "F31")

section(sc, "B47:I47", "ÉCHELLE DE NOTATION & RÈGLES DE DÉCISION", NAVY)
for i, (n, niv, desc) in enumerate(C.ECHELLE):
    rr = 48 + i
    calc(sc.cell(rr, 2, n), center=True, bold=True)
    put(sc, f"C{rr}:D{rr}", niv, F(9, True), FILL(CARD), Alignment(vertical="center", indent=1), BOX)
    put(sc, f"E{rr}:I{rr}", desc, F(9), FILL(CARD), Alignment(vertical="center", indent=1, wrap_text=True), BOX)
    sc.row_dimensions[rr].height = 20
put(sc, "B53:I53", '="Décision : ✅ GO si score ≥ "&SeuilGo&"  ·  🟠 À approfondir de "&SeuilApp&" à "&(SeuilGo-0.1)&"  ·  ❌ Non retenu sous "&SeuilApp&"  ·  ⛔ VETO si une note de 1 en Sécurité ou Bientraitance."',
    F(9, True, NAVY), FILL("E8EEF4"), Alignment(vertical="center", indent=1, wrap_text=True), BOX)
sc.row_dimensions[53].height = 28
sc.print_area = "B1:I53"
sc.page_setup.orientation = "landscape"; sc.page_setup.paperSize = sc.PAPERSIZE_A4
sc.page_setup.fitToWidth = 1; sc.page_setup.fitToHeight = 0
from openpyxl.worksheet.properties import PageSetupProperties
sc.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
sc.print_title_rows = "16:16"
sc.freeze_panes = "A8"

# ================================================================ COMPARATEUR
cp = ws_cmp
widths(cp, {"A": 2, "B": 7, "C": 26, "D": 16, "E": 9, "F": 34, "G": 9, "H": 9, "I": 9, "J": 9, "K": 9, "L": 9,
            "M": 28, "N": 12, "O": 24, "P": 12, "Q": 12, "R": 6, "S": 6})
bg(cp, 18, 50)
nav(cp, "B1", [("D1", "Scorecard", "🎯 Scorecard")])
banner(cp, "B2:Q2", "COMPARATEUR — classement des candidats évalués", "B3:Q3",
       "Choisissez la crèche et le poste : top 10 trié par score, détail par bloc, conformité, disponibilité.", LAV)
put(cp, "B5:C5", "Crèche ▸", F(10, True, NAVY), None, Alignment(vertical="center", horizontal="right"))
inp(cp["D5"], center=True); cp["D5"] = "Toutes"
dv_list(cp, "=lstCrechesToutes", "D5")
put(cp, "E5", "Poste ▸", F(10, True, NAVY), None, Alignment(vertical="center", horizontal="right"))
inp(cp["F5"], center=True); cp["F5"] = "CAP AEPE" if DEMO else "AP"
dv_list(cp, "=lstPostes", "F5")
cp.row_dimensions[5].height = 24
header_row(cp, 7, 2, ["Rang", "Candidat·e", "Crèche", "Score", "Verdict"] + [b[2] for b in C.BLOCS] +
           ["Étape", "Statut", "Dossier", "Disponible", "Prétentions"], NAVY, 34)
for kk in range(1, 11):
    rr = 7 + kk
    cp[f"R{rr}"] = (f'=IF($D$5="Toutes",IFERROR(MATCH("Toutes|"&$F$5&"|{kk}",Pipeline!$AJ${FIRST}:$AJ${LAST},0),""),'
                    f'IFERROR(MATCH($D$5&"|"&$F$5&"|{kk}",Pipeline!$AI${FIRST}:$AI${LAST},0),""))')
    cp[f"S{rr}"] = f'=IF(R{rr}="","",IFERROR(MATCH(INDEX(Pipeline!$A${FIRST}:$A${LAST},R{rr}),Grille_Notes!$A${FIRST}:$A${LAST},0),""))'
    def PI(col, txt=True):
        return f'IF($R{rr}="","",INDEX(Pipeline!${col}${FIRST}:${col}${LAST},$R{rr}){"&" + chr(34)*2 if txt else ""})'
    c = cp.cell(rr, 2, f'=IF(R{rr}="","",{kk})'); calc(c, center=True, bold=True)
    c = cp.cell(rr, 3, f'={PI("AE")}'); calc(c, bold=True, color=NAVY)
    c = cp.cell(rr, 4, f'={PI("F")}'); calc(c)
    c = cp.cell(rr, 5, f'={PI("T", False)}'); calc(c, "0.0", center=True, bold=True)
    c = cp.cell(rr, 6, f'={PI("U")}'); calc(c)
    for i in range(6):
        c = cp.cell(rr, 7 + i, f'=IF($S{rr}="","",INDEX(Grille_Notes!$S${FIRST}:$X${LAST},$S{rr},{i+1}))'); calc(c, "0.0", center=True)
    c = cp.cell(rr, 13, f'={PI("O")}'); calc(c)
    c = cp.cell(rr, 14, f'={PI("P")}'); calc(c, center=True)
    c = cp.cell(rr, 15, f'={PI("Z")}'); calc(c)
    c = cp.cell(rr, 16, f'=IF($R{rr}="","",IF(N(INDEX(Pipeline!$M${FIRST}:$M${LAST},$R{rr}))=0,"",INDEX(Pipeline!$M${FIRST}:$M${LAST},$R{rr})))'); calc(c, DATEF, center=True)
    c = cp.cell(rr, 17, f'=IF($R{rr}="","",IF(N(INDEX(Pipeline!$N${FIRST}:$N${LAST},$R{rr}))=0,"",INDEX(Pipeline!$N${FIRST}:$N${LAST},$R{rr})))'); calc(c, EURF, center=True)
    cp.row_dimensions[rr].height = 22
cp.column_dimensions["R"].hidden = True; cp.column_dimensions["S"].hidden = True
verdict_cf(cp, "F8:F17", "F8")
cp.conditional_formatting.add("E8:E17", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color=CORAL, showValue=True))
cp.conditional_formatting.add("G8:L17", ColorScaleRule(start_type="num", start_value=1, start_color="F4A6A0", mid_type="num", mid_value=3,
                                                       mid_color="FBE7A1", end_type="num", end_value=5, end_color="9ED8AE"))
cp.conditional_formatting.add("B8:Q8", FormulaRule(formula=['$B$8=1'], font=F(10, True, NAVY), border=Border(top=Side("medium", color=SUN), bottom=Side("medium", color=SUN))))
put(cp, "B19:Q19", '=IF(C8="","Aucun candidat évalué pour cette sélection : saisissez des notes dans la Grille.","🏆 Meilleur profil : "&C8&" — "&TEXT(E8,"0")&"/100 — "&F8)',
    F(11, True, NAVY), FILL("FFF6CC"), Alignment(vertical="center", indent=1), BOX)
cp.row_dimensions[19].height = 28
ch = bar3d("Score /100 des candidats classés", Reference(cp, min_col=3, min_row=8, max_row=17),
           [Reference(cp, min_col=5, min_row=7, max_row=17)], [CORAL], 24, 9, ymax=100)
cp.add_chart(ch, "B21")
cp.freeze_panes = "A8"

# ================================================================ ONBOARDING
ob = ws_onb
widths(ob, {"A": 2, "B": 12, "C": 50, **{L(4 + i): 15 for i in range(8)}})
bg(ob, 12, 30)
nav(ob, "B1", [("C1", "Plan", "📋 Plan")])
banner(ob, "B2:K2", "ONBOARDING — intégration des 8 recrues", "B3:K3",
       "Une colonne par poste du Plan (la recrue s'affiche dès l'embauche). Statuts : ✔ Fait · En cours · À faire · N/A", GREEN)
header_row(ob, 5, 2, ["Quand", "Tâche"] + [f'=Plan!A{7+i}&" · "&Plan!C{7+i}' for i in range(8)], NAVY, 24)
for i in range(8):
    c = ob.cell(6, 4 + i, f'=Plan!B{7+i}'); calc(c, center=True)
    c = ob.cell(7, 4 + i, f'=IF(Plan!Q{7+i}="","(à recruter)",Plan!Q{7+i})'); calc(c, center=True, bold=True, color=NAVY)
calc(ob.cell(6, 3, "Crèche"), bold=True); calc(ob.cell(7, 3, "Recrue"), bold=True)
for j, (quand, tache) in enumerate(C.ONBOARDING):
    rr = 8 + j
    calc(ob.cell(rr, 2, quand), center=True, bold=True)
    c = ob.cell(rr, 3, tache); c.font = F(9); c.fill = FILL(CARD); c.border = BOX; c.alignment = WRAPC
    for i in range(8):
        inp(ob.cell(rr, 4 + i, "À faire"), center=True)
    ob.row_dimensions[rr].height = 26 if len(tache) > 55 else 20
lr = 8 + len(C.ONBOARDING)
calc(ob.cell(lr, 3, "Avancement"), bold=True)
for i in range(8):
    col = L(4 + i)
    c = ob.cell(lr, 4 + i, f'=IFERROR(COUNTIF({col}8:{col}{lr-1},"✔ Fait")/(ROWS({col}8:{col}{lr-1})-COUNTIF({col}8:{col}{lr-1},"N/A")),0)')
    calc(c, "0%", center=True, bold=True)
ob.conditional_formatting.add(f"D{lr}:K{lr}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=GREEN, showValue=True))
dv_list(ob, "=lstOnb", f"D8:K{lr-1}")
ob.conditional_formatting.add(f"D8:K{lr-1}", FormulaRule(formula=['D8="✔ Fait"'], fill=FILL(GO_F), font=F(10, True, GO_T)))
ob.conditional_formatting.add(f"D8:K{lr-1}", FormulaRule(formula=['D8="En cours"'], fill=FILL(AP_F), font=F(10, False, AP_T)))
ob.conditional_formatting.add(f"D8:K{lr-1}", FormulaRule(formula=['D8="N/A"'], font=F(10, False, MUTED, True)))
ob.freeze_panes = "D8"

# ================================================================ DASHBOARD
d = ws_dash
widths(d, {"A": 2, **{L(c): 10.5 for c in range(2, 20)}, "T": 2})
bg(d, 20, 100)
nav(d, "B1", [("D1", "Pipeline", "👥 Pipeline"), ("F1", "Plan", "📋 Plan"), ("H1", "Comparateur", "⚖️ Comparateur")])
banner(d, "B2:S2", "TABLEAU DE BORD RECRUTEMENT — MICRO-CRÈCHES", "B3:S3",
       "Pilotage des 8 recrutements · tous les indicateurs et graphiques 3D se recalculent selon la crèche choisie", CORAL)
put(d, "B5:C5", "🏠 CRÈCHE ▸", F(10, True, NAVY), None, Alignment(vertical="center", horizontal="right"))
put(d, "D5:F5", "Toutes", F(11, True, INP_FONT), FILL(INP_FILL), CENTER, BOX)
dv_list(d, "=lstCrechesToutes", "D5", "Filtre", "Toutes, ou l'une des 2 crèches.")
put(d, "H5:I5", "Mis à jour le", F(9, False, MUTED), None, Alignment(vertical="center", horizontal="right"))
put(d, "J5:K5", "=TODAY()", F(10, True, NAVY), None, Alignment(vertical="center"), None, DATEF)
put(d, "L5:S5", "↺ Changez la crèche : KPI, entonnoir, postes, sources et radar se mettent à jour.", F(9, False, MUTED, True), None, Alignment(vertical="center"))
d.row_dimensions[5].height = 24

PF, PB = f"Pipeline!$F${FIRST}:$F${LAST}", f"Pipeline!$B${FIRST}:$B${LAST}"
PcolR = lambda col: f"Pipeline!${col}${FIRST}:${col}${LAST}"
cand = f'COUNTIFS({PF},critCreche,{PB},"?*")'
evals = f'COUNTIFS({PF},critCreche,{PcolR("T")},">=0")'
tiles = [
    ("POSTES POURVUS", 'SUMIFS(Plan!$O$7:$O$14,Plan!$B$7:$B$14,critCreche)&" / "&COUNTIFS(Plan!$B$7:$B$14,critCreche)',
     '"Couverture : "&ROUND(IFERROR(SUMIFS(Plan!$O$7:$O$14,Plan!$B$7:$B$14,critCreche)/COUNTIFS(Plan!$B$7:$B$14,critCreche),0)*100,0)&" %"', TEAL, None),
    ("CANDIDATS", cand, f'"dont "&COUNTIFS({PF},critCreche,{PcolR("P")},"En cours")&" en cours"', NAVY, "0"),
    ("ENTRETIENS ÉVALUÉS", evals, f'"Score moyen : "&IFERROR(ROUND(AVERAGEIFS({PcolR("T")},{PF},critCreche),0),"—")&" / 100"', LAV, "0"),
    ("TAUX DE « GO »", f'IFERROR(COUNTIFS({PF},critCreche,{PcolR("U")},"✅ GO*")/{evals},0)',
     f'COUNTIFS({PF},critCreche,{PcolR("U")},"✅ GO*")&" profil(s) GO"', GREEN, "0%"),
    ("DÉLAI MOYEN (JOURS)", f'IFERROR(ROUND(AVERAGEIFS({PcolR("AB")},{PF},critCreche,{PcolR("AF")},8),0),"—")', '"candidature → embauche"', OCHRE, "0"),
    ("ALERTES ACTIVES", f'COUNTIFS({PF},critCreche,{PcolR("AC")},"?*")', '"relances / dossiers à traiter"', CORAL, "0"),
]
for i, (lab, val, sub, col, fmt) in enumerate(tiles):
    c1, c3 = L(2 + 3 * i), L(4 + 3 * i)
    put(d, f"{c1}7:{c3}7", lab, F(9, True, "FFFFFF"), FILL(col), CENTER)
    put(d, f"{c1}8:{c3}9", "=" + val, F(24, True, col), FILL(CARD), CENTER, Border(left=THIN, right=THIN), fmt)
    put(d, f"{c1}10:{c3}10", "=" + sub, F(9, False, MUTED), FILL(CARD), CENTER, Border(left=THIN, right=THIN, bottom=THIN))
d.row_dimensions[7].height = 20; d.row_dimensions[8].height = 24; d.row_dimensions[9].height = 24; d.row_dimensions[10].height = 20
d.conditional_formatting.add("Q8:S9", CellIsRule(operator="equal", formula=["0"], font=F(24, True, GREEN)))

# --- tables de données (bas de page)
section(d, "B51:S51", "DONNÉES DES GRAPHIQUES & CONFORMITÉ RÉGLEMENTAIRE", NAVY)
header_row(d, 52, 2, ["Étape", "", "Atteint", "En cours", "Passage"], "4A6278", 24)
d.merge_cells("B52:C52")
for i, et in enumerate(C.ETAPES):
    rr = 53 + i
    put(d, f"B{rr}:C{rr}", et[4], F(9, True), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
    calc(d.cell(rr, 4, f'=COUNTIFS({PF},critCreche,{PcolR("AF")},">={i+1}")'), center=True)
    calc(d.cell(rr, 5, f'=COUNTIFS({PF},critCreche,{PcolR("AF")},{i+1},{PcolR("P")},"En cours")'), center=True)
    calc(d.cell(rr, 6, "" if i == 0 else f'=IFERROR(D{rr}/D{rr-1},"")'), "0%", center=True)
header_row(d, 52, 7, ["Poste", "À pourvoir", "Pourvus", "Candidats", "En cours", "Évalués", "Score moyen", "Meilleur"], "4A6278", 24)
for i, (code, *_r) in enumerate(C.POSTES):
    rr = 53 + i
    c = d.cell(rr, 7, code); c.font = F(9, True, "FFFFFF"); c.fill = FILL(POSTE_COL[code]); c.alignment = CENTER; c.border = BOX
    cp_ = f'{PF},critCreche,{PcolR("G")},G{rr}'
    vals = [f'=COUNTIFS(Plan!$B$7:$B$14,critCreche,Plan!$C$7:$C$14,G{rr})',
            f'=SUMIFS(Plan!$O$7:$O$14,Plan!$B$7:$B$14,critCreche,Plan!$C$7:$C$14,G{rr})',
            f'=COUNTIFS({cp_},{PB},"?*")', f'=COUNTIFS({cp_},{PcolR("P")},"En cours")',
            f'=COUNTIFS({cp_},{PcolR("T")},">=0")', f'=IFERROR(AVERAGEIFS({PcolR("T")},{cp_}),"—")',
            f'=IF(L{rr}=0,"—",_xlfn.MAXIFS({PcolR("T")},{cp_}))']
    for j, v in enumerate(vals):
        calc(d.cell(rr, 8 + j, v), "0" if j < 5 else "0.0", center=True)
header_row(d, 52, 16, ["Bloc", "EJE", "AP", "CAP AEPE"], "4A6278", 24)
for i, b in enumerate(C.BLOCS):
    rr = 53 + i
    c = d.cell(rr, 16, b[2]); calc(c, bold=True); c.font = F(8, True)
    for j, code in enumerate(("EJE", "AP", "CAP AEPE")):
        gcol = L(19 + i)
        calc(d.cell(rr, 17 + j, f'=IFERROR(AVERAGEIFS(Grille_Notes!${gcol}${FIRST}:${gcol}${LAST},Grille_Notes!$C${FIRST}:$C${LAST},critCreche,'
                                  f'Grille_Notes!$D${FIRST}:$D${LAST},"{code}"),0)'), "0.0", center=True)
header_row(d, 63, 2, ["Source", "", "Candidats", "Embauches", "Transfo."], "4A6278", 24)
d.merge_cells("B63:C63")
for i, src in enumerate(C.SOURCES):
    rr = 64 + i
    put(d, f"B{rr}:C{rr}", f"=Parametres!$C${52+i}", F(8, True), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
    calc(d.cell(rr, 4, f'=COUNTIFS({PF},critCreche,{PcolR("H")},B{rr},{PB},"?*")'), center=True)
    calc(d.cell(rr, 5, f'=COUNTIFS({PF},critCreche,{PcolR("H")},B{rr},{PcolR("AF")},8)'), center=True)
    calc(d.cell(rr, 6, f'=IF(D{rr}=0,"",E{rr}/D{rr})'), "0%", center=True)
header_row(d, 63, 7, ["Indicateur", "", "=Parametres!$C$6", "=Parametres!$C$7"], "4A6278", 30)
d.merge_cells("G63:H63")
comp = [("Candidats", lambda nm: f'=COUNTIFS({PF},{nm},{PB},"?*")'),
        ("Entretiens évalués", lambda nm: f'=COUNTIFS({PF},{nm},{PcolR("T")},">=0")'),
        ("Embauches", lambda nm: f'=COUNTIFS({PF},{nm},{PcolR("AF")},8)'),
        ("Postes ouverts", lambda nm: f'=COUNTIFS(Plan!$B$7:$B$14,{nm})')]
for i, (lab, fn) in enumerate(comp):
    rr = 64 + i
    put(d, f"G{rr}:H{rr}", lab, F(9, True), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
    calc(d.cell(rr, 9, fn("Parametres!$C$6")), center=True)
    calc(d.cell(rr, 10, fn("Parametres!$C$7")), center=True)

# Conformité
section(d, "L63:S63", "CONFORMITÉ DÉCRET 2025-304 (au 01/09/2026)", GREEN)
header_row(d, 64, 12, ["Crèche", "", "Cat. 1 prévu", "Cat. 1 réel", "≥ 1 diplômé cat. 1", "Direction ≥ 0,5 ETP", "Statut", ""], "4A6278", 30)
d.merge_cells("L64:M64"); d.merge_cells("R64:S64")
for i in range(2):
    rr = 65 + i; pr_ = 6 + i
    nm = f"Parametres!$C${pr_}"
    put(d, f"L{rr}:M{rr}", f"={nm}", F(9, True), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
    prevu = (f'(COUNTIFS(Plan!$B$7:$B$14,{nm},Plan!$E$7:$E$14,1)+Parametres!$J${pr_})/'
             f'(COUNTIFS(Plan!$B$7:$B$14,{nm})+Parametres!$J${pr_}+Parametres!$K${pr_})')
    cat1 = (f'(COUNTIFS({PF},{nm},{PcolR("AF")},8,{PcolR("G")},"EJE")+COUNTIFS({PF},{nm},{PcolR("AF")},8,{PcolR("G")},"AP")+Parametres!$J${pr_})')
    tot = f'(COUNTIFS({PF},{nm},{PcolR("AF")},8)+Parametres!$J${pr_}+Parametres!$K${pr_})'
    calc(d.cell(rr, 14, f"=IFERROR({prevu},0)"), "0%", center=True)
    calc(d.cell(rr, 15, f'=IF({tot}=0,"—",{cat1}/{tot})'), "0%", center=True)
    calc(d.cell(rr, 16, f'=IF({cat1}>0,"✅ Oui","❌ Non")'), center=True)
    c = d.cell(rr, 17, f'=IF(N(Parametres!$I${pr_})>=0.5,"✅ "&Parametres!$H${pr_},"⚠️ À désigner")'); calc(c, center=True); c.font = F(8, True)
    put(d, f"R{rr}:S{rr}", f'=IF(AND(ISNUMBER(O{rr}),N(O{rr})>=0.4,LEFT(P{rr},1)="✅",LEFT(Q{rr},1)="✅"),"✅ Conforme","⚠️ À lever")',
        F(10, True), FILL(CALC_FILL), CENTER, BOX)
    d.row_dimensions[rr].height = 22
d.conditional_formatting.add("R65:S66", FormulaRule(formula=['LEFT($R65,1)="✅"'], fill=FILL(GO_F), font=F(10, True, GO_T)))
d.conditional_formatting.add("R65:S66", FormulaRule(formula=['LEFT($R65,1)="⚠"'], fill=FILL(AP_F), font=F(10, True, AP_T)))
d.conditional_formatting.add("N65:O66", CellIsRule(operator="lessThan", formula=["0.4"], font=F(10, True, NO_T)))
put(d, "L67:S68", "Seuil : ≥ 40 % de professionnels de catégorie 1 et au moins 1 diplômé d'État dans l'équipe ; direction ≥ 0,5 ETP "
    "(Paramètres › Direction / ETP). « Réel » = embauchés du Pipeline + effectif déjà en poste (Paramètres).",
    F(8, False, MUTED, True), None, Alignment(vertical="top", wrap_text=True))

# Échéances
section(d, "B77:S77", "ÉCHÉANCES DES 8 POSTES", OCHRE)
header_row(d, 78, 2, ["ID", "Crèche", "", "Poste", "Prise de poste", "J restants", "Statut", "", "Recrue", "", "", "Score", "Alerte", "", ""], "4A6278", 24)
for ref in ("C78:D78", "H78:I78", "J78:L78", "N78:P78"):
    d.merge_cells(ref)
for i in range(8):
    rr = 79 + i; pr_ = 7 + i
    calc(d.cell(rr, 2, f"=Plan!A{pr_}"), center=True, bold=True)
    put(d, f"C{rr}:D{rr}", f"=Plan!B{pr_}", F(9), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
    calc(d.cell(rr, 5, f"=Plan!C{pr_}"), center=True, bold=True)
    calc(d.cell(rr, 6, f"=Plan!I{pr_}"), DATEF, center=True)
    calc(d.cell(rr, 7, f"=Plan!R{pr_}"), "0", center=True)
    put(d, f"H{rr}:I{rr}", f"=Plan!P{pr_}", F(9, True), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
    put(d, f"J{rr}:L{rr}", f'=IF(Plan!Q{pr_}="","—",Plan!Q{pr_})', F(9, True, NAVY), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
    calc(d.cell(rr, 13, f"=Plan!M{pr_}"), "0.0", center=True)
    put(d, f"N{rr}:P{rr}", f"=Plan!S{pr_}", F(9), FILL(CALC_FILL), Alignment(vertical="center"), BOX)
for poste, col in POSTE_COL.items():
    d.conditional_formatting.add("E79:E86", FormulaRule(formula=[f'E79="{poste}"'], fill=FILL(col), font=F(10, True, "FFFFFF")))
d.conditional_formatting.add("H79:I86", FormulaRule(formula=['LEFT($H79,1)="✅"'], fill=FILL(GO_F), font=F(10, True, GO_T)))
d.conditional_formatting.add("N79:P86", FormulaRule(formula=['ISNUMBER(SEARCH("dépassée",$N79))'], fill=FILL(NO_F), font=F(10, True, NO_T)))
d.conditional_formatting.add("N79:P86", FormulaRule(formula=['ISNUMBER(SEARCH("30 jours",$N79))'], fill=FILL(AP_F), font=F(10, True, AP_T)))

# --- graphiques
section(d, "B12:J12", "ENTONNOIR DE RECRUTEMENT", NAVY)
section(d, "K12:S12", "POSTES : À POURVOIR · POURVUS · CANDIDATS", NAVY)
section(d, "B31:G31", "SOURCES DE CANDIDATURES", NAVY)
section(d, "H31:M31", "PROFIL MOYEN PAR BLOC (/5)", NAVY)
section(d, "N31:S31", "COMPARATIF DES 2 CRÈCHES", NAVY)

ch1 = bar3d("Candidats ayant atteint chaque étape", Reference(d, min_col=2, min_row=53, max_row=60),
            [Reference(d, min_col=4, min_row=52, max_row=60)], [TEAL], 17.6, 8.6)
d.add_chart(ch1, "B13")
ch2 = bar3d("Avancement par poste", Reference(d, min_col=7, min_row=53, max_row=55),
            [Reference(d, min_col=8, min_row=52, max_row=55), Reference(d, min_col=9, min_row=52, max_row=55),
             Reference(d, min_col=10, min_row=52, max_row=55)], [OCHRE, GREEN, SKY], 17.6, 8.6)
d.add_chart(ch2, "K13")
pie = bar3d("Candidatures par source", Reference(d, min_col=2, min_row=64, max_row=74),
            [Reference(d, min_col=4, min_row=63, max_row=74), Reference(d, min_col=5, min_row=63, max_row=74)],
            [TEAL, CORAL], 11.6, 9, kind="bar")
d.add_chart(pie, "B32")
rad = RadarChart(); rad.type = "marker"
rad.title = ctitle("Moyenne des candidats évalués")
for j in range(3):
    rad.add_data(Reference(d, min_col=17 + j, min_row=52, max_row=58), titles_from_data=True)
rad.set_categories(Reference(d, min_col=16, min_row=53, max_row=58))
for s_, col in zip(rad.series, (LAV, TEAL, CORAL)):
    s_.graphicalProperties.line.solidFill = col; s_.graphicalProperties.line.width = 25000
    s_.marker.symbol = "circle"; s_.marker.size = 5
    s_.marker.graphicalProperties.solidFill = col; s_.marker.graphicalProperties.line.solidFill = col
rad.y_axis.scaling.min = 0; rad.y_axis.scaling.max = 5; rad.y_axis.majorUnit = 1
rad.width, rad.height = 11.6, 9
finish_chart(rad)
d.add_chart(rad, "H32")
ch5 = bar3d("Crèche 1 vs Crèche 2", Reference(d, min_col=7, min_row=64, max_row=67),
            [Reference(d, min_col=9, min_row=63, max_row=67), Reference(d, min_col=10, min_row=63, max_row=67)],
            [NAVY, CORAL], 11.6, 9)
d.add_chart(ch5, "N32")
d.freeze_panes = "A6"

# ================================================================ ACCUEIL
a = ws_acc
widths(a, {"A": 2, **{L(c): 12.5 for c in range(2, 14)}, "N": 2})
bg(a, 14, 60)
put(a, "B2:M2", "RECRUTEMENT · 2 MICRO-CRÈCHES", F(22, True, "FFFFFF"), FILL(NAVY), Alignment(vertical="center", indent=1))
a.row_dimensions[2].height = 44
put(a, "B3:M3", "Workflow complet · 8 postes : 2 EJE · 2 auxiliaires de puériculture · 4 CAP AEPE · scorecard interactive · dashboard 3D",
    F(11, False, "FFFFFF", True), FILL(NAVY), Alignment(vertical="center", indent=1))
a.row_dimensions[3].height = 24
section(a, "B5:M5", "NAVIGATION", CORAL)
tiles_nav = [("📊 Dashboard", "Dashboard", "KPI, entonnoir, sources, conformité", CORAL),
             ("📋 Plan de recrutement", "Plan", "8 postes, statuts, échéances", OCHRE),
             ("👥 Pipeline candidats", "Pipeline", "1 ligne = 1 candidature", TEAL),
             ("🎯 Scorecard", "Scorecard", "Fiche d'entretien interactive", LAV),
             ("✏️ Grille de notation", "Grille_Notes", "Saisie des notes 1 à 5", LAV),
             ("⚖️ Comparateur", "Comparateur", "Classement par crèche / poste", LAV),
             ("🚀 Onboarding", "Onboarding", "Intégration des recrues", GREEN),
             ("📄 Fiches de poste", "Fiches_Poste", "EJE · AP · CAP AEPE + annonces", SKY),
             ("❓ Banque de questions", "Questions", "36 questions + réponses attendues", SKY),
             ("☎ Scripts & modèles", "Scripts_Modeles", "Pré-qualif tél. + e-mails", SKY),
             ("⚙ Paramètres", "Parametres", "Crèches, pondérations, listes", MUTED)]
for i, (lab, sh, desc, col) in enumerate(tiles_nav):
    row = 6 + (i // 4) * 3; c1 = 2 + (i % 4) * 3
    ref1 = f"{L(c1)}{row}:{L(c1+2)}{row}"; ref2 = f"{L(c1)}{row+1}:{L(c1+2)}{row+1}"
    put(a, ref1, f'=HYPERLINK("#\'{sh}\'!A1","{lab}")', F(11, True, "FFFFFF"), FILL(col), CENTER)
    put(a, ref2, desc, F(9, False, MUTED), FILL(CARD), CENTER, Border(left=THIN, right=THIN, bottom=THIN))
    a.row_dimensions[row].height = 28; a.row_dimensions[row + 1].height = 20
section(a, "B15:M15", "LE WORKFLOW EN 8 ÉTAPES (délai max avant alerte de relance)", CORAL)
header_row(a, 16, 2, ["N°", "Étape", "", "", "Délai max", "Responsable", "", "Livrable / outil", "", "", "", ""], NAVY, 22)
for ref in ("C16:E16", "G16:H16", "I16:M16"):
    a.merge_cells(ref)
for i, (lib, sla, resp, liv, _) in enumerate(C.ETAPES):
    rr = 17 + i
    c = a.cell(rr, 2, i + 1); c.font = F(12, True, "FFFFFF"); c.fill = FILL([TEAL, TEAL, TEAL, LAV, LAV, GREEN, OCHRE, CORAL][i]); c.alignment = CENTER
    put(a, f"C{rr}:E{rr}", lib[3:], F(10, True, NAVY), FILL(CARD), Alignment(vertical="center", indent=1), BOX)
    put(a, f"F{rr}", f'=IF(Parametres!D{17+i}=0,"—",Parametres!D{17+i}&" j")', F(10), FILL(CARD), CENTER, BOX)
    put(a, f"G{rr}:H{rr}", f"=Parametres!E{17+i}", F(9), FILL(CARD), Alignment(vertical="center", indent=1, wrap_text=True), BOX)
    put(a, f"I{rr}:M{rr}", f"=Parametres!F{17+i}", F(9), FILL(CARD), Alignment(vertical="center", indent=1, wrap_text=True), BOX)
    a.row_dimensions[rr].height = 24
section(a, "B26:M26", "DÉMARRER EN 5 MINUTES", CORAL)
steps = [
    "1 · Paramètres : renommez les 2 crèches (avant toute saisie), ville, date de prise de poste cible, direction et ETP, effectif déjà en poste.",
    "2 · Fiches de poste : copiez les annonces prêtes à publier (Indeed, France Travail, Hellowork, écoles…).",
    "3 · Pipeline : saisissez chaque candidature dans les cellules bleues ; faites avancer l'étape et le statut. Les alertes 🔔 vous disent qui relancer.",
    "4 · Entretien : Scorecard › choisissez le candidat › imprimez la fiche. Après l'entretien, cliquez « Saisir ses notes » et notez 1 à 5 dans la Grille.",
    "5 · Décision : Comparateur + Dashboard ; vérifiez le dossier de conformité (attestation d'honorabilité !) avant toute proposition, puis Onboarding.",
]
if not DEMO:
    steps.append("ℹ La ligne C001 est un EXEMPLE : effacez ses cellules bleues dans Pipeline et Grille_Notes avant de démarrer.")
else:
    steps.append("ℹ Version DÉMO : 22 candidats fictifs pour voir le classeur « vivant ». Utilisez la version vierge pour vos vrais recrutements.")
for i, t in enumerate(steps):
    rr = 27 + i
    put(a, f"B{rr}:M{rr}", t, F(10, i == 5, NAVY if i < 5 else CORAL), FILL(CARD), Alignment(vertical="center", indent=1, wrap_text=True), BOTTOM)
    a.row_dimensions[rr].height = 22
section(a, "B34:M34", "RÈGLES DE DÉCISION DE LA SCORECARD", CORAL)
regles = [
    '="✅ GO : score pondéré ≥ "&SeuilGo&"/100   ·   🟠 À approfondir : "&SeuilApp&" à "&(SeuilGo-0.1)&"   ·   ❌ Non retenu : < "&SeuilApp',
    "⛔ VETO automatique : une note de 1 sur une question Sécurité (Q1-Q2) ou Bientraitance (Q3-Q4) écarte le candidat, quel que soit son score.",
    "12 questions par poste, 2 par bloc ; pondérations différentes selon le poste (ex. Sécurité 25 % pour l'AP, Éveil 25 % pour l'EJE) — modifiables dans Paramètres.",
    "Aucune proposition sans dossier « ✅ Complet » : diplôme vérifié, attestation d'honorabilité reçue, aptitude, références.",
]
for i, t in enumerate(regles):
    rr = 35 + i
    put(a, f"B{rr}:M{rr}", t, F(10, i == 1, NO_T if i == 1 else TXT), FILL(CARD), Alignment(vertical="center", indent=1, wrap_text=True), BOTTOM)
    a.row_dimensions[rr].height = 22
section(a, "B40:M40", "CADRE RÉGLEMENTAIRE INTÉGRÉ (micro-crèches, 2026)", GREEN)
for i, t in enumerate(C.REGLEMENTAIRE):
    rr = 41 + i
    put(a, f"B{rr}:M{rr}", t, F(9, i in (0, 4, 5), TXT), FILL(CARD), Alignment(vertical="center", indent=1, wrap_text=True), BOTTOM)
    a.row_dimensions[rr].height = 30 if len(t) > 150 else 20
put(a, "B47:M47", "Sources : décret n° 2025-304 du 1er avril 2025 (réponses ministérielles AN, JO 2026) · solidarites.gouv.fr (attestation d'honorabilité, "
    "fiches métiers EJE / AP / AEPE) · ONISEP · offres France Travail Île-de-France. À valider avec votre PMI et votre conseil juridique / social.",
    F(8, False, MUTED, True), None, Alignment(vertical="center", indent=1, wrap_text=True))
a.row_dimensions[47].height = 30
section(a, "B49:M49", "LÉGENDE", CORAL)
inp(a["B50"]); a["B50"] = "Saisie"
put(a, "C50:F50", "Cellule à remplir (texte bleu, fond bleu clair)", F(9), None, Alignment(vertical="center", indent=1))
calc(a["G50"]); a["G50"] = "Calcul"
put(a, "H50:M50", "Calcul automatique — ne pas modifier ·  ▼ = liste déroulante", F(9), None, Alignment(vertical="center", indent=1))

# ================================================================ EXEMPLE / DÉMO
today = dt.date(2026, 11, 25)
D = lambda m, dd: dt.date(2026, m, dd)
C1, C2 = "Micro-crèche 1", "Micro-crèche 2"
FULL_OK = ("Oui", "Reçue", "OK", "OK")
demo = [
    # nom, prénom, crèche, poste, source, date cand, diplôme, exp, dispo, prétentions, étape#, statut, dernière action, prochaine action, note tél, conformité, date embauche, comm, notes, évaluateur, forts, vigilance
    ("DUMONT", "Clara", C1, "EJE", "Indeed", D(10, 21), "DE Éducateur de jeunes enfants", 4, D(1, 5) if False else dt.date(2027, 1, 4), 2600, 8, "Embauché(e)", D(11, 17), "Lancer l'onboarding", 5, FULL_OK, D(11, 17), "",
     [5, 4, 5, 5, 4, 4, 5, 4, 4, 5, 4, 4], "Direction + gestionnaire", "Projet pédagogique très structuré, posture d'animatrice d'équipe", "Peu d'expérience de reprise de structure"),
    ("NASRI", "Yasmine", C1, "EJE", "École / organisme de formation", D(10, 28), "DE Éducateur de jeunes enfants", 1, dt.date(2027, 1, 4), 2300, 4, "Vivier", D(11, 10), "", 4, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "Profil à recontacter pour Crèche 2",
     [4, 3, 4, 3, 3, 4, 3, 4, 3, 3, 3, 4], "Direction", "Bonne connaissance du développement", "Manque d'assurance face aux familles"),
    ("LEROY", "Mathilde", C1, "EJE", "LinkedIn", D(11, 2), "DE Éducateur de jeunes enfants", 6, dt.date(2026, 12, 15), 2900, 4, "Refusé(e)", D(11, 14), "", 4, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "",
     [3, 4, 1, 3, 4, 4, 3, 3, 3, 4, 4, 3], "Direction", "Expérience solide", "Banalise les menaces de punition (VEO)"),
    ("TRAORÉ", "Awa", C1, "AP", "France Travail", D(10, 25), "DE Auxiliaire de puériculture", 5, dt.date(2027, 1, 4), 2350, 7, "En cours", D(11, 21), "Attendre retour signature", 5, FULL_OK, None, "",
     [5, 5, 4, 5, 4, 4, 4, 4, 5, 4, 4, 5], "Direction + EJE", "Excellente maîtrise des protocoles santé / PAI", "Prétentions en haut de fourchette"),
    ("MOREAU", "Julie", C1, "AP", "Indeed", D(11, 5), "DE Auxiliaire de puériculture", 2, dt.date(2027, 1, 18), 2100, 4, "En cours", D(11, 20), "Finir l'entretien (Q9-Q12)", 4, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "Entretien interrompu",
     [4, 3, 4, 4, 3, 3, 4, 3, None, None, None, None], "EJE", "Douceur, verbalisation des soins", "À compléter"),
    ("BENALI", "Sonia", C1, "AP", "Cooptation", D(11, 3), "DE Auxiliaire de puériculture", 8, dt.date(2027, 2, 1), 2500, 3, "Refusé(e)", D(11, 6), "", 2, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "Horaires de fermeture incompatibles", None, "", "", ""),
    ("PETIT", "Laura", C1, "AP", "Hellowork", D(11, 18), "DE Infirmier / Puériculteur", 3, dt.date(2027, 1, 4), 2400, 2, "En cours", D(11, 19), "Appeler pour pré-qualif", None, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "", None, "", "", ""),
    ("MARTINS", "Ana", C1, "CAP AEPE", "Indeed", D(10, 20), "CAP AEPE", 3, dt.date(2027, 1, 4), 1950, 8, "Embauché(e)", D(11, 14), "Onboarding", 4, FULL_OK, D(11, 14), "",
     [4, 4, 4, 5, 4, 4, 4, 3, 4, 4, 4, 4], "Direction", "Très à l'aise avec les tout-petits", "HACCP à consolider"),
    ("KONÉ", "Mariam", C1, "CAP AEPE", "France Travail", D(10, 30), "CAP AEPE", 2, dt.date(2027, 1, 4), 1900, 6, "En cours", D(11, 22), "Relancer attestation d'honorabilité", 4, ("Oui", "Demandée", "En attente", "OK"), None, "",
     [4, 4, 4, 4, 4, 3, 4, 4, 4, 4, 4, 4], "Direction + EJE", "Fiable, réponses sécurité impeccables", "Timide en réunion"),
    ("GIRARD", "Emma", C1, "CAP AEPE", "Candidature spontanée", D(11, 6), "CAP AEPE", 1, dt.date(2027, 1, 4), 1850, 4, "En cours", D(11, 24), "Organiser rencontre équipe", 3, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "",
     [3, 3, 3, 4, 3, 3, 3, 4, 3, 3, 4, 4], "EJE", "Motivée, projet de VAE AP", "Connaissances éveil à renforcer"),
    ("HAMIDI", "Nour", C1, "CAP AEPE", "Facebook / groupes locaux", D(11, 8), "Bac pro ASSP", 0.5, dt.date(2026, 12, 15), 1850, 4, "En cours", D(11, 18), "", 3, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "",
     [3, 2, 3, 3, 2, 3, 3, 3, 2, 3, 3, 3], "Direction", "Disponible rapidement", "Réponses imprécises sur l'urgence"),
    ("LAMBERT", "Chloé", C1, "CAP AEPE", "Indeed", D(11, 1), "CAP AEPE", 4, dt.date(2027, 1, 4), 2000, 5, "Désisté(e)", D(11, 16), "", 4, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "A accepté une autre offre",
     [4, 4, 4, 4, 3, 4, 4, 4, 4, 4, 3, 4], "Direction", "", ""),
    ("SY", "Aïssatou", C1, "CAP AEPE", "École / organisme de formation", D(11, 24), "CAP AEPE", 1, dt.date(2027, 1, 4), 1850, 1, "En cours", D(11, 24), "Présélection CV", None, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "", None, "", "", ""),
    ("ROUSSEAU", "Inès", C2, "EJE", "LinkedIn", D(11, 4), "DE Éducateur de jeunes enfants", 3, dt.date(2027, 1, 18), 2500, 4, "En cours", D(11, 23), "Rencontre équipe le 28/11", 5, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "",
     [4, 4, 5, 4, 5, 4, 4, 5, 4, 4, 4, 4], "Direction + gestionnaire", "Lien aux familles remarquable", "Disponible mi-janvier seulement"),
    ("FERREIRA", "Lucie", C2, "EJE", "Indeed", D(11, 16), "DE Éducateur de jeunes enfants", 2, dt.date(2027, 2, 1), 2400, 2, "En cours", D(11, 23), "Planifier appel", None, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "", None, "", "", ""),
    ("DIALLO", "Fatoumata", C2, "AP", "France Travail", D(11, 9), "DE Auxiliaire de puériculture", 6, dt.date(2027, 1, 4), 2300, 4, "En cours", D(11, 22), "Contacter références", 4, ("Oui", "Non demandée", "En attente", "Non contactées"), None, "",
     [5, 4, 4, 4, 4, 5, 4, 4, 4, 4, 5, 4], "Direction", "Grande sécurité affective, calme", ""),
    ("GARNIER", "Pauline", C2, "AP", "Agence d'intérim", D(11, 12), "DE Auxiliaire de puériculture", 1, dt.date(2027, 1, 4), 2200, 3, "En cours", D(11, 15), "Proposer entretien", 3, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "", None, "", "", ""),
    ("CHEVALIER", "Manon", C2, "AP", "Indeed", D(11, 22), "Titre pro Intervenant éducatif petite enfance", 2, dt.date(2027, 1, 4), 2100, 1, "En cours", D(11, 22), "Vérifier équivalence diplôme", None, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "", None, "", "", ""),
    ("N'GUESSAN", "Grace", C2, "CAP AEPE", "Cooptation", D(10, 24), "CAP AEPE", 5, dt.date(2027, 1, 4), 2000, 8, "Embauché(e)", D(11, 12), "Onboarding", 5, FULL_OK, D(11, 12), "",
     [5, 4, 5, 4, 4, 4, 4, 4, 5, 4, 4, 4], "Direction", "Expérience solide, leader naturel", ""),
    ("BLANC", "Océane", C2, "CAP AEPE", "Indeed", D(11, 2), "CAP AEPE", 2, dt.date(2027, 1, 4), 1900, 7, "En cours", D(11, 20), "Attendre attestation d'honorabilité", 4, ("Oui", "Demandée", "OK", "OK"), None, "",
     [4, 4, 4, 4, 3, 4, 4, 3, 4, 4, 4, 4], "Direction + EJE", "Sérieuse, ponctuelle", ""),
    ("MEZIANE", "Sarah", C2, "CAP AEPE", "Hellowork", D(11, 14), "CAP AEPE", 1, dt.date(2027, 1, 4), 1850, 3, "En cours", D(11, 21), "Entretien tél. prévu", None, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "", None, "", "", ""),
    ("LOPES", "Camille", C2, "CAP AEPE", "Indeed", D(11, 7), "Autre / en cours", 0, dt.date(2026, 12, 1), 1800, 2, "Refusé(e)", D(11, 9), "", None, ("En attente", "Non demandée", "En attente", "Non contactées"), None, "Diplôme non obtenu", None, "", "", ""),
]
if not DEMO:
    demo = [("EXEMPLE", "Candidat", C1, "AP", "Indeed", D(11, 20), "DE Auxiliaire de puériculture", 3, dt.date(2027, 1, 4), 2200, 4,
             "En cours", D(11, 24), "Organiser la rencontre équipe", 4, ("Oui", "Non demandée", "En attente", "Non contactées"), None,
             "Ligne d'exemple : à effacer",
             [4, 4, 4, 5, 3, 4, 4, 3, 4, 4, 4, 4], "Direction", "Exemple : protocoles santé maîtrisés", "Exemple : prétentions à négocier")]

for i, row in enumerate(demo):
    r = FIRST + i
    (nom, pre, cr, po, src, dcand, dip, exp, dispo, pret, et, stat, dla, nxt, ntel, conf, demb, comm, notes, evalr, forts, vig) = row
    tel = f"06 00 00 {i // 10:02d} {i % 100:02d}"
    import unicodedata
    strip = lambda t: unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode()
    mail = strip(f"{pre.lower()}.{nom.lower()}@exemple.fr").replace("'", "").replace(" ", "")
    for col, v in zip("BCDEFGHIJ", (nom, pre, tel, mail, cr, po, src, dcand, dip)):
        pp[f"{col}{r}"] = v
    pp[f"L{r}"] = exp; pp[f"M{r}"] = dispo; pp[f"N{r}"] = pret
    pp[f"O{r}"] = C.ETAPES[et - 1][0]; pp[f"P{r}"] = stat; pp[f"Q{r}"] = dla; pp[f"R{r}"] = nxt
    if ntel: pp[f"S{r}"] = ntel
    for col, v in zip("VWXY", conf):
        pp[f"{col}{r}"] = v
    if demb: pp[f"AA{r}"] = demb
    if comm: pp[f"AD{r}"] = comm
    if notes:
        g[f"E{r}"] = evalr; g[f"F{r}"] = dla if et >= 4 else None
        for j, n_ in enumerate(notes):
            if n_ is not None:
                g.cell(r, 7 + j, n_)
        g[f"AC{r}"] = forts; g[f"AD{r}"] = vig
if DEMO:
    sc["D5"] = "C004 · TRAORÉ Awa · AP · Micro-crèche 1"

from openpyxl.worksheet.properties import PageSetupProperties as _PSP
for _ws, _area, _orient in ((ws_dash, "A1:T87", "portrait"), (ws_acc, "A1:N50", "portrait"), (ws_cmp, "A1:Q40", "landscape"),
                            (ws_plan, "A1:S17", "landscape"), (ws_onb, "A1:K24", "landscape"), (ws_fp, "A1:F32", "landscape"),
                            (ws_q, "B1:I41", "landscape"), (ws_scr, "A1:F31", "portrait"), (ws_par, "A1:K62", "landscape")):
    _ws.print_area = _area
    _ws.page_setup.orientation = _orient; _ws.page_setup.paperSize = _ws.PAPERSIZE_A4
    _ws.page_setup.fitToWidth = 1; _ws.page_setup.fitToHeight = 0
    _ws.sheet_properties.pageSetUpPr = _PSP(fitToPage=True)
ws_q.print_title_rows = "5:5"
ws_pipe.page_setup.orientation = "landscape"; ws_pipe.page_setup.fitToWidth = 1; ws_pipe.page_setup.fitToHeight = 0
ws_pipe.sheet_properties.pageSetUpPr = _PSP(fitToPage=True); ws_pipe.print_title_rows = "6:6"
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("OK", OUT)
