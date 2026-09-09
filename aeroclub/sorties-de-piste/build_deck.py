# -*- coding: utf-8 -*-
"""Construit le support « Sorties de piste longitudinales » au format FV Vitesse."""
import copy, os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

SK = "/root/.claude/skills/synced/d089a006-0094-45a8-8dba-4638592a42e4_26bca67e-877d-461d-8e9e-dd2105052c40/pptx-fv/assets"
TEMPLATE = f"{SK}/template_FV_aviation.pptx"
LOGO_BB = f"{SK}/logo_fv_blanc_bleu.png"
LOGO_BLANC = f"{SK}/logo_fv_blanc.png"
TAMPON = f"{SK}/tampon_fv_gris.png"
AR_LOGO, AR_TAMPON = 870 / 360, 720 / 360

NAVY = RGBColor(0x1A, 0x3A, 0x5C)
SKY = RGBColor(0x4A, 0x9E, 0xD6)
LSKY = RGBColor(0xD2, 0xEB, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY = RGBColor(0xF5, 0xF7, 0xFA)
BGBLUE = RGBColor(0xF4, 0xF8, 0xFD)
GRAY = RGBColor(0x5A, 0x6A, 0x7E)
GRAY2 = RGBColor(0x6B, 0x7A, 0x8C)
GREEN = RGBColor(0x2E, 0x7D, 0x4F)
AMBER = RGBColor(0xD6, 0x89, 0x10)
RED = RGBColor(0xC0, 0x39, 0x2B)
LGREEN = RGBColor(0xE3, 0xF1, 0xE8)
LAMBER = RGBColor(0xFB, 0xF0, 0xDC)
LRED = RGBColor(0xF7, 0xE1, 0xDE)

FOOTER_TXT = "Formation PPL · LAPL  —  Frédéric Vincenot · FI"

prs = Presentation(TEMPLATE)
# Purge des slides du template (on garde le master et le thème)
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst):
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)
BLANK = prs.slide_layouts[6]
page = [0]


# ─── helpers ────────────────────────────────────────────────────────────────
def bg(s, color):
    f = s.background.fill
    f.solid()
    f.fore_color.rgb = color


def rect(s, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid()
    r.fill.fore_color.rgb = color
    r.line.fill.background()
    r.shadow.inherit = False
    return r


def txt(s, text, x, y, w, h, size, bold=False, color=NAVY, align=PP_ALIGN.LEFT,
        font="Calibri", italic=False, anchor=None, wrap=True):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    if anchor:
        tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        f = r.font
        f.size, f.bold, f.italic, f.name = Pt(size), bold, italic, font
        f.color.rgb = color
    return tb


def bullets(s, items, x, y, w, h, size=20, sub_size=15, color=NAVY, sub_color=GRAY,
            space=8):
    """items : liste de str ou (str, [sous-points])."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for it in items:
        main, subs = (it, []) if isinstance(it, str) else it
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(space)
        r = p.add_run()
        r.text = "▸  " + main
        r.font.size, r.font.bold, r.font.name = Pt(size), True, "Calibri"
        r.font.color.rgb = color
        for sb in subs:
            p = tf.add_paragraph()
            p.space_after = Pt(space)
            r = p.add_run()
            r.text = "       " + sb
            r.font.size, r.font.name = Pt(sub_size), "Calibri Light"
            r.font.color.rgb = sub_color
    return tb


def circle(s, x, y, d, color, label, size=20):
    o = rect(s, x, y, d, d, color, MSO_SHAPE.OVAL)
    tf = o.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size, r.font.bold, r.font.name = Pt(size), True, "Calibri"
    r.font.color.rgb = WHITE
    return o


def notes(s, text):
    s.notes_slide.notes_text_frame.text = text


def content_slide(title, background=WHITE):
    s = prs.slides.add_slide(BLANK)
    page[0] += 1
    bg(s, background)
    rect(s, 0, 0, 13.33, 1.05, NAVY)
    rect(s, 0, 1.05, 13.33, 0.06, SKY)
    txt(s, title, 0.5, 0.17, 12.33, 0.78, 28, True, WHITE, font="Calibri Light",
        anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 0, 7.12, 13.33, 0.38, NAVY)
    txt(s, FOOTER_TXT, 0.4, 7.17, 9.0, 0.30, 9, color=LSKY)
    h = 0.35
    s.shapes.add_picture(TAMPON, Inches(12.33), Inches(6.68), Inches(h * AR_TAMPON), Inches(h))
    txt(s, str(page[0]), 12.78, 7.17, 0.40, 0.30, 9, color=LSKY, align=PP_ALIGN.RIGHT)
    return s


def title_slide(kicker, title, subtitle, meta):
    s = prs.slides.add_slide(BLANK)
    page[0] += 1
    bg(s, NAVY)
    rect(s, 0, 0, 13.33, 0.08, SKY)
    rect(s, 2.65, 0.65, 0.03, 1.70, SKY)
    h = 0.85
    s.shapes.add_picture(LOGO_BB, Inches(0.45), Inches(1.10), Inches(h * AR_LOGO), Inches(h))
    txt(s, kicker, 2.90, 0.75, 9.0, 0.5, 11, color=SKY)
    txt(s, "FV — Instructeur de Vol", 2.90, 1.25, 9.0, 0.5, 13, color=LSKY, font="Calibri Light")
    txt(s, title, 0.60, 2.45, 12.2, 1.6, 46, True, WHITE, font="Calibri Light")
    txt(s, subtitle, 0.60, 4.00, 11.5, 0.8, 22, color=SKY, font="Calibri Light")
    rect(s, 0.60, 4.85, 11.80, 0.03, SKY)
    txt(s, meta, 0.60, 4.95, 9.0, 0.5, 12, color=LSKY)
    rect(s, 0, 7.42, 13.33, 0.08, SKY)
    return s


def section_slide(num, title, desc):
    s = prs.slides.add_slide(BLANK)
    page[0] += 1
    bg(s, SKY)
    rect(s, 0, 0, 0.18, 7.5, NAVY)
    rect(s, 5.5, 0, 7.83, 7.5, NAVY)
    txt(s, "SECTION", 0.5, 1.20, 5.0, 0.5, 13, color=WHITE)
    txt(s, num, 0.5, 1.60, 5.0, 1.4, 72, True, WHITE, font="Calibri Light")
    h = 0.80
    s.shapes.add_picture(LOGO_BLANC, Inches(0.5), Inches(3.20), Inches(h * AR_LOGO), Inches(h))
    rect(s, 5.5, 2.50, 7.83, 0.04, SKY)
    txt(s, title, 5.9, 1.30, 7.0, 1.2, 34, True, WHITE, font="Calibri Light",
        anchor=MSO_ANCHOR.BOTTOM)
    txt(s, desc, 5.9, 2.75, 6.8, 2.0, 16, color=LSKY, font="Calibri Light")
    txt(s, str(page[0]), 12.78, 7.17, 0.40, 0.30, 9, color=LSKY, align=PP_ALIGN.RIGHT)
    return s


def objectives_slide(items):
    s = content_slide("Objectifs de la séance")
    txt(s, "À la fin de cette séance, vous saurez…", 0.7, 1.30, 11.0, 0.4, 13, color=GRAY2, italic=True)
    for i, it in enumerate(items):
        y = 1.75 + i * 1.05
        circle(s, 0.70, y, 0.55, SKY, str(i + 1))
        txt(s, it, 1.55, y + 0.05, 10.8, 0.5, 18, bold=(i == 0), anchor=MSO_ANCHOR.MIDDLE)
    return s


def keyfigures_slide(title, figs, note):
    """figs : liste de (valeur, libellé, couleur accent)."""
    s = content_slide(title)
    xs = [0.70, 4.85, 9.00]
    for (val, lab, col), x in zip(figs, xs):
        rect(s, x, 2.00, 3.9, 0.06, col)
        txt(s, val, x, 2.35, 3.9, 1.3, 50, True, NAVY, PP_ALIGN.CENTER, "Calibri Light")
        txt(s, lab, x, 3.75, 3.9, 1.1, 15, color=GRAY2, align=PP_ALIGN.CENTER)
    txt(s, note, 0.7, 5.60, 11.9, 0.9, 13, color=GRAY2, italic=True)
    return s


def two_col_slide(title, left, right, lcol=SKY, rcol=NAVY, size=17):
    s = content_slide(title, BGBLUE)
    for (head, items), x, col in ((left, 0.5, lcol), (right, 6.9, rcol)):
        rect(s, x, 1.35, 5.9, 5.3, WHITE)
        rect(s, x, 1.35, 5.9, 0.55, col)
        txt(s, head, x + 0.15, 1.40, 5.6, 0.45, 16, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
        bullets(s, items, x + 0.2, 2.10, 5.5, 4.3, size=size, sub_size=13, space=10)
    return s


def case_slide(title, ident, chrono, factors, lesson):
    s = content_slide(title, BGBLUE)
    # colonne gauche : déroulement
    rect(s, 0.5, 1.35, 6.6, 5.3, WHITE)
    rect(s, 0.5, 1.35, 6.6, 0.55, NAVY)
    txt(s, ident, 0.65, 1.40, 6.3, 0.45, 15, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    tb = s.shapes.add_textbox(Inches(0.7), Inches(2.0), Inches(6.2), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, c in enumerate(chrono):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(6)
        r = p.add_run()
        r.text = f"{i + 1}.  "
        r.font.size, r.font.bold, r.font.name = Pt(15), True, "Calibri"
        r.font.color.rgb = SKY
        r = p.add_run()
        r.text = c
        r.font.size, r.font.name = Pt(15), "Calibri"
        r.font.color.rgb = NAVY
    # colonne droite haut : facteurs (ambre)
    rect(s, 7.4, 1.35, 5.45, 3.05, LAMBER)
    rect(s, 7.4, 1.35, 5.45, 0.55, AMBER)
    txt(s, "Facteurs en amont", 7.55, 1.40, 5.1, 0.45, 15, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    bullets(s, factors, 7.55, 1.98, 5.2, 2.4, size=13, space=3, color=NAVY)
    # colonne droite bas : leçon (vert)
    rect(s, 7.4, 4.60, 5.45, 2.05, LGREEN)
    rect(s, 7.4, 4.60, 5.45, 0.55, GREEN)
    txt(s, "La décision qui manquait", 7.55, 4.65, 5.1, 0.45, 15, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, lesson, 7.6, 5.22, 5.1, 1.4, 14, True, GREEN, anchor=MSO_ANCHOR.MIDDLE)
    return s


def definition_slide(title, definition, cards):
    s = content_slide(title)
    rect(s, 0.7, 1.50, 11.9, 1.75, LSKY)
    rect(s, 0.7, 1.50, 0.12, 1.75, SKY)
    txt(s, definition, 1.1, 1.70, 11.2, 1.4, 20, color=NAVY, font="Calibri Light",
        anchor=MSO_ANCHOR.MIDDLE)
    xs = [0.70, 4.75, 8.80]
    for (head, body, col), x in zip(cards, xs):
        rect(s, x, 3.65, 3.7, 2.85, NAVY)
        rect(s, x, 3.65, 3.7, 0.50, col)
        txt(s, head, x + 0.15, 3.73, 3.5, 0.4, 13, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
        txt(s, body, x + 0.15, 4.25, 3.45, 2.2, 13, color=LSKY, font="Calibri Light")
    return s


def qcm_slide(title, question, answers, good, hint):
    s = content_slide(title)
    q = rect(s, 0.7, 1.45, 11.9, 1.10, LSKY)
    tf = q.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.25)
    r = tf.paragraphs[0].add_run()
    r.text = question
    r.font.size, r.font.bold, r.font.name = Pt(18), True, "Calibri"
    r.font.color.rgb = NAVY
    pos = [(0.70, 2.95), (6.85, 2.95), (0.70, 4.45), (6.85, 4.45)]
    for i, ((x, y), a) in enumerate(zip(pos, answers)):
        ok = i == good
        rect(s, x, y, 5.75, 1.15, LGREEN if ok else LGRAY)
        circle(s, x + 0.18, y + 0.32, 0.5, GREEN if ok else NAVY, "ABCD"[i], 16)
        txt(s, a, x + 0.85, y + 0.15, 4.8, 0.85, 14, bold=ok, color=GREEN if ok else NAVY,
            anchor=MSO_ANCHOR.MIDDLE)
    txt(s, hint, 0.7, 5.85, 11.0, 0.9, 12, color=GRAY2, italic=True)
    return s


def remember_slide(points):
    s = content_slide("Ce qu'il faut retenir", LGRAY)
    for i, pt in enumerate(points):
        y = 1.60 + i * 1.25
        rect(s, 0.7, y, 11.9, 1.05, WHITE)
        circle(s, 0.95, y + 0.27, 0.5, GREEN, "✓", 18)
        txt(s, pt, 1.70, y + 0.15, 10.6, 0.75, 17, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
    return s


def questions_slide(sub):
    s = prs.slides.add_slide(BLANK)
    page[0] += 1
    bg(s, NAVY)
    rect(s, 0, 0, 13.33, 0.08, SKY)
    rect(s, 0, 7.42, 13.33, 0.08, SKY)
    h = 1.0
    s.shapes.add_picture(LOGO_BB, Inches((13.33 - h * AR_LOGO) / 2), Inches(1.05), Inches(h * AR_LOGO), Inches(h))
    txt(s, "Questions ?", 0.6, 2.75, 12.1, 1.2, 54, True, WHITE, PP_ALIGN.CENTER, "Calibri Light")
    txt(s, sub, 0.6, 3.90, 12.1, 0.7, 20, color=SKY, align=PP_ALIGN.CENTER, font="Calibri Light")
    rect(s, 3.5, 4.70, 6.3, 0.03, SKY)
    txt(s, "FV · Instructeur de Vol · Formation PPL / LAPL", 0.6, 4.85, 12.1, 0.5, 13, color=LSKY,
        align=PP_ALIGN.CENTER)
    return s


# ═══════════════════════════════════════════════════════════════════════════
# CONTENU
# ═══════════════════════════════════════════════════════════════════════════

# 1. Titre
s = title_slide("SÉCURITÉ DES VOLS · AÉROCLUB",
                "Sorties de piste longitudinales",
                "Pourquoi continue-t-on un atterrissage qu'il fallait abandonner ?",
                "Réunion sécurité · Pilotes brevetés et élèves · D'après aeroVFR (sept. 2026) et rapports BEA")
notes(s, "Source principale : aeroVFR, « Sorties de piste longitudinales » volets 1 et 2, septembre 2026. "
         "Cas BEA : F-GOOF Pressignac-Vicq (BEA2023-0218), DA40 La Grand'Combe, F-BXZG Fumel-Montayral (2014).")

# 2. Objectifs
s = objectives_slide([
    "Reconnaître les facteurs qui précèdent une approche non stabilisée",
    "Analyser trois accidents récents du BEA sur des avions de club",
    "Appliquer une porte de décision et un repère de remise de gaz sur nos pistes",
    "Calculer une distance d'atterrissage majorée avant un vol vers un terrain court",
])

# 3. Accroche : chiffres
s = keyfigures_slide("Un toucher à mi-piste : que reste-t-il ?", [
    ("700 m", "Longueur de la piste (cas typique des rapports BEA)", SKY),
    ("350 m", "Ce qu'il reste pour freiner après un toucher à mi-piste", RED),
    ("?  m", "Distance de roulement de votre avion, herbe, vent nul, freinage normal", AMBER),
], "Question à l'assistance : qui connaît, de tête, la distance d'atterrissage de l'avion qu'il pilote le plus souvent ? "
   "Puis : qui l'a déjà recalculée avec de l'herbe mouillée ou une composante de vent arrière ?")
notes(s, "aeroVFR cite des rapports où le pilote accepte un toucher à mi-piste sur 700 m, ne laissant que 350 m pour freiner. "
         "Ouvrir la discussion avant de donner des chiffres.")

# 4. Section 1
section_slide("01", "Le constat", "Ce que disent les rapports du BEA, et pourquoi la cause « officielle » n'est qu'un symptôme.")

# 5. Ce que disent les rapports
s = content_slide("Ce que disent les rapports du BEA")
bullets(s, [
    ("À l'atterrissage bien plus qu'au décollage",
     ["Au décollage : panne moteur en accélération, verrière ouverte, puis décélération tardive ou insuffisante."]),
    ("La cause citée presque à chaque fois : finale non stabilisée",
     ["Trop haut, trop vite, ou les deux. Toucher long, freinage insuffisant, sortie en bout de piste."]),
    ("Le point commun : la remise de gaz n'a pas été décidée",
     ["Le pilote voit les paramètres dégradés et continue quand même. Le rapport le dit sobrement : « le pilote poursuit l'atterrissage »."]),
    ("Pas un problème de petits avions",
     ["Symposium DSAC 2023 : sur 115 accidents mondiaux d'avions commerciaux en 2022, 16 sont des sorties de piste."]),
], 0.6, 1.35, 11.9, 5.4, size=22, sub_size=16, space=12)
notes(s, "Symposium sécurité DSAC du 7 décembre 2023, consacré à la prévention des sorties longitudinales de piste. "
         "Chiffre à revérifier dans le livret du symposium avant projection.")

# 6. Symptôme vs facteurs
s = two_col_slide("La finale non stabilisée est un symptôme",
                  ("Ce que dit le rapport", [
                      "Approche trop haute et trop rapide",
                      "Toucher long, au-delà du premier tiers",
                      "Rebond(s), freinage tardif",
                      "Absence de remise de gaz",
                      "Sortie longitudinale, dégâts, parfois blessés",
                  ]),
                  ("Ce qui s'est passé avant", [
                      "Terrain inconnu, VAC survolée, pente et relief non anticipés",
                      "Circuit raccourci ou « contraint » : pas le temps de stabiliser",
                      "Vent mal évalué : composante arrière acceptée sans calcul",
                      "Précipitation : un souci banal traité comme une urgence",
                      "Peu d'expérience récente, aucune marge définie avant le vol",
                  ]), lcol=RED, rcol=AMBER)
notes(s, "Thèse centrale d'aeroVFR : le BEA décrit la déstabilisation, mais celle-ci résulte de facteurs antérieurs. "
         "C'est sur ces facteurs que le club peut agir.")

# 7. Section 2
section_slide("02", "Trois cas du BEA", "Des avions de club, des pistes de 700 à 800 mètres, et à chaque fois une décision de remise de gaz qui n'a pas été prise.")

# 8. Cas 1
s = case_slide("Cas 1 · DR400 à Pressignac-Vicq (2023)",
               "Robin DR400 F-GOOF · Piste 19 non revêtue · 800 m · pente montante 5 %",
               ["Premier posé : rebonds, toucher dur, l'avion part à gauche. Remise de gaz.",
                "Le pilote raccourcit le circuit pour éviter une ferme, alors que la VAC prévoit un circuit large.",
                "Seconde approche en 19 : QFU imposé, vent 340° / 9 kt, donc composante arrière.",
                "Vitesse excessive, trajectoire décalée de l'axe et au-dessus du plan. Pas de stabilisation possible.",
                "Toucher à environ 500 m après le seuil, à gauche de l'axe. Sortie de piste, clôture, arbres."],
               ["Terrain privé court, QFU imposé, pente : contraintes cumulées",
                "Circuit raccourci « pour bien faire »",
                "Composante de vent arrière non prise en compte",
                "Une première remise de gaz réussie, puis plus rien"],
               "Seconde remise de gaz à 300 ft, dès que l'axe et la vitesse ne sont pas tenus. Une remise de gaz n'est pas un échec, c'est une procédure.")
notes(s, "Rapport BEA2023-0218, accident du 14 juin 2023 à Pressignac (24), aérodrome de Rebeyrotte LF2432. "
         "Lire le rapport complet avant la séance : la chronologie des rebonds diffère légèrement selon les extraits.")

# 9. Cas 2
s = case_slide("Cas 2 · DA40 à La Grand'Combe",
               "Diamond DA40 · Piste 35 · 780 m · terrain au sommet d'une colline, entouré de relief",
               ["Première approche en 17 : vitesse impossible à réduire en finale. Remise de gaz.",
                "Passage de reconnaissance pour vérifier le vent.",
                "Seconde approche en 35. Toucher dans le premier tiers de la piste.",
                "Rebond sur une bosse de la piste.",
                "Sortie longitudinale, l'avion s'immobilise en contrebas."],
               ["Terrain « complexe » selon le BEA : relief, turbulence, piste bosselée",
                "Avion à aile lisse, peu tolérant à l'excès de vitesse",
                "La première remise de gaz avait consommé une partie de l'attention",
                "Après le rebond, le pilote poursuit"],
               "Après un rebond franc, remise de gaz immédiate. Le second toucher risque un nouveau rebond, et la piste restante diminue très vite.")
notes(s, "Rapport BEA « Approche non stabilisée, atterrissage long, rebond, sortie longitudinale de piste ». "
         "Point pédagogique : un toucher dans le premier tiers ne garantit rien si l'énergie n'est pas dissipée.")

# 10. Cas 3
s = case_slide("Cas 3 · Cessna F150M à Fumel-Montayral (2014)",
               "Reims-Cessna F150M F-BXZG · Piste 35 · 720 m · vent plein travers 10 kt avec rafales",
               ["Le pilote totalise 1 heure et 3 atterrissages dans les 8 derniers mois.",
                "Préparation du vol sans analyse des paramètres météo.",
                "Finale entamée avec une vitesse excessive.",
                "À l'arrondi, une rafale fait remonter l'avion. Second arrondi, même phénomène.",
                "Hésitation à remettre les gaz, puis décision de continuer. Toucher à mi-piste, talus, train avant arraché."],
               ["Expérience récente très faible, non compensée par une marge",
                "Météo non intégrée à la décision de partir",
                "Vent travers avec composante arrière et rafales",
                "Excès d'énergie impossible à absorber à l'arrondi"],
               "Au premier arrondi remonté par une rafale, remise de gaz. Et avant cela : renoncer au vol, ou choisir une piste plus longue.")
notes(s, "Rapport BEA, accident du 29 juin 2014 à Fumel-Montayral (47). Cas plus ancien mais très pédagogique sur le lien "
         "entre expérience récente et marge de décision.")

# 11. Section 3
section_slide("03", "Les facteurs en amont", "Ce qui, avant même la finale, met le pilote dans une situation qu'il ne saura plus rattraper.")

# 12. Facteurs
s = content_slide("Six facteurs qui précèdent la déstabilisation")
bullets(s, [
    ("Préparation insuffisante du terrain d'arrivée",
     ["VAC lue en vol, topographie non regardée, consignes de turbulence ou de pente ignorées."]),
    ("Circuit contraint ou raccourci",
     ["Éviter une ferme, gagner du temps, suivre un autre avion : la finale devient trop courte pour stabiliser."]),
    ("Vent mal évalué",
     ["Manche à air non surveillée, approche à contre-QFU par vent « calme », composante arrière acceptée."]),
    ("Syndrome de précipitation",
     ["Un souci banal est traité comme une urgence : on veut « faire vite » et on supprime les marges."]),
    ("Surface et pente non intégrées",
     ["Herbe mouillée : freinage insuffisant sur 685 m et fossé. Pente descendante : roulement allongé."]),
    ("Expérience récente faible",
     ["Sans marge supplémentaire décidée à l'avance, le pilote n'a plus de ressource quand la situation se dégrade."]),
], 0.6, 1.30, 11.9, 5.6, size=19, sub_size=14, space=7)
notes(s, "Le cas « herbe mouillée » : approche à contre-QFU par vent calme, seul dans le circuit, atterrissage long après hésitation "
         "à remettre les gaz, freinage insuffisant sur 685 m, fossé. Le pilote avait déjà hésité à décoller à cause d'un ciel menaçant.")

# 13. Définition : syndrome de précipitation
s = definition_slide("Le syndrome de précipitation",
                     "« Un problème en vol est traité comme une urgence, alors qu'un tour de piste normal serait la bonne solution. "
                     "Sous une pression de temps injustifiée, le pilote veut faire vite : il raccourcit le circuit, "
                     "ou se pose à contre-QFU sans regarder la composante de vent arrière. »",
                     [("Les signes", "Un souci mineur (verrière, porte, passager malade, carburant « juste », météo qui se dégrade) "
                                     "et soudain tout s'accélère : radio précipitée, circuit tronqué, check-list oubliée.", SKY),
                      ("Le mémo", "« Est-ce que l'avion vole ? » Si oui, j'ai le temps d'un tour de piste complet. "
                                  "Prendre de la hauteur, respirer, refaire la préparation de l'approche.", GREEN),
                      ("Le piège", "Se dire qu'un circuit court « fera gagner une minute ». Il fait surtout perdre "
                                   "la seule chose qui permet de stabiliser : la distance en finale.", AMBER)])
notes(s, "Notion reprise du volet 2 d'aeroVFR. La verrière ouverte au décollage est l'exemple classique : l'avion vole, "
         "il faut juste faire un tour de piste normal.")

# 14. Chiffres : la physique
s = keyfigures_slide("La physique ne négocie pas", [
    ("+ 20 %", "de distance d'atterrissage pour 10 % de vitesse en trop (l'énergie varie avec le carré de la vitesse)", AMBER),
    ("× 1,5", "distance de roulement avec 10 kt de vent arrière (ordre de grandeur donné par les manuels Cessna)", RED),
    ("× 1,3", "sur herbe mouillée ou haute, par rapport à une piste en dur sèche (majorations usuelles)", AMBER),
], "Ces coefficients se cumulent. Ils sont indicatifs : la référence reste le manuel de vol de l'avion, "
   "puis une marge personnelle par-dessus. Le calcul se fait au sol, avant le départ, jamais en finale.")
notes(s, "Vérifier ces ordres de grandeur avec les manuels de vol des avions du club (DR400, Cessna, etc.) et adapter la slide. "
         "L'important est le message : les majorations se multiplient entre elles.")

# 15. Section 4
section_slide("04", "Les règles et leur application", "Une porte de décision, un repère de remise de gaz, un calcul de distance : trois habitudes à installer au club.")

# 16. Règles
s = content_slide("Cinq règles à appliquer sur chaque atterrissage")
bullets(s, [
    ("La porte des 300 ft sol : stabilisé ou remise de gaz",
     ["Axe, plan, vitesse, configuration, moteur. Si un seul paramètre n'est pas tenu, on remet les gaz. Sans débat."]),
    ("Vitesse d'approche : 1,3 Vs, corrigée seulement pour le vent fort ou turbulent",
     ["C'est le paramètre sur lequel le pilote a le plus d'influence. Chaque nœud en trop se paie en mètres."]),
    ("Une zone de toucher et un repère physique de remise de gaz",
     ["Toucher visé dans le premier tiers. Si les roues ne sont pas au sol avant le repère, on remet les gaz."]),
    ("Rebond franc = remise de gaz immédiate",
     ["Petit rebond : conserver l'assiette. Rebond franc : plein gaz, assiette de montée, on recommence."]),
    ("Majorer les distances quand le vent est incertain, la piste mouillée ou en pente",
     ["Manche à air surveillée jusqu'en finale. Composante arrière : on change de QFU ou de terrain."]),
], 0.6, 1.30, 11.9, 5.6, size=20, sub_size=15, space=9)
notes(s, "Doctrine reprise d'aeroVFR 2015 « Pour éviter les sorties longitudinales » et du guide DGAC sur l'approche stabilisée. "
         "En aviation commerciale la porte est à 1 000 ft, en aviation générale à 300 ft sol.")

# 17. Grille de décision
s = two_col_slide("À 300 ft sol : je continue ou je remets les gaz ?",
                  ("Je continue seulement si tout est vrai", [
                      "Sur l'axe, sans correction importante en cours",
                      "Sur le plan, avec un point d'aboutissement stable",
                      "Vitesse : 1,3 Vs, écart de 5 kt au plus",
                      "Configuration finale sortie, avion compensé",
                      "Piste dégagée, vent conforme à ce qui était prévu",
                      "Zone de toucher et repère de remise de gaz identifiés",
                  ]),
                  ("Je remets les gaz dès que", [
                      "Un paramètre sort de la tolérance et n'y revient pas",
                      "L'arrondi est remonté par une rafale",
                      "Rebond franc au toucher",
                      "Les roues ne sont pas au sol avant le repère",
                      "Un doute, quel qu'il soit, sur le vent ou la piste",
                      "J'ai déjà remis les gaz une fois : même règle la seconde fois",
                  ]), lcol=GREEN, rcol=RED)
notes(s, "Insister sur la dernière ligne de droite : dans deux des trois cas, le pilote avait déjà remis les gaz une fois, "
         "et n'a pas refait le même choix à la seconde approche.")

# 18. Application locale
s = content_slide("Application à nos pistes et à nos navigations")
bullets(s, [
    ("Nos pistes : longueur, revêtement, pente, repère de remise de gaz",
     ["[À compléter : pour chaque QFU, un repère visible depuis le cockpit avant lequel les roues doivent être au sol.]"]),
    ("Nos terrains de navigation habituels de moins de 800 m",
     ["[À compléter : liste des terrains courts fréquentés par le club, avec leurs particularités VAC.]"]),
    ("Exercice au tableau : distance d'atterrissage majorée",
     ["Distance manuel (piste dure, sèche, vent nul) × herbe × mouillée × vent arrière × marge personnelle = ?",
      "Comparer le résultat à la longueur de piste. Si l'écart est faible, le terrain est hors limites ce jour-là."]),
    ("Engagement de séance",
     ["Chacun repart avec une valeur de distance pour son avion habituel et un repère de remise de gaz sur notre piste principale."]),
], 0.6, 1.30, 11.9, 5.6, size=20, sub_size=15, space=9)
notes(s, "Slide à personnaliser avec les données du club avant la séance. Prévoir le manuel de vol de deux avions du club "
         "pour faire l'exercice en direct.")

# 19. QCM
s = qcm_slide("Question de révision",
              "Finale sur une piste en herbe de 750 m. À 300 ft sol, vous êtes à Vref + 15 kt et légèrement au-dessus du plan. Que faites-vous ?",
              ["Réduire les gaz, sortir le dernier cran de volets et rattraper le plan avant l'arrondi",
               "Remettre les gaz maintenant, refaire un tour de piste complet et une finale stabilisée",
               "Continuer, poser dans le premier tiers et freiner énergiquement",
               "Faire une glissade pour perdre la hauteur en excès et poser normalement"],
              1,
              "Correction : deux paramètres hors tolérance à la porte de décision. Rattraper en finale courte, c'est exactement le scénario "
              "des trois cas BEA. La réponse B est la seule qui restaure les marges.")

# 20. Débat
s = content_slide("Pour le débat")
bullets(s, [
    ("Qui a déjà continué un atterrissage qu'il aurait dû abandonner ? Pourquoi ?",
     ["Sans jugement : la fatigue, l'orgueil, la présence d'un passager, le « ça va passer » sont des facteurs partagés par tous."]),
    ("Que se passe-t-il au club après une remise de gaz ?",
     ["Est-elle vue comme un échec ou comme une bonne décision ? La culture du club influence la décision en finale."]),
    ("Quels terrains de nos navigations méritent une préparation particulière ?",
     ["Pente, bosse, relief, QFU imposé, herbe : les repérer maintenant, pas au moment de s'y poser."]),
], 0.6, 1.40, 11.9, 5.4, size=22, sub_size=16, space=18)

# 21. À retenir
remember_slide([
    "La finale non stabilisée est un symptôme : les causes sont dans la préparation et le circuit.",
    "À 300 ft sol, un paramètre hors tolérance suffit : remise de gaz, sans débat.",
    "Une zone de toucher, un repère physique, un rebond franc : trois déclencheurs de remise de gaz.",
    "Distance d'atterrissage majorée calculée au sol. Vent arrière, herbe mouillée et pente se multiplient.",
])

# 22. Sources
s = content_slide("Sources et lectures")
bullets(s, [
    ("aeroVFR, « Sorties de piste longitudinales » (1) et (2), septembre 2026",
     ["aerovfr.com/2026/09/sorties-de-piste-longitudinales-1/ et …-2/"]),
    ("Rapports BEA",
     ["F-GOOF, DR400, Pressignac (24), 14 juin 2023, BEA2023-0218",
      "DA40, La Grand'Combe : « Approche non stabilisée, atterrissage long, rebond, sortie longitudinale de piste »",
      "F-BXZG, Reims-Cessna F150M, Fumel-Montayral (47), 29 juin 2014"]),
    ("Doctrine",
     ["aeroVFR, « Pour éviter les sorties longitudinales à l'atterrissage » (2015) ; « Finale non stabilisée, prise de décision et remise de gaz » (2023) ; « La menace des circuits contraints » (2025)",
      "DGAC, Guide de transition « Approche stabilisée »",
      "DSAC, Symposium sécurité 2023 « Prévenir les sorties de piste » ; BEA, étude sorties de piste en aviation générale (2006)"]),
], 0.6, 1.30, 11.9, 5.6, size=17, sub_size=13, space=6)

# 23. Questions
questions_slide("Une remise de gaz n'est jamais une erreur")

out = "/tmp/claude-0/-home-user-trip-moto-juin-2026/4f4a1dd2-0483-5b95-b2d1-ab9c271fc642/scratchpad/Sorties_de_piste_longitudinales_FV.pptx"
prs.save(out)
print("saved", out, "slides:", len(prs.slides))
