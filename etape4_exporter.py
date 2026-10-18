"""
ÉTAPE 4 — Construction du classeur Excel de dépouillement.

    python etape4_exporter.py
    python etape4_exporter.py --inclure-non-verifies   # si tu n'as pas fini l'étape 3

Onglets produits :
  Données brutes       1 ligne par enseignant, 1 colonne par variable (+ CSV jumeau)
  Dictionnaire         définition de chaque variable (à mettre en annexe du mémoire)
  Scores recodés       items inversés remis à l'endroit + score moyen par section
  Stats échelles       N, moyenne, écart-type, médiane, distribution 1-5, % d'accord
  Stats catégorielles  effectifs et pourcentages de toutes les questions à choix
  Synthèse sections    score moyen par section + alpha de Cronbach
  Contrôles            incohérences et non-réponses à examiner
  Graphiques           6 graphiques Excel natifs, modifiables
Tout est écrit en FORMULES : si tu corriges une valeur dans « Données brutes »,
les statistiques et les graphiques se recalculent tout seuls.
"""

import argparse
import json
import statistics
import sys

import config
from codebook import (CHAMPS, PAR_CODE, SECTIONS_SCORE, LIKERT_LABELS,
                      champs_de_section, colonnes_excel, _slug)

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

POLICE = "Arial"
TITRE = Font(name=POLICE, bold=True, color="FFFFFF", size=10)
FOND = PatternFill("solid", fgColor="1F4E79")
FOND_DOUX = PatternFill("solid", fgColor="DDEBF7")
NORMAL = Font(name=POLICE, size=10)
GRAS = Font(name=POLICE, size=10, bold=True)
BORD = Border(*[Side(style="thin", color="BFBFBF")] * 4)


# ==========================================================================
# 1. Chargement des questionnaires dépouillés
# ==========================================================================
def charger_tout(inclure_non_verifies: bool):
    fiches, non_verifies, exclus = [], 0, 0
    for chemin in sorted(config.DOSSIER_JSON.glob("q*.json")):
        qid = int(chemin.stem[1:])
        valide = config.DOSSIER_VALIDE / f"q{qid:05d}.json"
        if valide.exists():
            d = json.loads(valide.read_text(encoding="utf-8"))
        else:
            d = json.loads(chemin.read_text(encoding="utf-8"))
            non_verifies += 1
            if not inclure_non_verifies:
                continue
        if d.get("exclu"):
            exclus += 1
            continue
        fiches.append(d)
    return fiches, non_verifies, exclus


def ligne_plate(d: dict) -> dict:
    """Transforme un JSON de questionnaire en une ligne de tableau."""
    r = d.get("reponses", {})
    ligne = {"id_questionnaire": d["id_questionnaire"],
             "pages_source": "-".join(str(p) for p in d.get("pages_source", []))}
    for c in CHAMPS:
        v = r.get(c.code)
        if c.type == "multi":
            choisis = v or []
            for opt in c.options:
                ligne[f"{c.code}__{_slug(opt)}"] = 1 if opt in choisis else 0
            ligne[c.code] = " | ".join(choisis)
        elif c.type == "likert":
            ligne[c.code] = int(v) if isinstance(v, (int, float)) else None
        else:
            ligne[c.code] = v if v not in ("", None) else None
    ligne["a_verifier"] = "oui" if d.get("a_verifier") else ""
    ligne["verifie"] = "oui" if d.get("verifie") else ""
    ligne["remarques"] = " ; ".join(d.get("anomalies", [])) or d.get("commentaire", "")
    return ligne


# ==========================================================================
# 2. Petits utilitaires de mise en forme
# ==========================================================================
def entete(ws, valeurs, ligne=1, largeurs=None):
    for j, v in enumerate(valeurs, start=1):
        cel = ws.cell(row=ligne, column=j, value=v)
        cel.font, cel.fill = TITRE, FOND
        cel.alignment = Alignment(vertical="center", wrap_text=True)
        cel.border = BORD
    ws.row_dimensions[ligne].height = 32
    if largeurs:
        for j, l in enumerate(largeurs, start=1):
            ws.column_dimensions[get_column_letter(j)].width = l


def styler_zone(ws, min_row, max_row, max_col, fmt=None):
    for r in range(min_row, max_row + 1):
        for c in range(1, max_col + 1):
            cel = ws.cell(row=r, column=c)
            cel.font = NORMAL
            cel.border = BORD
            if fmt and c >= fmt[0]:
                cel.number_format = fmt[1]


def ref(feuille, col_lettre, l1, l2):
    return f"'{feuille}'!${col_lettre}${l1}:${col_lettre}${l2}"


# ==========================================================================
# 3. Les onglets
# ==========================================================================
def onglet_brut(wb, lignes):
    ws = wb.create_sheet("Données brutes")
    cols = colonnes_excel()
    largeurs = [16 if c in ("id_questionnaire", "pages_source") else 14 for c in cols]
    entete(ws, cols, largeurs=largeurs)
    for i, ligne in enumerate(lignes, start=2):
        for j, code in enumerate(cols, start=1):
            ws.cell(row=i, column=j, value=ligne.get(code))
    styler_zone(ws, 2, len(lignes) + 1, len(cols))
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(lignes) + 1}"
    return {code: get_column_letter(j) for j, code in enumerate(cols, start=1)}


def onglet_dictionnaire(wb):
    ws = wb.create_sheet("Dictionnaire")
    entete(ws, ["Code", "Libellé de la question / affirmation", "Section",
                "Page", "Type", "Modalités", "Item inversé"],
           largeurs=[16, 70, 34, 6, 10, 46, 12])
    r = 2
    for c in CHAMPS:
        ws.cell(row=r, column=1, value=c.code)
        ws.cell(row=r, column=2, value=c.libelle).alignment = Alignment(wrap_text=True)
        ws.cell(row=r, column=3, value=c.section)
        ws.cell(row=r, column=4, value=c.page)
        ws.cell(row=r, column=5, value={"likert": "échelle 1-5", "single": "choix unique",
                                        "multi": "choix multiple", "text": "texte"}[c.type])
        modal = " / ".join(c.options) if c.options else (
            " / ".join(f"{k}={v}" for k, v in LIKERT_LABELS.items())
            if c.type == "likert" else "")
        ws.cell(row=r, column=6, value=modal).alignment = Alignment(wrap_text=True)
        ws.cell(row=r, column=7, value="oui" if c.inverse else "")
        r += 1
    styler_zone(ws, 2, r - 1, 7)
    ws.freeze_panes = "A2"


def onglet_scores(wb, colmap, n):
    """Items recodés (6-x pour les items inversés) + score moyen par section."""
    ws = wb.create_sheet("Scores recodés")
    derniere = n + 1
    colonnes, groupes = ["id_questionnaire"], []
    for section in SECTIONS_SCORE:
        items = [c for c in champs_de_section(section) if c.type == "likert"]
        if not items:
            continue
        debut = len(colonnes) + 1
        colonnes += [c.code for c in items]
        groupes.append((section, debut, len(colonnes)))
    colonnes_scores = [f"Score {s.split('.')[0]}" for s, _, _ in groupes] + ["Score global"]
    entete(ws, colonnes + colonnes_scores,
           largeurs=[16] + [10] * (len(colonnes) - 1) + [13] * len(colonnes_scores))

    for i in range(2, derniere + 1):
        ws.cell(row=i, column=1,
                value=f"='Données brutes'!{colmap['id_questionnaire']}{i}")
        for j, code in enumerate(colonnes[1:], start=2):
            src = f"'Données brutes'!{colmap[code]}{i}"
            f = (f'=IF({src}="","",6-{src})' if PAR_CODE[code].inverse
                 else f'=IF({src}="","",{src})')
            ws.cell(row=i, column=j, value=f)
        for k, (_, d, f_) in enumerate(groupes):
            cel = ws.cell(row=i, column=len(colonnes) + 1 + k)
            cel.value = (f'=IFERROR(AVERAGE({get_column_letter(d)}{i}:'
                         f'{get_column_letter(f_)}{i}),"")')
        ws.cell(row=i, column=len(colonnes) + 1 + len(groupes),
                value=f'=IFERROR(AVERAGE(B{i}:{get_column_letter(len(colonnes))}{i}),"")')

    styler_zone(ws, 2, derniere, len(colonnes) + len(colonnes_scores),
                fmt=(len(colonnes) + 1, "0.00"))
    ws.freeze_panes = "B2"
    return groupes, len(colonnes)


def onglet_stats_likert(wb, colmap, n):
    ws = wb.create_sheet("Stats échelles")
    derniere = n + 1
    cols = ["Section", "Code", "Affirmation", "Inversé", "N", "Moyenne", "Écart-type",
            "Médiane"] + [f"n{k}" for k in LIKERT_LABELS] + \
           [f"%{k}" for k in LIKERT_LABELS] + ["% accord (4-5)", "% désaccord (1-2)",
                                               "% neutre (3)", "Non-réponses"]
    entete(ws, cols, largeurs=[30, 9, 62, 8, 7, 10, 11, 9] + [6] * 5 + [8] * 5 + [13, 15, 11, 13])

    r = 2
    items = [c for c in CHAMPS if c.type == "likert"]
    for c in items:
        plage = ref("Données brutes", colmap[c.code], 2, derniere)
        ws.cell(row=r, column=1, value=c.section)
        ws.cell(row=r, column=2, value=c.code)
        ws.cell(row=r, column=3, value=c.libelle).alignment = Alignment(wrap_text=True)
        ws.cell(row=r, column=4, value="oui" if c.inverse else "")
        ws.cell(row=r, column=5, value=f"=COUNT({plage})")
        ws.cell(row=r, column=6, value=f'=IFERROR(AVERAGE({plage}),"")')
        ws.cell(row=r, column=7, value=f'=IFERROR(STDEV({plage}),"")')
        ws.cell(row=r, column=8, value=f'=IFERROR(MEDIAN({plage}),"")')
        for k in range(1, 6):
            ws.cell(row=r, column=8 + k, value=f"=COUNTIF({plage},{k})")
            ws.cell(row=r, column=13 + k,
                    value=f'=IFERROR(COUNTIF({plage},{k})/COUNT({plage}),"")')
        ws.cell(row=r, column=19, value=f'=IFERROR(COUNTIF({plage},">=4")/COUNT({plage}),"")')
        ws.cell(row=r, column=20, value=f'=IFERROR(COUNTIF({plage},"<=2")/COUNT({plage}),"")')
        ws.cell(row=r, column=21, value=f'=IFERROR(COUNTIF({plage},3)/COUNT({plage}),"")')
        ws.cell(row=r, column=22, value=f"={n}-COUNT({plage})")
        r += 1

    styler_zone(ws, 2, r - 1, 22)
    for rr in range(2, r):
        for cc in (6, 7, 8):
            ws.cell(row=rr, column=cc).number_format = "0.00"
        for cc in list(range(14, 22)):
            ws.cell(row=rr, column=cc).number_format = "0.0%"
    ws.freeze_panes = "D2"
    return r - 1  # dernière ligne


def onglet_stats_categorielles(wb, colmap, n):
    ws = wb.create_sheet("Stats catégorielles")
    entete(ws, ["Variable", "Question", "Modalité", "Effectif", "Pourcentage"],
           largeurs=[18, 58, 46, 12, 14])
    derniere = n + 1
    r = 2
    reperes = {}
    for c in CHAMPS:
        if c.type not in ("single", "multi"):
            continue
        reperes[c.code] = (r, r + len(c.options) - 1)
        for opt in c.options:
            ws.cell(row=r, column=1, value=c.code)
            ws.cell(row=r, column=2, value=c.libelle).alignment = Alignment(wrap_text=True)
            ws.cell(row=r, column=3, value=opt)
            if c.type == "single":
                plage = ref("Données brutes", colmap[c.code], 2, derniere)
                ws.cell(row=r, column=4, value=f'=COUNTIF({plage},"{opt}")')
                ws.cell(row=r, column=5,
                        value=f'=IFERROR(COUNTIF({plage},"{opt}")/COUNTA({plage}),"")')
            else:
                plage = ref("Données brutes", colmap[f"{c.code}__{_slug(opt)}"], 2, derniere)
                ws.cell(row=r, column=4, value=f"=SUM({plage})")
                ws.cell(row=r, column=5, value=f'=IFERROR(SUM({plage})/{n},"")')
            ws.cell(row=r, column=5).number_format = "0.0%"
            r += 1
        r += 1  # ligne vide entre deux variables
    styler_zone(ws, 2, r - 1, 5)
    ws.freeze_panes = "C2"
    return reperes


def alpha_cronbach(matrice):
    """Alpha de Cronbach sur une liste de listes (1 liste = 1 répondant)."""
    complets = [l for l in matrice if all(v is not None for v in l)]
    k = len(complets[0]) if complets else 0
    if k < 2 or len(complets) < 3:
        return None
    variances = [statistics.pvariance([l[j] for l in complets]) for j in range(k)]
    total = statistics.pvariance([sum(l) for l in complets])
    if total == 0:
        return None
    return k / (k - 1) * (1 - sum(variances) / total)


def onglet_synthese(wb, groupes, n_cols_scores, lignes, n):
    ws = wb.create_sheet("Synthèse sections")
    entete(ws, ["Section", "Nb d'items", "Score moyen (1-5)", "Écart-type",
                "N répondants", "Alpha de Cronbach"],
           largeurs=[46, 12, 18, 13, 14, 18])
    derniere = n + 1
    r = 2
    for k, (section, deb, fin) in enumerate(groupes):
        col = get_column_letter(n_cols_scores + 1 + k)
        plage = ref("Scores recodés", col, 2, derniere)
        items = [c for c in champs_de_section(section) if c.type == "likert"]
        ws.cell(row=r, column=1, value=section)
        ws.cell(row=r, column=2, value=len(items))
        ws.cell(row=r, column=3, value=f'=IFERROR(AVERAGE({plage}),"")')
        ws.cell(row=r, column=4, value=f'=IFERROR(STDEV({plage}),"")')
        ws.cell(row=r, column=5, value=f"=COUNT({plage})")
        matrice = [[(6 - l[c.code]) if (c.inverse and l[c.code] is not None) else l[c.code]
                    for c in items] for l in lignes]
        a = alpha_cronbach(matrice)
        ws.cell(row=r, column=6, value=round(a, 3) if a is not None else "n.d.")
        for cc in (3, 4):
            ws.cell(row=r, column=cc).number_format = "0.00"
        r += 1
    styler_zone(ws, 2, r - 1, 6)
    ws.cell(row=r + 1, column=1,
            value="Les scores sont calculés après recodage des items inversés (6 - x) : "
                  "un score élevé va toujours dans le sens favorable.").font = NORMAL
    ws.cell(row=r + 2, column=1,
            value="L'alpha de Cronbach est calculé en Python au moment de l'export "
                  "(sur les répondants sans donnée manquante) et ne se recalcule pas "
                  "dans Excel : relance etape4_exporter.py après correction.").font = NORMAL
    return r - 1


def onglet_ouvertes(wb, lignes):
    ws = wb.create_sheet("Réponses ouvertes")
    codes = ["q50a", "q50b", "q50c", "q50d"]
    entete(ws, ["id_questionnaire"] + [PAR_CODE[c].libelle for c in codes],
           largeurs=[16, 55, 55, 55, 55])
    r = 2
    for l in lignes:
        if not any(l.get(c) for c in codes):
            continue
        ws.cell(row=r, column=1, value=l["id_questionnaire"])
        for j, c in enumerate(codes, start=2):
            cel = ws.cell(row=r, column=j, value=l.get(c))
            cel.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    styler_zone(ws, 2, max(r - 1, 2), 5)
    ws.freeze_panes = "B2"


def onglet_controles(wb, lignes, non_verifies, exclus):
    ws = wb.create_sheet("Contrôles")
    entete(ws, ["Contrôle", "Nb concernés", "Questionnaires concernés"],
           largeurs=[62, 14, 90])

    def ids(test):
        return [l["id_questionnaire"] for l in lignes if test(l)]

    ame_items = ["ame_a", "ame_b", "ame_c"]
    controles = [
        ("Volet AME rempli alors que le statut n'est pas AME",
         ids(lambda l: l.get("statut") != "AME" and any(l.get(c) for c in ame_items))),
        ("Statut AME mais volet AME vide",
         ids(lambda l: l.get("statut") == "AME" and not any(l.get(c) for c in ame_items))),
        ("Pas de formation en 12 mois mais nombre de jours renseigné",
         ids(lambda l: l.get("formation_12mois") == "Non" and l.get("nb_jours_formation"))),
        ("Pas de formation à distance mais satisfaction renseignée",
         ids(lambda l: l.get("formation_distance") == "Non"
             and l.get("satisfaction_distance") not in (None, "Sans objet"))),
        ("« Aucun accompagnement reçu » coché avec d'autres modalités (q15)",
         ids(lambda l: l.get("q15__aucun_accompagnement_recu") == 1
             and sum(l.get(f"q15__{_slug(o)}", 0) for o in PAR_CODE["q15"].options) > 1)),
        ("« Pas d'obstacle majeur » coché avec d'autres obstacles (q42)",
         ids(lambda l: l.get("q42__pas_d_obstacle_majeur") == 1
             and sum(l.get(f"q42__{_slug(o)}", 0) for o in PAR_CODE["q42"].options) > 1)),
        ("Plus de 20 % des items d'échelle sans réponse",
         ids(lambda l: sum(1 for c in CHAMPS if c.type == "likert"
                           and l.get(c.code) is None) > 0.2 * 44)),
        ("Signalé « à vérifier » par la lecture automatique",
         ids(lambda l: l.get("a_verifier") == "oui")),
        ("Non validé manuellement (étape 3)",
         ids(lambda l: l.get("verifie") != "oui")),
    ]
    r = 2
    for libelle, lst in controles:
        ws.cell(row=r, column=1, value=libelle)
        ws.cell(row=r, column=2, value=len(lst))
        txt = ", ".join(str(i) for i in lst[:60]) + (" ..." if len(lst) > 60 else "")
        ws.cell(row=r, column=3, value=txt).alignment = Alignment(wrap_text=True)
        if lst:
            ws.cell(row=r, column=2).fill = FOND_DOUX
        r += 1
    ws.cell(row=r + 1, column=1, value=f"Questionnaires exclus (illisibles) : {exclus}")
    ws.cell(row=r + 2, column=1,
            value=f"Questionnaires non vérifiés au moment de l'export : {non_verifies}")
    styler_zone(ws, 2, r - 1, 3)


# ==========================================================================
# 4. Graphiques
# ==========================================================================
def onglet_graphiques(wb, derniere_likert, reperes_cat, nb_sections):
    ws = wb.create_sheet("Graphiques")
    ws.column_dimensions["A"].width = 3

    # --- Moyenne par item -------------------------------------------------
    g1 = BarChart()
    g1.type, g1.style = "bar", 10
    g1.title = "Score moyen par affirmation (1 = pas du tout d'accord, 5 = tout à fait)"
    g1.y_axis.title, g1.x_axis.title = "Affirmation", "Moyenne"
    g1.add_data(Reference(wb["Stats échelles"], min_col=6, min_row=1,
                          max_row=derniere_likert), titles_from_data=True)
    g1.set_categories(Reference(wb["Stats échelles"], min_col=3, min_row=2,
                                max_row=derniere_likert))
    g1.height, g1.width = 44, 26
    ws.add_chart(g1, "B2")

    # --- Distribution 1-5 empilée ----------------------------------------
    g2 = BarChart()
    g2.type, g2.grouping, g2.overlap = "bar", "percentStacked", 100
    g2.title = "Répartition des réponses par affirmation (en %)"
    g2.add_data(Reference(wb["Stats échelles"], min_col=14, max_col=18, min_row=1,
                          max_row=derniere_likert), titles_from_data=True)
    g2.set_categories(Reference(wb["Stats échelles"], min_col=3, min_row=2,
                                max_row=derniere_likert))
    g2.height, g2.width = 44, 26
    ws.add_chart(g2, "S2")

    # --- Score par section ------------------------------------------------
    g3 = BarChart()
    g3.type, g3.style = "col", 11
    g3.title = "Score moyen par section (items inversés recodés)"
    g3.y_axis.title = "Score moyen /5"
    g3.add_data(Reference(wb["Synthèse sections"], min_col=3, min_row=1,
                          max_row=nb_sections), titles_from_data=True)
    g3.set_categories(Reference(wb["Synthèse sections"], min_col=1, min_row=2,
                                max_row=nb_sections))
    g3.height, g3.width = 11, 26
    ws.add_chart(g3, "B92")

    # --- Variables signalétiques -----------------------------------------
    ancre = ["S92", "S113", "S134", "B113", "B134", "B155"]
    a_tracer = [("sexe", "Répartition par sexe", "pie"),
                ("age", "Répartition par tranche d'âge", "col"),
                ("anciennete", "Ancienneté dans l'enseignement", "col"),
                ("statut", "Statut professionnel", "col"),
                ("discipline", "Discipline enseignée", "col"),
                ("q47", "Type de formation préféré", "col")]
    feuille = wb["Stats catégorielles"]
    for (code, titre, genre), pos in zip(a_tracer, ancre):
        if code not in reperes_cat:
            continue
        d, f = reperes_cat[code]
        g = PieChart() if genre == "pie" else BarChart()
        if genre != "pie":
            g.type = "col"
        g.title = titre
        g.add_data(Reference(feuille, min_col=4, min_row=d, max_row=f))
        g.set_categories(Reference(feuille, min_col=3, min_row=d, max_row=f))
        g.height, g.width = 10, 14
        ws.add_chart(g, pos)


# ==========================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--inclure-non-verifies", action="store_true")
    args = ap.parse_args()

    fiches, non_verifies, exclus = charger_tout(args.inclure_non_verifies)
    if not fiches:
        sys.exit("Aucun questionnaire à exporter. "
                 "Lance etape2_lire.py puis etape3_verifier.py "
                 "(ou ajoute --inclure-non-verifies).")
    if non_verifies and not args.inclure_non_verifies:
        print(f"Attention : {non_verifies} questionnaire(s) non vérifié(s) ont été ignorés.")

    lignes = [ligne_plate(d) for d in fiches]
    n = len(lignes)
    print(f"{n} questionnaires exportés.")

    wb = Workbook()
    wb.remove(wb.active)
    colmap = onglet_brut(wb, lignes)
    onglet_dictionnaire(wb)
    groupes, n_cols_scores = onglet_scores(wb, colmap, n)
    derniere_likert = onglet_stats_likert(wb, colmap, n)
    reperes = onglet_stats_categorielles(wb, colmap, n)
    nb_sections = onglet_synthese(wb, groupes, n_cols_scores, lignes, n)
    onglet_ouvertes(wb, lignes)
    onglet_controles(wb, lignes, non_verifies, exclus)
    onglet_graphiques(wb, derniere_likert, reperes, nb_sections)

    config.SORTIE_EXCEL.parent.mkdir(parents=True, exist_ok=True)
    wb.save(config.SORTIE_EXCEL)
    print(f"Excel  -> {config.SORTIE_EXCEL}")

    # CSV jumeau, pratique pour pandas / R / SPSS
    import csv
    cols = colonnes_excel()
    with open(config.SORTIE_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(lignes)
    print(f"CSV    -> {config.SORTIE_CSV}")


if __name__ == "__main__":
    main()
