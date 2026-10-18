"""
Structure du questionnaire « La formation continue des enseignants du secondaire
général au Bénin ». C'est LE fichier de référence : tout le reste (extraction,
vérification, export Excel) se construit à partir d'ici.

Si tu modifies un libellé ou une modalité, tout le pipeline suit automatiquement.
"""

from dataclasses import dataclass, field
from typing import List, Optional

# --------------------------------------------------------------------------
# Échelle de Likert utilisée dans tout le questionnaire
# --------------------------------------------------------------------------
LIKERT_LABELS = {
    1: "Pas du tout d'accord",
    2: "Peu d'accord",
    3: "Neutre",
    4: "D'accord",
    5: "Tout à fait d'accord",
}


@dataclass
class Champ:
    code: str                       # nom de la colonne Excel
    libelle: str                    # texte exact de la question / affirmation
    page: int                       # page du questionnaire (1 à 4)
    section: str                    # section du questionnaire
    type: str                       # "likert" | "single" | "multi" | "text"
    options: List[str] = field(default_factory=list)
    inverse: bool = False           # item formulé négativement (recodage 6-x)
    note: Optional[str] = None      # aide pour le modèle de lecture


def L(code, libelle, page, section, inverse=False, note=None):
    """Raccourci pour un item d'échelle 1-5."""
    return Champ(code, libelle, page, section, "likert", inverse=inverse, note=note)


# --------------------------------------------------------------------------
# Sections (l'ordre définit l'ordre des colonnes et des onglets)
# --------------------------------------------------------------------------
S_INFO = "I. Informations générales"
S_AME = "I bis. Volet spécifique AME"
S_POL = "II. Politique et organisation de la formation continue"
S_PERT = "III. Pertinence, approche pédagogique et réflexivité"
S_ACC = "IV. Pratiques réelles d'accompagnement"
S_CONTR = "V. Contraintes et obstacles majeurs"
S_DUR = "VI. Développement professionnel durable"
S_NUM = "VII. Numérique, évaluation des formateurs et auto-évaluation"
S_TECH = "VIII. Conditions technologiques (hybride / à distance)"
S_PERSP = "IX. Perspectives, modalités souhaitées et propositions"

# Sections dont on calcule un score moyen (uniquement des items 1-5)
SECTIONS_SCORE = [S_AME, S_POL, S_PERT, S_ACC, S_CONTR, S_DUR, S_NUM, S_TECH]


# --------------------------------------------------------------------------
# PAGE 1 — Informations générales
# --------------------------------------------------------------------------
PAGE1 = [
    Champ("etablissement", "Établissement", 1, S_INFO, "text",
          note="Manuscrit. Recopier tel quel, sans corriger l'orthographe."),
    Champ("ville", "Ville", 1, S_INFO, "text", note="Manuscrit."),
    Champ("departement", "Département", 1, S_INFO, "text",
          note="Manuscrit, souvent en majuscules (ex : OUEME, ATLANTIQUE, BORGOU)."),
    Champ("type_etablissement", "Type d'établissement", 1, S_INFO, "single",
          ["Public", "Privé"]),
    Champ("sexe", "Sexe", 1, S_INFO, "single", ["Masculin", "Féminin"]),
    Champ("age", "Âge", 1, S_INFO, "single",
          ["moins de 30 ans", "30-39 ans", "40-49 ans", "50 ans et plus"]),
    Champ("anciennete", "Ancienneté dans l'enseignement", 1, S_INFO, "single",
          ["moins de 5 ans", "5-10 ans", "11-20 ans", "plus de 20 ans"]),
    Champ("statut", "Statut professionnel", 1, S_INFO, "single",
          ["FE", "ACDPE", "AME", "Autre"]),
    Champ("statut_autre", "Statut professionnel — précision « Autre »", 1, S_INFO, "text"),
    Champ("diplome_pro", "Diplôme professionnel", 1, S_INFO, "single",
          ["CAPES", "BAPES", "Licence Pro", "Master Pro"]),
    Champ("diplome_academique", "Diplôme académique", 1, S_INFO, "single",
          ["BAC", "Licence", "Maîtrise", "Master"]),
    Champ("discipline", "Discipline enseignée (groupe)", 1, S_INFO, "single",
          ["Lettres", "Sciences et Technologies", "Sciences Humaines", "Autre"],
          note="Lettres = Français, Anglais, Allemand, Espagnol, Philosophie. "
               "Sciences et Technologies = Maths, PCT, SVT, Informatique. "
               "Sciences Humaines = Histoire-Géo, Éducation Civique, EPS."),
    Champ("discipline_autre", "Discipline — précision « Autre »", 1, S_INFO, "text"),
    Champ("formation_12mois",
          "Avez-vous participé à au moins une activité de formation continue au cours des 12 derniers mois ?",
          1, S_INFO, "single", ["Oui", "Non"]),
    Champ("nb_jours_formation",
          "Si oui, combien de jours de formation avez-vous suivis (cumulés) ?",
          1, S_INFO, "single",
          ["1 à 2 jours", "3 à 5 jours", "6 à 10 jours", "Plus de 10 jours"]),
    Champ("formation_distance",
          "Avez-vous déjà suivi une formation à distance ou hybride, même partiellement ?",
          1, S_INFO, "single", ["Oui", "Non"]),
    Champ("satisfaction_distance",
          "Si oui, niveau de satisfaction globale vis-à-vis de cette expérience",
          1, S_INFO, "single", ["1", "2", "3", "4", "5", "Sans objet"]),
    Champ("equipements",
          "Quel(s) équipement(s) personnel(s) utilisez-vous ou pourriez-vous utiliser pour vous former ?",
          1, S_INFO, "multi",
          ["Smartphone", "Ordinateur/tablette", "Connexion Internet à domicile",
           "Aucun de ces équipements"],
          note="Plusieurs réponses possibles."),
]

# --------------------------------------------------------------------------
# PAGE 2 — Volet AME, sections II, III et début de IV
# --------------------------------------------------------------------------
PAGE2 = [
    L("ame_a", "Je me sens suffisamment accompagné(e) dans mon insertion professionnelle en tant qu'AME.",
      2, S_AME, note="À remplir uniquement si statut = AME. Sinon laisser null."),
    L("ame_b", "Mon statut d'AME limite mon accès aux mêmes formations que les enseignants titulaires.",
      2, S_AME, inverse=True),
    L("ame_c", "Un accompagnement spécifique aux AME, au-delà des formations générales, serait utile à mon intégration professionnelle.",
      2, S_AME),

    L("q1", "Je suis régulièrement informé à l'avance des opportunités de formation.", 2, S_POL),
    L("q2", "Les formations sont organisées de façon ponctuelle, sans continuité dans l'année.", 2, S_POL, inverse=True),
    L("q3", "Les thèmes de formation sont choisis sans me consulter (décision descendante).", 2, S_POL, inverse=True),
    L("q4", "Les Animations Pédagogiques de Zone (APZ) répondent à mes besoins réels.", 2, S_POL),
    L("q5", "Les Animations Pédagogiques d'Établissement (APE/APH) sont utiles pour ma pratique quotidienne.", 2, S_POL),
    L("q6", "Les formations dépendent trop de projets extérieurs (ONG, bailleurs, projets temporaires).", 2, S_POL, inverse=True),
    Champ("q7", "Qui initie généralement les formations que vous suivez ?", 2, S_POL, "multi",
          ["Ministère", "Inspection", "Établissement", "ONG/Partenaires",
           "Initiative personnelle", "Autre"],
          note="Plusieurs choix possibles."),

    L("q8", "Les formations répondent à mes besoins réels rencontrés en classe.", 2, S_PERT),
    L("q9", "Les contenus sont trop théoriques et éloignés de ma réalité de terrain.", 2, S_PERT, inverse=True),
    L("q10", "Les formations sont adaptées à l'Approche Par Compétences (APC) attendue dans les programmes.", 2, S_PERT),
    L("q11", "Les formateurs valorisent mon expérience professionnelle.", 2, S_PERT),
    L("q12", "On utilise des méthodes actives (ateliers, études de cas, résolution de problèmes) pendant les formations.", 2, S_PERT),
    L("q13", "Je repars avec des outils ou fiches pratiques immédiatement utilisables en classe.", 2, S_PERT),
    L("q14", "Les formations m'amènent à analyser et réfléchir sur ma propre pratique de classe.", 2, S_PERT),

    Champ("q15", "Quelle(s) forme(s) d'accompagnement avez-vous reçue(s) après vos dernières formations ?",
          2, S_ACC, "multi",
          ["Visite de classe en situation", "Entretien individuel avec un encadreur",
           "Mentorat/tutorat par un pair",
           "Réunion collective de suivi (conseil d'enseignement, animation pédagogique)",
           "Aucun accompagnement reçu"],
          note="Plusieurs réponses possibles."),
    Champ("q16", "À quelle fréquence avez-vous bénéficié d'un accompagnement de proximité au cours de la dernière année scolaire ?",
          2, S_ACC, "single",
          ["Jamais", "1 à 2 fois", "3 à 5 fois", "Plus de 5 fois"]),
    L("q17", "Un suivi rigoureux est assuré après chaque formation (visite en classe, entretien individuel).", 2, S_ACC),
    L("q18", "Je bénéficie d'un accompagnement de proximité (conseiller pédagogique, tuteur, pair mentor).", 2, S_ACC),
    L("q19", "Les inspecteurs ou encadreurs assurent un suivi efficace après la formation.", 2, S_ACC),
]

# --------------------------------------------------------------------------
# PAGE 3 — Suite IV, sections V, VI, VII
# --------------------------------------------------------------------------
PAGE3 = [
    L("q20", "Les effets des formations sur les apprentissages des élèves sont régulièrement évalués.", 3, S_ACC),
    L("q21", "Sans accompagnement post-formation, j'ai du mal à mettre en œuvre ce que j'ai appris.", 3, S_ACC, inverse=True),
    L("q22", "Un système de tutorat ou mentorat entre enseignants serait utile.", 3, S_ACC),
    L("q23", "L'accompagnement reçu porte sur des éléments concrets de ma pratique de classe, et non sur de simples formalités administratives.", 3, S_ACC),
    L("q24", "Les acteurs chargés de m'accompagner (inspecteur, conseiller pédagogique, pair mentor) sont clairement identifiés et disponibles.", 3, S_ACC),

    L("q25", "Ma charge de travail limite fortement ma participation aux formations.", 3, S_CONTR, inverse=True),
    L("q26", "Les contraintes financières (transport, frais divers) freinent mon accès aux formations.", 3, S_CONTR, inverse=True),
    L("q27", "Le manque de moyens matériels (salles, équipements, documents) est un frein important.", 3, S_CONTR, inverse=True),
    L("q28", "L'éloignement géographique des lieux de formation est un obstacle majeur.", 3, S_CONTR, inverse=True),
    L("q29", "Le manque de motivation personnelle ou collective impacte mon engagement.", 3, S_CONTR, inverse=True),
    L("q30", "Les animateurs ou formateurs sont parfois insuffisamment formés.", 3, S_CONTR, inverse=True),
    L("q31", "Le message des formations est déformé ou affaibli lorsqu'il passe de l'inspecteur au conseiller pédagogique puis à moi.", 3, S_CONTR, inverse=True),

    L("q32", "La formation continue actuelle produit des changements durables dans ma pratique enseignante.", 3, S_DUR),
    L("q33", "Les formations ne dépendent pas uniquement de projets temporaires (elles pourraient continuer sans eux).", 3, S_DUR),
    L("q34", "La formation continue contribue réellement à un enseignement de qualité.", 3, S_DUR),
    L("q35", "Je me sens capable d'adapter ma pratique sans accompagnement extérieur après une formation.", 3, S_DUR),

    L("q36", "Les formations intègrent efficacement les outils numériques (TIC, plateformes, ressources en ligne).", 3, S_NUM),
    L("q37", "J'ai accès à des ressources numériques pour me former.", 3, S_NUM),
    L("q38", "Les formateurs sont qualifiés pour les contenus qu'ils dispensent.", 3, S_NUM),
    L("q39", "Je peux évaluer anonymement les formateurs après chaque session.", 3, S_NUM),
    L("q40", "Je pratique une auto-évaluation régulière de mes compétences après les formations.", 3, S_NUM),

    Champ("q41", "À quelle fréquence utilisez-vous le numérique pour vous former ?", 3, S_NUM, "single",
          ["Jamais", "Rarement", "Parfois", "Souvent", "Très souvent"]),
    Champ("q42", "Quels obstacles numériques rencontrez-vous ?", 3, S_NUM, "multi",
          ["Accès limité à Internet", "Matériel insuffisant",
           "Manque de compétences numériques", "Pas d'obstacle majeur"],
          note="Plusieurs choix possibles."),
]

# --------------------------------------------------------------------------
# PAGE 4 — Sections VIII et IX
# --------------------------------------------------------------------------
PAGE4 = [
    L("q43", "Je me sens capable d'utiliser seul(e) une plateforme de formation en ligne.", 4, S_TECH),
    L("q44", "Je suis à l'aise pour suivre un module de formation à distance sans accompagnement présentiel.", 4, S_TECH),
    L("q45", "Une formation hybride (présentiel + distanciel) serait adaptée à mes conditions de travail actuelles.", 4, S_TECH),
    L("q46", "Je serais prêt(e) à m'investir activement dans un dispositif de formation hybride si celui-ci était mis en place.", 4, S_TECH),

    Champ("q47", "Quel type de formation préférez-vous ?", 4, S_PERSP, "single",
          ["Présentiel pur", "À distance (en ligne)", "Mixte (présentiel + à distance)", "Autre"],
          note="Une seule réponse."),
    Champ("q48", "Quelle durée de formation préférez-vous ?", 4, S_PERSP, "single",
          ["Une journée complète", "Deux jours consécutifs",
           "Modules répartis sur plusieurs semaines",
           "Formation longue (plusieurs semaines continues)"]),
    Champ("q49", "Souhaitez-vous un système de tutorat post-formation ?", 4, S_PERSP, "single",
          ["Oui, absolument", "Plutôt oui", "Indifférent", "Plutôt non", "Non, pas nécessaire"]),

    Champ("q50a", "Principal obstacle actuel à un accompagnement réel et efficace des enseignants au Bénin",
          4, S_PERSP, "text",
          note="Réponse manuscrite. Transcrire fidèlement ; si illisible, mettre le "
               "texte lisible suivi de [illisible]. Si rien n'est écrit, null."),
    Champ("q50b", "Amélioration concrète et prioritaire proposée pour rapprocher les dispositifs prescrits des pratiques réelles",
          4, S_PERSP, "text", note="Réponse manuscrite."),
    Champ("q50c", "Enjeu principal que l'accompagnement des enseignants doit absolument résoudre",
          4, S_PERSP, "text", note="Réponse manuscrite."),
    Champ("q50d", "Conditions pour que la formation à distance/hybride devienne un levier crédible",
          4, S_PERSP, "text", note="Réponse manuscrite."),
]

CHAMPS: List[Champ] = PAGE1 + PAGE2 + PAGE3 + PAGE4
PAR_PAGE = {1: PAGE1, 2: PAGE2, 3: PAGE3, 4: PAGE4}
PAR_CODE = {c.code: c for c in CHAMPS}

NB_PAGES = 4
CODES_LIKERT = [c.code for c in CHAMPS if c.type == "likert"]
CODES_TEXTE = [c.code for c in CHAMPS if c.type == "text"]


def champs_de_section(section: str) -> List[Champ]:
    return [c for c in CHAMPS if c.section == section]


def colonnes_excel() -> List[str]:
    """Colonnes du tableau de données brutes (les 'multi' sont éclatés en 0/1)."""
    cols = ["id_questionnaire", "pages_source"]
    for c in CHAMPS:
        if c.type == "multi":
            for opt in c.options:
                cols.append(f"{c.code}__{_slug(opt)}")
            cols.append(c.code)          # version texte concaténée
        else:
            cols.append(c.code)
    cols += ["a_verifier", "verifie", "remarques"]
    return cols


def _slug(txt: str) -> str:
    import re
    import unicodedata
    txt = unicodedata.normalize("NFKD", txt).encode("ascii", "ignore").decode()
    txt = re.sub(r"[^A-Za-z0-9]+", "_", txt).strip("_").lower()
    return txt[:28]


if __name__ == "__main__":
    print(f"{len(CHAMPS)} champs, dont {len(CODES_LIKERT)} items d'échelle 1-5.")
    for p in range(1, NB_PAGES + 1):
        print(f"  page {p} : {len(PAR_PAGE[p])} champs")
    print(f"{len(colonnes_excel())} colonnes dans l'onglet « Données brutes ».")
