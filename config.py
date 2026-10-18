"""Paramètres du dépouillement. C'est le seul fichier que tu devrais avoir à modifier."""

from pathlib import Path

# --- Dossiers -------------------------------------------------------------
BASE = Path(__file__).parent
PDF_SOURCE = BASE / "scans" / "questionnaires.pdf"   # le gros PDF de tous les questionnaires
DOSSIER_IMAGES = BASE / "travail" / "images"         # 1 image par page
DOSSIER_JSON = BASE / "travail" / "lecture"          # 1 JSON par questionnaire (sortie IA)
DOSSIER_VALIDE = BASE / "travail" / "valide"         # 1 JSON par questionnaire vérifié
FICHIER_GROUPES = BASE / "travail" / "groupes.json"  # découpage pages -> questionnaires
SORTIE_EXCEL = BASE / "resultats" / "depouillement.xlsx"
SORTIE_CSV = BASE / "resultats" / "donnees_brutes.csv"

# --- Lecture automatique --------------------------------------------------
MODELE = "claude-haiku-4-5-20251001"         # lecture des cases cochées
MODELE_RAPIDE = "claude-haiku-4-5-20251001"   # détection du numéro de page (option --detect)
DPI = 200                           # résolution de rendu des pages du PDF
LARGEUR_MAX = 1568                  # px : au-delà, l'API redimensionne de toute façon
QUALITE_JPEG = 85
NB_THREADS = 6                      # requêtes en parallèle
MAX_TENTATIVES = 5                  # retries en cas d'erreur 429 / 529

# --- Structure du lot -----------------------------------------------------
PAGES_PAR_QUESTIONNAIRE = 4

# --- Règles de qualité ----------------------------------------------------
# Un questionnaire est marqué « à vérifier » si le modèle signale au moins
# ce nombre d'incertitudes, ou s'il manque des réponses obligatoires.
SEUIL_INCERTITUDES = 1

for d in (DOSSIER_IMAGES, DOSSIER_JSON, DOSSIER_VALIDE,
          SORTIE_EXCEL.parent, PDF_SOURCE.parent):
    d.mkdir(parents=True, exist_ok=True)
