"""
ÉTAPE 1 — Découper le gros PDF et rendre chaque page en image.

    python etape1_preparer.py                 # découpage fixe : 4 pages = 1 questionnaire
    python etape1_preparer.py --detect        # détecte le n° de page (1-4) et recale le découpage

Le mode --detect coûte quelques centimes d'API mais évite la catastrophe classique :
un questionnaire scanné avec 3 ou 5 pages décale TOUT le reste du lot.
"""

import argparse
import base64
import io
import json
import sys
from concurrent.futures import ThreadPoolExecutor

import config

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF manquant.  ->  pip install --user pymupdf")

from PIL import Image


def rendre_pages() -> int:
    """Convertit chaque page du PDF en JPEG. Retourne le nombre de pages."""
    if not config.PDF_SOURCE.exists():
        sys.exit(f"PDF introuvable : {config.PDF_SOURCE}\n"
                 f"Place ton gros PDF à cet emplacement exact.")

    doc = fitz.open(config.PDF_SOURCE)
    zoom = config.DPI / 72
    matrice = fitz.Matrix(zoom, zoom)
    total = doc.page_count
    print(f"{total} pages dans le PDF  ->  rendu en images (DPI={config.DPI})")

    for i, page in enumerate(doc):
        chemin = config.DOSSIER_IMAGES / f"page_{i + 1:05d}.jpg"
        if chemin.exists():
            continue
        pix = page.get_pixmap(matrix=matrice)
        img = Image.open(io.BytesIO(pix.tobytes("ppm")))
        if img.width > config.LARGEUR_MAX:
            ratio = config.LARGEUR_MAX / img.width
            img = img.resize((config.LARGEUR_MAX, int(img.height * ratio)),
                             Image.LANCZOS)
        img.convert("RGB").save(chemin, "JPEG", quality=config.QUALITE_JPEG,
                                optimize=True)
        if (i + 1) % 50 == 0:
            print(f"  ... {i + 1}/{total}")
    doc.close()
    print(f"Images prêtes dans {config.DOSSIER_IMAGES}")
    return total


def image_b64(numero_page: int) -> str:
    chemin = config.DOSSIER_IMAGES / f"page_{numero_page:05d}.jpg"
    return base64.standard_b64encode(chemin.read_bytes()).decode()


def detecter_numero_page(client, numero_page: int) -> int:
    """Demande au modèle rapide de lire le numéro de page (1 à 4)."""
    from anthropic import Anthropic  # noqa: F401  (client déjà fourni)
    reponse = client.messages.create(
        model=config.MODELE_RAPIDE,
        max_tokens=16,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": "image/jpeg",
                            "data": image_b64(numero_page)}},
                {"type": "text", "text":
                    "Ce scan est une page d'un questionnaire de 4 pages.\n"
                    "Page 1 : titre « QUESTIONNAIRE AUX ENSEIGNANTS » + « I. INFORMATIONS GÉNÉRALES ».\n"
                    "Page 2 : « I bis. VOLET SPÉCIFIQUE — ENSEIGNANTS AME » en haut.\n"
                    "Page 3 : commence par l'affirmation n° 20, contient les sections V, VI, VII.\n"
                    "Page 4 : « VIII. CONDITIONS TECHNOLOGIQUES » + questions ouvertes.\n"
                    "Réponds uniquement par le chiffre : 1, 2, 3 ou 4."}
            ]
        }]
    )
    txt = "".join(b.text for b in reponse.content if b.type == "text").strip()
    for car in txt:
        if car in "1234":
            return int(car)
    return 0


def grouper_fixe(total: int) -> list:
    n = config.PAGES_PAR_QUESTIONNAIRE
    groupes = []
    for i in range(0, total, n):
        pages = list(range(i + 1, min(i + n, total) + 1))
        groupes.append({"id": len(groupes) + 1, "pages": pages,
                        "complet": len(pages) == n})
    return groupes


def grouper_detecte(total: int) -> list:
    from anthropic import Anthropic
    client = Anthropic()
    print("Détection du numéro de page sur chaque scan...")
    with ThreadPoolExecutor(max_workers=config.NB_THREADS) as ex:
        numeros = list(ex.map(lambda p: detecter_numero_page(client, p),
                              range(1, total + 1)))

    groupes, courant = [], []
    for page, num in zip(range(1, total + 1), numeros):
        if num == 1 and courant:
            groupes.append(courant)
            courant = []
        courant.append({"page": page, "num_detecte": num})
    if courant:
        groupes.append(courant)

    resultat = []
    for i, g in enumerate(groupes, start=1):
        pages = [x["page"] for x in g]
        detectes = [x["num_detecte"] for x in g]
        complet = detectes == [1, 2, 3, 4]
        resultat.append({"id": i, "pages": pages, "num_detectes": detectes,
                         "complet": complet})
    anomalies = [g for g in resultat if not g["complet"]]
    print(f"{len(resultat)} questionnaires détectés, {len(anomalies)} anomalie(s).")
    for g in anomalies[:20]:
        print(f"  ! questionnaire {g['id']} : pages {g['pages']} "
              f"-> numéros lus {g['num_detectes']}")
    return resultat


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--detect", action="store_true",
                    help="détecter le numéro de page au lieu de découper tous les 4")
    args = ap.parse_args()

    total = rendre_pages()
    groupes = grouper_detecte(total) if args.detect else grouper_fixe(total)

    config.FICHIER_GROUPES.parent.mkdir(parents=True, exist_ok=True)
    config.FICHIER_GROUPES.write_text(
        json.dumps(groupes, ensure_ascii=False, indent=2), encoding="utf-8")

    incomplets = [g["id"] for g in groupes if not g["complet"]]
    print(f"\n{len(groupes)} questionnaires  ->  {config.FICHIER_GROUPES}")
    if incomplets:
        print(f"À contrôler à la main : {incomplets[:30]}"
              f"{' ...' if len(incomplets) > 30 else ''}")


if __name__ == "__main__":
    main()
