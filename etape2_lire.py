"""
ÉTAPE 2 — Lecture automatique des questionnaires scannés.

    set ANTHROPIC_API_KEY=sk-ant-...        (Windows, cmd)
    $env:ANTHROPIC_API_KEY="sk-ant-..."     (Windows, PowerShell)

    python etape2_lire.py --limite 5     # commence TOUJOURS par un essai sur 5
    python etape2_lire.py                # tout le lot
    python etape2_lire.py --refaire 37 41  # relit uniquement ces questionnaires

Chaque page est envoyée séparément avec la liste exacte des champs à y lire :
le modèle ne peut donc pas inventer une variable ni se tromper de section.
La sortie est un JSON par questionnaire dans travail/lecture/ — le script est
interruptible et reprend où il s'est arrêté (rien n'est refacturé deux fois).
"""

import argparse
import json
import random
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import config
from codebook import PAR_PAGE, NB_PAGES, LIKERT_LABELS
from etape1_preparer import image_b64

try:
    from anthropic import Anthropic
except ImportError:
    sys.exit("SDK manquant.  ->  pip install --user anthropic")


# --------------------------------------------------------------------------
# Construction du schéma d'outil (garantit un JSON valide en sortie)
# --------------------------------------------------------------------------
def schema_page(page: int) -> dict:
    props, codes = {}, []
    for c in PAR_PAGE[page]:
        codes.append(c.code)
        if c.type == "likert":
            props[c.code] = {"type": ["integer", "null"], "minimum": 1, "maximum": 5,
                             "description": c.libelle}
        elif c.type == "single":
            props[c.code] = {"type": ["string", "null"], "enum": c.options + [None],
                             "description": c.libelle}
        elif c.type == "multi":
            props[c.code] = {"type": "array", "items": {"enum": c.options},
                             "description": c.libelle + " (plusieurs réponses)"}
        else:
            props[c.code] = {"type": ["string", "null"], "description": c.libelle}
    return {
        "type": "object",
        "properties": {
            "page_detectee": {"type": "integer", "minimum": 1, "maximum": 4,
                              "description": "Numéro de page réellement visible sur le scan."},
            "reponses": {"type": "object", "properties": props,
                         "required": codes, "additionalProperties": False},
            "incertitudes": {"type": "array", "items": {"enum": codes},
                             "description": "Codes dont la lecture n'est pas certaine."},
            "remarques": {"type": ["string", "null"],
                          "description": "Anomalie éventuelle : page illisible, "
                                         "double coche, rature, cadre coupé..."},
        },
        "required": ["page_detectee", "reponses", "incertitudes"],
        "additionalProperties": False,
    }


def consigne_page(page: int) -> str:
    lignes = [
        "Tu dépouilles un questionnaire papier scanné (recherche universitaire, Bénin).",
        f"Ce scan est la PAGE {page} sur {NB_PAGES}.",
        "",
        "Règles de lecture :",
        "- Les cases sont cochées à la main : croix, coche, trait, case noircie. "
        "Une marque même débordante ou décalée compte comme cochée.",
        "- Pour les grilles 1 à 5, la colonne coche se lit par la position horizontale "
        "de la marque : 1 = Pas du tout d'accord ... 5 = Tout à fait d'accord.",
        "- Si une ligne n'a AUCUNE marque, réponds null pour ce code (ne devine jamais).",
        "- Si deux cases sont cochées sur une même ligne à choix unique, choisis la marque "
        "la plus nette et ajoute le code à « incertitudes ».",
        "- Ajoute le code à « incertitudes » dès que tu hésites entre deux colonnes voisines, "
        "que la marque est ambiguë ou l'écriture illisible. Mieux vaut trop signaler que pas assez.",
        "- Le texte manuscrit est transcrit tel quel, sans reformuler ni corriger.",
        "",
        f"Échelle : {', '.join(f'{k} = {v}' for k, v in LIKERT_LABELS.items())}",
        "",
        "Champs à lire sur cette page, dans l'ordre où ils apparaissent :",
    ]
    for c in PAR_PAGE[page]:
        détail = ""
        if c.type == "likert":
            détail = " [échelle 1-5]"
        elif c.type == "single":
            détail = f" [une seule : {' / '.join(c.options)}]"
        elif c.type == "multi":
            détail = f" [plusieurs : {' / '.join(c.options)}]"
        else:
            détail = " [texte manuscrit]"
        lignes.append(f"  • {c.code} — {c.libelle}{détail}"
                      + (f"  ({c.note})" if c.note else ""))
    lignes.append("")
    lignes.append("Appelle l'outil enregistrer_reponses avec TOUS les codes ci-dessus.")
    return "\n".join(lignes)


CONSIGNES = {p: consigne_page(p) for p in range(1, NB_PAGES + 1)}
SCHEMAS = {p: schema_page(p) for p in range(1, NB_PAGES + 1)}

_usage = {"in": 0, "out": 0, "cache_read": 0, "cache_write": 0}
_verrou = threading.Lock()


def lire_page(client, num_page_pdf: int, page_attendue: int) -> dict:
    """Envoie une page à l'API et renvoie le dict de réponses."""
    for tentative in range(config.MAX_TENTATIVES):
        try:
            rep = client.messages.create(
                model=config.MODELE,
                max_tokens=2000,
                system=[{"type": "text", "text": CONSIGNES[page_attendue],
                         "cache_control": {"type": "ephemeral"}}],
                tools=[{"name": "enregistrer_reponses",
                        "description": "Enregistre les réponses lues sur le scan.",
                        "input_schema": SCHEMAS[page_attendue]}],
                tool_choice={"type": "tool", "name": "enregistrer_reponses"},
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64",
                                                 "media_type": "image/jpeg",
                                                 "data": image_b64(num_page_pdf)}},
                    {"type": "text", "text": "Dépouille ce scan."},
                ]}],
            )
            with _verrou:
                u = rep.usage
                _usage["in"] += u.input_tokens
                _usage["out"] += u.output_tokens
                _usage["cache_read"] += getattr(u, "cache_read_input_tokens", 0) or 0
                _usage["cache_write"] += getattr(u, "cache_creation_input_tokens", 0) or 0
            for bloc in rep.content:
                if bloc.type == "tool_use":
                    return bloc.input
            raise RuntimeError("pas de tool_use dans la réponse")
        except Exception as err:  # 429, 529, réseau, JSON incomplet...
            if tentative == config.MAX_TENTATIVES - 1:
                return {"erreur": f"{type(err).__name__}: {err}",
                        "reponses": {}, "incertitudes": [], "page_detectee": 0}
            time.sleep(min(2 ** tentative, 30) + random.random())


def traiter_questionnaire(client, groupe: dict) -> dict:
    qid = groupe["id"]
    resultat = {"id_questionnaire": qid, "pages_source": groupe["pages"],
                "reponses": {}, "incertitudes": [], "anomalies": []}

    for position, num_page_pdf in enumerate(groupe["pages"], start=1):
        page_attendue = min(position, NB_PAGES)
        brut = lire_page(client, num_page_pdf, page_attendue)
        if brut.get("erreur"):
            resultat["anomalies"].append(f"page {position} : {brut['erreur']}")
            continue
        if brut.get("page_detectee") not in (0, page_attendue):
            resultat["anomalies"].append(
                f"page {position} : le scan semble être la page "
                f"{brut['page_detectee']} du questionnaire (décalage possible)")
        resultat["reponses"].update(brut.get("reponses", {}))
        resultat["incertitudes"] += brut.get("incertitudes", [])
        if brut.get("remarques"):
            resultat["anomalies"].append(f"page {position} : {brut['remarques']}")

    resultat["a_verifier"] = bool(
        len(resultat["incertitudes"]) >= config.SEUIL_INCERTITUDES
        or resultat["anomalies"]
        or not groupe.get("complet", True))
    resultat["verifie"] = False
    return resultat


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limite", type=int, help="ne traiter que les N premiers")
    ap.add_argument("--refaire", type=int, nargs="*", help="relire ces id de questionnaires")
    args = ap.parse_args()

    if not config.FICHIER_GROUPES.exists():
        sys.exit("Lance d'abord : python etape1_preparer.py")
    groupes = json.loads(config.FICHIER_GROUPES.read_text(encoding="utf-8"))

    if args.refaire:
        groupes = [g for g in groupes if g["id"] in args.refaire]
    else:
        groupes = [g for g in groupes
                   if not (config.DOSSIER_JSON / f"q{g['id']:05d}.json").exists()]
        if args.limite:
            groupes = groupes[:args.limite]

    if not groupes:
        print("Rien à lire : tout est déjà dépouillé.")
        return

    print(f"Lecture de {len(groupes)} questionnaire(s) "
          f"({len(groupes) * NB_PAGES} pages) avec {config.MODELE}...")
    client = Anthropic()
    debut = time.time()
    faits = 0

    with ThreadPoolExecutor(max_workers=config.NB_THREADS) as ex:
        futurs = {ex.submit(traiter_questionnaire, client, g): g for g in groupes}
        for f in as_completed(futurs):
            res = f.result()
            chemin = config.DOSSIER_JSON / f"q{res['id_questionnaire']:05d}.json"
            chemin.write_text(json.dumps(res, ensure_ascii=False, indent=2),
                              encoding="utf-8")
            faits += 1
            drapeau = "  A VERIFIER" if res["a_verifier"] else ""
            print(f"  [{faits}/{len(groupes)}] q{res['id_questionnaire']:05d} "
                  f"({len(res['incertitudes'])} incertitude(s)){drapeau}")

    duree = time.time() - debut
    print(f"\nTerminé en {duree/60:.1f} min.")
    print(f"Tokens — entrée {_usage['in']:,} | sortie {_usage['out']:,} | "
          f"cache lu {_usage['cache_read']:,} | cache écrit {_usage['cache_write']:,}")
    print("Multiplie par le tarif du modèle (console.anthropic.com) pour le coût réel.")
    print("Étape suivante :  streamlit run etape3_verifier.py")


if __name__ == "__main__":
    main()
