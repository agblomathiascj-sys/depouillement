"""
Génère des questionnaires FICTIFS pour tester les étapes 3 et 4 sans API ni scans.

    python demo_donnees_fictives.py 40

Supprime ensuite le dossier travail/lecture avant de lancer le vrai dépouillement.
"""

import json
import random
import sys

import config
from codebook import CHAMPS

VILLES = ["Cotonou", "Porto-Novo", "Parakou", "Abomey", "Natitingou", "Lokossa"]
DEPTS = ["LITTORAL", "OUEME", "BORGOU", "ZOU", "ATACORA", "MONO"]


def un_questionnaire(qid: int) -> dict:
    r, incert = {}, []
    statut = random.choice(["FE", "ACDPE", "AME", "AME"])
    for c in CHAMPS:
        if c.code == "statut":
            r[c.code] = statut
        elif c.type == "likert":
            if c.code.startswith("ame_") and statut != "AME":
                r[c.code] = None
            else:
                r[c.code] = None if random.random() < 0.03 else random.choices(
                    [1, 2, 3, 4, 5], weights=[1, 2, 3, 3, 2])[0]
        elif c.type == "single":
            r[c.code] = None if random.random() < 0.05 else random.choice(c.options)
        elif c.type == "multi":
            r[c.code] = random.sample(c.options, k=random.randint(1, min(2, len(c.options))))
        elif c.code == "ville":
            r[c.code] = random.choice(VILLES)
        elif c.code == "departement":
            r[c.code] = random.choice(DEPTS)
        elif c.code == "etablissement":
            r[c.code] = f"CEG {random.randint(1, 9)} {random.choice(VILLES)}"
        elif c.code.startswith("q50"):
            r[c.code] = random.choice([None, "Manque de moyens et de suivi de proximité.",
                                       "Motiver les enseignants et former les encadreurs."])
        else:
            r[c.code] = None
        if c.type == "likert" and random.random() < 0.02:
            incert.append(c.code)
    return {"id_questionnaire": qid,
            "pages_source": [(qid - 1) * 4 + i for i in range(1, 5)],
            "reponses": r, "incertitudes": incert, "anomalies": [],
            "a_verifier": bool(incert), "verifie": False}


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 30
    for i in range(1, n + 1):
        chemin = config.DOSSIER_JSON / f"q{i:05d}.json"
        chemin.write_text(json.dumps(un_questionnaire(i), ensure_ascii=False, indent=2),
                          encoding="utf-8")
    print(f"{n} questionnaires fictifs écrits dans {config.DOSSIER_JSON}")
