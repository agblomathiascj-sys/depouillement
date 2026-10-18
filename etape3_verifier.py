"""
ÉTAPE 3 — Écran de vérification humaine.

    streamlit run etape3_verifier.py

Le scan s'affiche à gauche, les réponses lues par l'IA à droite, modifiables.
Les champs signalés comme incertains sont surlignés : par défaut tu ne vois
que les questionnaires qui en contiennent, ce qui ramène le contrôle manuel
à quelques minutes par lot plutôt qu'à 400 saisies complètes.
"""

import json

import streamlit as st

import config
from codebook import PAR_PAGE, NB_PAGES, LIKERT_LABELS

st.set_page_config(page_title="Dépouillement", layout="wide")

SANS_REPONSE = "— (non renseigné)"


# --------------------------------------------------------------------------
def charger(qid: int) -> dict:
    """Charge la version validée si elle existe, sinon la lecture IA."""
    val = config.DOSSIER_VALIDE / f"q{qid:05d}.json"
    src = val if val.exists() else config.DOSSIER_JSON / f"q{qid:05d}.json"
    return json.loads(src.read_text(encoding="utf-8"))


@st.cache_data(show_spinner=False)
def inventaire() -> list:
    fiches = []
    for chemin in sorted(config.DOSSIER_JSON.glob("q*.json")):
        d = json.loads(chemin.read_text(encoding="utf-8"))
        qid = d["id_questionnaire"]
        valide = (config.DOSSIER_VALIDE / f"q{qid:05d}.json").exists()
        fiches.append({"id": qid, "a_verifier": d.get("a_verifier", False),
                       "verifie": valide,
                       "n_incert": len(d.get("incertitudes", []))})
    return fiches


def sauvegarder(donnees: dict):
    donnees["verifie"] = True
    donnees["a_verifier"] = False
    chemin = config.DOSSIER_VALIDE / f"q{donnees['id_questionnaire']:05d}.json"
    chemin.write_text(json.dumps(donnees, ensure_ascii=False, indent=2),
                      encoding="utf-8")
    inventaire.clear()


# --------------------------------------------------------------------------
fiches = inventaire()
if not fiches:
    st.error("Aucune lecture trouvée. Lance d'abord : python etape2_lire.py")
    st.stop()

with st.sidebar:
    st.header("Dépouillement")
    total = len(fiches)
    faits = sum(f["verifie"] for f in fiches)
    st.progress(faits / total, text=f"{faits} / {total} validés")

    filtre = st.radio("Afficher", ["À vérifier seulement", "Non validés", "Tous"])
    if filtre == "À vérifier seulement":
        liste = [f for f in fiches if f["a_verifier"] and not f["verifie"]]
    elif filtre == "Non validés":
        liste = [f for f in fiches if not f["verifie"]]
    else:
        liste = fiches

    if not liste:
        st.success("Plus rien à vérifier dans ce filtre.")
        st.stop()

    etiquettes = [f"q{f['id']:05d}" + (" ✓" if f["verifie"] else "")
                  + (f"  ({f['n_incert']})" if f["n_incert"] else "")
                  for f in liste]
    idx = st.selectbox("Questionnaire", range(len(liste)),
                       format_func=lambda i: etiquettes[i],
                       index=min(st.session_state.get("idx", 0), len(liste) - 1))
    st.session_state["idx"] = idx
    qid = liste[idx]["id"]

    st.caption("Astuce : garde le filtre « À vérifier seulement », "
               "valide, passe au suivant.")

donnees = charger(qid)
reponses = donnees.get("reponses", {})
incertains = set(donnees.get("incertitudes", []))

st.subheader(f"Questionnaire q{qid:05d}")
if donnees.get("anomalies"):
    for a in donnees["anomalies"]:
        st.warning(a)

pages_source = donnees.get("pages_source", [])
onglets = st.tabs([f"Page {i}" for i in range(1, len(pages_source) + 1)])

nouvelles = dict(reponses)

for position, onglet in enumerate(onglets, start=1):
    page = min(position, NB_PAGES)
    with onglet:
        col_img, col_form = st.columns([1.1, 1])

        with col_img:
            chemin = config.DOSSIER_IMAGES / f"page_{pages_source[position-1]:05d}.jpg"
            if chemin.exists():
                try:
                    st.image(str(chemin), use_container_width=True)
                except TypeError:            # Streamlit < 1.40
                    st.image(str(chemin), use_column_width=True)
            else:
                st.error(f"Image manquante : {chemin.name}")

        with col_form:
            for c in PAR_PAGE[page]:
                marque = "🔴 " if c.code in incertains else ""
                valeur = reponses.get(c.code)
                cle = f"{qid}_{c.code}"

                if c.type == "likert":
                    choix = [SANS_REPONSE] + [f"{k} — {v}" for k, v in LIKERT_LABELS.items()]
                    defaut = 0 if valeur in (None, "") else int(valeur)
                    sel = st.selectbox(f"{marque}**{c.code}** · {c.libelle}",
                                       choix, index=defaut, key=cle)
                    nouvelles[c.code] = None if sel == SANS_REPONSE else int(sel.split(" — ")[0])

                elif c.type == "single":
                    choix = [SANS_REPONSE] + c.options
                    defaut = choix.index(valeur) if valeur in choix else 0
                    sel = st.selectbox(f"{marque}**{c.code}** · {c.libelle}",
                                       choix, index=defaut, key=cle)
                    nouvelles[c.code] = None if sel == SANS_REPONSE else sel

                elif c.type == "multi":
                    defaut = [v for v in (valeur or []) if v in c.options]
                    nouvelles[c.code] = st.multiselect(
                        f"{marque}**{c.code}** · {c.libelle}",
                        c.options, default=defaut, key=cle)

                else:
                    nouvelles[c.code] = st.text_area(
                        f"{marque}**{c.code}** · {c.libelle}",
                        value=valeur or "", height=80, key=cle) or None

st.divider()
commentaire = st.text_input("Remarque sur ce questionnaire (facultatif)",
                            value=donnees.get("commentaire", ""))

c1, c2, c3 = st.columns([1, 1, 3])
if c1.button("✅ Valider et suivant", type="primary", use_container_width=True):
    donnees["reponses"] = nouvelles
    donnees["commentaire"] = commentaire
    sauvegarder(donnees)
    st.session_state["idx"] = st.session_state.get("idx", 0)  # la liste se recalcule
    st.rerun()

if c2.button("🚩 Écarter (illisible)", use_container_width=True):
    donnees["reponses"] = nouvelles
    donnees["commentaire"] = commentaire or "écarté"
    donnees["exclu"] = True
    sauvegarder(donnees)
    st.rerun()

c3.caption("Une fois tout validé :  python etape4_exporter.py")
