"""Fabrique les documents d'essai, entièrement fictifs, à partir des gabarits (copiés dans tests/fixtures/).

Chaque document contient des écarts volontaires, attendus par tests/moteur.test.js.
Usage : python3 tests/fabriquer-fixtures.py
"""
from pathlib import Path

from docx import Document

ICI = Path(__file__).resolve().parent
GABARITS = ICI / "fixtures"
SORTIE = ICI / "fixtures"


def supprimer(p):
    p._element.getparent().remove(p._element)


def paragraphe_apres(p, texte, style=None):
    nouveau = p._parent.add_paragraph(texte, style=style)  # ajouté en fin de document…
    p._element.addnext(nouveau._element)                     # … puis déplacé
    return nouveau


def notice():
    doc = Document(GABARITS / "gabarit-demande-agrement.docx")
    ps = doc.paragraphs
    remplacements = {
        "[Intitulé du sujet du mémoire]": "Sujet fictif de démonstration",
        "[Prénom NOM du candidat]": "Prénom NOM",
        "[Adresse personnelle : au minimum code postal et ville]": "75000 Ville fictive",
        "[Mois et année de dépôt]": "novembre 2026",
    }
    for p in ps:
        if p.text in remplacements:
            p.text = remplacements[p.text]
        elif p.text.startswith("Aucune référence (nom, logo)"):
            supprimer(p)

    sections = {}
    for p in doc.paragraphs:
        if p.style.name == "Titre notice":
            sections[int(p.text[0])] = p
    textes = {
        1: "Le sujet est fictif. La problématique est formulée sans point d'interrogation.",
        2: "Les objectifs; ensuite  le Cabinet Fictif Durand, 1000€ et l'Etat (parenthèse ouverte. TODO",
        3: "Comparaison avec deux mémoires fictifs.",
        4: "La NPMQ s'applique ; la lutte contre le blanchiment (LAB) aussi. A partir de là, 12 dossiers.",
        5: "Les outils sont présentés ci-dessous.",
        7: "Le périmètre exclut les sujets fictifs.",
    }
    for n, titre in sections.items():
        texte = titre._element.getnext()
        from docx.text.paragraph import Paragraph
        para = Paragraph(texte, titre._parent)
        if n == 6:
            supprimer(para)
            supprimer(titre)
        else:
            para.text = textes[n]
    # Section 5 sans tableau (le tableau, sa légende et sa source sont retirés).
    t5 = doc.tables[1]
    leg = t5._tbl.getprevious()
    src = t5._tbl.getnext()
    for el in (leg, t5._tbl, src):
        el.getparent().remove(el)
    # Tableau de la section 3 rempli.
    t3 = doc.tables[0]
    t3.cell(1, 0).text = "Mémoire fictif X (2020)"
    t3.cell(1, 1).text = "Apport fictif"

    for p in doc.paragraphs:
        precedent = p._element.getprevious()
        consigne = precedent.xpath("string(.)") if precedent is not None else ""
        if p.text == "[…]" and consigne.startswith("Partie"):
            lignes = (["Partie 1 – Analyse fictive (50 p.)", "Chapitre 1 – Premier chapitre (25 p.)",
                       "Section 1 – Interdite ici (10 p.)", "Chapitre 2 – Second chapitre"]
                      if consigne.startswith("Parties et chapitres") else
                      ["Introduction (4 p.)", "Partie 1 – A (45 p.)", "Chapitre 1 – B (20 p.)", "Section 1 – C (10 p.)",
                       "§ 1 – D, OUT-01 Fiche fictive (5 p.)", "§ 2 – E (5 p.)", "Section 2 – F (10 p.)",
                       "§ 1 – G (10 p.)", "Chapitre 3 – H (25 p.)", "Partie 2 – I (45 p.)", "Chapitre 1 – J (45 p.)",
                       "I. Ligne proscrite (2 p.)", "Conclusion (3 p.)"])
            p.text = lignes[0]
            dernier = p
            for l in lignes[1:]:
                dernier = paragraphe_apres(dernier, l)

    # Bibliographie : « Ouvrage » déplacée avant la catégorie 2 ; catégorie 2 hors ordre alphabétique.
    cats = {p.text[0]: p for p in doc.paragraphs if p.style.name == "Titre catégorie"}
    entree = lambda c: c._element.getnext()
    from docx.text.paragraph import Paragraph
    Paragraph(entree(cats["1"]), cats["1"]._parent).text = "Code de commerce, article fictif"
    e2 = Paragraph(entree(cats["2"]), cats["2"]._parent)
    e2.text = "ZOLA Émile, (2020), Article fictif, Revue fictive, 1, 1-2"
    paragraphe_apres(e2, "ALPHA Anne, (2021), Autre article fictif, Revue fictive, 2, 3-4")
    for n in "35":
        Paragraph(entree(cats[n]), cats[n]._parent).text = "AUTEUR Fictif, (2022), Titre fictif"
    Paragraph(entree(cats["4"]), cats["4"]._parent).text = "AUTEUR Fictif, (2019), Ouvrage fictif, Éditeur, 1, 100"
    ouvrage, sa_ref = cats["4"]._element, entree(cats["4"])
    cats["2"]._element.addprevious(ouvrage)
    cats["2"]._element.addprevious(sa_ref)
    doc.core_properties.author = "Prénom NOM"
    doc.save(SORTIE / "notice-fictive.docx")


def memoire():
    doc = Document(GABARITS / "gabarit-memoire.docx")
    for p in doc.paragraphs:
        if p.text == "[Texte du candidat]":
            p.text = "Texte fictif de démonstration, voir annexe 2 pour un renvoi sans annexe."
        elif p.text == "[Intitulé de l'annexe]":
            p.text = "Tableau fictif"
    doc.core_properties.author = "Prénom NOM"
    doc.save(SORTIE / "memoire-fictif.docx")


if __name__ == "__main__":
    SORTIE.mkdir(exist_ok=True)
    notice()
    memoire()
    print("fixtures écrites dans", SORTIE)
