"""Gabarits Word du mémoire et de la demande d'agrément (format de la Note du jury, SRC-FC-001, p. 11-14).

Les gabarits ne contiennent que des titres imposés par la Note et des repères entre crochets :
aucune ligne du mémoire ni de la notice n'y est écrite.
"""
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

from commun import (champ, el, maj_champs_a_l_ouverture, mise_en_page, nouveau_document, numero_de_page,
                    style_paragraphe, texte_balise)

NIVEAUX = [  # nomenclature Partie / Chapitre / Section / Paragraphe (Note, p. 14)
    ("Heading 1", "Partie %1 –", 16, True),
    ("Heading 2", "Chapitre %2 –", 14, False),
    ("Heading 3", "Section %3 –", 12, False),
    ("Heading 4", "§ %4 –", 12, False),
]
ID_PLAN, ID_ANNEXE = 90, 91


def liste_numerotee(doc, ident, niveaux):
    """Liste à plusieurs niveaux liée aux styles ; chaque niveau repart à 1 sous son parent."""
    num = doc.part.numbering_part.element
    abstrait = el("w:abstractNum", **{"w:abstractNumId": ident})
    abstrait.append(el("w:multiLevelType", **{"w:val": "multilevel"}))
    for i, (style_id, texte) in enumerate(niveaux):
        lvl = el("w:lvl", **{"w:ilvl": i})
        for e in (el("w:start", **{"w:val": 1}), el("w:numFmt", **{"w:val": "decimal"}),
                  el("w:pStyle", **{"w:val": style_id}), el("w:suff", **{"w:val": "space"}),
                  el("w:lvlText", **{"w:val": texte}), el("w:lvlJc", **{"w:val": "left"})):
            lvl.append(e)
        ppr = el("w:pPr")
        ppr.append(el("w:ind", **{"w:left": 0, "w:firstLine": 0}))
        lvl.append(ppr)
        abstrait.append(lvl)
    existants = num.findall(qn("w:abstractNum"))
    if existants:
        existants[-1].addnext(abstrait)
    else:
        num.insert(0, abstrait)
    n = el("w:num", **{"w:numId": ident})
    n.append(el("w:abstractNumId", **{"w:val": ident}))
    num.append(n)


def lier_style(st, ident, niveau):
    numpr = st.element.get_or_add_pPr().get_or_add_numPr()
    numpr.get_or_add_ilvl().val = niveau
    numpr.get_or_add_numId().val = ident


def styles_titres(doc):
    for i, (nom, _, taille, saut) in enumerate(NIVEAUX):
        st = style_paragraphe(doc, nom, taille=taille, gras=True, italique=(i == 3), avant=18 if i < 2 else 12,
                              apres=6, aligne=WD_ALIGN_PARAGRAPH.LEFT, suivant_lie=True, saut_avant=saut)
        lier_style(st, ID_PLAN, i)
    liste_numerotee(doc, ID_PLAN, [(doc.styles[n].style_id, t) for n, t, _, _ in NIVEAUX])
    # Titres sans numéro (sommaire, introduction, conclusion…) : niveau 1 de la table des matières.
    style_paragraphe(doc, "Titre hors plan", taille=16, gras=True, avant=18, apres=12, aligne=WD_ALIGN_PARAGRAPH.LEFT,
                     suivant_lie=True, saut_avant=True, niveau_plan=0)
    style_paragraphe(doc, "Titre notice", taille=13, gras=True, avant=12, apres=6, aligne=WD_ALIGN_PARAGRAPH.LEFT,
                     suivant_lie=True, niveau_plan=1)
    style_paragraphe(doc, "Titre catégorie", taille=12, gras=True, avant=12, apres=6, aligne=WD_ALIGN_PARAGRAPH.LEFT,
                     suivant_lie=True)
    annexe = style_paragraphe(doc, "Titre annexe", taille=14, gras=True, avant=18, apres=12,
                              aligne=WD_ALIGN_PARAGRAPH.LEFT, suivant_lie=True, saut_avant=True)
    liste_numerotee(doc, ID_ANNEXE, [(annexe.style_id, "Annexe %1 –")])
    lier_style(annexe, ID_ANNEXE, 0)
    style_paragraphe(doc, "Page de garde", taille=14, avant=0, apres=12, aligne=WD_ALIGN_PARAGRAPH.CENTER)


def repere(doc, texte, style="Normal"):
    return texte_balise(doc.add_paragraph(style=style), texte)


def table_des_matieres(doc, titre, instruction):
    doc.add_paragraph(titre, style="Titre hors plan")
    champ(doc.add_paragraph(), instruction, "[Clic droit puis « Mettre à jour les champs »]")


_signet = [0]


def signet(paragraphe, nom):
    _signet[0] += 1
    debut = el("w:bookmarkStart", **{"w:id": _signet[0], "w:name": nom})
    fin = el("w:bookmarkEnd", **{"w:id": _signet[0]})
    paragraphe._p.insert(1 if paragraphe._p.pPr is not None else 0, debut)
    paragraphe._p.append(fin)


def lien_interne(paragraphe, texte, cible):
    """Renvoi cliquable dans le PDF (Note, p. 8 : renvois automatisés et retour au corps du mémoire)."""
    h = el("w:hyperlink", **{"w:anchor": cible, "w:history": 1})
    r = el("w:r")
    rpr = el("w:rPr")
    rpr.append(el("w:color", **{"w:val": "1F4E79"}))
    rpr.append(el("w:u", **{"w:val": "single"}))
    r.append(rpr)
    t = el("w:t", **{"xml:space": "preserve"})
    t.text = texte
    r.append(t)
    h.append(r)
    paragraphe._p.append(h)


def tableau_exemple(doc, colonnes, legende="[Titre du tableau]"):
    p = doc.add_paragraph(style="Caption")
    p.add_run("Tableau ")
    champ(p, "SEQ Tableau \\* ARABIC", "1")
    p.add_run(f" – {legende}")
    t = doc.add_table(rows=2, cols=len(colonnes))
    t.style = "Table Grid"
    for i, c in enumerate(colonnes):
        t.cell(0, i).text = c
        t.cell(1, i).text = "[…]"
        for ligne in (0, 1):
            for par in t.cell(ligne, i).paragraphs:
                par.style = "Texte tableau"
        t.cell(0, i).paragraphs[0].runs[0].bold = True
    doc.add_paragraph("Source : [origine du tableau, ou « travail personnel »]", style="Source")


def page_de_garde(doc, lignes):
    for _ in range(6):
        doc.add_paragraph(style="Page de garde")
    for texte, taille, gras in lignes:
        p = doc.add_paragraph(style="Page de garde")
        r = p.add_run(texte)
        r.font.size, r.bold = Pt(taille), gras
    p = doc.add_paragraph(style="Consigne")
    p.add_run("Aucune référence (nom, logo) au cabinet sur cette page (Note du jury, p. 14). "
              "Cette page n'est pas numérotée ; la numérotation commence à 1 à la page suivante.")
    # La page de garde est la section 1, sans pied de page ; tout le reste est la section 2.
    numero_de_page(doc.add_section(WD_SECTION.NEW_PAGE), depart=1)


def bibliographie(doc):
    doc.add_paragraph("Bibliographie", style="Titre hors plan")
    repere(doc, "Ordre alphabétique du premier auteur dans chaque catégorie (Note du jury, p. 10-11).", "Consigne")
    for titre, forme in [
        ("1. Référence réglementaire", "…"),
        ("2. Article extrait d'une revue scientifique ou professionnelle",
         "NOM Prénom, (Année), Titre de l'article, Titre de la Revue, numéro, pages"),
        ("3. Article professionnel ou étude consultables sur un site internet",
         "NOM DU SITE, (Année), Titre de l'article, lien URL cliquable, (date de dernière connexion : j/m/année)"),
        ("4. Ouvrage", "NOM Prénom, (année), Titre de l'ouvrage, Éditeur, N° édition, Nombre de pages"),
        ("5. Mémoire d'expertise comptable", "NOM Prénom, (année), Titre du mémoire, session de soutenance, nombre de pages"),
        ("6. Site internet", "page d'accueil seulement, en lien direct avec le sujet"),
    ]:
        doc.add_paragraph(titre, style="Titre catégorie")
        repere(doc, f"[{forme}]")


def gabarit_memoire(chemin):
    doc = nouveau_document()
    styles_titres(doc)
    page_de_garde(doc, [
        ("[Mentions de la page de garde du mémoire : voir la Note du jury en vigueur, p. 14]", 12, False),
        ("[Intitulé du sujet agréé]", 18, True),
        ("[Prénom NOM du candidat]", 14, False),
        ("[Session de soutenance]", 12, False),
    ])
    table_des_matieres(doc, "Sommaire", 'TOC \\o "1-2" \\h \\z \\u')
    repere(doc, "Sommaire paginé, parties et chapitres uniquement, sur une page (Note du jury, p. 14).", "Consigne")

    doc.add_paragraph("Note de synthèse", style="Titre hors plan")
    repere(doc, "Deux ou trois pages résumant le mémoire (Note du jury, p. 15). Emplacement à confirmer.", "Consigne")
    repere(doc, "[Texte du candidat]")

    doc.add_paragraph("Introduction", style="Titre hors plan")
    repere(doc, "[Texte du candidat]")

    for partie in (1, 2):
        doc.add_paragraph("[Intitulé de la partie]", style="Heading 1")
        repere(doc, "[Introduction et annonce du plan de la partie]")
        doc.add_paragraph("[Intitulé du chapitre]", style="Heading 2")
        doc.add_paragraph("[Intitulé de la section]", style="Heading 3")
        doc.add_paragraph("[Intitulé du paragraphe]", style="Heading 4")
        p = repere(doc, "[Texte du candidat]")
        if partie == 1:
            p.add_run(" (exemple de renvoi : ")
            lien_interne(p, "voir annexe 1", "annexe_1")
            p.add_run(")")
            signet(p, "appel_annexe_1")
            tableau_exemple(doc, ["[Colonne]", "[Colonne]", "[Colonne]"])
        doc.add_paragraph("[Intitulé du paragraphe]", style="Heading 4")
        repere(doc, "[Texte du candidat]")
    repere(doc, "Les numéros « Partie n – », « Chapitre n – », « Section n – », « § n – » sont automatiques "
                "et repartent à 1 sous chaque niveau supérieur. Notes de bas de page : Références > Insérer une note "
                "(numérotation continue, 10 points).", "Consigne")

    doc.add_paragraph("Conclusion", style="Titre hors plan")
    repere(doc, "[Texte du candidat]")
    bibliographie(doc)

    table_des_matieres(doc, "Liste des annexes", 'TOC \\h \\z \\t "Titre annexe,1"')
    p = doc.add_paragraph("[Intitulé de l'annexe]", style="Titre annexe")
    signet(p, "annexe_1")
    repere(doc, "[Contenu]")
    p = doc.add_paragraph(style="Consigne")
    lien_interne(p, "Retour au texte", "appel_annexe_1")
    repere(doc, "Annexes : 100 pages au plus, numérotées 1, 2…, un retour au corps du mémoire à la fin de chacune "
                "(Note du jury, p. 8 et 15). Pour un nouveau renvoi : signet à l'appel et sur le titre de l'annexe, "
                "puis Insertion > Lien > Emplacement dans ce document.", "Consigne")

    table_des_matieres(doc, "Table des matières", 'TOC \\o "1-4" \\h \\z \\u')
    maj_champs_a_l_ouverture(doc)
    doc.save(chemin)


SECTIONS_NOTICE = [  # titres et ordre imposés (Note du jury, p. 12)
    "1. Intérêt du sujet et problématique",
    "2. Objectifs poursuivis",
    "3. Différenciation avec les autres mémoires d'expertise comptable",
    "4. Possibilités d'atteindre les apports",
    "5. Apports du mémoire à la profession",
    "6. Illustrations pratiques",
    "7. Limites du mémoire",
]


def gabarit_demande_agrement(chemin):
    doc = nouveau_document()
    styles_titres(doc)
    page_de_garde(doc, [
        ("Demande d'agrément du sujet de mémoire présentée en vue de l'obtention du diplôme d'expertise comptable",
         14, True),
        ("[Intitulé du sujet du mémoire]", 18, True),
        ("[Prénom NOM du candidat]", 12, False),
        ("[Adresse personnelle : au minimum code postal et ville]", 12, False),
        ("[Mois et année de dépôt]", 12, False),
    ])
    repere(doc, "Seulement après une décision 4.2 : tableau des changements, deux pages au plus, avant la notice ; "
                "modifications écrites en rouge (Note du jury, p. 13). Sinon, supprimer ce paragraphe.", "Consigne")
    table_des_matieres(doc, "Table des matières", 'TOC \\o "1-2" \\h \\z \\u')

    doc.add_paragraph("Notice explicative", style="Titre hors plan")
    repere(doc, "Dix pages au plus, sept sections dans cet ordre (Note du jury, p. 11-12).", "Consigne")
    for titre in SECTIONS_NOTICE:
        doc.add_paragraph(titre, style="Titre notice")
        repere(doc, "[Texte du candidat]")
        if titre.startswith("3."):
            tableau_exemple(doc, ["Mémoire (Bibliodec)", "Apports de mon projet"], "[Différenciation]")
        if titre.startswith("5."):
            tableau_exemple(doc, ["N° d'ordre dans le plan", "Nom", "Nature", "Objectifs principaux"], "[Outils proposés]")

    doc.add_paragraph("Plan synthétique", style="Titre hors plan")
    repere(doc, "Parties et chapitres seulement, avec le nombre prévisionnel de pages (Note du jury, p. 12).", "Consigne")
    repere(doc, "[…]")
    doc.add_paragraph("Plan détaillé", style="Titre hors plan")
    repere(doc, "Partie / Chapitre / Section / Paragraphe, introduction et conclusion, pages prévisionnelles à chaque "
                "niveau, numéro et nom des outils dans les paragraphes concernés (Note du jury, p. 12).", "Consigne")
    repere(doc, "[…]")
    bibliographie(doc)
    doc.add_paragraph("Annexe 1 de la charte d'usage de l'IA", style="Titre hors plan")
    repere(doc, "Complétée et signée (Note du jury, p. 12 et annexe 1).", "Consigne")
    maj_champs_a_l_ouverture(doc)
    doc.save(chemin)
