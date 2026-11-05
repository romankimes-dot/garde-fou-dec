"""Briques communes aux gabarits et aux trames : format du jury (Note du jury, p. 14, SRC-FC-001)."""
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# La Note n'impose pas de police : Times New Roman (rendue par Liberation Serif, de mêmes métriques, sous LibreOffice).
POLICE = "Times New Roman"
TAILLE = 12          # « taille des caractères 12 »
INTERLIGNE = 1.5     # « interligne 1,5 »
TAILLE_NOTE = 10     # notes « en caractères plus petits que ceux du texte »
TAILLE_TABLEAU = 10  # tableaux : taille et interligne inférieurs admis
GRIS = RGBColor(0x59, 0x59, 0x59)


def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def poser_police(rpr_parent, police=POLICE):
    """Police explicite sur les quatre plages, sans police de thème (qui l'emporterait)."""
    rpr = rpr_parent.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = el("w:rFonts")
        rpr.insert(0, fonts)
    for a in list(fonts.attrib):
        del fonts.attrib[a]
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(a), police)


def mise_en_page(section, haut=2.5, bas=2.5):
    """A4 ; marges de 2,5 cm à gauche et à droite ; marges « normales » (2,5 cm) en haut et en bas."""
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(2.5)
    section.top_margin, section.bottom_margin = Cm(haut), Cm(bas)
    section.header_distance = section.footer_distance = Cm(1.25)


def style_paragraphe(doc, nom, base="Normal", taille=None, gras=False, italique=False, interligne=None,
                     avant=None, apres=None, aligne=None, suivant_lie=False, saut_avant=False,
                     niveau_plan=None, couleur=None, style_id=None):
    styles = doc.styles
    if nom in [s.name for s in styles]:
        st = styles[nom]
    else:
        st = styles.add_style(nom, WD_STYLE_TYPE.PARAGRAPH)
        if style_id:
            st.element.set(qn("w:styleId"), style_id)
        st.base_style = styles[base] if base else None
        st.quick_style = True
    poser_police(st.element)
    f, p = st.font, st.paragraph_format
    f.size = Pt(taille or TAILLE)
    f.bold, f.italic = gras, italique
    f.color.rgb = couleur or RGBColor(0, 0, 0)
    p.line_spacing = interligne or INTERLIGNE
    if avant is not None:
        p.space_before = Pt(avant)
    if apres is not None:
        p.space_after = Pt(apres)
    if aligne is not None:
        p.alignment = aligne
    p.keep_with_next = suivant_lie
    p.page_break_before = saut_avant
    if niveau_plan is not None:
        ppr = st.element.get_or_add_pPr()
        for old in ppr.findall(qn("w:outlineLvl")):
            ppr.remove(old)
        # outlineLvl se place après jc / textAlignment, avant divId : en fin de pPr suffit hors rPr.
        ppr.append(el("w:outlineLvl", **{"w:val": niveau_plan}))
    return st


def styles_de_base(doc):
    """Normal, langue française, notes de bas de page, légendes, texte de tableau, pied de page."""
    rpr_defaut = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    for old in rpr_defaut.findall(qn("w:lang")):
        rpr_defaut.remove(old)
    rpr_defaut.append(el("w:lang", **{"w:val": "fr-FR", "w:eastAsia": "fr-FR", "w:bidi": "ar-SA"}))

    style_paragraphe(doc, "Normal", base=None, apres=6, aligne=WD_ALIGN_PARAGRAPH.JUSTIFY)
    style_paragraphe(doc, "footnote text", taille=TAILLE_NOTE, interligne=1.0, apres=0, style_id="FootnoteText")
    ref = doc.styles.add_style("footnote reference", WD_STYLE_TYPE.CHARACTER)
    ref.element.set(qn("w:styleId"), "FootnoteReference")
    ref.font.superscript = True
    style_paragraphe(doc, "Caption", taille=TAILLE_TABLEAU, gras=True, interligne=1.0, avant=6, apres=3,
                     aligne=WD_ALIGN_PARAGRAPH.LEFT, suivant_lie=True)
    style_paragraphe(doc, "Source", taille=9, italique=True, interligne=1.0, apres=12, aligne=WD_ALIGN_PARAGRAPH.LEFT)
    style_paragraphe(doc, "Texte tableau", taille=TAILLE_TABLEAU, interligne=1.0, apres=0, aligne=WD_ALIGN_PARAGRAPH.LEFT)
    style_paragraphe(doc, "Consigne", taille=10, italique=True, interligne=1.0, apres=4,
                     aligne=WD_ALIGN_PARAGRAPH.LEFT, couleur=GRIS)
    style_paragraphe(doc, "Footer", interligne=1.0, apres=0, aligne=WD_ALIGN_PARAGRAPH.RIGHT)


def champ(paragraphe, instruction, affiche=""):
    """Champ Word (PAGE, TOC, SEQ…) ; le texte affiché est remplacé à la mise à jour des champs."""
    def run_avec(*enfants):
        r = paragraphe.add_run()
        for e in enfants:
            r._r.append(e)
        return r
    run_avec(el("w:fldChar", **{"w:fldCharType": "begin", "w:dirty": "true"}))
    instr = el("w:instrText", **{"xml:space": "preserve"})
    instr.text = f" {instruction} "
    run_avec(instr)
    run_avec(el("w:fldChar", **{"w:fldCharType": "separate"}))
    paragraphe.add_run(affiche)
    run_avec(el("w:fldChar", **{"w:fldCharType": "end"}))


def numero_de_page(section, depart=1):
    """Numéro en bas à droite ; la numérotation repart à `depart` dans cette section."""
    section.footer.is_linked_to_previous = False
    p = section.footer.paragraphs[0]
    p.style = "Footer"
    # Alignement à droite posé sur le paragraphe lui-même (en plus du style) : Word en ligne renvoie parfois le
    # pied de page sans ses styles, et Garde-fou (A19) doit pouvoir lire l'alignement sans eux.
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    champ(p, "PAGE", "1")
    sect = section._sectPr
    for old in sect.findall(qn("w:pgNumType")):
        sect.remove(old)
    pg = el("w:pgNumType", **{"w:start": depart})
    suivant = None
    for tag in ("w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
                "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange"):
        suivant = sect.find(qn(tag))
        if suivant is not None:
            break
    if suivant is not None:
        suivant.addprevious(pg)
    else:
        sect.append(pg)


def maj_champs_a_l_ouverture(doc):
    """Word propose de mettre à jour sommaire, table des matières et numéros à l'ouverture."""
    settings = doc.settings.element
    for old in settings.findall(qn("w:updateFields")):
        settings.remove(old)
    settings.append(el("w:updateFields", **{"w:val": "true"}))


def texte_balise(paragraphe, texte, taille=None):
    """Ajoute un texte où **…** est en gras."""
    morceaux = texte.split("**")
    for i, m in enumerate(morceaux):
        if m:
            r = paragraphe.add_run(m)
            r.bold = i % 2 == 1
            if taille:
                r.font.size = Pt(taille)
    return paragraphe


def ombrer(cellule, couleur="E7E6E6"):
    tcpr = cellule._tc.get_or_add_tcPr()
    tcpr.append(el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": couleur}))


def nouveau_document():
    doc = Document()
    styles_de_base(doc)
    mise_en_page(doc.sections[0])
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""
    doc.core_properties.comments = ""
    return doc


def saut_de_page(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
