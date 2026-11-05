// Lecture d'un fichier .docx (Word, « Télécharger une copie ») vers le modèle commun du moteur.
// Rien ne quitte le navigateur : le fichier est lu en mémoire.

import { lireZip } from './zip.js';
import { analyserXml, enfants, premier, val, chercher, tous, texteBrut } from './xml.js';

// Styles des gabarits du dépôt (scripts/mise-en-forme/gabarits.py) : nom en minuscules -> [rôle, niveau].
export const STYLES_TITRES = {
  'titre hors plan': ['zone', 1],
  'titre notice': ['section-notice', 2],
  'titre catégorie': ['categorie', 3],
  'titre annexe': ['annexe', 1],
};
export const STYLES_CONSIGNE = new Set(['consigne']);

const TWIPS_CM = 566.929;

function lireStyles(xml) {
  const styles = new Map();
  const defauts = { taille: null, interligne: null };
  if (!xml) return { styles, defauts, defaut: null };
  const r = analyserXml(xml);
  const racine = premier(r, 'w:styles');
  const dd = premier(racine, 'w:docDefaults');
  const sz = chercher(premier(dd, 'w:rPrDefault'), 'w:sz');
  if (sz) defauts.taille = Number(val(sz)) / 2;
  const sp = chercher(premier(dd, 'w:pPrDefault'), 'w:spacing');
  if (sp && val(sp, 'w:line')) defauts.interligne = interligne(sp);
  let defaut = null;
  for (const s of enfants(racine, 'w:style')) {
    const id = s.a['w:styleId'];
    const pPr = premier(s, 'w:pPr');
    const rPr = premier(s, 'w:rPr');
    const numPr = premier(pPr, 'w:numPr');
    const spacing = premier(pPr, 'w:spacing');
    const st = {
      id,
      type: s.a['w:type'],
      nom: (val(premier(s, 'w:name')) || id || '').toLowerCase(),
      base: val(premier(s, 'w:basedOn')) || null,
      plan: premier(pPr, 'w:outlineLvl') ? Number(val(premier(pPr, 'w:outlineLvl'))) : null,
      numId: numPr && premier(numPr, 'w:numId') ? val(premier(numPr, 'w:numId')) : null,
      ilvl: numPr && premier(numPr, 'w:ilvl') ? Number(val(premier(numPr, 'w:ilvl'))) : null,
      saut: !!premier(pPr, 'w:pageBreakBefore') && val(premier(pPr, 'w:pageBreakBefore')) !== '0',
      taille: premier(rPr, 'w:sz') ? Number(val(premier(rPr, 'w:sz'))) / 2 : null,
      interligne: spacing && val(spacing, 'w:line') ? interligne(spacing) : null,
      jc: premier(pPr, 'w:jc') ? val(premier(pPr, 'w:jc')) : null,
    };
    styles.set(id, st);
    if (s.a['w:default'] === '1' && st.type === 'paragraph') defaut = id;
  }
  return { styles, defauts, defaut };
}

/** Interligne en multiple de ligne ; un interligne fixe (exact, au moins) est rendu en texte « 18 pt fixe ». */
function interligne(spacing) {
  const regle = val(spacing, 'w:lineRule') || 'auto';
  const ligne = Number(val(spacing, 'w:line'));
  return regle === 'auto' ? ligne / 240 : `${ligne / 20} pt ${regle === 'exact' ? 'exactement' : 'au moins'}`;
}

/** Valeur d'une propriété en remontant la chaîne basedOn. */
function propriete(styles, id, cle) {
  const vus = new Set();
  while (id && !vus.has(id)) {
    vus.add(id);
    const st = styles.get(id);
    if (!st) return null;
    if (st[cle] !== null && st[cle] !== undefined && st[cle] !== false) return st[cle];
    id = st.base;
  }
  return null;
}

function lireNumerotation(xml) {
  const nums = new Map();
  const abstraits = new Map();
  if (!xml) return { nums, abstraits };
  const r = premier(analyserXml(xml), 'w:numbering');
  for (const a of enfants(r, 'w:abstractNum')) {
    const niveaux = [];
    for (const l of enfants(a, 'w:lvl')) {
      niveaux[Number(l.a['w:ilvl'])] = {
        debut: premier(l, 'w:start') ? Number(val(premier(l, 'w:start'))) : 1,
        format: val(premier(l, 'w:numFmt')) || 'decimal',
        texte: val(premier(l, 'w:lvlText')) ?? '',
        suffixe: val(premier(l, 'w:suff')) || 'tab',
      };
    }
    abstraits.set(a.a['w:abstractNumId'], niveaux);
  }
  for (const n of enfants(r, 'w:num')) nums.set(n.a['w:numId'], val(premier(n, 'w:abstractNumId')));
  return { nums, abstraits };
}

const ROMAINS = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
  [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
function romain(n) {
  let s = '';
  for (const [v, l] of ROMAINS) while (n >= v) { s += l; n -= v; }
  return s;
}
function formater(n, format) {
  switch (format) {
    case 'upperRoman': return romain(n);
    case 'lowerRoman': return romain(n).toLowerCase();
    case 'upperLetter': return String.fromCharCode(64 + ((n - 1) % 26) + 1);
    case 'lowerLetter': return String.fromCharCode(96 + ((n - 1) % 26) + 1);
    case 'none': return '';
    default: return String(n);
  }
}

/** Numéro affiché par Word pour un paragraphe numéroté (« Partie 1 – », « I. », « 1.1 »…). */
function creerNumeroteur({ nums, abstraits }) {
  const compteurs = new Map();
  return (numId, ilvl) => {
    const abs = nums.get(numId);
    const niveaux = abstraits.get(abs);
    if (!niveaux || numId === '0') return '';
    const lvl = niveaux[ilvl];
    if (!lvl) return '';
    if (!compteurs.has(abs)) compteurs.set(abs, niveaux.map((l) => (l ? l.debut - 1 : 0)));
    const c = compteurs.get(abs);
    c[ilvl] += 1;
    for (let k = ilvl + 1; k < c.length; k++) c[k] = niveaux[k] ? niveaux[k].debut - 1 : 0;
    if (lvl.format === 'bullet') return '•';
    const texte = lvl.texte.replace(/%(\d)/g, (_, k) => {
      const i = Number(k) - 1;
      const v = Math.max(c[i] ?? 1, niveaux[i] ? niveaux[i].debut : 1);
      return formater(v, niveaux[i]?.format || 'decimal');
    });
    return texte + (lvl.suffixe === 'space' ? ' ' : lvl.suffixe === 'nothing' ? '' : '\t');
  };
}

/** Pied de page : contient-il un champ PAGE, et est-il aligné à droite ? */
function analyserPied(xml, styles) {
  const paras = tous(analyserXml(xml), 'w:p');
  const avecPage = paras.find((p) => /\bPAGE\b/.test(tous(p, 'w:instrText').map(texteBrut).join(' ') +
    ' ' + tous(p, 'w:fldSimple').map((f) => f.a['w:instr']).join(' ')));
  const pPr = avecPage ? premier(avecPage, 'w:pPr') : null;
  const jc = avecPage ? (val(premier(pPr, 'w:jc')) || propriete(styles, val(premier(pPr, 'w:pStyle')), 'jc')) : null;
  return { numero: !!avecPage, droite: jc === 'right' || jc === 'end' };
}

/** Pied de page rendu par Office.js (footer.getOoxml()) : paquet plat dont le corps est le pied. */
export async function piedPlat(xml) {
  const paquet = paquetPlat(xml);
  const { styles } = lireStyles(await paquet.texte('word/styles.xml'));
  const corps = await paquet.texte('word/document.xml');
  return corps ? analyserPied(corps, styles) : { numero: false, droite: false };
}

function lireRelations(xml) {
  const rel = new Map();
  if (!xml) return rel;
  for (const r of tous(analyserXml(xml), 'Relationship')) rel.set(r.a.Id, r.a.Target);
  return rel;
}

/** Texte, liens, notes et révisions d'un paragraphe. */
function lireContenu(p, compte) {
  let texte = '';
  const liens = [];
  let image = false;
  let notes = 0;
  let saut = false;
  let rouge = false;
  const champs = [];
  const tailles = new Set();

  const ajouter = (t, lien) => {
    texte += t;
    if (lien) lien.t += t;
    for (const ch of champs) if (ch.resultat) ch.t += t;
  };

  const marcher = (el, lien) => {
    for (const x of el.c || []) {
      if (typeof x !== 'object') continue;
      switch (x.n) {
        case 'w:pPr': case 'w:delText': case 'w:instrText':
          if (x.n === 'w:instrText' && champs.length) champs[champs.length - 1].instr += texteBrut(x);
          break;
        case 'w:rPr': break;
        case 'w:r': {
          // Taille de chaque passage écrit : directe (w:sz) ou celle du style (« style »).
          const sz = premier(premier(x, 'w:rPr'), 'w:sz');
          const couleur = val(premier(premier(x, 'w:rPr'), 'w:color')) || '';
          const avant = texte.length;
          marcher(x, lien);
          if (texte.slice(avant).trim()) {
            tailles.add(sz ? Number(val(sz)) / 2 : 'style');
            if (/^[c-f][0-9a-f][0-3][0-9a-f][0-3][0-9a-f]$/i.test(couleur)) rouge = true; // rouge dominant
          }
          break;
        }
        case 'mc:AlternateContent': marcher(premier(x, 'mc:Choice') || { c: [] }, lien); break;
        case 'w:t': ajouter(x.c.filter((y) => typeof y === 'string').join(''), lien); break;
        case 'w:tab': ajouter('\t', lien); break;
        case 'w:br': case 'w:cr':
          if (val(x, 'w:type') === 'page') saut = true;
          ajouter(' ', lien);
          break;
        case 'w:noBreakHyphen': ajouter('‑', lien); break;
        case 'w:sym': ajouter('□', lien); break;
        case 'w:hyperlink': {
          const l = { t: '' };
          marcher(x, l);
          liens.push(l.t);
          break;
        }
        case 'w:fldSimple': {
          const instr = x.a['w:instr'] || '';
          if (/\b(HYPERLINK|REF|PAGEREF|NOTEREF)\b/.test(instr)) {
            const l = { t: '' };
            marcher(x, l);
            if (!/PAGEREF/.test(instr)) liens.push(l.t);
          } else marcher(x, lien);
          break;
        }
        case 'w:fldChar': {
          const type = x.a['w:fldCharType'];
          if (type === 'begin') champs.push({ instr: '', resultat: false, t: '' });
          else if (type === 'separate' && champs.length) champs[champs.length - 1].resultat = true;
          else if (type === 'end' && champs.length) {
            const ch = champs.pop();
            if (/\b(HYPERLINK|REF|NOTEREF)\b/.test(ch.instr) && ch.t) liens.push(ch.t);
          }
          break;
        }
        case 'w:drawing': case 'w:pict': case 'w:object': {
          // Image véritable (photo, logo) ou simple forme ; le texte des zones de texte est lu.
          if (tous(x, 'pic:pic').length || tous(x, 'v:imagedata').length || tous(x, 'a:blip').length) image = true;
          for (const zone of tous(x, 'w:txbxContent')) {
            const t = texteDesParagraphes(zone).join(' ');
            if (t.trim()) ajouter(` ${t} `, lien);
          }
          break;
        }
        case 'w:footnoteReference': notes++; break;
        case 'w:ins': case 'w:moveTo': compte.revisions++; marcher(x, lien); break;
        case 'w:del': case 'w:moveFrom': compte.revisions++; break;
        default: marcher(x, lien);
      }
    }
  };
  marcher(p, null);
  return { texte, liens, image, notes, saut, tailles, rouge };
}

function texteDesParagraphes(racine) {
  return tous(racine, 'w:p').map((p) => lireContenu(p, { revisions: 0 }).texte).filter((t) => t.trim());
}

export async function lireDocx(donnees) {
  return lirePaquet(await lireZip(donnees), 'docx');
}

/**
 * Paquet « Flat OPC » rendu par Office.js (body.getOoxml()) : mêmes parties qu'un .docx, dans un seul XML.
 * Seules les parties XML sont gardées ; les images (binaires) sont ignorées.
 */
export function paquetPlat(xml) {
  const parties = new Map();
  const re = /<pkg:part\b[^>]*pkg:name="([^"]+)"[^>]*>([\s\S]*?)<\/pkg:part>/g;
  let m;
  while ((m = re.exec(xml))) {
    const d = /<pkg:xmlData>([\s\S]*)<\/pkg:xmlData>/.exec(m[2]);
    if (d) parties.set(m[1].replace(/^\//, ''), d[1]);
  }
  return { noms: [...parties.keys()], async texte(nom) { return parties.get(nom) ?? null; } };
}

export async function lirePaquet(zip, source = 'docx') {
  const docXml = await zip.texte('word/document.xml');
  if (!docXml) throw new Error("Ce fichier n'est pas un document Word (.docx).");
  const { styles, defauts, defaut } = lireStyles(await zip.texte('word/styles.xml'));
  const numeroter = creerNumeroteur(lireNumerotation(await zip.texte('word/numbering.xml')));
  const relations = lireRelations(await zip.texte('word/_rels/document.xml.rels'));

  const compte = { revisions: 0 };
  const blocs = [];
  const sections = [];
  let tableaux = 0;
  let sautEnAttente = false;

  const nomStyle = (id) => (styles.get(id)?.nom || '').toLowerCase();

  function paragraphe(p, cellule) {
    const pPr = premier(p, 'w:pPr');
    const styleId = val(premier(pPr, 'w:pStyle')) || defaut;
    const nom = nomStyle(styleId);
    const contenu = lireContenu(p, compte);

    // Numérotation : directe, sinon héritée du style.
    let numId = null;
    let ilvl = 0;
    const numPr = premier(pPr, 'w:numPr');
    if (numPr && premier(numPr, 'w:numId')) {
      numId = val(premier(numPr, 'w:numId'));
      ilvl = premier(numPr, 'w:ilvl') ? Number(val(premier(numPr, 'w:ilvl'))) : 0;
    } else {
      numId = propriete(styles, styleId, 'numId');
      ilvl = propriete(styles, styleId, 'ilvl') ?? 0;
    }
    const label = numId ? numeroter(numId, ilvl) : '';

    let role = null;
    let niveau = 0;
    const titreNomme = /^heading (\d)$/.exec(nom);
    if (STYLES_TITRES[nom]) [role, niveau] = STYLES_TITRES[nom];
    else if (titreNomme) { role = 'plan'; niveau = Number(titreNomme[1]); }
    else if (/^toc \d$/.test(nom)) role = 'toc';
    else {
      const direct = premier(pPr, 'w:outlineLvl');
      const plan = direct ? Number(val(direct)) : propriete(styles, styleId, 'plan');
      if (plan !== null && plan !== undefined && plan < 9) { role = 'autre'; niveau = plan + 1; }
    }

    const spacing = premier(pPr, 'w:spacing');
    const bloc = {
      i: blocs.length,
      texte: contenu.texte,
      label: label.trim() ? label : '',
      titre: niveau > 0 && !cellule,
      niveau: cellule ? 0 : niveau,
      role: cellule && role !== 'toc' ? null : role,
      style: nom || null,
      tableau: cellule ? cellule.tableau : null,
      ligne: cellule ? cellule.ligne : null,
      colonne: cellule ? cellule.colonne : null,
      liens: contenu.liens,
      image: contenu.image,
      notes: contenu.notes,
      rouge: contenu.rouge,
      consigne: STYLES_CONSIGNE.has(nom),
      sautAvant: sautEnAttente || !!propriete(styles, styleId, 'saut'),
      garde: false,
      direct: {
        tailles: [...contenu.tailles].filter((t) => t !== 'style'),
        interligne: spacing && val(spacing, 'w:line') ? interligne(spacing) : null,
      },
      // Valeurs affichées par Word : directe, sinon style (chaîne basedOn), sinon valeurs par défaut du document.
      effectif: {
        tailles: [...new Set([...contenu.tailles].map((t) => (t === 'style' ? (propriete(styles, styleId, 'taille') ?? defauts.taille) : t)))],
        interligne: (spacing && val(spacing, 'w:line') ? interligne(spacing) : null)
          ?? propriete(styles, styleId, 'interligne') ?? defauts.interligne,
      },
    };
    sautEnAttente = contenu.saut;
    blocs.push(bloc);
    const sectPr = premier(pPr, 'w:sectPr');
    if (sectPr) sections.push({ sectPr, dernier: bloc.i });
  }

  function tableau(tbl, parent) {
    const k = parent ? parent.tableau : tableaux++;
    enfants(tbl, 'w:tr').forEach((tr, ligne) => {
      enfants(tr, 'w:tc').forEach((tc, colonne) => {
        for (const x of tc.c) {
          if (typeof x !== 'object') continue;
          if (x.n === 'w:p') paragraphe(x, { tableau: k, ligne, colonne });
          else if (x.n === 'w:tbl') tableau(x, { tableau: k });
        }
      });
    });
  }

  function corps(el) {
    for (const x of el.c || []) {
      if (typeof x !== 'object') continue;
      if (x.n === 'w:p') paragraphe(x, null);
      else if (x.n === 'w:tbl') tableau(x, null);
      else if (x.n === 'w:sdt') corps(premier(x, 'w:sdtContent') || { c: [] });
      else if (x.n === 'w:customXml') corps(x);
      else if (x.n === 'w:sectPr') sections.push({ sectPr: x, dernier: blocs.length - 1 });
    }
  }
  const body = chercher(analyserXml(docXml), 'w:body');
  corps(body);

  // Page de garde : première section quand le document en compte plusieurs.
  if (sections.length > 1) for (let i = 0; i <= sections[0].dernier; i++) blocs[i].garde = true;

  // Pieds de page : numéro de page à droite (Note du jury, p. 14).
  const lireSection = async (s, n) => {
    const mar = premier(s.sectPr, 'w:pgMar');
    const pieds = [];
    for (const ref of enfants(s.sectPr, 'w:footerReference')) {
      const cible = relations.get(ref.a['r:id']);
      const xml = cible ? await zip.texte('word/' + cible.replace(/^\/?word\//, '')) : null;
      if (!xml) continue;
      pieds.push({ type: ref.a['w:type'] || 'default', ...analyserPied(xml, styles) });
    }
    const numType = premier(s.sectPr, 'w:pgNumType');
    const titlePg = premier(s.sectPr, 'w:titlePg');
    return {
      n,
      premierBloc: n === 0 ? 0 : sections[n - 1].dernier + 1,
      gauche: mar ? Number(mar.a['w:left']) / TWIPS_CM : null,
      droite: mar ? Number(mar.a['w:right']) / TWIPS_CM : null,
      haut: mar ? Math.abs(Number(mar.a['w:top'])) / TWIPS_CM : null,
      bas: mar ? Math.abs(Number(mar.a['w:bottom'])) / TWIPS_CM : null,
      premierePageDistincte: !!titlePg && val(titlePg) !== '0' && val(titlePg) !== 'false',
      debutNumero: numType && numType.a['w:start'] ? Number(numType.a['w:start']) : null,
      pieds,
    };
  };
  const miseEnPageSections = [];
  for (let n = 0; n < sections.length; n++) miseEnPageSections.push(await lireSection(sections[n], n));
  // Une section sans pied de page d'un type reprend celui de la section précédente (comportement de Word).
  for (let n = 1; n < miseEnPageSections.length; n++) {
    const s = miseEnPageSections[n];
    for (const f of miseEnPageSections[n - 1].pieds) if (!s.pieds.some((x) => x.type === f.type)) s.pieds.push({ ...f, herite: true });
  }

  const normal = defaut ? styles.get(defaut) : null;
  // Notes de bas de page : numérotation (continue ou par page) et taille du style « footnote text ».
  const reglages = await zip.texte('word/settings.xml');
  const repriseParPage = /<w:footnotePr>[\s\S]*?<w:numRestart w:val="eachPage"/.test(reglages || '')
    || sections.some((s) => tous(premier(s.sectPr, 'w:footnotePr'), 'w:numRestart').some((r) => val(r) === 'eachPage'));
  const styleNote = [...styles.values()].find((st) => st.nom === 'footnote text');
  const entetes = [];
  for (const nom of zip.noms.filter((n) => /^word\/(header|footer)\d*\.xml$/.test(n))) {
    entetes.push(...texteDesParagraphes(analyserXml(await zip.texte(nom))));
  }
  const notes = [];
  const notesXml = await zip.texte('word/footnotes.xml');
  if (notesXml) {
    for (const f of tous(analyserXml(notesXml), 'w:footnote')) {
      if (f.a['w:type'] === 'separator' || f.a['w:type'] === 'continuationSeparator') continue;
      const t = texteDesParagraphes(f).join(' ');
      if (t.trim()) notes.push(t);
    }
  }
  const commentaires = [];
  const comXml = await zip.texte('word/comments.xml');
  if (comXml) for (const c of tous(analyserXml(comXml), 'w:comment')) commentaires.push(texteDesParagraphes(c).join(' '));

  const proprietes = {};
  for (const [fichier, champs] of [['docProps/core.xml', { 'dc:creator': 'Auteur', 'cp:lastModifiedBy': 'Modifié par',
    'dc:title': 'Titre', 'dc:subject': 'Objet', 'cp:keywords': 'Mots clés', 'dc:description': 'Commentaires' }],
  ['docProps/app.xml', { Company: 'Société', Manager: 'Responsable' }]]) {
    const xml = await zip.texte(fichier);
    if (!xml) continue;
    const r = analyserXml(xml);
    for (const [balise, libelle] of Object.entries(champs)) {
      const el = chercher(r, balise);
      const t = el ? texteBrut(el).trim() : '';
      if (t) proprietes[libelle] = t;
    }
  }

  return {
    source,
    blocs,
    tableaux,
    notes,
    commentaires,
    revisions: compte.revisions,
    entetes,
    proprietes,
    labelsLus: true,
    miseEnPage: {
      sections: miseEnPageSections,
      normal: {
        taille: (normal && propriete(styles, normal.id, 'taille')) ?? defauts.taille,
        interligne: (normal && propriete(styles, normal.id, 'interligne')) ?? defauts.interligne,
      },
      notes: {
        repriseParPage,
        taille: styleNote ? (propriete(styles, styleNote.id, 'taille') ?? defauts.taille) : null,
      },
    },
    diagnostic: { paragraphes: blocs.length, tableaux, sections: sections.length },
  };
}
