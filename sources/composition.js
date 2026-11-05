// Mise en forme : le texte du document ouvert (collé sans mise en forme) est versé dans un nouveau document créé
// depuis le gabarit au format du jury (styles, page de garde, table des matières, marges, numéros de page).
// Le document d'origine n'est pas modifié. Aucun texte n'est écrit par l'outil : il range et met en forme.

import { lireZip, ecrireZip } from './zip.js';
import { analyserStructure, lirePlan } from '../moteur/structure.js';
import { lireDocx } from './docx.js';
import { pagesDesTitres } from '../moteur/moteur.js';
import { cellulesTapees } from '../moteur/controles.js';
import { normaliser } from '../moteur/texte.js';

const ZONES_CHAMP = new Set(['tdm', 'sommaire', 'liste-annexes']); // tables produites par Word (champs)
const PREFIXE_PLAN = /^(partie|chapitre|section|§|paragraphe)\s*\d+(?:\.\d+)*\s*[-–—:.]?\s*/i;
const NIVEAU_PLAN = { introduction: 0, conclusion: 0, partie: 1, chapitre: 2, section: 3, '§': 4, paragraphe: 4 };
// « Introduction de la partie », « Conclusion partielle du chapitre »… (texte normalisé, sans accents).
export const PARTIEL = /^(introduction|conclusion)(?: partielle| generale)? (?:de la|du|de) (partie|chapitre)\b/;
const PAGES_EN_FIN = /\s*[(–—-]?\s*(\d+(?:[,.]\d+)?)\s*(?:pages?|p\.)\s*\)?\s*$/i;
const PAGES_DANS = /\s*\((\d+(?:[,.]\d+)?)\s*(?:pages?|p\.)\)/i;
const OUTIL_EN_FIN = /\s+[–—-]\s+(outils?\s+(?:n°\s*)?\d+\b.*)$/i;
const LEGENDE = /^(tableau|figure|graphique|sch[ée]ma)\s*(?:n°\s*)?\d+\s*[-–—:.]/i;
const SOURCE = /^(sources?\s*:|r[ée]alis[ée] par l'auteur)/i;
const PUCE = /^[-–•▪◦]\s+\S/;

/**
 * Ligne d'un plan de la notice : niveau (0 introduction ou conclusion, 1 à 4 P/C/S/§), intitulé, pages et outil,
 * tels que tapés (les pages passent dans la colonne « pages » du tableau du plan).
 */
export function lignePlan(texte) {
  const n = normaliser(texte).replace(/^[•▪◦-]\s*/, '');
  const m = /^(introduction|conclusion|partie|chapitre|section|§|paragraphe)(?![a-z])/.exec(n);
  if (!m) return null;
  let titre = texte.replace(/^[•▪◦-]\s*/, '');
  // Introduction ou conclusion partielle d'une partie ou d'un chapitre (Note, p. 15) : sous-ligne du niveau au-dessus,
  // comptée dans sa somme de pages (A12), en italique et sans fond dans le tableau.
  const partiel = PARTIEL.exec(n);
  let outil = null;
  const o = OUTIL_EN_FIN.exec(titre);
  if (o) { outil = o[1]; titre = titre.slice(0, o.index); }
  let pages = null;
  const p = PAGES_EN_FIN.exec(titre) || PAGES_DANS.exec(titre);
  if (p) { pages = p[1].replace('.', ','); titre = (titre.slice(0, p.index) + titre.slice(p.index + p[0].length)).replace(/[^\S\u00a0\u202f]+/g, ' ').trim(); }
  if (partiel) return { niveau: partiel[2] === 'partie' ? 2 : 3, titre, pages, outil, partiel: true };
  return { niveau: NIVEAU_PLAN[m[1]], titre, pages, outil };
}

/**
 * Plan de mise en forme : pour chaque paragraphe du texte, le style du gabarit (ou l'omission, avec sa raison).
 * doc : modèle lu (office.js ou docx.js). Rend { lignes: [{ i, texte, style, gras?, champ?, raison? }], avertissements }.
 */
export function planifier(doc, regles, profil) {
  const copie = structuredClone(doc);
  const S = analyserStructure(copie, regles);
  const lignes = [];
  const avertissements = [];
  const tableaux = new Set();
  let images = 0;
  const dansListe = (i) => [...ZONES_CHAMP].some((z) => S.dans(z, i));

  copie.blocs.forEach((b, k) => {
    const texte = (b.intitule || b.texte || '').replace(/\t/g, ' ').replace(/\s+$/, '');
    if (b.garde) { lignes.push({ i: b.i, texte, omis: 'page de garde (remplie par les champs ci-dessus)' }); return; }
    if (b.consigne) { lignes.push({ i: b.i, texte, omis: 'consigne du gabarit' }); return; }
    if (b.tableau !== null) {
      // Le tableau est repris en bloc à l'emplacement de sa première cellule.
      if (!tableaux.has(b.tableau)) lignes.push({ i: b.i, texte, tableau: b.tableau });
      tableaux.add(b.tableau);
      return;
    }
    if (b.image) images++;
    if (!b.titre && dansListe(b.i)) { lignes.push({ i: b.i, texte, omis: 'entrée de table (refaite par Word)' }); return; }
    if (!texte.trim()) return;

    if (b.titre && b.zone && ZONES_CHAMP.has(b.zone)) { lignes.push({ i: b.i, texte, champ: b.zone }); return; }
    if (b.titre && b.zone) { lignes.push({ i: b.i, texte, style: 'Titrehorsplan' }); return; }
    const dansPlan = S.dans('plan-det', b.i) || S.dans('plan-synth', b.i);
    // Plans de la notice : tableau « intitulé | pages » (niveaux en retrait, parties et chapitres marqués).
    if (profil === 'notice' && dansPlan) {
      const pl = lignePlan(texte);
      if (pl) { lignes.push({ i: b.i, texte, planLigne: pl }); return; }
    }
    // Tableau tapé en texte (cellules séparées par « | ») : un vrai tableau, première ligne en en-tête.
    const cellules = !dansPlan && cellulesTapees(texte);
    if (cellules) {
      const dernier = lignes[lignes.length - 1];
      if (dernier && dernier.tableauTexte && dernier.tableauTexte[0].length === cellules.length) dernier.tableauTexte.push(cellules);
      else lignes.push({ i: b.i, texte, tableauTexte: [cellules] });
      return;
    }
    if (LEGENDE.test(texte) && texte.length <= 200) { lignes.push({ i: b.i, texte, style: 'Caption' }); return; }
    if (SOURCE.test(texte) && texte.length <= 300) { lignes.push({ i: b.i, texte, style: 'Source' }); return; }
    if (S.dans('biblio', b.i) && !b.titre) { lignes.push({ i: b.i, texte, style: 'Referencebiblio' }); return; }
    if (PUCE.test(texte) && !dansPlan) { lignes.push({ i: b.i, texte, style: 'Listetirets' }); return; }
    if (b.titre && b.role === 'section-notice') { lignes.push({ i: b.i, texte, style: 'Titrenotice' }); return; }
    if (b.titre && b.role === 'categorie') { lignes.push({ i: b.i, texte, style: 'Titrecatégorie' }); return; }
    // Annexe : numéro automatique du gabarit (« Annexe n – »), le numéro tapé ou hérité est retiré.
    if (b.titre && b.role === 'annexe') { lignes.push({ i: b.i, texte: texte.replace(/^annexe\s*(?:n°\s*)?\d+\s*[-–—:.]?\s*/i, ''), style: 'Titreannexe' }); return; }
    // Texte simple du mémoire : « Partie 1 – … », « Chapitre 2 – … » reconnus à leur début.
    const pl = !dansPlan && texte.length <= 160 && !/[.;,]$/.test(texte) ? lirePlan(texte) : null;
    if (profil === 'memoire' && !dansPlan && (pl || (b.titre && b.role === 'plan'))) {
      const niveau = pl ? pl.niveau : b.niveau;
      // Titres du mémoire : numéros automatiques du gabarit (« Partie n – »), le numéro tapé est retiré.
      lignes.push({ i: b.i, texte: texte.replace(PREFIXE_PLAN, ''), style: `Heading${Math.min(4, Math.max(1, niveau || 1))}` });
      return;
    }
    // Notice : intitulé de section numéroté « 3. … » (même hors des sept intitulés de la Note) en style de section.
    const dansNotice = S.notice && b.i >= S.notice.debut && b.i < S.notice.fin;
    if (profil === 'notice' && dansNotice && /^\d\s*[.)]\s+\S/.test(texte) && texte.length <= 120 && !/[.;,:]$/.test(texte)) {
      lignes.push({ i: b.i, texte, style: 'Titrenotice' });
      return;
    }
    // Sous-titre court de la notice (titre stylé, ou ligne courte suivie d'un paragraphe plus long) : en gras.
    const suivant = copie.blocs.slice(k + 1).find((x) => x.texte.trim());
    const sousTitre = !dansPlan && !/^\[/.test(texte) && S.notice && b.i >= S.notice.debut && b.i < S.notice.fin
      && (b.titre || (texte.length <= 120 && !/[.:;,]$/.test(texte) && suivant && suivant.texte.length > texte.length * 1.5 && !/^[•▪◦\-–]/.test(texte)));
    // Sous-titre de la notice : style « Sous-titre notice » (niveau 3, dans la table des matières) ; numéro « 1.1 »
    // retiré (découpage proscrit par la Note, p. 14), comme le propose A06.
    if (sousTitre) { lignes.push({ i: b.i, texte: texte.replace(/^\d+(?:\.\d+)+\.?\s+/, ''), style: 'Soustitrenotice' }); return; }
    lignes.push({ i: b.i, texte, style: 'Normal' });
  });
  ajouterReperes(lignes, copie, S, profil);
  const liens = copie.blocs.filter((b) => !b.garde && b.tableau === null && (b.liens || []).length).length;
  if (liens) avertissements.push(`${liens} paragraphe(s) avec lien hypertexte : le texte est repris, pas le lien (renvois aux annexes, URL) ; à recréer (Insertion, Lien).`);
  if (images) avertissements.push(`${images} image(s) non reprise(s) : les recopier dans le nouveau document.`);
  const sections = copie.blocs.filter((b) => b.role === 'section-notice').length;
  if (profil === 'notice' && sections < 7) avertissements.push(`${sections} section(s) sur 7 portent l'intitulé exact de la Note : après création, Garde-fou proposera de renommer les autres (règle A09).`);
  return { lignes, avertissements, tableaux: tableaux.size };
}

// Un repère « [Tableau n à insérer ici] » tient la place du tableau (pas de second repère dans la même section).
const estTableau = (l) => l && (l.tableau !== undefined || l.tableauTexte || l.planLigne || (l.repere && /^\[Tableau/.test(l.texte)));
const repere = (texte) => ({ i: null, texte, style: 'Consigne', repere: true });

/**
 * Repères « [Tableau n à insérer ici] » (style Consigne, signalés par F01 jusqu'à leur remplacement) :
 * légende « Tableau n – … » sans tableau à sa suite ; sections 3 et 5 de la notice sans tableau (Note, p. 12) ;
 * annexe n° 1 de la charte IA absente (notice). L'outil n'écrit que le repère, jamais le contenu.
 */
function ajouterReperes(lignes, doc, S, profil) {
  const visibles = () => lignes.filter((l) => !l.omis);
  // Légende sans tableau.
  for (const l of visibles()) {
    const m = l.style === 'Caption' && /^tableau\s*(?:n°\s*)?(\d+)/i.exec(l.texte);
    if (!m) continue;
    const v = visibles();
    const suivant = v[v.indexOf(l) + 1];
    if (!estTableau(suivant)) lignes.splice(lignes.indexOf(l) + 1, 0, repere(`[Tableau ${m[1]} à insérer ici]`));
  }
  if (profil !== 'notice') return;
  // Sections 3 et 5 : tableau exigé.
  const objets = { 3: 'tableau de différenciation avec les mémoires les plus proches', 5: 'tableau des outils (numéro d\'ordre dans le plan, nom, nature, objectifs)' };
  const sections = doc.blocs.filter((b) => b.role === 'section-notice' && objets[b.sectionNotice]);
  for (const b of sections.reverse()) {
    const debut = lignes.findIndex((l) => l.i === b.i);
    if (debut < 0) continue;
    let fin = lignes.findIndex((l, k) => k > debut && !l.omis && ['Titrenotice', 'Titrehorsplan'].includes(l.style));
    if (fin < 0) fin = lignes.length;
    if (lignes.slice(debut + 1, fin).some(estTableau)) continue;
    const avant = lignes.slice(0, debut).filter(estTableau).length;
    lignes.splice(fin, 0, repere(`[Tableau ${avant + 1} à insérer ici : ${objets[b.sectionNotice]}]`));
  }
  // Annexe n° 1 de la charte IA.
  if (!S.zone('charte') && !doc.blocs.some((b) => b.role === 'annexe')) {
    lignes.push(repere('[Annexe 1 à insérer ici : modèle officiel « Annexe n° 1 – Usage de l\'IA » de la charte, complété, daté et signé]'));
  }
}

/**
 * Tableau repris d'un autre document : liens, images, numérotations et révisions retirés (leurs références
 * n'existent pas dans le gabarit) ; texte, cellules et mise en forme des caractères conservés.
 */
export function nettoyerTableau(xml) {
  return xml
    .replace(/<mc:AlternateContent>[\s\S]*?<\/mc:AlternateContent>/g, '')
    .replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, '')
    .replace(/<w:pict>[\s\S]*?<\/w:pict>/g, '')
    .replace(/<w:object[\s\S]*?<\/w:object>/g, '')
    .replace(/<w:del\b[\s\S]*?<\/w:del>/g, '')
    .replace(/<\/?w:ins\b[^>]*>/g, '')
    .replace(/<\/?w:hyperlink\b[^>]*>/g, '')
    .replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, '')
    .replace(/<w:(footnoteReference|endnoteReference|commentReference)\b[^>]*\/>/g, '')
    .replace(/<w:(commentRangeStart|commentRangeEnd)\b[^>]*\/>/g, '')
    .replace(/<w:pStyle w:val="[^"]*"\/>/g, '')
    .replace(/\s(r:id|r:embed|r:link)="[^"]*"/g, '');
}

/** Tableaux (XML) d'un paquet Open XML : premier niveau du corps, dans l'ordre. */
export function tableauxDuCorps(documentXml) {
  const ouverture = documentXml.indexOf('<w:body>') + '<w:body>'.length;
  const fermeture = documentXml.lastIndexOf('</w:body>');
  return elementsDuCorps(documentXml.slice(ouverture, fermeture)).filter((x) => x.startsWith('<w:tbl'));
}

// ---------------------------------------------------------------- assemblage du .docx

// Typographie posée à la création (le collage « texte seulement » de Word en ligne remplace les insécables par des
// espaces ordinaires) : insécable avant : ; ! ? et à l'intérieur de « », apostrophe courbe entre deux lettres.
// Mise en forme seulement : aucun mot ajouté ni retiré. Les URL sont laissées telles quelles.
const NBSP = '\u00a0';
export function typographier(t) {
  if (!t || /https?:\/\/|www\./i.test(t)) return t;
  return t
    .replace(/ (?=[:;!?])/g, NBSP)
    .replace(/(?<=«) /g, NBSP)
    .replace(/ (?=»)/g, NBSP)
    .replace(/(?<=\p{L})'(?=\p{L})/gu, '’');
}

const echapper = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function paragraphe(texte, style, gras = false) {
  const rPr = gras ? '<w:rPr><w:b/></w:rPr>' : '';
  const runs = typographier(texte).split('\n').map((ligne, k) => `${k ? '<w:r><w:br/></w:r>' : ''}<w:r>${rPr}<w:t xml:space="preserve">${echapper(ligne)}</w:t></w:r>`).join('');
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${runs}</w:p>`;
}

/** Éléments de premier niveau du corps (paragraphes, tableaux, propriétés de section…), en XML. */
function elementsDuCorps(corps) {
  const elements = [];
  const re = /<(\/?)(w:p|w:tbl|w:sdt|w:sectPr)(?=[\s>/])[^>]*?(\/?)>/g;
  let profondeur = 0;
  let debut = -1;
  let m;
  while ((m = re.exec(corps))) {
    const fermant = m[1] === '/';
    const vide = m[3] === '/';
    if (!fermant && profondeur === 0) debut = m.index;
    if (!fermant && !vide) profondeur++;
    if (fermant) profondeur--;
    if (profondeur === 0 && debut >= 0) {
      elements.push(corps.slice(debut, m.index + m[0].length));
      debut = -1;
    }
  }
  return elements;
}

const texteDe = (xml) => (xml.match(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join('');
const styleDe = (xml) => (/<w:pStyle w:val="([^"]+)"/.exec(xml) || [])[1] || '';

/**
 * Compose le nouveau document. champs : { intitule, auteur, adresse, date } (notice) ou
 * { intitule, auteur, adresse, session } (mémoire), plus { tableaux, regles } (regles : tables pré-remplies). Rend un Uint8Array (.docx).
 */
export async function composer(gabarit, plan, profil, champs) {
  const zip = await lireZip(gabarit);
  const docXml = await zip.texte('word/document.xml');
  const ouverture = docXml.indexOf('<w:body>') + '<w:body>'.length;
  const fermeture = docXml.lastIndexOf('</w:body>');
  const elements = elementsDuCorps(docXml.slice(ouverture, fermeture));

  // Page de garde : jusqu'au paragraphe qui porte la fin de la première section ; consignes retirées.
  const finGarde = elements.findIndex((x) => x.startsWith('<w:p') && x.includes('<w:sectPr'));
  const remplacer = (xml) => {
    const t = texteDe(xml);
    // Notice (p. 11) : mention, intitulé, nom, adresse, mois et année. Mémoire (p. 14) : NOM – Prénom – adresse,
    // « Diplôme d'expertise comptable – Session… », « MÉMOIRE », TITRE en majuscules.
    const valeurs = profil === 'memoire' ? {
      '[Mentions de la page de garde du mémoire : voir la Note du jury en vigueur, p. 14]':
        champs.session ? `Diplôme d'expertise comptable – ${champs.session}\nMÉMOIRE` : null,
      '[Intitulé du sujet agréé]': champs.intitule && champs.intitule.toUpperCase(),
      '[Prénom NOM du candidat]': champs.auteur && [champs.auteur, champs.adresse].filter(Boolean).join(' – '),
      '[Session de soutenance]': champs.session ? '' : null,
    } : {
      '[Intitulé du sujet du mémoire]': champs.intitule,
      '[Prénom NOM du candidat]': champs.auteur,
      '[Adresse personnelle : au minimum code postal et ville]': champs.adresse,
      '[Mois et année de dépôt]': champs.date,
    };
    if (!(t in valeurs) || valeurs[t] === null || valeurs[t] === undefined) return xml;
    if (valeurs[t] === '') return '';
    return paragraphe(valeurs[t], styleDe(xml) || 'Pagedegarde');
  };
  const garde = elements.slice(0, finGarde + 1).filter((x) => styleDe(x) !== 'Consigne').map(remplacer);
  const sectionFinale = elements[elements.length - 1].startsWith('<w:sectPr') ? elements[elements.length - 1] : '';

  // Tables du gabarit (sommaire, liste des annexes, table des matières) : titre (style « Titre hors table », hors de la
  // table elle-même) et instruction du champ.
  const champsGabarit = {};
  elements.forEach((x, k) => {
    if (/\bTOC\b/.test(x) && k > 0 && styleDe(elements[k - 1]) === 'Titrehorsplan') {
      const titre = texteDe(elements[k - 1]).toLowerCase();
      const zone = titre.startsWith('sommaire') ? 'sommaire' : titre.startsWith('liste des annexes') ? 'liste-annexes' : 'tdm';
      let instr = (/<w:instrText[^>]*>([^<]*)</.exec(x) || [])[1] || ' TOC \\o "1-2" \\h \\z \\u ';
      // Notice : sous-titres (niveau 3) dans la table.
      if (profil === 'notice' && zone === 'tdm') instr = instr.replace(/\\o "1-2"/, '\\o "1-3"');
      // Annexes (style « Titre annexe ») dans la table des matières : mémoire et notice (annexe 1 de la charte IA).
      if (zone === 'tdm' && !/\\t /.test(instr)) instr = instr.replace(/\s*$/, ' \\t "Titre annexe,1" ');
      // Titre du sommaire et de la table des matières : style hors table (sinon la table se liste elle-même).
      const titreXml = zone === 'liste-annexes' ? elements[k - 1] : elements[k - 1].replace('<w:pStyle w:val="Titrehorsplan"/>', '<w:pStyle w:val="Titrehorstable"/>');
      champsGabarit[zone] = { titre: titreXml, instr };
    }
  });

  const corps = [];
  const poses = new Set();
  let lignesPlan = [];
  const fermerPlan = () => { if (lignesPlan.length) corps.push(tableauDuPlan(lignesPlan)); lignesPlan = []; };
  for (const l of plan.lignes) {
    if (l.omis) continue;
    if (l.planLigne) { lignesPlan.push(l.planLigne); continue; }
    fermerPlan();
    if (l.tableauTexte) { corps.push(tableauTape(l.tableauTexte)); continue; }
    if (l.tableau !== undefined) {
      const xml = (champs.tableaux || [])[l.tableau];
      corps.push(xml ? nettoyerTableau(xml) : paragraphe(`[Tableau ${l.tableau + 1} à recopier ici]`, 'Consigne'));
      continue;
    }
    if (l.champ) {
      if (champsGabarit[l.champ] && !poses.has(l.champ)) { corps.push({ champ: l.champ }); poses.add(l.champ); }
      continue;
    }
    // « Source : … » sous un tableau : un léger espace au-dessus (le style du gabarit n'en prévoit pas).
    const xml = paragraphe(l.texte, l.style, l.gras);
    corps.push(l.style === 'Source' ? xml.replace('<w:pStyle w:val="Source"/>', '<w:pStyle w:val="Source"/><w:spacing w:before="60"/>') : xml);
  }
  fermerPlan();
  // Un paragraphe vide après un tableau suivi d'un autre tableau, d'une table ou de la fin du corps (Word l'exige) ;
  // sinon le paragraphe suivant (« Source : … ») suit le tableau sans ligne blanche.
  for (let k = corps.length - 1; k >= 0; k--) {
    const x = corps[k];
    const suivant = corps[k + 1];
    if (typeof x === 'string' && x.startsWith('<w:tbl') && (suivant === undefined || typeof suivant !== 'string' || suivant.startsWith('<w:tbl'))) corps.splice(k + 1, 0, '<w:p/>');
  }
  // Tables attendues absentes du texte : sommaire en tête (mémoire), table des matières en tête (notice) ou en fin (mémoire).
  const tete = [];
  if (profil === 'notice' && !poses.has('tdm') && champsGabarit.tdm) tete.push({ champ: 'tdm' });
  if (profil === 'memoire' && !poses.has('sommaire') && champsGabarit.sommaire) tete.push({ champ: 'sommaire' });
  const fin = profil === 'memoire' && !poses.has('tdm') && champsGabarit.tdm ? [{ champ: 'tdm' }] : [];

  const styles = ajouterStyles(await zip.texte('word/styles.xml'));
  const assembler = (entrees) => {
    const morceaux = [...garde, ...tete, ...corps, ...fin, sectionFinale]
      .map((x) => (typeof x === 'string' ? x : champsGabarit[x.champ].titre + champTable(champsGabarit[x.champ].instr, entrees[x.champ] || [])));
    const document = docXml.slice(0, ouverture) + morceaux.join('') + docXml.slice(fermeture);
    return ecrire(zip, { 'word/document.xml': document, 'word/styles.xml': styles });
  };
  // Première passe : document sans entrées ; seconde : entrées des tables avec la page estimée de chaque titre
  // (Word les recalcule au clic « Mettre à jour la table », ou par le bouton du complément).
  const brouillon = await assembler({});
  if (!champs.regles) return brouillon;
  return assembler(entreesDesTables(await lireDocx(brouillon), champs.regles, champsGabarit));
}

// ---------------------------------------------------------------- tableaux composés

const LARGEUR_TEXTE = 9072; // A4, marges de 2,5 cm : largeur utile en vingtièmes de point
// Couleurs : bleu ardoise, lisible imprimé en noir et blanc ; jamais de rouge (réservé aux modifications après
// une décision 4.2, Note p. 13).
export const COULEURS = { titre: '1F3864', sousTitre: '2E5C8A', fondFort: 'DCE6F2', fondLeger: 'F2F2F2' };
const GRIS_FONCE = COULEURS.fondFort;
const GRIS_CLAIR = COULEURS.fondLeger;

/** Paragraphe de cellule (style « Texte tableau ») : { texte, gras, italique, souligne, retrait, centre }. */
function paragrapheCellule({ texte, gras, italique, souligne, retrait = 0, centre = false, lie = false }) {
  const rPr = `${gras ? '<w:b/>' : ''}${italique ? '<w:i/>' : ''}${souligne ? '<w:u w:val="single"/>' : ''}`;
  const pPr = `<w:pStyle w:val="Textetableau"/>${lie ? '<w:keepNext/>' : ''}${retrait ? `<w:ind w:left="${retrait}"/>` : ''}${centre ? '<w:jc w:val="center"/>' : ''}`;
  return `<w:p><w:pPr>${pPr}</w:pPr><w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${echapper(typographier(texte))}</w:t></w:r></w:p>`;
}

// Dernière rangée liée au paragraphe suivant : la mention « Source : … » ne part pas seule sur la page suivante.
function celluleXml(paragraphes, largeur, fond, lie = false) {
  const tcPr = `<w:tcW w:w="${largeur}" w:type="dxa"/>${fond ? `<w:shd w:val="clear" w:color="auto" w:fill="${fond}"/>` : ''}<w:vAlign w:val="center"/>`;
  const p = paragraphes.map((x) => paragrapheCellule({ ...x, lie })).join('') || (lie ? '<w:p><w:pPr><w:keepNext/></w:pPr></w:p>' : '<w:p/>');
  return `<w:tc><w:tcPr>${tcPr}</w:tcPr>${p}</w:tc>`;
}

function tableauXml(largeurs, rangees) {
  const grille = largeurs.map((w) => `<w:gridCol w:w="${w}"/>`).join('');
  const tblPr = `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${LARGEUR_TEXTE}" w:type="dxa"/><w:tblLayout w:type="fixed"/>`
    + '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="85" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar>'
    + '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>';
  const tr = rangees.map((r, k) => `<w:tr><w:trPr><w:cantSplit/>${r.entete ? '<w:tblHeader/>' : ''}</w:trPr>`
    + r.cellules.map((c, j) => celluleXml(c.paragraphes, largeurs[j], c.fond, k === rangees.length - 1 && rangees.length > 1)).join('') + '</w:tr>').join('');
  return `<w:tbl>${tblPr}<w:tblGrid>${grille}</w:tblGrid>${tr}</w:tbl>`;
}

/** Tableau tapé en texte (« a | b | c ») : colonnes au prorata du texte, première ligne en en-tête grisée. */
export function tableauTape(rangees) {
  const n = rangees[0].length;
  // Largeur de colonne : au prorata du texte (plafonné), jamais moins que le plus long mot de la colonne.
  const mot = (t) => Math.max(0, ...(t || '').split(/\s+/).map((x) => x.length));
  const poids = Array.from({ length: n }, (_, k) => Math.max(1, ...rangees.map((r) => Math.min(60, (r[k] || '').length))));
  const minimums = Array.from({ length: n }, (_, k) => Math.min(2000, (Math.max(...rangees.map((r) => mot(r[k]))) + 2) * 115));
  const reste = Math.max(0, LARGEUR_TEXTE - minimums.reduce((a, b) => a + b, 0));
  const somme = poids.reduce((a, b) => a + b, 0);
  const largeurs = poids.map((w, k) => minimums[k] + Math.round((w / somme) * reste));
  largeurs[n - 1] += LARGEUR_TEXTE - largeurs.reduce((a, b) => a + b, 0);
  return tableauXml(largeurs, rangees.map((r, k) => ({
    entete: k === 0,
    cellules: r.map((texte) => ({ fond: k === 0 ? GRIS_FONCE : null, paragraphes: [{ texte, gras: k === 0, centre: k === 0 }] })),
  })));
}

/**
 * Plan synthétique ou détaillé en tableau « intitulé | pages » : introduction, conclusion et parties en gras sur fond
 * gris foncé, chapitres en gras sur fond gris clair, sections en retrait, paragraphes en italique, outils soulignés.
 * Le nombre de pages tapé passe dans la seconde colonne (lu tel quel par les contrôles A12 et A13).
 */
export function tableauDuPlan(lignes) {
  const pages = 1100;
  return tableauXml([LARGEUR_TEXTE - pages, pages], lignes.map((l) => {
    const fond = l.partiel ? null : l.niveau <= 1 ? GRIS_FONCE : l.niveau === 2 ? GRIS_CLAIR : null;
    const gras = l.niveau <= 2 && !l.partiel;
    const retrait = [0, 0, 284, 567, 851][l.niveau];
    const paragraphes = [{ texte: l.titre, gras, italique: l.niveau === 4 || !!l.partiel, retrait }];
    if (l.outil) paragraphes.push({ texte: l.outil, souligne: true, retrait: retrait + 284 });
    return { cellules: [{ fond, paragraphes }, { fond, paragraphes: l.pages ? [{ texte: `${l.pages} ${Number(l.pages.replace(',', '.')) <= 1 ? 'page' : 'pages'}`, gras: l.niveau <= 1, centre: true }] : [] }] };
  }));
}

/** Entrées des tables d'un document composé : { tdm: [{ niveau, texte, page }], sommaire, 'liste-annexes' }. */
function entreesDesTables(doc, regles, champsGabarit) {
  const pages = pagesDesTitres(doc, regles);
  const niveauDe = (b) => {
    const m = /^heading (\d)$/.exec(b.style || '');
    if (m) return Number(m[1]);
    return { 'titre hors plan': 1, 'titre notice': 2, 'sous-titre notice': 3 }[b.style] || null;
  };
  // Pages numérotées à partir de 1 après la page de garde (section 2 du gabarit).
  const garde = Math.max(1, ...doc.blocs.filter((b) => b.garde).map((b) => pages.get(b.i) || 1));
  const titres = doc.blocs.filter((b) => b.titre && !b.garde && b.tableau === null && pages.has(b.i));
  const entree = (b, niveau) => ({ niveau, texte: `${b.label || ''}${b.texte}`.replace(/[^\S\u00a0\u202f]+/g, ' ').trim(), page: Math.max(1, pages.get(b.i) - garde) });
  const resultat = {};
  for (const [zone, { instr }] of Object.entries(champsGabarit)) {
    // Niveaux de plan (\\o "1-n") et styles nommés (\\t "Titre annexe,1").
    const o = /\\o "1-(\d)"/.exec(instr);
    const max = o ? Number(o[1]) : 0;
    const nommes = /\\t "Titre annexe,(\d)"/.exec(instr);
    resultat[zone] = titres.map((b) => ({ b, n: b.style === 'titre annexe' ? (nommes ? Number(nommes[1]) : null) : niveauDe(b) && niveauDe(b) <= max ? niveauDe(b) : null }))
      .filter((x) => x.n).map((x) => entree(x.b, x.n));
  }
  return resultat;
}

/** Champ de table (TOC) : entrées pré-remplies, ou consigne si aucune. */
function champTable(instr, entrees) {
  const debut = `<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve">${echapper(instr)}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const fin = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  if (!entrees.length) return `<w:p>${debut}<w:r><w:t>[Clic droit sur la table, « Mettre à jour la table »]</w:t></w:r>${fin}</w:p>`;
  return entrees.map((e, k) => `<w:p><w:pPr><w:pStyle w:val="TOC${e.niveau}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9060"/></w:tabs></w:pPr>${k ? '' : debut}`
    + `<w:r><w:t xml:space="preserve">${echapper(e.texte)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${e.page}</w:t></w:r>${k === entrees.length - 1 ? fin : ''}</w:p>`).join('');
}

/**
 * Styles ajoutés au gabarit s'ils manquent : « Titre hors table » (titre d'une table, absent de la table elle-même),
 * « Sous-titre notice » (niveau 3), entrées de table « toc 1 » à « toc 4 ».
 */
export function ajouterStyles(xml) {
  if (!xml) return xml;
  const ajouts = [];
  const existe = (id) => xml.includes(`w:styleId="${id}"`);
  if (!existe('Titrehorstable')) {
    const base = (/<w:style\b[^>]*w:styleId="Titrehorsplan"[^>]*>[\s\S]*?<\/w:style>/.exec(xml) || [])[0];
    if (base) {
      let copie = base.replace('w:styleId="Titrehorsplan"', 'w:styleId="Titrehorstable"').replace(/<w:name w:val="[^"]*"\/>/, '<w:name w:val="Titre hors table"/>');
      copie = /<w:outlineLvl /.test(copie) ? copie.replace(/<w:outlineLvl w:val="\d"\/>/, '<w:outlineLvl w:val="9"/>') : copie.replace('<w:pPr>', '<w:pPr><w:outlineLvl w:val="9"/>');
      ajouts.push(copie);
    }
  }
  // Listes à tirets (le tiret tapé est gardé) et références de la bibliographie : retrait suspendu.
  if (!existe('Listetirets')) {
    ajouts.push('<w:style w:type="paragraph" w:customStyle="1" w:styleId="Listetirets"><w:name w:val="Liste tirets"/><w:basedOn w:val="Normal"/>'
      + '<w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="567" w:hanging="283"/></w:pPr></w:style>');
  }
  if (!existe('Referencebiblio')) {
    ajouts.push('<w:style w:type="paragraph" w:customStyle="1" w:styleId="Referencebiblio"><w:name w:val="Référence bibliographique"/><w:basedOn w:val="Normal"/>'
      + '<w:qFormat/><w:pPr><w:spacing w:after="120"/><w:ind w:left="567" w:hanging="567"/><w:jc w:val="left"/></w:pPr></w:style>');
  }
  if (!existe('Soustitrenotice')) {
    ajouts.push('<w:style w:type="paragraph" w:customStyle="1" w:styleId="Soustitrenotice"><w:name w:val="Sous-titre notice"/><w:basedOn w:val="Normal"/>'
      + '<w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>');
  }
  for (const n of [1, 2, 3, 4]) {
    if (!existe(`TOC${n}`)) {
      ajouts.push(`<w:style w:type="paragraph" w:styleId="TOC${n}"><w:name w:val="toc ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>`
        + `<w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:before="0" w:after="60"/><w:ind w:left="${(n - 1) * 240}"/></w:pPr>${n === 1 ? '<w:rPr><w:b/></w:rPr>' : ''}</w:style>`);
    }
  }
  let sortie = ajouts.length ? xml.replace('</w:styles>', `${ajouts.join('')}</w:styles>`) : xml;
  for (const id of ['Titrehorsplan', 'Titrehorstable', 'Titrenotice', 'Titreannexe', 'Heading1', 'Heading2']) sortie = colorer(sortie, id, COULEURS.titre);
  for (const id of ['Soustitrenotice', 'Titrecatégorie', 'Heading3', 'Heading4']) sortie = colorer(sortie, id, COULEURS.sousTitre);
  return sortie;
}

/** Couleur du texte d'un style (remplace la couleur existante, sinon l'ajoute en fin des propriétés de caractère). */
function colorer(xml, id, couleur) {
  return xml.replace(new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`), (st) => {
    if (/<w:color\b[^>]*\/>/.test(st)) return st.replace(/<w:color\b[^>]*\/>/, `<w:color w:val="${couleur}"/>`);
    const rPr = st.lastIndexOf('</w:rPr>');
    return rPr >= 0 ? `${st.slice(0, rPr)}<w:color w:val="${couleur}"/>${st.slice(rPr)}` : st.replace('</w:style>', `<w:rPr><w:color w:val="${couleur}"/></w:rPr></w:style>`);
  });
}

/** Réécrit le .docx du gabarit avec les parties remplacées. */
async function ecrire(zip, remplacees) {
  const utf8 = new TextEncoder();
  const fichiers = [];
  for (const nom of zip.noms) {
    if (nom.endsWith('/')) continue;
    fichiers.push({ nom, donnees: nom in remplacees ? utf8.encode(remplacees[nom]) : await zip.octets(nom) });
  }
  return ecrireZip(fichiers);
}

/** Uint8Array → base64 (par morceaux, sans dépasser la pile). */
export function enBase64(u8) {
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
  return btoa(s);
}
