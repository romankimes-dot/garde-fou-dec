// Contrôles déterministes, un par règle de regles.json.
// Chaque contrôle reçoit (doc, S, ctx) et rend { occurrences: [{ i, extrait, message, niveau? }], infos: [] }.
// `niveau: 'avertissement'` sur une occurrence l'abaisse quand la règle est bloquante.

import { normaliser, cleTitre, extrait, masquerUrl, nombre, pages } from './texte.js';
import { estTexte, caracteres, lirePlan } from './structure.js';

const AV = 'avertissement';
const occ = (b, message, opts = {}) => ({
  i: b ? b.i : null,
  extrait: b ? extrait(b.intitule, opts.pos || 0) : opts.extrait || '',
  message,
  niveau: opts.niveau,
});

// ---------------------------------------------------------------- A. Conformité à la Note du jury

function A01(doc, S, { p }) {
  const n = S.notice;
  if (!n) return { occurrences: [], infos: ['Notice explicative non trouvée : volume non calculé.'] };
  const car = caracteres(S.blocsDe(n.debut, n.fin));
  const pg = car / p.caracteres_par_page;
  const infos = [`Notice explicative : ${nombre(car)} caractères, soit au moins ${pages(pg)} pages (limite : ${p.notice_pages_max}).`];
  const o = [];
  if (pg > p.notice_pages_max) {
    o.push(occ(n.titre, `Au moins ${pages(pg)} pages estimées, au-delà de ${p.notice_pages_max}. Tableaux et sauts de page en ajoutent encore.`));
  } else if (pg > p.notice_pages_max * 0.9) {
    o.push(occ(n.titre, `${pages(pg)} pages estimées : proche de la limite ; vérifier sur le PDF.`, { niveau: AV }));
  }
  return { occurrences: o, infos };
}

function A02(doc, S, { p }) {
  if (!S.corps) return { occurrences: [], infos: ['Introduction ou conclusion non trouvée : volume du corps non calculé.'] };
  const car = caracteres(S.blocsDe(S.corps.debut, S.corps.fin));
  const pg = car / p.caracteres_par_page;
  const t = doc.blocs[S.corps.debut];
  const o = [];
  if (pg > p.memoire_pages_max) o.push(occ(t, `Au moins ${pages(pg)} pages de l'introduction à la conclusion, au-delà de ${p.memoire_pages_max}.`));
  else if (pg < p.memoire_pages_min) {
    o.push(occ(t, `${pages(pg)} pages estimées, en deçà de ${p.memoire_pages_min} : estimation basse (tableaux, figures et sauts de page en ajoutent) ; vérifier sur le PDF.`, { niveau: AV }));
  }
  return { occurrences: o, infos: [`Corps du mémoire : ${nombre(car)} caractères, soit au moins ${pages(pg)} pages (fourchette ${p.memoire_pages_min} à ${p.memoire_pages_max}).`] };
}

function A03(doc, S, { p }) {
  if (!S.annexes.length) return { occurrences: [], infos: ['Aucune annexe repérée.'] };
  const car = S.annexes.reduce((n, a) => n + caracteres(S.blocsDe(a.titre.i, a.fin)), 0);
  const pg = car / p.caracteres_par_page;
  const o = pg > p.annexes_pages_max ? [occ(S.annexes[0].titre, `Annexes : au moins ${pages(pg)} pages, au-delà de ${p.annexes_pages_max}.`)] : [];
  return { occurrences: o, infos: [`${S.annexes.length} annexe(s) : au moins ${pages(pg)} pages de texte (limite : ${p.annexes_pages_max}).`] };
}

function A04(doc, S, { p }) {
  const z = S.zone('synthese');
  if (!z) return { occurrences: [occ(null, 'Note de synthèse absente (titre « Note de synthèse » non trouvé).')], infos: [] };
  const pg = caracteres(S.blocsDe(z.debut, z.fin)) / p.caracteres_par_page;
  const o = [];
  if (pg < p.synthese_pages_min) o.push(occ(z.titre, `${pages(pg)} page(s) estimée(s), moins de ${p.synthese_pages_min}.`, { niveau: AV }));
  if (pg > p.synthese_pages_max) o.push(occ(z.titre, `Au moins ${pages(pg)} pages, plus de ${p.synthese_pages_max}.`));
  return { occurrences: o, infos: [] };
}

function partiesDuCorps(doc, S) {
  if (!S.corps) return [];
  return S.titres.filter((t) => t.role === 'plan' && t.niveau === 1 && t.i >= S.corps.debut && t.i < S.corps.fin);
}

function A05(doc, S, { p }) {
  const o = [];
  if (!S.corps) return { occurrences: o, infos: [] };
  const parties = partiesDuCorps(doc, S).map((t) => ({ t, car: caracteres(S.blocsDe(t.i, t.fin)) }));
  if (parties.length >= 2) {
    const max = Math.max(...parties.map((x) => x.car));
    const min = Math.min(...parties.map((x) => x.car));
    if (min > 0 && max / min > p.ecart_parties_max) {
      const detail = parties.map((x, k) => `P${k + 1} : ${pages(x.car / p.caracteres_par_page)} p.`).join(', ');
      o.push(occ(parties[0].t, `Parties déséquilibrées (${detail}) : rapport ${pages(max / min)}, seuil ${p.ecart_parties_max}.`));
    }
  }
  const total = caracteres(S.blocsDe(S.corps.debut, S.corps.fin));
  for (const nom of ['intro', 'conclusion']) {
    const z = S.zone(nom);
    if (!z || !total) continue;
    const part = caracteres(S.blocsDe(z.titre.i, z.fin)) / total;
    if (part > p.intro_conclusion_part_max) {
      o.push(occ(z.titre, `${nom === 'intro' ? 'Introduction' : 'Conclusion'} : ${Math.round(part * 100)} % du corps, au-delà de ${Math.round(p.intro_conclusion_part_max * 100)} %.`));
    }
  }
  return { occurrences: o, infos: [] };
}

const INTERDIT = /^\s*(?:[IVXLC]{1,6}\s*[.)\-–]|\d+(?:\.\d+)+\.?(?=\s)|[A-Za-z]\s*[.)](?=\s)|\d+\s*[.)](?=\s))/;
const MOT_ATTENDU = { 1: 'Partie', 2: 'Chapitre', 3: 'Section', 4: '§' };

/** Lignes du plan : titres du plan (mémoire) ou lignes des plans de la notice. */
function lignesDuPlan(doc, S, profil) {
  if (profil === 'notice') {
    const lignes = [];
    for (const nom of ['plan-det']) {
      const z = S.zone(nom);
      if (!z) continue;
      for (const b of S.blocsDe(z.debut, z.fin)) if (!b.consigne && b.intitule && b.tableau === null) lignes.push(b);
    }
    return lignes;
  }
  return S.titres.filter((t) => t.role === 'plan' && (!S.corps || (t.i >= S.corps.debut && t.i < S.corps.fin)));
}

function A06(doc, S, { profil }) {
  const o = [];
  const cibles = new Set([...S.titres.filter((t) => t.role === 'plan')]);
  for (const nom of ['plan-synth', 'plan-det']) {
    const z = S.zone(nom);
    if (z) for (const b of S.blocsDe(z.debut, z.fin)) if (!b.consigne && b.intitule && b.tableau === null) cibles.add(b);
  }
  let sansNumero = 0;
  const decimaux = [];
  for (const b of [...cibles].sort((x, y) => x.i - y.i)) {
    const lab = b.intitule.replace(/^[•▪◦\-–]\s*/, '');
    if (INTERDIT.test(lab)) {
      o.push(occ(b, 'Numérotation proscrite (I, 1.1, A, a)…) : utiliser Partie / Chapitre / Section / §.'));
      continue;
    }
    const pl = lirePlan(lab);
    if (pl && pl.decimal) decimaux.push({ b, pl });
    if (!b.titre || b.role !== 'plan' || b.niveau < 1 || b.niveau > 4) continue;
    if (!pl) {
      if (doc.labelsLus) sansNumero++;
      if (doc.labelsLus && sansNumero <= 10) {
        o.push(occ(b, `Titre de niveau ${b.niveau} sans « ${MOT_ATTENDU[b.niveau]} n – » : style de titre du gabarit à appliquer.`, { niveau: AV }));
      }
    } else if (pl.niveau !== b.niveau && profil === 'memoire') {
      o.push(occ(b, `Titre de niveau ${b.niveau} numéroté « ${pl.mot} » : « ${MOT_ATTENDU[b.niveau]} » attendu.`));
    }
  }
  // Numéro décimal (« Section 1.2 ») : une seule alerte pour l'ensemble, sur le premier.
  if (decimaux.length) {
    const { b, pl } = decimaux[0];
    o.push(occ(b, `${decimaux.length} intitulé(s) à numéro décimal (« ${nomMot(pl.mot)} ${pl.brut} ») : numéro simple attendu (« ${nomMot(pl.mot)} ${pl.numero} »), la Note proscrivant le découpage 1.1.`, { niveau: AV }));
  }
  const infos = doc.labelsLus ? [] : ['Numéros automatiques non lus : seuls les numéros tapés sont contrôlés. Contrôle complet sur le .docx.'];
  return { occurrences: o, infos };
}

/** Arbre P/C/S/§ à partir des lignes du plan. */
function arbreDuPlan(lignes) {
  const racine = { niveau: 0, enfants: [], b: null };
  const pile = [racine];
  const ruptures = [];
  // Deux usages admis : numérotation reprise à 1 sous chaque niveau supérieur, ou continue sur tout le plan.
  const dernier = {};
  for (const b of lignes) {
    const pl = lirePlan(b.intitule);
    if (!pl) continue;
    while (pile.length > 1 && pile[pile.length - 1].niveau >= pl.niveau) pile.pop();
    const parent = pile[pile.length - 1];
    const noeud = { niveau: pl.niveau, numero: pl.numero, mot: pl.mot, b, enfants: [] };
    const freres = parent.enfants.filter((e) => e.niveau === pl.niveau);
    const attendu = freres.length ? freres[freres.length - 1].numero + 1 : 1;
    const continu = (dernier[pl.niveau] || 0) + 1;
    dernier[pl.niveau] = pl.numero;
    if (pl.numero !== attendu && pl.numero !== continu) ruptures.push({ noeud, attendu });
    if (parent.niveau && pl.niveau > parent.niveau + 1) ruptures.push({ noeud, saut: true });
    parent.enfants.push(noeud);
    pile.push(noeud);
  }
  return { racine, ruptures };
}

const nomMot = (mot) => (mot === '§' ? '§' : mot[0].toUpperCase() + mot.slice(1));

function A07(doc, S, { p, profil }) {
  const o = [];
  const { racine, ruptures } = arbreDuPlan(lignesDuPlan(doc, S, profil));
  for (const r of ruptures) {
    if (r.saut) o.push(occ(r.noeud.b, 'Niveau sauté : ce titre suit directement un niveau plus élevé de deux crans.'));
    else o.push(occ(r.noeud.b, `« ${nomMot(r.noeud.mot)} ${r.noeud.numero} » : ${r.attendu === 1 ? 'la numérotation doit repartir à 1' : `« ${r.attendu} » attendu`}.`));
  }
  const parties = racine.enfants.filter((e) => e.niveau === 1);
  if (!parties.length) {
    const ou = profil === 'notice' ? 'dans le plan détaillé' : 'entre l\'introduction et la conclusion';
    return { occurrences: [occ(null, `Aucune partie « Partie n » trouvée ${ou}.`)], infos: [] };
  }
  if (parties.length < p.parties_min || parties.length > p.parties_max) {
    o.push(occ(parties[0].b, `${parties.length} partie(s) : le plan en comporte ${p.parties_min} ou ${p.parties_max}.`));
  }
  for (const pa of parties) {
    const ch = pa.enfants.filter((e) => e.niveau === 2);
    if (ch.length < p.chapitres_min || ch.length > p.chapitres_max) {
      o.push(occ(pa.b, `Partie ${pa.numero} : ${ch.length} chapitre(s), ${p.chapitres_min} ou ${p.chapitres_max} attendus.`));
    }
  }
  return { occurrences: o, infos: [] };
}

function A08(doc, S, { profil }) {
  const o = [];
  const { racine } = arbreDuPlan(lignesDuPlan(doc, S, profil));
  const visiter = (n) => {
    for (const niveau of [2, 3, 4]) {
      const sous = n.enfants.filter((e) => e.niveau === niveau);
      if (sous.length === 1 && n.niveau >= 1) {
        o.push(occ(sous[0].b, `« ${nomMot(sous[0].mot)} ${sous[0].numero} » est seul sous « ${nomMot(n.mot)} ${n.numero} ».`));
      }
    }
    n.enfants.forEach(visiter);
  };
  visiter(racine);
  return { occurrences: o, infos: [] };
}

function sectionsNotice(S) {
  const r = [];
  for (const t of S.titres) {
    if (t.role !== 'section-notice') continue;
    if (S.notice && S.notice.titre && (t.i < S.notice.debut || t.i >= S.notice.fin)) continue;
    if (!r[t.sectionNotice]) r[t.sectionNotice] = t;
  }
  return r;
}

function A09(doc, S, { p }) {
  const o = [];
  const s = sectionsNotice(S);
  let dernier = -1;
  p.notice_sections.forEach((def, k) => {
    const t = s[k + 1];
    if (!t) { o.push(occ(null, `Section ${k + 1} « ${def.titre} » absente.`)); return; }
    if (t.i < dernier) o.push(occ(t, `Section ${k + 1} placée avant la section ${k} : ordre imposé.`));
    dernier = Math.max(dernier, t.i);
    const contenu = S.blocsDe(t.i + 1, t.fin).filter((b) => (estTexte(b) || b.tableau !== null) && b.texte.trim());
    if (!contenu.length) o.push(occ(t, `Section ${k + 1} vide.`));
  });
  return { occurrences: o, infos: [] };
}

function A10(doc, S) {
  const o = [];
  const s = sectionsNotice(S);
  for (const [n, objet] of [[3, 'tableau de différenciation avec les mémoires les plus proches'], [5, 'tableau des outils (numéro d\'ordre dans le plan, nom, nature, objectifs)']]) {
    const t = s[n];
    if (!t) continue;
    if (!S.blocsDe(t.i + 1, t.fin).some((b) => b.tableau !== null)) o.push(occ(t, `Section ${n} sans tableau : ${objet} attendu.`));
  }
  return { occurrences: o, infos: [] };
}

function A11(doc, S) {
  const t = sectionsNotice(S)[1];
  if (!t) return { occurrences: [], infos: [] };
  const questions = [];
  for (const b of S.blocsDe(t.i + 1, t.fin)) {
    if (!estTexte(b)) continue;
    for (const m of b.texte.matchAll(/\?/g)) questions.push({ b, pos: m.index });
  }
  if (!questions.length) return { occurrences: [occ(t, 'Aucune phrase interrogative dans la section 1 : la problématique se formule en une question unique.')], infos: [] };
  if (questions.length > 1) {
    return { occurrences: [occ(questions[1].b, `${questions.length} questions dans la section 1 : vérifier que la problématique est une question unique.`, { pos: questions[1].pos, niveau: AV })], infos: [] };
  }
  return { occurrences: [], infos: [] };
}

const PAGES_PLAN = /(\d+(?:[,.]\d+)?)\s*(?:p\b\.?|pages?\b)/gi;
const lirePages = (t) => {
  const m = [...t.matchAll(PAGES_PLAN)];
  return m.length ? Number(m[m.length - 1][1].replace(',', '.')) : null;
};
const TYPE_LIGNE = /^(introduction|conclusion|partie|chapitre|section|§|paragraphe)\b/;

function A12(doc, S, { p }) {
  const z = S.zone('plan-det');
  if (!z) return { occurrences: [occ(null, 'Plan détaillé absent (titre « Plan détaillé » non trouvé).')], infos: [] };
  const o = [];
  const lignes = S.blocsDe(z.debut, z.fin).filter((b) => !b.consigne && b.intitule && b.tableau === null);
  let intro = null; let conclusion = null; let total = 0;
  const valeurs = [];
  for (const b of lignes) {
    const n = normaliser(b.intitule).replace(/^[•▪◦-]\s*/, '');
    const m = TYPE_LIGNE.exec(n);
    if (!m) continue;
    const pg = lirePages(b.intitule);
    if (m[1] === 'introduction') intro = b;
    if (m[1] === 'conclusion' && !/partielle/.test(n)) conclusion = b;
    if (pg === null) o.push(occ(b, 'Nombre prévisionnel de pages absent sur cette ligne.'));
    const pl = lirePlan(b.intitule);
    valeurs.push({ b, pl, pg, type: m[1] });
    if (pg !== null && (m[1] === 'introduction' || m[1] === 'partie' || (m[1] === 'conclusion' && b === conclusion))) total += pg;
  }
  if (!intro) o.push(occ(z.titre, 'Introduction absente du plan détaillé (avec son nombre de pages).'));
  if (!conclusion) o.push(occ(z.titre, 'Conclusion absente du plan détaillé (avec son nombre de pages).'));

  // Somme des sous-niveaux égale au niveau supérieur.
  for (let k = 0; k < valeurs.length; k++) {
    const v = valeurs[k];
    if (!v.pl || v.pg === null || v.pl.niveau >= 4) continue;
    let somme = 0; let vus = 0;
    for (let j = k + 1; j < valeurs.length; j++) {
      const w = valeurs[j];
      if (!w.pl || w.pl.niveau <= v.pl.niveau) break;
      if (w.pl.niveau === v.pl.niveau + 1 && w.pg !== null) { somme += w.pg; vus++; }
    }
    if (vus && Math.abs(somme - v.pg) > 0.5) {
      o.push(occ(v.b, `${v.pg} pages annoncées, ${somme} pour la somme des sous-niveaux.`, { niveau: AV }));
    }
  }
  if (total && (total < p.memoire_pages_min || total > p.memoire_pages_max)) {
    o.push(occ(z.titre, `Total prévisionnel ${total} pages (introduction, parties, conclusion) : hors de la fourchette ${p.memoire_pages_min} à ${p.memoire_pages_max}.`, { niveau: AV }));
  }

  // Outils : numéro et nom dans les paragraphes, et rapprochement avec le tableau de la section 5.
  const motif = new RegExp(p.outil_motif, 'g');
  const texte = lignes.map((b) => b.intitule).join('\n');
  const cites = new Set([...texte.matchAll(motif)].map((m) => normaliser(m[0])));
  if (!cites.size) o.push(occ(z.titre, 'Aucun outil cité dans le plan détaillé (numéro et nom dans les paragraphes concernés).'));
  const s5 = sectionsNotice(S)[5];
  if (s5) {
    const premiereColonne = S.blocsDe(s5.i + 1, s5.fin).filter((b) => b.tableau !== null && b.colonne === 0 && b.ligne > 0);
    for (const b of premiereColonne) {
      for (const m of b.texte.matchAll(motif)) {
        if (!cites.has(normaliser(m[0]))) o.push(occ(b, `« ${m[0]} » figure au tableau des outils mais pas dans le plan détaillé.`, { niveau: AV }));
      }
    }
  }
  return { occurrences: o, infos: total ? [`Plan détaillé : ${total} pages prévisionnelles (introduction, parties, conclusion).`] : [] };
}

function A13(doc, S) {
  const z = S.zone('plan-synth');
  if (!z) return { occurrences: [occ(null, 'Plan synthétique absent (titre « Plan synthétique » non trouvé).')], infos: [] };
  const o = [];
  for (const b of S.blocsDe(z.debut, z.fin)) {
    if (b.consigne || !b.intitule || b.tableau !== null) continue;
    const pl = lirePlan(b.intitule);
    if (!pl) continue;
    if (pl.niveau > 2) o.push(occ(b, 'Le plan synthétique ne fait apparaître que les parties et chapitres.'));
    else if (lirePages(b.intitule) === null) o.push(occ(b, 'Nombre prévisionnel de pages absent.'));
  }
  return { occurrences: o, infos: [] };
}

const NOMS_PIECES = {
  tdm: 'Table des matières', notice: 'Notice explicative', 'plan-synth': 'Plan synthétique', 'plan-det': 'Plan détaillé',
  biblio: 'Bibliographie', charte: 'Annexe 1 de la charte IA', sommaire: 'Sommaire', intro: 'Introduction',
  conclusion: 'Conclusion', 'liste-annexes': 'Liste des annexes', synthese: 'Note de synthèse',
};

function A14(doc, S, { profil }) {
  const o = [];
  const ordre = profil === 'notice'
    ? ['tdm', 'notice', 'plan-synth', 'plan-det', 'biblio', 'charte']
    : ['sommaire', 'intro', 'conclusion', 'biblio'];
  const trouves = [];
  for (const nom of ordre) {
    const z = S.zone(nom);
    if (!z) {
      if (!(profil === 'notice' && nom === 'notice' && S.notice)) o.push(occ(null, `Pièce absente : « ${NOMS_PIECES[nom]} ».`));
      continue;
    }
    trouves.push({ nom, i: z.titre.i, b: z.titre });
  }
  for (let k = 1; k < trouves.length; k++) {
    if (trouves[k].i < trouves[k - 1].i) {
      o.push(occ(trouves[k].b, `« ${NOMS_PIECES[trouves[k].nom]} » placé avant « ${NOMS_PIECES[trouves[k - 1].nom]} ».`));
    }
  }
  if (profil === 'memoire') {
    const tdm = S.zone('tdm');
    if (!tdm) o.push(occ(null, 'Pièce absente : « Table des matières » (détaillée, en fin de mémoire).'));
    else if (S.titres.some((t) => t.niveau === 1 && t.i > tdm.titre.i && !t.zone)) {
      o.push(occ(tdm.titre, 'La table des matières détaillée se place à la fin du mémoire.'));
    }
    if (S.annexes.length) {
      const liste = S.zone('liste-annexes');
      if (!liste) o.push(occ(S.annexes[0].titre, 'Liste des annexes absente avant les annexes.'));
      else if (liste.titre.i > S.annexes[0].titre.i) o.push(occ(liste.titre, 'La liste des annexes précède les annexes.'));
    }
    if (S.corps) {
      for (const pa of S.titres.filter((t) => t.role === 'plan' && t.niveau === 1)) {
        if (pa.i < S.corps.debut || pa.i >= S.corps.fin) o.push(occ(pa, 'Partie placée hors de l\'introduction et de la conclusion.'));
      }
    }
  }
  return { occurrences: o, infos: [] };
}

const MOIS = /(janvier|fevrier|mars|avril|mai|juin|juillet|aout|septembre|octobre|novembre|decembre)\s+20\d\d/;

function A15(doc, S, { p, profil }) {
  const o = [];
  const garde = doc.blocs.filter((b) => b.garde);
  if (!garde.length) return { occurrences: [occ(null, 'Page de garde non repérée (rien avant le premier titre).', { niveau: AV })], infos: [] };
  const texte = normaliser(garde.map((b) => b.intitule).join(' '));
  if (profil === 'notice') {
    if (!texte.includes(normaliser(p.mention_page_de_garde_notice))) {
      o.push(occ(garde[0], `Mention exacte absente : « ${p.mention_page_de_garde_notice} ».`));
    }
    if (!/\b\d{5}\b|\b\d{2}\s\d{3}\b/.test(texte)) o.push(occ(garde[0], 'Code postal non trouvé (adresse personnelle : au minimum code postal et ville).', { niveau: AV }));
    if (!MOIS.test(texte)) o.push(occ(garde[0], 'Mois et année de dépôt non trouvés (« novembre 2026 »).', { niveau: AV }));
  }
  const image = garde.find((b) => b.image);
  if (image) o.push(occ(image, 'Image sur la page de garde : vérifier qu\'il ne s\'agit pas d\'un logo du cabinet.', { niveau: AV }));
  return { occurrences: o, infos: [] };
}

function A16(doc, S, { termes }) {
  const liste = (termes || []).map((t) => t.trim()).filter(Boolean);
  if (!liste.length) {
    return { occurrences: [occ(null, 'Liste des termes du cabinet vide : la saisir dans Réglages (elle reste dans ce navigateur).', { niveau: AV })], infos: [] };
  }
  const motifs = liste.map((t) => ({ t, re: new RegExp(`(?<![\\p{L}\\p{N}])${normaliser(t).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![\\p{L}\\p{N}])`, 'u') }));
  const o = [];
  const tester = (texte, b, ou) => {
    const n = normaliser(texte);
    for (const { t, re } of motifs) {
      const m = re.exec(n);
      if (m) o.push(b ? occ(b, `Terme du cabinet « ${t} »${ou ? ` (${ou})` : ''}.`, { pos: m.index }) : occ(null, `Terme du cabinet « ${t} » : ${ou}.`, { extrait: extrait(texte, m.index) }));
    }
  };
  for (const b of doc.blocs) tester(b.intitule, b, null);
  for (const t of doc.entetes || []) tester(t, null, 'en-tête ou pied de page');
  for (const t of doc.notes || []) tester(t, null, 'note de bas de page');
  for (const t of doc.commentaires || []) tester(t, null, 'commentaire');
  for (const [k, v] of Object.entries(doc.proprietes || {})) tester(v, null, `propriétés du fichier, champ « ${k} »`);
  return { occurrences: o, infos: [] };
}

function A17(doc, S, { profil }) {
  if (profil === 'notice') {
    const z = S.zone('charte');
    if (!z) return { occurrences: [occ(null, 'Annexe 1 de la charte IA absente (titre « Annexe 1 de la charte d\'usage de l\'IA »).')], infos: [] };
    const contenu = S.blocsDe(z.debut, z.fin).some((b) => (estTexte(b) && b.texte.trim()) || b.image || b.tableau !== null);
    return { occurrences: contenu ? [] : [occ(z.titre, 'Annexe 1 de la charte IA vide : tableau des usages complété et signé attendu.')], infos: [] };
  }
  const a1 = S.annexes[0];
  if (!a1) return { occurrences: [occ(null, 'Aucune annexe : l\'annexe 1 doit être l\'annexe n° 1 de la charte IA.')], infos: [] };
  const ia = /intelligence artificielle|\bia\b/.test(normaliser(a1.titre.intitule));
  return { occurrences: ia ? [] : [occ(a1.titre, 'La première annexe doit être l\'annexe n° 1 de la charte IA (usage de l\'IA).')], infos: [] };
}

function A18(doc, S, { p }) {
  const z = S.zone('biblio');
  if (!z) return { occurrences: [occ(null, 'Bibliographie absente.')], infos: [] };
  const o = [];
  const cats = [];
  let courante = null;
  let horsCategorie = null;
  for (const b of S.blocsDe(z.debut, z.fin)) {
    if (b.consigne || !b.texte.trim()) continue;
    if (b.titre && b.role === 'categorie') {
      courante = { b, n: b.categorie || null, entrees: [] };
      if (!courante.n) o.push(occ(b, 'Catégorie non prévue par la Note (six catégories fixées).'));
      cats.push(courante);
    } else if (!b.titre) {
      if (courante) courante.entrees.push(b);
      else if (!horsCategorie) horsCategorie = b;
    }
  }
  if (horsCategorie) o.push(occ(horsCategorie, 'Référence placée avant toute catégorie.', { niveau: AV }));
  const numeros = cats.filter((c) => c.n).map((c) => c.n);
  for (let k = 1; k < numeros.length; k++) {
    if (numeros[k] <= numeros[k - 1]) {
      const c = cats.filter((x) => x.n)[k];
      o.push(occ(c.b, `« ${p.bibliographie_categories[c.n - 1].titre} » placée après « ${p.bibliographie_categories[numeros[k - 1] - 1].titre} » : ordre des six catégories.`));
    }
  }
  let total = 0;
  for (const c of cats) {
    total += c.entrees.length;
    if (!c.entrees.length) { o.push(occ(c.b, 'Catégorie vide : la supprimer ou la compléter.', { niveau: AV })); continue; }
    const def = c.n ? p.bibliographie_categories[c.n - 1] : null;
    if (!def || !def.alpha) continue;
    for (let k = 1; k < c.entrees.length; k++) {
      const a = normaliser(c.entrees[k - 1].texte);
      const b = normaliser(c.entrees[k].texte);
      if (a.localeCompare(b, 'fr') > 0) {
        o.push(occ(c.entrees[k], 'Ordre alphabétique du premier auteur non respecté dans la catégorie.', { niveau: AV }));
        break;
      }
    }
  }
  return { occurrences: o, infos: [`Bibliographie : ${total} référence(s) dans ${cats.length} catégorie(s) (la Note juge une trentaine suffisante).`] };
}

function A19(doc, S, { p }) {
  const mp = doc.miseEnPage;
  if (!mp) return { occurrences: [], infos: ['Mise en page contrôlable sur le .docx seulement.'] };
  const o = [];
  const cm = (v) => v.toLocaleString('fr-FR', { maximumFractionDigits: 2 });
  mp.sections.forEach((s, k) => {
    for (const [cote, v] of [['gauche', s.gauche], ['droite', s.droite]]) {
      if (v !== null && Math.abs(v - 2.5) > 0.05) o.push(occ(doc.blocs[s.premierBloc] || null, `Section ${k + 1} du document : marge ${cote} de ${cm(v)} cm, 2,5 cm attendus.`));
    }
  });
  const apresGarde = mp.sections.length > 1 ? mp.sections.slice(1) : mp.sections;
  if (!apresGarde.some((s) => s.pieds.some((f) => f.numero))) o.push(occ(null, 'Aucun numéro de page dans les pieds de page.'));
  else if (!apresGarde.some((s) => s.pieds.some((f) => f.numero && f.droite))) o.push(occ(null, 'Numéro de page non aligné à droite.'));
  if (mp.normal.taille && mp.normal.taille !== 12) o.push(occ(null, `Style Normal en taille ${mp.normal.taille}, 12 attendu.`));
  if (mp.normal.interligne && Math.abs(mp.normal.interligne - 1.5) > 0.01) o.push(occ(null, `Style Normal en interligne ${cm(mp.normal.interligne)}, 1,5 attendu.`));
  const exemptes = new Set(p.styles_exemptes);
  let n = 0;
  for (const b of doc.blocs) {
    if (!estTexte(b) || b.tableau !== null || exemptes.has(b.style) || !b.direct) continue;
    const tailles = b.direct.tailles.filter((t) => t !== 12);
    const il = b.direct.interligne;
    if (tailles.length || (il && Math.abs(il - 1.5) > 0.01)) {
      n++;
      if (n <= 15) o.push(occ(b, `Mise en forme directe : ${tailles.length ? `taille ${tailles.join(', ')}` : ''}${tailles.length && il ? ' ; ' : ''}${il && Math.abs(il - 1.5) > 0.01 ? `interligne ${cm(il)}` : ''}.`));
    }
  }
  if (n > 15) o.push(occ(null, `… et ${n - 15} autre(s) paragraphe(s) en mise en forme directe.`));
  return { occurrences: o, infos: [] };
}

// ---------------------------------------------------------------- B. Structure

const ENTREE_PAGE = /[\s.…·_-]*\d+\s*$/;

function entreesTable(S, nom) {
  const z = S.zone(nom);
  if (!z) return null;
  return S.blocsDe(z.debut, z.fin).filter((b) => !b.consigne && b.texte.trim() && (b.role === 'toc' || !b.titre));
}

function B01(doc, S, { profil }) {
  const o = [];
  const titresTexte = S.titres.filter((t) => !S.dans('plan-det', t.i) && !S.dans('plan-synth', t.i) && !S.dans('sommaire', t.i) && !S.dans('tdm', t.i));
  const cles = new Set(titresTexte.map((t) => cleTitre(t.intitule)));
  const tables = profil === 'notice' ? [['tdm', 2]] : [['sommaire', 2], ['tdm', 4]];
  for (const [nom, profondeur] of tables) {
    const entrees = entreesTable(S, nom);
    if (!entrees) continue;
    const z = S.zone(nom);
    const utiles = entrees.filter((b) => !/^\[.*\]$/.test(b.texte.trim()));
    if (!utiles.length) { o.push(occ(z.titre, `« ${NOMS_PIECES[nom] || nom} » vide : clic droit, « Mettre à jour les champs ».`)); continue; }
    const lues = new Set(utiles.map((b) => cleTitre(b.intitule.replace(ENTREE_PAGE, ''))));
    for (const b of utiles) {
      const c = cleTitre(b.intitule.replace(ENTREE_PAGE, ''));
      if (c && !cles.has(c)) o.push(occ(b, 'Entrée sans titre correspondant : table à mettre à jour.'));
    }
    for (const t of titresTexte) {
      if (t.niveau > profondeur || t.zone === nom || (profil === 'memoire' && nom === 'sommaire' && t.role !== 'plan')) continue;
      if (t.role === 'categorie') continue;
      if (!lues.has(cleTitre(t.intitule))) o.push(occ(t, `Titre absent de « ${NOMS_PIECES[nom] || nom} » : table à mettre à jour.`));
    }
  }
  return { occurrences: o, infos: [] };
}

function B02(doc, S, { p }) {
  const entrees = entreesTable(S, 'sommaire');
  if (!entrees) return { occurrences: [], infos: [] };
  const o = [];
  const profonds = new Set(S.titres.filter((t) => t.role === 'plan' && t.niveau > 2).map((t) => cleTitre(t.intitule)));
  for (const b of entrees) {
    const c = cleTitre(b.intitule.replace(ENTREE_PAGE, ''));
    const pl = lirePlan(b.intitule);
    if ((pl && pl.niveau > 2) || profonds.has(c)) o.push(occ(b, 'Le sommaire ne reprend que les parties et chapitres.'));
  }
  const car = caracteres(entrees);
  if (car > p.caracteres_par_page) o.push(occ(S.zone('sommaire').titre, `Sommaire de ${nombre(car)} caractères : plus d'une page.`));
  return { occurrences: o, infos: [] };
}

function B03(doc, S, { p }) {
  const o = [];
  for (const t of S.titres) {
    const mots = cleTitre(t.intitule).split(/\s+/).filter(Boolean).length;
    if (mots > p.titre_mots_max) o.push(occ(t, `${mots} mots (seuil : ${p.titre_mots_max}).`));
  }
  return { occurrences: o, infos: [] };
}

function B04(doc, S) {
  const o = [];
  for (const t of S.titres) {
    if (t.role !== 'plan' || t.niveau > 3 || (S.corps && (t.i < S.corps.debut || t.i >= S.corps.fin))) continue;
    const suivant = doc.blocs.slice(t.i + 1).find((b) => !b.consigne && b.texte.trim());
    if (suivant && suivant.titre && suivant.niveau > t.niveau) o.push(occ(t, 'Titre suivi directement d\'un sous-titre : introduction et annonce du plan attendues.'));
  }
  return { occurrences: o, infos: [] };
}

function B05(doc, S) {
  const o = [];
  for (const pa of partiesDuCorps(doc, S)) {
    const conclusion = S.blocsDe(pa.i + 1, pa.fin).some((b) => /^conclusion/.test(cleTitre(b.intitule)) && !b.consigne);
    if (!conclusion) o.push(occ(pa, 'Pas de conclusion partielle repérée (titre ou paragraphe commençant par « Conclusion »).'));
  }
  return { occurrences: o, infos: [] };
}

const RENVOI_ANNEXE = /\bannexe\s*(?:n°\s*)?(\d+)\b(?!\s*de la charte)/gi;

function B06(doc, S) {
  const o = [];
  const numeros = new Set(S.annexes.map((a) => a.numero).filter((n) => n !== null));
  const cites = new Set();
  const horsAnnexes = (i) => !S.annexes.some((a) => i >= a.titre.i && i < a.fin) && !S.dans('liste-annexes', i)
    && !S.dans('tdm', i) && !S.dans('sommaire', i);
  for (const b of doc.blocs) {
    if (!estTexte(b) || !horsAnnexes(b.i)) continue;
    for (const m of b.texte.matchAll(RENVOI_ANNEXE)) {
      const n = Number(m[1]);
      cites.add(n);
      if (!numeros.has(n)) { o.push(occ(b, `Renvoi à l'annexe ${n}, qui n'existe pas.`, { pos: m.index })); continue; }
      const cliquable = b.liens.some((l) => new RegExp(`annexe\\s*(n°\\s*)?${n}\\b`, 'i').test(normaliser(l)));
      if (!cliquable) o.push(occ(b, `Renvoi à l'annexe ${n} non cliquable : lien vers le signet de l'annexe attendu.`, { pos: m.index }));
    }
  }
  for (const a of S.annexes) {
    if (a.numero === 1 && /intelligence artificielle|\bia\b/.test(normaliser(a.titre.intitule))) continue;
    if (a.numero !== null && !cites.has(a.numero)) o.push(occ(a.titre, `Annexe ${a.numero} jamais citée dans le mémoire.`, { niveau: AV }));
    const retour = S.blocsDe(a.debut, a.fin).some((b) => b.liens.length);
    if (!retour) o.push(occ(a.titre, 'Pas de lien de retour au corps du mémoire dans cette annexe.'));
  }
  return { occurrences: o, infos: [] };
}

function B07(doc, S, { profil }) {
  const o = [];
  const { racine } = arbreDuPlan(lignesDuPlan(doc, S, profil));
  const parties = racine.enfants.filter((e) => e.niveau === 1);
  if (!parties.length) return { occurrences: o, infos: [] };
  const maxChap = Math.max(0, ...parties.map((pa) => pa.enfants.filter((e) => e.niveau === 2).length));
  for (const b of doc.blocs) {
    if (!estTexte(b) || S.dans('plan-det', b.i) || S.dans('plan-synth', b.i) || S.dans('sommaire', b.i) || S.dans('tdm', b.i)) continue;
    for (const m of b.texte.matchAll(/\b(partie|chapitre)\s+(\d+)\b/gi)) {
      const n = Number(m[2]);
      const max = m[1].toLowerCase() === 'partie' ? parties.length : maxChap;
      if (n > max) o.push(occ(b, `« ${m[0]} » : le plan n'en compte que ${max}.`, { pos: m.index }));
    }
  }
  return { occurrences: o, infos: [] };
}

const SIGLE = /(?<![\p{L}\p{N}])([A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ0-9&]{1,9})(?![\p{L}\p{N}])/gu;
const ROMAIN = /^[IVXLCDM]+$/;
const NON_SIGLES = new Set(['TODO', 'FIXME', 'XXX', 'SOURCE', 'TROUVER']);

function B08(doc, S, { p }) {
  const o = [];
  const exemptes = new Set(p.sigles_exemptes);
  const liste = new Set();
  const zoneAbr = S.zone('abreviations');
  if (zoneAbr) {
    for (const b of S.blocsDe(zoneAbr.debut, zoneAbr.fin)) {
      const m = /^\s*([A-ZÀ-ÖØ-Þ][A-ZÀ-ÖØ-Þ0-9&]{1,9})\b/u.exec(b.texte);
      if (m && (b.tableau === null || b.colonne === 0)) liste.add(m[1]);
    }
  }
  const vus = new Map();
  for (const b of doc.blocs) {
    // Intitulé en capitales (titre tapé, ligne de plan) : ses mots ne sont pas des sigles.
    const lettres = b.texte.match(/\p{L}/gu) || [];
    if (lettres.length >= 8 && (b.texte.match(/\p{Lu}/gu) || []).length / lettres.length > 0.6) continue;
    if (!estTexte(b) || S.dans('biblio', b.i) || S.dans('abreviations', b.i) || S.dans('sommaire', b.i) || S.dans('tdm', b.i)) continue;
    const t = b.texte;
    for (const m of t.matchAll(SIGLE)) {
      const s = m[1];
      const apres = t.slice(m.index + s.length);
      if (ROMAIN.test(s) || exemptes.has(s) || NON_SIGLES.has(s) || /^-\d/.test(apres)
        || (s.match(/\p{Lu}/gu) || []).length < 2 || vus.has(s)) continue;
      vus.set(s, b);
      const avant = t.slice(0, m.index).trimEnd();
      const defini = (avant.endsWith('(') && /^\s*\)/.test(apres)) || /^\s*\(/.test(apres) || liste.has(s);
      if (!defini) o.push(occ(b, `« ${s} » : première occurrence sans développement (forme développée, puis sigle entre parenthèses).`, { pos: m.index }));
    }
  }
  if (zoneAbr) {
    for (const s of vus.keys()) if (!liste.has(s)) o.push(occ(vus.get(s), `« ${s} » absent de la liste des abréviations.`));
    for (const s of liste) if (!vus.has(s)) o.push(occ(zoneAbr.titre, `« ${s} » figure dans la liste des abréviations sans être employé.`));
  }
  return { occurrences: o, infos: [] };
}

// ---------------------------------------------------------------- F. Propreté

const RESTES = /\bTODO\b|\bFIXME\b|\bXXX\b|SOURCE À TROUVER|\?\?|\[(?:[^\]]{2,}|…|\.\.\.)\]/g;

function dansGuillemets(t, pos) {
  const avant = t.slice(0, pos);
  return (avant.match(/«/g) || []).length > (avant.match(/»/g) || []).length;
}

function F01(doc, S) {
  const o = [];
  let consignes = 0;
  for (const b of doc.blocs) {
    if (b.consigne && b.texte.trim()) {
      consignes++;
      if (consignes <= 10) o.push(occ(b, 'Consigne du gabarit à supprimer.'));
      continue;
    }
    for (const m of b.intitule.matchAll(RESTES)) {
      if (m[0].startsWith('[') && /^\[(…|\.\.\.)\]$/.test(m[0]) && dansGuillemets(b.intitule, m.index)) continue;
      o.push(occ(b, `Reste de travail : « ${m[0].length > 40 ? m[0].slice(0, 40) + '…' : m[0]} ».`, { pos: m.index }));
    }
  }
  if (consignes > 10) o.push(occ(null, `… et ${consignes - 10} autre(s) consigne(s).`));
  for (const t of doc.notes || []) for (const m of t.matchAll(RESTES)) o.push(occ(null, `Reste de travail dans une note : « ${m[0]} ».`, { extrait: extrait(t, m.index) }));
  return { occurrences: o, infos: [] };
}

function F02(doc) {
  const o = [];
  if (doc.commentaires.length) o.push(occ(null, `${doc.commentaires.length} commentaire(s) dans le document.`));
  if (doc.revisions) o.push(occ(null, `${doc.revisions} modification(s) suivie(s) non acceptée(s) ou refusée(s).`));
  const infos = doc.revisions === null || doc.revisions === undefined
    ? ['Modifications suivies non lisibles avec cette version de Word : contrôle complet sur le .docx.'] : [];
  return { occurrences: o, infos };
}

/** Blocs soumis aux contrôles de typographie. */
function blocsTypo(doc, S) {
  return doc.blocs.filter((b) => (estTexte(b) || (b.titre && b.role !== 'toc')) && !b.consigne
    && !S.dans('sommaire', b.i) && !S.dans('tdm', b.i) && !S.dans('liste-annexes', b.i));
}

function F03(doc, S) {
  const o = [];
  const paires = [['«', '»'], ['(', ')'], ['[', ']'], ['“', '”']];
  for (const b of blocsTypo(doc, S)) {
    const t = b.intitule.replace(/^\s*[\p{L}\d]{1,3}\)\s/u, '');
    const ecarts = paires.filter(([a, f]) => t.split(a).length !== t.split(f).length).map(([a, f]) => `${a} ${f}`);
    if (ecarts.length) o.push(occ(b, `Signes non appariés : ${ecarts.join(', ')}.`));
  }
  return { occurrences: o, infos: [] };
}

const INSECABLE = /[  ]/;

function F04(doc, S) {
  const o = [];
  for (const b of blocsTypo(doc, S)) {
    const t = masquerUrl(b.intitule);
    const ecarts = new Map();
    const noter = (cle, pos) => { if (!ecarts.has(cle)) ecarts.set(cle, pos); };
    for (const m of t.matchAll(/(.)([;:!?])/gu)) {
      const avant = m[1];
      if (INSECABLE.test(avant) || /[;:!?»)\]…]/.test(avant)) continue;
      if (m[2] === ':' && /\d/.test(avant) && /\d/.test(t[m.index + 2] || '')) continue;
      if (avant === ' ') noter(`espace ordinaire avant « ${m[2]} »`, m.index);
      else if (/[\p{L}\p{N}”)]/u.test(avant)) noter(`pas d'espace avant « ${m[2]} »`, m.index);
    }
    for (const m of t.matchAll(/«(.)/g)) if (!INSECABLE.test(m[1])) noter('espace insécable attendue après « « »', m.index);
    for (const m of t.matchAll(/(.)»/g)) if (!INSECABLE.test(m[1])) noter('espace insécable attendue avant « » »', m.index);
    const droit = t.indexOf('"');
    if (droit >= 0) noter('guillemets droits " " : utiliser « »', droit);
    const anglais = t.search(/[“”]/);
    if (anglais >= 0) noter('guillemets anglais : « » en premier niveau', anglais);
    if (ecarts.size) {
      const [premiere] = ecarts.values();
      o.push(occ(b, [...ecarts.keys()].join(' ; ') + '.', { pos: premiere }));
    }
  }
  return { occurrences: o, infos: [] };
}

function F05(doc, S) {
  const o = [];
  for (const b of blocsTypo(doc, S)) {
    const t = masquerUrl(b.texte).replace(/\t/g, '');
    const ecarts = [];
    let pos = null;
    const chercher = (re, msg) => { const m = re.exec(t); if (m) { ecarts.push(msg); if (pos === null) pos = m.index; } };
    chercher(/\S[   ]{2,}\S/, 'espaces doublés');
    chercher(/ [,.](?!\.\.)/, 'espace avant une virgule ou un point');
    chercher(/,,|;;|::|(?<!\.)\.\.(?!\.)|,\.|\.,/, 'ponctuation doublée');
    if (ecarts.length) o.push(occ(b, ecarts.join(' ; ') + '.', { pos }));
  }
  return { occurrences: o, infos: [] };
}

function F06(doc, S, { p }) {
  const o = [];
  const mots = new RegExp(`(?<![\\p{L}])(${p.majuscules_accentuees.join('|')})(?![\\p{L}])`, 'gu');
  for (const b of blocsTypo(doc, S)) {
    const t = b.intitule;
    const trouves = [];
    let pos = null;
    for (const m of t.matchAll(mots)) { trouves.push(m[1]); pos ??= m.index; }
    for (const m of t.matchAll(/(?:^|[.!?…]\s+)(A)\s+(?=\p{Ll})/gu)) { trouves.push('A → À'); pos ??= m.index; }
    if (trouves.length) o.push(occ(b, `Majuscule sans accent : ${[...new Set(trouves)].join(', ')}.`, { pos: pos || 0 }));
  }
  return { occurrences: o, infos: [] };
}

function F07(doc, S) {
  const o = [];
  for (const b of blocsTypo(doc, S)) {
    if (S.dans('biblio', b.i)) continue;
    const t = masquerUrl(b.texte);
    const ecarts = [];
    let pos = null;
    const chercher = (re, msg) => { const m = re.exec(t); if (m) { ecarts.push(msg); if (pos === null) pos = m.index; } };
    chercher(/\d(%|€)/, 'pas d\'espace avant % ou €');
    chercher(/\d (%|€)/, 'espace insécable attendue avant % ou €');
    chercher(/(?<![\d.,/\-–\p{L}])\d{5,}(?![\d/\-–\p{L}])/u, 'milliers non séparés (1 000 000)');
    chercher(/(?<![\d.])\d{1,3}(?:\.\d{3})+(?![\d.])/, 'point comme séparateur de milliers');
    if (ecarts.length) o.push(occ(b, ecarts.join(' ; ') + '.', { pos }));
  }
  return { occurrences: o, infos: [] };
}

function F08(doc, S) {
  const o = [];
  for (const b of doc.blocs) {
    if (!estTexte(b) || b.tableau !== null || S.dans('biblio', b.i) || S.dans('plan-det', b.i) || S.dans('plan-synth', b.i)) continue;
    const t = b.texte;
    if (/^\s*\d+\s*[.)°]\s/.test(t)) continue;
    let pos = /^\s*\d/.test(t) ? 0 : null;
    if (pos === null) {
      for (const m of t.matchAll(/([.!?…])\s+(\d)/g)) {
        const mot = (t.slice(0, m.index).match(/(\S+)$/) || [''])[0].replace(/[^\p{L}]/gu, '');
        if (m[1] === '.' && mot.length <= 3) continue; // « p. 12 », « art. 4 », « n. 3 »
        pos = m.index + m[0].length - 1;
        break;
      }
    }
    if (pos !== null) o.push(occ(b, 'Phrase commençant par un chiffre.', { pos }));
  }
  return { occurrences: o, infos: [] };
}

export const CONTROLES = {
  A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, A11, A12, A13, A14, A15, A16, A17, A18, A19,
  B01, B02, B03, B04, B05, B06, B07, B08,
  F01, F02, F03, F04, F05, F06, F07, F08,
};
