// Contrôles déterministes, un par règle de regles.json.
// Chaque contrôle reçoit (doc, S, ctx) et rend { occurrences: [{ i, extrait, message, niveau? }], infos: [] }.
// `niveau: 'avertissement'` sur une occurrence l'abaisse quand la règle est bloquante.

import { normaliser, cleTitre, extrait, masquerUrl, nombre, pages } from './texte.js';
import { estTexte, caracteres, lirePlan, lignesDeZone } from './structure.js';

const AV = 'avertissement';
// `cibles` : paragraphes visés par une alerte regroupée.
// `corrections` : corrections possibles dans Word, [{ type, i, texte?, style? }] (voir sources/correction.js).
const occ = (b, message, opts = {}) => ({
  i: b ? b.i : null,
  extrait: b ? extrait(b.intitule, opts.pos || 0) : opts.extrait || '',
  message,
  niveau: opts.niveau,
  cibles: opts.cibles || (b ? [b.i] : []),
  corrections: opts.corrections || [],
  plage: opts.plage || null,
  passage: b && opts.pos !== undefined ? passages(b, opts.pos) : null,
});

/** Morceau du texte à sélectionner dans Word autour d'une position (sans espace en bord, sans caractère spécial). */
function passage(b, pos) {
  const decalage = (b.intitule.length - b.intitule.trimStart().length) + ((b.label || '').replace(/\t/g, ' ').trim() ? (b.label || '').length : 0);
  const debut = b.texte.indexOf(b.intitule.slice(pos, pos + 12).trim());
  const i = debut >= 0 ? debut : pos - decalage;
  if (i < 0 || i >= b.texte.length) return null;
  let t = b.texte.slice(i, i + 40);
  if (i + 40 < b.texte.length) t = t.replace(/\s+\S*$/, '');
  t = t.trim();
  return t.length >= 2 && !/[\^\t]/.test(t) ? t : null;
}
/** Passages à essayer pour la sélection : le plus court d'abord (le caractère fautif et ses voisins), puis plus large. */
function passages(b, pos) {
  const long = passage(b, pos);
  const decalage = (b.intitule.length - b.intitule.trimStart().length) + ((b.label || '').replace(/\t/g, ' ').trim() ? (b.label || '').length : 0);
  const j = b.texte.indexOf(b.intitule.slice(pos, pos + 12).trim());
  const i = j >= 0 ? j : pos - decalage;
  const a = Math.max(0, i - 3);
  const court = i >= 0 ? b.texte.slice(a, i + 5).trim() : '';
  // Rang de l'occurrence dans le paragraphe : le même morceau peut figurer plus haut.
  const rang = (t, debut) => (debut > 0 ? b.texte.slice(0, debut + t.length - 1).split(t).length - 1 : 0);
  const liste = [];
  if (court.length >= 2 && !/[\^\t]/.test(court)) liste.push({ texte: court, rang: rang(court, b.texte.indexOf(court, a)) });
  if (long && long !== court) liste.push({ texte: long, rang: 0 });
  return liste.length ? liste : null;
}
const corr = (type, blocs, extra = {}) => blocs.map((b) => ({ type, i: b.i, ...extra }));

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
  for (const t of S.titres) {
    if (/annexes? (dites? )?[«\s]*(outils|informatives?)/.test(normaliser(t.intitule))) {
      o.push(occ(t, 'Distinction entre annexes « outils » et annexes « informatives » à proscrire (Note, p. 8) : une seule suite d\'annexes numérotées.'));
    }
  }
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
      lignes.push(...lignesDeZone(S, z));
    }
    return lignes;
  }
  return S.titres.filter((t) => t.role === 'plan' && (!S.corps || (t.i >= S.corps.debut && t.i < S.corps.fin)));
}

function A06(doc, S, { profil }) {
  const o = [];
  // Notice : la nomenclature vise le plan proposé (plans synthétique et détaillé), pas les titres de la notice elle-même.
  const cibles = new Set(profil === 'notice' ? [] : S.titres.filter((t) => t.role === 'plan'));
  // Notice : sous-titres numérotés « 1.1 » sous les sections ; la Note proscrit ce découpage, on retire le numéro.
  if (profil === 'notice') {
    const sous = S.titres.filter((t) => t.role !== 'zone' && t.role !== 'categorie' && !S.dans('plan-det', t.i) && !S.dans('plan-synth', t.i)
      && /^\s*\d+(\.\d+)+\.?\s/.test(t.intitule));
    if (sous.length) {
      o.push(occ(sous[0], `${sous.length} sous-titre(s) de la notice numéroté(s) « 1.1 » : découpage proscrit (« I, 1.1, 1.1.1 ou A, a)… », p. 14) ; garder des sous-titres sans numéro.`,
        { cibles: sous.map((t) => t.i), corrections: corr('sans-numero', sous.filter((t) => /^\s*\d+(\.\d+)+/.test(t.texte))) }));
    }
  }
  for (const nom of ['plan-synth', 'plan-det']) {
    const z = S.zone(nom);
    if (z) for (const b of lignesDeZone(S, z)) cibles.add(b);
  }
  let sansNumero = 0;
  const decimaux = [];
  for (const b of [...cibles].sort((x, y) => x.i - y.i)) {
    const lab = b.intitule.replace(/^[•▪◦\-–]\s*/, '');
    if (INTERDIT.test(lab)) {
      o.push(occ(b, 'Numérotation proscrite (« I, 1.1, 1.1.1 ou A, a)… », p. 14) : utiliser Partie / Chapitre / Section / Paragraphe.'));
      continue;
    }
    const pl = lirePlan(lab);
    if (pl && pl.decimal) decimaux.push({ b, pl, tape: /(partie|chapitre|section|paragraphe|§)\s*\d+(\.\d+)+/i.test(b.texte) });
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
    const tapes = decimaux.filter((d) => d.tape).map((d) => d.b);
    const automatiques = decimaux.length - tapes.length;
    o.push(occ(b, `${decimaux.length} intitulé(s) à numéro décimal (« ${nomMot(pl.mot)} ${pl.brut} ») : numéro simple attendu (« ${nomMot(pl.mot)} ${pl.numero} ») ; « 1.1, 1.1.1 » proscrits (p. 14).`
      + (automatiques ? ` ${automatiques} numéro(s) automatique(s) : corriger la liste à plusieurs niveaux.` : ''),
    { cibles: decimaux.map((d) => d.b.i), corrections: corr('numero-simple', tapes) }));
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
  const usages = {};
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
    else if (attendu !== continu) {
      // Les deux usages sont admis, mais un seul par niveau dans tout le plan.
      const usage = pl.numero === attendu ? 'reprise' : 'continue';
      if (!usages[pl.niveau]) usages[pl.niveau] = usage;
      else if (usages[pl.niveau] !== usage) ruptures.push({ noeud, melange: usages[pl.niveau] });
    }
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
    else if (r.melange) {
      o.push(occ(r.noeud.b, `« ${nomMot(r.noeud.mot)} ${r.noeud.numero} » : numérotation ${r.melange === 'reprise' ? 'reprise à 1 sous chaque niveau supérieur' : 'continue'} ailleurs dans le plan ; un seul usage pour tout le plan.`));
    }
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

/** Premier titre (hors sections reconnues) dont l'intitulé approche un des motifs, du plus précis au plus large. */
function titreProche(candidats, proches, pris) {
  for (const motif of proches || []) {
    const re = new RegExp(motif);
    const t = candidats.find((c) => !pris.has(c.i) && re.test(cleTitre(c.intitule)));
    if (t) return t;
  }
  return null;
}

function A09(doc, S, { p }) {
  const o = [];
  const s = sectionsNotice(S);
  let dernier = -1;
  // Titres de la notice qui pourraient tenir lieu des sections manquantes (niveau le plus haut d'abord).
  const dansNotice = S.titres.filter((t) => t.role !== 'zone' && t.role !== 'section-notice'
    && (!S.notice || (t.i >= S.notice.debut && t.i < S.notice.fin)))
    .sort((a, b) => a.niveau - b.niveau || a.i - b.i);
  const pris = new Set();
  p.notice_sections.forEach((def, k) => {
    const t = s[k + 1];
    if (!t) {
      const proche = titreProche(dansNotice, def.proches, pris);
      if (proche) {
        pris.add(proche.i);
        const ref = s.find(Boolean);
        const texte = proche.label ? def.titre : `${k + 1}. ${def.titre}`;
        o.push(occ(proche, `Section ${k + 1} « ${def.titre} » absente ; titre approchant : à renommer « ${k + 1}. ${def.titre} » si son contenu y répond.`,
          { corrections: [{ type: 'renommer', i: proche.i, texte, style: ref?.styleBuiltIn || null }] }));
      } else o.push(occ(null, `Section ${k + 1} « ${def.titre} » absente : à rédiger (ordre imposé, ${k ? `après la section ${k}` : 'en tête de la notice'}).`));
      return;
    }
    if (t.i < dernier) o.push(occ(t, `Section ${k + 1} placée avant la section ${k} : ordre imposé.`));
    dernier = Math.max(dernier, t.i);
    const contenu = S.blocsDe(t.i + 1, t.fin).filter((b) => (estTexte(b) || b.tableau !== null) && b.texte.trim());
    if (!contenu.length) o.push(occ(t, `Section ${k + 1} vide.`));
  });
  return { occurrences: o, infos: [] };
}

/**
 * Sections de la notice, complétées par les titres approchants pour celles qui manquent (contrôles de contenu :
 * un tableau, une question). `approche` est vrai pour une section retrouvée par un titre approchant.
 */
function sectionsOuProches(S, p) {
  const s = sectionsNotice(S);
  const dansNotice = S.titres.filter((t) => t.role !== 'zone' && t.role !== 'section-notice'
    && (!S.notice || (t.i >= S.notice.debut && t.i < S.notice.fin)))
    .sort((a, b) => a.niveau - b.niveau || a.i - b.i);
  const pris = new Set();
  const r = [];
  p.notice_sections.forEach((def, k) => {
    if (s[k + 1]) { r[k + 1] = s[k + 1]; return; }
    const t = titreProche(dansNotice, def.proches, pris);
    if (t) { pris.add(t.i); r[k + 1] = Object.assign(Object.create(t), { approche: true }); }
  });
  return r;
}

/** Ligne de tableau tapée en texte : au moins deux cellules séparées par « | ». */
export const cellulesTapees = (t) => {
  const c = t.split('|').map((x) => x.trim());
  return c.length >= 2 && c.filter(Boolean).length >= 2 ? c : null;
};

function A10(doc, S, { p }) {
  const o = [];
  const s = sectionsOuProches(S, p);
  const infos = [];
  for (const [n, objet] of [[3, 'tableau de différenciation avec les mémoires les plus proches'], [5, 'tableau des outils (numéro d\'ordre dans le plan, nom, nature, objectifs)']]) {
    const t = s[n];
    if (!t) { infos.push(`Section ${n} non trouvée : présence du ${objet} non contrôlée.`); continue; }
    if (t.approche) infos.push(`Section ${n} contrôlée sous le titre approchant « ${t.intitule.slice(0, 60)} ».`);
    const blocs = S.blocsDe(t.i + 1, t.fin);
    if (blocs.some((b) => b.tableau !== null)) continue;
    // Tableau tapé en texte, cellules séparées par « | » (collage depuis un éditeur de texte) : à convertir.
    const tape = blocs.filter((b) => b.tableau === null && cellulesTapees(b.texte));
    if (tape.length >= 2) {
      o.push(occ(tape[0], `Section ${n} : ${objet} tapé en texte (cellules séparées par « | »), pas en tableau Word.`, { cibles: tape.map((b) => b.i) }));
    } else o.push(occ(t, `Section ${n} sans tableau : ${objet} attendu.`));
  }
  return { occurrences: o, infos };
}

function A11(doc, S, { p }) {
  const t = sectionsOuProches(S, p)[1];
  if (!t) return { occurrences: [], infos: ['Section 1 non trouvée : problématique (question unique) non contrôlée.'] };
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
const TYPE_LIGNE = /^(introduction|conclusion|partie|chapitre|section|§|paragraphe)(?![a-z])/;
// « Introduction de la partie (0,5 page) », « Conclusion partielle du chapitre » : sous-ligne du niveau au-dessus
// (Note, p. 15), comptée dans la somme de ses pages ; ni l'introduction ni la conclusion du mémoire.
const PARTIEL = /^(introduction|conclusion)(?: partielle| generale)? (?:de la|du|de) (partie|chapitre)\b/;
function lirePlanOuPartiel(intitule) {
  const n = normaliser(intitule).replace(/^[•▪◦-]\s*/, '');
  const m = PARTIEL.exec(n);
  if (m) return { mot: m[1], niveau: m[2] === 'partie' ? 2 : 3, numero: null, decimal: false, brut: '', partiel: true };
  return lirePlan(intitule);
}

function A12(doc, S, { p }) {
  const z = S.zone('plan-det');
  if (!z) return { occurrences: [occ(null, 'Plan détaillé absent (titre « Plan détaillé » non trouvé).')], infos: [] };
  const o = [];
  const lignes = lignesDeZone(S, z);
  let intro = null; let conclusion = null; let total = 0;
  const valeurs = [];
  const sansPages = [];
  for (const b of lignes) {
    const n = normaliser(b.intitule).replace(/^[•▪◦-]\s*/, '');
    const m = TYPE_LIGNE.exec(n);
    if (!m) continue;
    const pg = lirePages(b.intitule);
    const pl = lirePlanOuPartiel(b.intitule);
    const partiel = !!(pl && pl.partiel);
    if (m[1] === 'introduction' && !partiel) intro = b;
    if (m[1] === 'conclusion' && !partiel) conclusion = b;
    if (pg === null) sansPages.push({ b, type: m[1] });
    valeurs.push({ b, pl, pg, type: m[1] });
    if (pg !== null && !partiel && (m[1] === 'introduction' || m[1] === 'partie' || (m[1] === 'conclusion' && b === conclusion))) total += pg;
  }
  // Lignes sans pages : une alerte par ligne, ou une seule par niveau au-delà de trois.
  const parType = new Map();
  for (const x of sansPages) parType.set(x.type, [...(parType.get(x.type) || []), x.b]);
  for (const [type, blocs] of parType) {
    if (blocs.length <= 3) blocs.forEach((b) => o.push(occ(b, 'Nombre prévisionnel de pages absent sur cette ligne.')));
    else o.push(occ(blocs[0], `${blocs.length} lignes « ${nomMot(type)} » sans nombre prévisionnel de pages (exigé à chaque niveau).`, { cibles: blocs.map((b) => b.i) }));
  }
  if (valeurs.some((v) => v.pl && v.pl.niveau === 3) && !valeurs.some((v) => v.pl && v.pl.niveau === 4)) {
    o.push(occ(z.titre, 'Aucun paragraphe dans le plan détaillé : découpage « Partie/chapitre/section/paragraphe », pages prévisionnelles et outils au niveau des paragraphes (p. 12).'));
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
  for (const b of lignesDeZone(S, z)) {
    const pl = lirePlanOuPartiel(b.intitule);
    if (!pl) continue;
    if (pl.niveau > 2 || (pl.partiel && pl.niveau === 3)) o.push(occ(b, 'Le plan synthétique ne fait apparaître que les parties et chapitres.'));
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
  } else {
    // Page de garde du mémoire (p. 14).
    if (!/diplome d'expertise comptable/.test(texte)) o.push(occ(garde[0], 'Mention « Diplôme d\'expertise comptable » absente.'));
    if (!/session d.(printemps|automne)/.test(texte) || !/20\d\d/.test(texte)) o.push(occ(garde[0], 'Session absente : « Session de printemps (mai) » ou « Session d\'automne (novembre) » suivie de l\'année.'));
    if (!/\bmemoire\b/.test(texte)) o.push(occ(garde[0], 'Mention « MÉMOIRE » absente.'));
    if (!/\b\d{5}\b|\b\d{2}\s\d{3}\b/.test(texte)) o.push(occ(garde[0], 'Code postal non trouvé (NOM – Prénom – adresse : au minimum code postal et ville).', { niveau: AV }));
    const titre = garde.find((b) => {
      const n = normaliser(b.texte);
      const lettres = b.texte.match(/\p{L}/gu) || [];
      return lettres.length >= 15 && (b.texte.match(/\p{Lu}/gu) || []).length / lettres.length > 0.9
        && !/diplome d'expertise|session d|^memoire$/.test(n);
    });
    if (!titre) o.push(occ(garde[0], 'Titre en majuscules non repéré : « TITRE en majuscules pour tous les mots du titre ».', { niveau: AV }));
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

/** Contenu de l'annexe n° 1 de la charte IA (modèle officiel) : tableau des usages, engagement, date, nom, signature. */
function verifierCharte(blocs, titre) {
  const o = [];
  const cellules = blocs.filter((b) => b.tableau !== null);
  const tableau = cellules.find((b) => /type d.usage/.test(normaliser(b.texte)));
  if (!tableau) {
    o.push(occ(titre, 'Tableau récapitulatif des usages absent (colonnes : type d\'usage A, B, C ou D ; bref descriptif ; pages concernées).'));
  } else {
    const lignes = new Map();
    for (const b of cellules.filter((x) => x.tableau === tableau.tableau && x.ligne > 0)) {
      if (!lignes.has(b.ligne)) lignes.set(b.ligne, []);
      lignes.get(b.ligne).push(b);
    }
    const remplies = [...lignes.values()].filter((l) => l.some((b) => b.texte.trim()));
    const exemple = cellules.find((b) => b.tableau === tableau.tableau && /^\s*ex\s*:/i.test(b.texte));
    if (exemple) o.push(occ(exemple, 'Exemple du modèle encore présent (« Ex : … ») : le remplacer par les usages réels.'));
    if (!remplies.length) o.push(occ(tableau, 'Aucun usage déclaré dans le tableau : recenser de manière exhaustive les parties ayant bénéficié de l\'IA.'));
    for (const l of remplies) {
      const type = l.find((b) => b.colonne === 0);
      if (type && !/type\s*[a-d]\b|^\s*[a-d]\s*$/i.test(type.texte)) o.push(occ(type, 'Type d\'usage attendu : A, B, C ou D (typologie de la charte).', { niveau: AV }));
      const pagesCol = l.find((b) => b.colonne === 2);
      if (pagesCol && !pagesCol.texte.trim()) o.push(occ(pagesCol, 'Pages concernées non renseignées.', { niveau: AV }));
    }
  }
  const texte = normaliser(blocs.map((b) => b.texte).join(' '));
  if (!/je certifie avoir pris connaissance/.test(texte)) o.push(occ(titre, 'Engagement du candidat absent (« Je certifie avoir pris connaissance… »).'));
  const ligneDate = blocs.find((b) => /date\s*:/.test(normaliser(b.texte)));
  if (ligneDate && /jj\/mm\/aaaa/i.test(ligneDate.texte)) o.push(occ(ligneDate, 'Date non renseignée (JJ/MM/AAAA).'));
  if (ligneDate && /nom\s*[-–]\s*prenom\s*:\s*(signature|$)/.test(normaliser(ligneDate.texte))) o.push(occ(ligneDate, 'Nom – prénom non renseignés.'));
  const signe = blocs.some((b) => b.image) || (ligneDate && /signature\s*:\s*\S/.test(normaliser(ligneDate.texte)));
  if (!signe) o.push(occ(ligneDate || titre, 'Signature non repérée (image de la signature) : annexe à signer (Note, p. 12).', { niveau: AV }));
  return o;
}

function A17(doc, S, { profil }) {
  if (profil === 'notice') {
    const z = S.zone('charte');
    if (!z) return { occurrences: [occ(null, 'Annexe n° 1 de la charte IA absente (« Usage de l\'Intelligence Artificielle pour la rédaction du rapport de stage ou du mémoire »), complétée, signée et insérée à la fin de la notice.')], infos: [] };
    return { occurrences: verifierCharte(S.blocsDe(z.titre.i, z.fin), z.titre), infos: [] };
  }
  const a1 = S.annexes[0];
  if (!a1) return { occurrences: [occ(null, 'Aucune annexe : l\'annexe 1 doit être l\'annexe n° 1 de la charte IA.')], infos: [] };
  const ia = /intelligence artificielle|\bia\b/.test(normaliser(a1.titre.intitule));
  if (!ia) return { occurrences: [occ(a1.titre, 'La première annexe doit être l\'annexe n° 1 « Usage de l\'Intelligence Artificielle pour la rédaction du rapport de stage ou du mémoire » (p. 15).')], infos: [] };
  return { occurrences: verifierCharte(S.blocsDe(a1.titre.i, a1.fin), a1.titre), infos: [] };
}

// Présentation des références par catégorie (Note, p. 10-11), exemples : « ROCHER Aurélien, (2025), Titre, Revue, n°, p. ».
const FORMES = {
  annee: [/\(\s*\d{4}\s*\)/, 'année entre parenthèses après l\'auteur, « (2025) »'],
  nom: [/^\s*[\p{Lu}][\p{Lu}'’\- ]+(?=[\s,])/u, 'NOM du premier auteur en capitales'],
  virgule: [null, 'pas de virgule entre NOM et Prénom (« ROCHER Aurélien, (2025), … »)'],
  url: [/(https?:\/\/|www\.)\S+/i, 'lien URL absent'],
  cliquable: [null, 'lien URL non cliquable'],
  consultation: [/consultation/i, '« (date de la dernière consultation : j/m/année) » absente'],
  pages: [/\d+\s*p\b\.?|\d+\s*pages/i, 'nombre de pages absent (« 180 p. »)'],
  session: [/\b(mai|novembre|printemps|automne)\b/i, 'session de soutenance absente (mai ou novembre)'],
  accueil: [null, 'site internet : page d\'accueil uniquement'],
};
const FORMES_CATEGORIE = {
  2: ['nom', 'virgule', 'annee'], 3: ['annee', 'url', 'cliquable', 'consultation'], 4: ['nom', 'virgule', 'annee', 'pages'],
  5: ['nom', 'virgule', 'annee', 'session', 'pages'], 6: ['url', 'cliquable', 'consultation', 'accueil'],
};

function formeReferences(references, n, o) {
  const ecarts = new Map();
  for (const b of references) {
    const t = b.texte;
    const url = (t.match(FORMES.url[0]) || [''])[0].replace(/[).,;]+$/, '');
    for (const f of FORMES_CATEGORIE[n] || []) {
      let ok;
      if (f === 'virgule') ok = !/^\s*[\p{Lu}][\p{Lu}'’\-]+,\s*\p{Lu}\p{Ll}/u.test(t);
      else if (f === 'cliquable') ok = !url || b.liens.some((l) => l.includes(url.replace(/^https?:\/\//, '').slice(0, 20)));
      else if (f === 'accueil') ok = !url || !/^(https?:\/\/)?[^/]+\/[^/\s]+/.test(url);
      else ok = FORMES[f][0].test(t);
      if (ok) continue;
      if (!ecarts.has(f)) ecarts.set(f, []);
      ecarts.get(f).push(b);
    }
  }
  for (const [f, blocs] of ecarts) {
    o.push(occ(blocs[0], `${blocs.length} référence(s) : ${FORMES[f][1]}.`, { niveau: AV, cibles: blocs.map((b) => b.i) }));
  }
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
      if (!courante.n) {
        const cle = cleTitre(b.intitule);
        const k = p.bibliographie_categories.findIndex((c) => (c.proches || []).some((m) => new RegExp(m).test(cle)));
        courante.proche = k >= 0 ? k + 1 : null;
        const nouveau = k >= 0 ? (b.label ? p.bibliographie_categories[k].titre : `${k + 1}. ${p.bibliographie_categories[k].titre}`) : null;
        o.push(occ(b, k >= 0
          ? `Intitulé hors des six catégories de la Note : à remplacer par « ${p.bibliographie_categories[k].titre} » (ou répartir ses références).`
          : 'Intitulé hors des six catégories de la Note : répartir ses références entre les six catégories.',
        { corrections: nouveau ? [{ type: 'renommer', i: b.i, texte: nouveau }] : [] }));
      }
      else {
        // Catégorie reconnue : intitulé et numéro exacts de la Note (« 5. Mémoire d'expertise comptable »).
        const def = p.bibliographie_categories[courante.n - 1];
        const exact = b.label ? def.titre : `${courante.n}. ${def.titre}`;
        const lu = normaliser((b.label || '') + b.texte).replace(/[\s.:]+$/, '');
        const num = /^(\d+)\s*[.)-]/.exec(lu);
        if (lu !== normaliser(exact) || (num && Number(num[1]) !== courante.n)) {
          o.push(occ(b, num && Number(num[1]) !== courante.n
            ? `Numéro « ${num[1]} » : la catégorie « ${def.titre} » est la n° ${courante.n} de la Note.`
            : `Intitulé à reprendre à l'identique de la Note : « ${courante.n}. ${def.titre} ».`,
          { niveau: AV, corrections: [{ type: 'renommer', i: b.i, texte: exact }] }));
        }
      }
      // Même catégorie sous deux intitulés : à fusionner.
      const n = courante.n || courante.proche;
      const double = n && cats.find((c) => (c.n || c.proche) === n);
      if (double) {
        courante.double = true;
        o.push(occ(b, `Catégorie en double avec ¶ ${double.b.i + 1} (« ${p.bibliographie_categories[n - 1].titre} ») : regrouper ses références sous un seul intitulé.`));
      }
      cats.push(courante);
    } else if (!b.titre) {
      if (courante) courante.entrees.push(b);
      else if (!horsCategorie) horsCategorie = b;
    }
  }
  if (horsCategorie) o.push(occ(horsCategorie, 'Référence placée avant toute catégorie.', { niveau: AV }));
  const ordonnees = cats.filter((c) => c.n && !c.double);
  const numeros = ordonnees.map((c) => c.n);
  for (let k = 1; k < numeros.length; k++) {
    if (numeros[k] < numeros[k - 1]) {
      const c = ordonnees[k];
      o.push(occ(c.b, `« ${p.bibliographie_categories[c.n - 1].titre} » placée après « ${p.bibliographie_categories[numeros[k - 1] - 1].titre} » : ordre des six catégories.`));
    }
  }
  let total = 0;
  for (const c of cats) {
    total += c.entrees.length;
    if (!c.entrees.length) { o.push(occ(c.b, 'Catégorie vide : la supprimer ou la compléter.', { niveau: AV })); continue; }
    const def = c.n || c.proche ? p.bibliographie_categories[(c.n || c.proche) - 1] : null;
    if (!def) continue;
    const references = c.entrees.filter((b) => !/^\s*\(/.test(b.texte));
    formeReferences(references, c.n || c.proche, o);
    if (!def.alpha) continue;
    for (let k = 1; k < references.length; k++) {
      const a = normaliser(references[k - 1].texte);
      const b = normaliser(references[k].texte);
      if (a.localeCompare(b, 'fr') > 0) {
        o.push(occ(references[k], 'Ordre alphabétique (A à Z) du premier auteur non respecté dans la catégorie : sélectionner les références, puis Accueil, Trier.',
          { niveau: AV, plage: { debut: references[0].i, fin: references[references.length - 1].i } }));
        break;
      }
    }
  }
  const exactes = new Set(cats.map((c) => c.n).filter(Boolean));
  const ordre = p.bibliographie_categories.map((c, k) => `${k + 1}. ${c.titre}${exactes.has(k + 1) ? ' ✓' : ''}`).join(' ; ');
  return { occurrences: o, infos: [`Bibliographie : ${total} référence(s) dans ${cats.length} catégorie(s) (la Note juge une trentaine suffisante). Catégories et ordre de la Note : ${ordre}.`] };
}

function A19(doc, S, { p }) {
  const mp = doc.miseEnPage;
  if (!mp) {
    return { occurrences: [], infos: ['Mise en page non lue dans ce Word : contrôler le .docx (Fichier, Enregistrer sous, Télécharger une copie).'] };
  }
  const o = [];
  const cm = (v) => v.toLocaleString('fr-FR', { maximumFractionDigits: 2 });
  // Marges : une alerte par écart distinct, sur la première section concernée.
  const marges = new Map();
  mp.sections.forEach((s, k) => {
    for (const [cote, v] of [['gauche', s.gauche], ['droite', s.droite]]) {
      if (v === null || v === undefined || Math.abs(v - p.marge_cm) <= 0.05) continue;
      const cle = `${cote}|${cm(v)}`;
      if (!marges.has(cle)) marges.set(cle, { cote, v, sections: [], b: doc.blocs[s.premierBloc] || null });
      marges.get(cle).sections.push(k + 1);
    }
  });
  for (const x of marges.values()) {
    o.push(occ(x.b, `Marge ${x.cote} de ${cm(x.v)} cm, ${cm(p.marge_cm)} cm attendus (section${x.sections.length > 1 ? 's' : ''} ${x.sections.join(', ')} du document).`));
  }
  // « Marges normales en haut et en bas » : préréglage « Normales » de Word, 2,5 cm.
  const hb = mp.sections.filter((s) => [s.haut, s.bas].some((v) => v !== null && v !== undefined && Math.abs(v - p.marge_haut_bas_cm) > 0.05));
  if (hb.length) {
    const sHb = hb[0];
    o.push(occ(doc.blocs[sHb.premierBloc] || null, `Marges haut et bas de ${cm(sHb.haut)} et ${cm(sHb.bas)} cm : « marges normales » attendues (préréglage Normales de Word : ${cm(p.marge_haut_bas_cm)} cm).`, { niveau: AV }));
  }
  // Page de garde non numérotée : numéros « depuis la première page après la page de garde ».
  const s0 = mp.sections[0];
  if (s0 && s0.pieds) {
    const defaut = s0.pieds.find((f) => f.type === 'default');
    const premiere = s0.pieds.find((f) => f.type === 'first');
    const numerotee = s0.premierePageDistincte ? !!premiere?.numero : !!defaut?.numero;
    if (numerotee) o.push(occ(doc.blocs[0] || null, 'La page de garde porte un numéro : Insertion, Pied de page, Modifier, cocher « Première page différente ».', { niveau: AV }));
  }
  if (mp.sections.length && mp.sections.every((s) => s.pieds !== null)) {
    const apresGarde = mp.sections.length > 1 ? mp.sections.slice(1) : mp.sections;
    if (!apresGarde.some((s) => s.pieds.some((f) => f.numero))) o.push(occ(null, 'Aucun numéro de page dans les pieds de page : Insertion, Numéro de page, Bas de page, à droite.'));
    else if (!apresGarde.some((s) => s.pieds.some((f) => f.numero && f.droite))) o.push(occ(null, 'Numéro de page non aligné à droite.'));
  }

  // Texte courant : taille et interligne affichés, regroupés par style.
  const exemptes = new Set(p.styles_exemptes);
  const groupes = new Map();
  let lus = 0;
  for (const b of mp.paragraphes === false ? [] : doc.blocs) {
    if (!estTexte(b) || b.tableau !== null || exemptes.has(b.style) || !b.effectif) continue;
    lus++;
    const tailles = b.effectif.tailles.filter((t) => t !== null && t !== p.taille_texte);
    const il = b.effectif.interligne;
    const ecartIl = il !== null && il !== undefined && (typeof il === 'string' || Math.abs(il - p.interligne) > 0.01);
    if (!tailles.length && !ecartIl) continue;
    const motif = [tailles.length ? `taille ${tailles.join(', ')}` : '', ecartIl ? `interligne ${typeof il === 'string' ? `fixe ${il}` : cm(il)}` : ''].filter(Boolean).join(', ');
    const cle = `${b.style || 'sans style'}|${motif}`;
    if (!groupes.has(cle)) groupes.set(cle, { style: b.style || 'sans style', motif, blocs: [] });
    groupes.get(cle).blocs.push(b);
  }
  for (const g of [...groupes.values()].sort((a, b) => b.blocs.length - a.blocs.length)) {
    o.push(occ(g.blocs[0], `${g.blocs.length} paragraphe(s) de style « ${g.style} » en ${g.motif} (attendu : taille ${p.taille_texte}, interligne ${cm(p.interligne)}).`, { cibles: g.blocs.map((b) => b.i) }));
  }
  const infos = [`Mise en page : ${mp.sections.length} section(s) lue(s), ${lus} paragraphe(s) de texte courant contrôlé(s) ; tableaux, légendes et notes non contrôlés.`];
  if (mp.sections.some((s) => s.gauche === null)) infos.push('Marges non lues dans ce Word : les contrôler sur le .docx.');
  if (mp.paragraphes === false) infos.push('Taille et interligne paragraphe par paragraphe : contrôler le .docx (lecture du document ouvert incomplète).');
  if (mp.normal.taille && mp.normal.taille !== p.taille_texte) {
    o.push(occ(null, `Style Normal en taille ${mp.normal.taille} : la passer à ${p.taille_texte} règle d'un coup les paragraphes sans mise en forme directe.`,
      { niveau: AV, corrections: [{ type: 'style-normal', i: -1 }] }));
  }
  return { occurrences: o, infos };
}

function A20(doc, S, { p }) {
  const n = doc.miseEnPage?.notes;
  const nb = (doc.notes || []).length;
  if (!n) return { occurrences: [], infos: [`${nb} note(s) de bas de page ; numérotation et taille contrôlées sur le .docx.`] };
  const o = [];
  if (nb && n.repriseParPage) o.push(occ(null, 'Notes de bas de page numérotées page par page : numérotation de 1 à x attendue (Références, Notes de bas de page, Numérotation : Continue).'));
  if (nb && n.taille && n.taille >= p.taille_texte) o.push(occ(null, `Notes de bas de page en taille ${n.taille} : caractères plus petits que ceux du texte attendus (style « Note de bas de page »).`));
  return { occurrences: o, infos: [`${nb} note(s) de bas de page.`] };
}

function A21(doc, S, { p, profil }) {
  const o = [];
  const infos = [];
  if (profil === 'notice') {
    const z = S.zone('changements');
    if (!z) {
      infos.push('Non applicable (aucun tableau des changements) : après une décision 4.2 ou 5.2 seulement, tableau des changements principaux (deux pages au plus, entre la page de garde et la notice) et modifications écrites en rouge (p. 13).');
    } else {
      const notice = S.zone('notice');
      if (notice && z.titre.i > notice.titre.i) o.push(occ(z.titre, 'Le tableau des changements se place entre la page de garde et la notice explicative.'));
      const blocs = S.blocsDe(z.debut, z.fin);
      if (!blocs.some((b) => b.tableau !== null)) o.push(occ(z.titre, 'Les changements se présentent sous la forme d\'un tableau.'));
      if (caracteres(blocs) > 2 * p.caracteres_par_page) o.push(occ(z.titre, 'Tableau des changements : deux pages au maximum.'));
      if (doc.source === 'docx' || doc.miseEnPage) {
        if (!doc.blocs.some((b) => b.rouge && !S.dans('changements', b.i))) o.push(occ(z.titre, 'Aucun passage en rouge repéré : toutes les modifications s\'écrivent en rouge.', { niveau: AV }));
      }
    }
  } else {
    const z = S.zone('liminaire');
    if (!z) {
      infos.push('Non applicable (aucune note liminaire) : pour une deuxième présentation ou plus seulement, note liminaire en tableau, en début de mémoire, explicitant les modifications (p. 15).');
    } else {
      const debut = Math.min(...['sommaire', 'intro'].map((nom) => S.zone(nom)?.titre.i ?? Infinity));
      if (z.titre.i > debut) o.push(occ(z.titre, 'La note liminaire se place en début de mémoire.'));
      if (!S.blocsDe(z.debut, z.fin).some((b) => b.tableau !== null)) o.push(occ(z.titre, 'La note liminaire se présente sous la forme d\'un tableau.'));
    }
  }
  return { occurrences: o, infos };
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
    if (!utiles.length) { o.push(occ(z.titre, `« ${NOMS_PIECES[nom] || nom} » vide : clic droit sur la table, « Mettre à jour la table ».`, { corrections: [{ type: 'tables', i: z.titre.i }] })); continue; }
    const lues = new Set(utiles.map((b) => cleTitre(b.intitule.replace(ENTREE_PAGE, ''))));
    for (const b of utiles) {
      const c = cleTitre(b.intitule.replace(ENTREE_PAGE, ''));
      if (c && !cles.has(c)) o.push(occ(b, 'Entrée sans titre correspondant : table à mettre à jour.', { corrections: [{ type: 'tables', i: b.i }] }));
    }
    for (const t of titresTexte) {
      // Titres des tables elles-mêmes (sommaire, table des matières) : hors des tables.
      if (t.niveau > profondeur || t.zone === nom || t.zone === 'sommaire' || t.zone === 'tdm' || (profil === 'memoire' && nom === 'sommaire' && t.role !== 'plan')) continue;
      if (t.role === 'categorie') continue;
      if (!lues.has(cleTitre(t.intitule))) o.push(occ(t, `Titre absent de « ${NOMS_PIECES[nom] || nom} » : table à mettre à jour.`, { corrections: [{ type: 'tables', i: t.i }] }));
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
const NON_SIGLES = new Set(['TODO', 'FIXME', 'XXX', 'SOURCE', 'TROUVER', 'TOTAL', 'NB']);

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
    // Plans de la notice : lignes de titres (comme le sommaire), le sigle se développe dans le texte des sections.
    if (!estTexte(b) || S.dans('biblio', b.i) || S.dans('abreviations', b.i) || S.dans('sommaire', b.i) || S.dans('tdm', b.i)
      || S.dans('plan-synth', b.i) || S.dans('plan-det', b.i)) continue;
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

/** Titres numérotés « 1. », « 1.1 » (notice, annexes…) : doublons et sauts. */
function B09(doc, S) {
  const o = [];
  const vus = new Map();
  const dernierSous = new Map();
  for (const t of S.titres) {
    if (t.zone || t.role === 'categorie') continue;
    const m = /^(\d+(?:\.\d+)*)(?=[.)\s-])/.exec(normaliser(t.intitule).replace(/^[•▪◦]\s*/, ''));
    if (!m) continue;
    const num = m[1];
    if (vus.has(num)) { o.push(occ(t, `Numéro « ${num} » déjà utilisé (¶ ${vus.get(num) + 1}).`)); continue; }
    vus.set(num, t.i);
    const parts = num.split('.').map(Number);
    const parent = parts.slice(0, -1).join('.');
    const attendu = (dernierSous.get(parent) || 0) + 1;
    if (parts[parts.length - 1] !== attendu) o.push(occ(t, `Numéro « ${num} » : « ${[...parts.slice(0, -1), attendu].join('.')} » attendu.`));
    dernierSous.set(parent, parts[parts.length - 1]);
    dernierSous.set(num, 0);
  }
  return { occurrences: o, infos: [] };
}

/** Tableaux et figures (p. 14) : titre, de préférence numéroté, et origine. */
function B10(doc, S) {
  const o = [];
  const legende = (b) => b && b.tableau === null && (/caption|l[ée]gende/.test(b.style || '')
    || /^(tableau|tab\.|figure|graphique|sch[ée]ma|illustration)\b/i.test(b.texte.trim()));
  const titreCourt = (b) => b && b.tableau === null && !b.titre && b.texte.trim().length <= 120 && !/[.:;,]\s*$/.test(b.texte.trim())
    && /^\p{Lu}/u.test(b.texte.trim());
  const voisin = (i, pas) => {
    for (let k = i + pas; k >= 0 && k < doc.blocs.length; k += pas) {
      const b = doc.blocs[k];
      if (b.texte.trim() || b.tableau !== null) return b;
    }
    return null;
  };
  const tableaux = new Map();
  for (const b of doc.blocs) {
    if (b.tableau === null || b.garde) continue;
    if (!tableaux.has(b.tableau)) tableaux.set(b.tableau, []);
    tableaux.get(b.tableau).push(b);
  }
  const sansTitre = [];
  const sansNumero = [];
  const sansSource = [];
  for (const cellules of tableaux.values()) {
    // Plan synthétique ou détaillé présenté en tableau « Intitulé | Pages » : la pièce porte déjà son titre.
    if (S.dans('plan-synth', cellules[0].i) || S.dans('plan-det', cellules[0].i)) continue;
    const avant = voisin(cellules[0].i, -1);
    const apres = voisin(cellules[cellules.length - 1].i, 1);
    const titre = [avant, apres].find((b) => legende(b)) || (titreCourt(avant) ? avant : null);
    if (!titre) sansTitre.push(cellules[0]);
    else if (!/\d/.test(titre.texte)) sansNumero.push(titre);
    const officiel = S.dans('charte', cellules[0].i);
    const source = [apres, apres && voisin(apres.i, 1), cellules[cellules.length - 1]].some((b) => b && /\bsources?\b|d'apr[eè]s|r[ée]alis[ée] par/i.test(b.texte));
    if (!source && !officiel) sansSource.push(cellules[0]);
  }
  if (sansTitre.length) o.push(occ(sansTitre[0], `${sansTitre.length} tableau(x) sans titre (« tout tableau avec un titre », p. 14) : Références, Insérer une légende.`, { cibles: sansTitre.map((b) => b.i) }));
  if (sansNumero.length) o.push(occ(sansNumero[0], `${sansNumero.length} titre(s) de tableau non numéroté(s) (numérotation recommandée : « Tableau 1 – … »).`, { niveau: AV, cibles: sansNumero.map((b) => b.i) }));
  if (sansSource.length) o.push(occ(sansSource[0], `${sansSource.length} tableau(x) sans mention d'origine (« Source : … » ou « Réalisé par l'auteur ») sous le tableau.`, { niveau: AV, cibles: sansSource.map((b) => b.i) }));
  // Figures : images hors tableaux (une image dans une cellule illustre le tableau, qui porte déjà un titre).
  const figures = doc.blocs.filter((b) => b.image && !b.garde && b.tableau === null && !S.dans('charte', b.i)
    && ![b, voisin(b.i, -1), voisin(b.i, 1)].some((x) => x && /^(figure|graphique|sch[ée]ma|illustration)\s*(n°\s*)?\d+/i.test(x.texte.trim())));
  if (figures.length) o.push(occ(figures[0], `${figures.length} figure(s) sans numéro (« Figure 1 – … ») ni origine (p. 14).`, { niveau: AV, cibles: figures.map((b) => b.i) }));
  return { occurrences: o, infos: [`${tableaux.size} tableau(x) et ${doc.blocs.filter((b) => b.image && !b.garde && b.tableau === null).length} figure(s) hors page de garde.`] };
}

// ---------------------------------------------------------------- F. Propreté

const RESTES = /\bTODO\b|\bFIXME\b|\bXXX\b|SOURCE À TROUVER|\?\?|\[(?:[^\]]{2,}|…|\.\.\.)\]/g;

// Anonymisation volontaire (secret professionnel) : pas un reste de travail.
const ANONYME = /^\[(?:anonymi[sz][eé]e?s?|confidentiel(?:le)?|nom masqué|x+|\*+)\]$/i;

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
      if (ANONYME.test(m[0])) continue;
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
      const cles = [...ecarts.keys()];
      const corrections = [];
      if (cles.some((k) => k.startsWith('espace ordinaire') || k.startsWith('espace insécable'))) corrections.push({ type: 'insecables', i: b.i });
      if (ecarts.has('guillemets droits " " : utiliser « »')) corrections.push({ type: 'guillemets', i: b.i });
      o.push(occ(b, cles.join(' ; ') + '.', { pos: premiere, corrections }));
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
    chercher(/,,|;;|::|(?<!\.)\.\.(?!\.)|,\.|(?<!(?:^|[\s(])(?:al|etc|cf|coll|vol|art|éd|ed|p|pp|n|ibid|cit|op|dir))\.,/, 'ponctuation doublée');
    if (ecarts.length) o.push(occ(b, ecarts.join(' ; ') + '.', { pos, corrections: ecarts.some((e) => e === 'espaces doublés') || / ,/.test(t) ? [{ type: 'espaces', i: b.i }] : [] }));
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
    if (trouves.length) o.push(occ(b, `Majuscule sans accent : ${[...new Set(trouves)].join(', ')}.`, { pos: pos || 0, corrections: trouves.some((x) => x !== 'A → À') ? [{ type: 'accents', i: b.i }] : [] }));
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
    if (ecarts.length) o.push(occ(b, ecarts.join(' ; ') + '.', { pos, corrections: /\d ?[%€]/.test(t) ? [{ type: 'unites', i: b.i }] : [] }));
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

function F09(doc, S) {
  const o = [];
  for (const b of blocsTypo(doc, S)) {
    const t = masquerUrl(b.texte);
    const m = [...t.matchAll(/(?<=\p{L})'(?=\p{L})/gu)];
    if (m.length) o.push(occ(b, `${m.length} apostrophe(s) droite(s) ' : utiliser ’.`, { pos: m[0].index, corrections: [{ type: 'apostrophes', i: b.i }] }));
  }
  return { occurrences: o, infos: [] };
}

export const CONTROLES = {
  A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, A11, A12, A13, A14, A15, A16, A17, A18, A19, A20, A21,
  B01, B02, B03, B04, B05, B06, B07, B08, B09, B10,
  F01, F02, F03, F04, F05, F06, F07, F08, F09,
};
