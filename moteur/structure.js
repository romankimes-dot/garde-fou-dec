// Repérage de la structure : titres, pièces (zones), page de garde, plan, annexes.
// Commun aux deux sources (Word par Office.js et .docx) : quand le style n'est pas connu, le titre est reconnu à son texte.

import { normaliser, cleTitre } from './texte.js';

export const ZONES = [
  ['sommaire', /^sommaire$/],
  ['tdm', /^table des matieres$/],
  ['notice', /^notice explicative/],
  ['plan-synth', /^(proposition de )?plan synthetique/],
  ['plan-det', /^(proposition de )?plan detaille/],
  ['biblio', /^bibliographie/],
  ['charte', /^annexe (n° ?)?1 de la charte|^usage de l'?intelligence artificielle/],
  ['synthese', /^note de synthese/],
  ['intro', /^introduction( generale)?$/],
  ['conclusion', /^conclusion( generale)?$/],
  ['abreviations', /^(liste des |table des )?(abreviations|sigles)/],
  ['liste-annexes', /^liste des annexes/],
  ['figures', /^(table|liste) des (figures|tableaux|illustrations)/],
  ['changements', /^(tableau des )?(principaux )?changements|^modifications apportees/],
  ['liminaire', /^note liminaire/],
];

const PLAN = /^(partie|chapitre|section|§|paragraphe)\s*(\d+(?:\.\d+)*)\b/;
const MOT_NIVEAU = { partie: 1, chapitre: 2, section: 3, '§': 4, paragraphe: 4 };
// Pièces faites de lignes de plan ou d'entrées : leurs lignes ne sont pas des titres du document.
const LISTES = new Set(['plan-synth', 'plan-det', 'sommaire', 'tdm', 'liste-annexes', 'figures', 'abreviations']);
const ANNEXE = /^annexe\s*(n°\s*)?(\d+)\s*[-:.]/;

export function zoneDe(texte) {
  const cle = cleTitre(texte);
  const complet = normaliser(texte).replace(/[\s.:;,]+$/, '');
  for (const [nom, motif] of ZONES) if (motif.test(cle) || motif.test(complet)) return nom;
  return null;
}

/**
 * Niveau P/C/S/§ et numéro lus dans un intitulé (« Chapitre 2 – … » → { niveau: 2, numero: 2 }).
 * Numéro décimal (« Section 1.2 ») : le dernier nombre fait foi, `decimal` le signale.
 */
export function lirePlan(intitule) {
  const m = PLAN.exec(normaliser(intitule).replace(/^[•▪◦]\s*/, ''));
  if (!m) return null;
  const nombres = m[2].split('.');
  return { mot: m[1], niveau: MOT_NIVEAU[m[1]], numero: Number(nombres[nombres.length - 1]), decimal: nombres.length > 1, brut: m[2] };
}

function capitales(t) {
  const lettres = t.match(/\p{L}/gu) || [];
  return lettres.length >= 4 && (t.match(/\p{Lu}/gu) || []).length / lettres.length > 0.8;
}

function ressembleTitre(b) {
  const t = b.texte.trim();
  return t.length > 0 && t.length <= 160 && !/[.;,]$/.test(t) && b.tableau === null && !ENTREE_DE_TABLE.test(t);
}
// Entrée de sommaire ou de table : se termine par un numéro de page après une tabulation ou des points.
const ENTREE_DE_TABLE = /(\t|\.{3,}|…)\s*\d+\s*$/;

export function analyserStructure(doc, regles) {
  const motifsSections = regles.parametres.notice_sections.map((s) => new RegExp(s.motif));
  const motifsCategories = regles.parametres.bibliographie_categories.map((c) => new RegExp(c.motif));
  const consignes = (regles.consignes || []).map(normaliser).filter((c) => c.length > 20);
  let zoneCourante = null;

  for (const b of doc.blocs) {
    b.intitule = ((b.label || '') + b.texte).replace(/\t/g, ' ').trim();
    const cle = cleTitre(b.intitule);
    const n = normaliser(b.intitule);
    if (!b.consigne && consignes.length && n.length > 20 && consignes.some((c) => n.includes(c) || c.includes(n))) {
      b.consigne = true;
    }
    if (b.consigne || b.role === 'toc' || b.tableau !== null) continue;

    // Rôle d'un titre connu par son style ou par aria-level, sinon par son texte.
    const zone = zoneDe(b.intitule);
    const plan = lirePlan(b.intitule);
    const section = motifsSections.findIndex((m) => m.test(cle));
    const categorie = motifsCategories.findIndex((m) => m.test(cle));
    const annexe = ANNEXE.test(normaliser(b.intitule)) && !zone;
    // Dans un plan ou une table, une ligne sans style de titre n'est un titre que si elle ouvre une autre pièce
    // (y compris le plan détaillé qui suit le plan synthétique).
    const court = ressembleTitre(b) && (!LISTES.has(zoneCourante) || (zone && zone !== zoneCourante));

    if (b.titre || court) {
      if (zone) { b.titre = true; b.zone = zone; b.role = b.role === 'section-notice' ? b.role : 'zone'; b.niveau = 1; }
      else if (b.role === 'annexe' || (annexe && court)) { b.titre = true; b.role = 'annexe'; b.niveau = 1; }
      else if (section >= 0 && (b.titre || /^\d\s*[.)-]/.test(n))) {
        b.titre = true; b.role = 'section-notice'; b.niveau = 2; b.sectionNotice = section + 1;
      } else if (categorie >= 0 && (b.titre || b.role === 'categorie' || /^\d\s*[.)-]/.test(n) || cle.length < 45)) {
        b.titre = true; b.role = 'categorie'; b.niveau = 3; b.categorie = categorie + 1;
      } else if (b.role === 'categorie') { b.titre = true; b.niveau = 3; }
      else if (zoneCourante === 'biblio' && !b.titre && (/^\d\s*[.)-]/.test(n) || capitales(b.texte))) {
        // Intitulé de catégorie hors des six de la Note : repéré pour le signaler, pas pris pour une référence.
        b.titre = true; b.role = 'categorie'; b.niveau = 3; b.categorie = null;
      }
      else if (b.titre && (b.role === 'plan' || b.role === 'autre' || !b.role)) {
        b.role = 'plan';
        if (plan) b.niveau = b.niveau || plan.niveau;
      }
    }
    // Style « Titre notice » sur un intitulé qui n'est pas l'une des sept sections : titre ordinaire (A09 le signale).
    if (b.role === 'section-notice' && !b.sectionNotice) { b.role = 'plan'; b.niveau = 2; }
    if (b.titre && b.role === 'plan') b.plan = plan;
    if (b.titre && b.zone) zoneCourante = b.zone;
    else if (b.titre && (b.role === 'annexe' || (b.niveau === 1 && !LISTES.has(zoneCourante)))) zoneCourante = null;
  }

  const titres = doc.blocs.filter((b) => b.titre);

  // Page de garde : section 1 du .docx, sinon ce qui précède le premier titre.
  if (!doc.blocs.some((b) => b.garde)) {
    const premierTitre = titres[0] ? titres[0].i : 0;
    for (let i = 0; i < premierTitre; i++) doc.blocs[i].garde = true;
  }

  // Étendue de chaque titre : jusqu'au titre suivant de niveau inférieur ou égal.
  for (let k = 0; k < titres.length; k++) {
    const t = titres[k];
    let fin = doc.blocs.length;
    for (let j = k + 1; j < titres.length; j++) {
      if (titres[j].niveau <= t.niveau) { fin = titres[j].i; break; }
    }
    // Plans et tables : jusqu'à la pièce suivante, même si des lignes du plan portent un style de titre.
    if (LISTES.has(t.zone)) {
      const suivante = titres.slice(k + 1).find((x) => x.zone || x.role === 'annexe');
      fin = suivante ? suivante.i : doc.blocs.length;
    }
    t.fin = fin;
  }

  const zones = new Map();
  for (const t of titres) {
    if (!t.zone) continue;
    if (!zones.has(t.zone)) zones.set(t.zone, []);
    zones.get(t.zone).push({ titre: t, debut: t.i + 1, fin: t.fin });
  }
  const annexes = titres.filter((t) => t.role === 'annexe').map((t) => {
    const m = ANNEXE.exec(normaliser(t.intitule));
    return { titre: t, numero: m ? Number(m[2]) : null, debut: t.i + 1, fin: t.fin };
  });

  const zone = (nom) => (zones.get(nom) || [])[0] || null;
  const dans = (nom, i) => (zones.get(nom) || []).some((z) => i >= z.debut && i < z.fin);
  const blocsDe = (debut, fin) => doc.blocs.slice(debut, fin);

  // Corps du mémoire : du début de l'introduction à la fin de la conclusion.
  const intro = zone('intro');
  const conclusion = zone('conclusion');
  const corps = intro && conclusion && conclusion.fin > intro.debut
    ? { debut: intro.titre.i, fin: conclusion.fin } : null;

  // Notice explicative : sous le titre « Notice explicative », sinon de la section 1 à la fin de la section 7.
  let notice = zone('notice');
  if (!notice) {
    const s = titres.filter((t) => t.role === 'section-notice');
    if (s.length) notice = { titre: null, debut: s[0].i, fin: s[s.length - 1].fin };
  }

  return { titres, zones, zone, dans, annexes, corps, notice, blocsDe };
}

/** Bloc de texte courant : ni titre, ni consigne, ni table des matières, ni page de garde. */
export function estTexte(b) {
  return !b.titre && !b.consigne && b.role !== 'toc' && !b.garde && b.texte.trim().length > 0;
}

/** Caractères espaces compris d'un ensemble de blocs (consignes exclues). */
export function caracteres(blocs) {
  return blocs.reduce((n, b) => n + (b.consigne ? 0 : b.intitule.length), 0);
}

const PAGES_CELLULE = /^(\d+(?:[,.]\d+)?)\s*(?:p\.?|pages?)?$/i;
const LIGNE_DE_PLAN = /^(introduction|conclusion|partie|chapitre|section|§|paragraphe)(?![a-z])/;

/**
 * Lignes d'un plan de la notice (zone « plan-synth » ou « plan-det »), qu'il soit en paragraphes ou en tableau
 * « Intitulé | Pages » : chaque paragraphe de la première colonne est une ligne ; le nombre de la colonne des pages
 * est rattaché à la première ligne du plan de la rangée, sous la forme « (n pages) » lue par les contrôles.
 */
export function lignesDeZone(S, z) {
  const lignes = [];
  const rangees = new Map();
  for (const b of S.blocsDe(z.debut, z.fin)) {
    if (b.consigne || !b.intitule) continue;
    if (b.tableau === null) { lignes.push(b); continue; }
    const cle = `${b.tableau}:${b.ligne}`;
    if (!rangees.has(cle)) { rangees.set(cle, { premiere: [], pages: null }); lignes.push(rangees.get(cle)); }
    const r = rangees.get(cle);
    if (b.colonne === 0) r.premiere.push(b);
    else if (PAGES_CELLULE.test(b.intitule.trim())) r.pages = PAGES_CELLULE.exec(b.intitule.trim())[1];
  }
  return lignes.flatMap((x) => {
    if (!x.premiere) return [x];
    let pose = x.pages === null;
    return x.premiere.map((b) => {
      if (pose || !LIGNE_DE_PLAN.test(normaliser(b.intitule).replace(/^[•▪◦-]\s*/, ''))) return b;
      pose = true;
      return Object.assign(Object.create(b), { intitule: `${b.intitule} (${x.pages} pages)` });
    });
  });
}
