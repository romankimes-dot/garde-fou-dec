// Lanceur : applique les règles du profil, calcule le feu et la pagination estimée.

import { analyserStructure, caracteres } from './structure.js';
import { CONTROLES } from './controles.js';
import { normaliser } from './texte.js';

export const cleOccurrence = (regle, o) => `${regle}|${normaliser(o.extrait).slice(0, 80)}|${o.message.slice(0, 50)}`;

/**
 * @param doc      modèle lu (docx.js ou office.js)
 * @param regles   contenu de regles.json, avec `consignes` (consignes.json) le cas échéant
 * @param options  { profil: 'notice'|'memoire', termes: [], ignores: Set }
 */
export function controler(doc, regles, options) {
  const { profil } = options;
  const ignores = options.ignores || new Set();
  const p = regles.parametres;
  const S = analyserStructure(doc, regles);
  const ctx = { p, profil, termes: options.termes || [] };

  const resultats = [];
  const infos = [];
  let bloquants = 0;
  let avertissements = 0;
  let ignoresCompte = 0;
  let restes = 0; // restes de travail (catégorie R) : comptés à part, à nettoyer avant l'export

  for (const r of regles.regles) {
    if (!r.profils.includes(profil)) continue;
    const niveau = r.niveau;
    if (niveau === 'off') continue;
    const f = CONTROLES[r.id];
    if (!f) continue;
    if (r.source === 'docx' && !doc.miseEnPage) {
      resultats.push({ regle: r, niveau, occurrences: [], horsSource: true });
      continue;
    }
    let sortie;
    try {
      sortie = f(doc, S, ctx);
    } catch (e) {
      sortie = { occurrences: [{ i: null, extrait: '', message: `Contrôle en erreur : ${e.message}`, niveau: 'avertissement' }], infos: [] };
    }
    infos.push(...(sortie.infos || []).map((t) => ({ regle: r.id, texte: t })));
    const occurrences = [];
    const ignorees = [];
    for (const o of sortie.occurrences) {
      o.niveau = niveau === 'bloquant' && o.niveau === 'avertissement' ? 'avertissement' : niveau;
      o.cle = cleOccurrence(r.id, o);
      if (ignores.has(o.cle)) { ignoresCompte++; ignorees.push(o); continue; }
      if (r.categorie === 'R') restes++;
      else if (o.niveau === 'bloquant') bloquants++;
      else avertissements++;
      occurrences.push(o);
    }
    resultats.push({ regle: r, niveau, occurrences, ignorees });
  }

  if (profil === 'memoire' && S.zone('notice') && !S.corps) infos.unshift({ regle: '-', texte: 'Ce document ressemble à une notice (titre « Notice explicative », ni introduction ni conclusion) : choisir « Demande d’agrément » en haut du volet.' });
  if (profil === 'notice' && !S.notice && S.corps) infos.unshift({ regle: '-', texte: 'Ce document ressemble à un mémoire (introduction et conclusion, pas de notice explicative) : choisir « Mémoire » en haut du volet.' });

  return {
    profil,
    source: doc.source,
    feu: bloquants ? 'rouge' : avertissements || restes ? 'orange' : 'vert',
    comptes: { bloquants, avertissements, restes, ignores: ignoresCompte },
    resultats,
    infos,
    pagination: pagination(doc, S, p.caracteres_par_page),
    diagnostic: doc.diagnostic,
    titres: S.titres.map((t) => ({ i: t.i, role: t.role, niveau: t.niveau, zone: t.zone || null, section: t.sectionNotice || null })),
  };
}

/**
 * Pagination estimée, de deux façons :
 * - « pages » : mesure de la Note, 2 400 caractères par page (sert aux limites de volume) ;
 * - « page » : position à l'affichage, en lignes (90 caractères par ligne, 32 lignes par page en taille 12 et
 *   interligne 1,5 avec marges de 2,5 cm ; tableaux ligne par ligne ; titres et paragraphes vides comptés ;
 *   saut de page à chaque titre de niveau 1 et à chaque saut manuel). Page de garde comprise.
 */
const CAR_LIGNE = 90;
const LIGNES_PAGE = 32;

/** Page estimée de chaque titre (page physique, page de garde = 1) : sert à pré-remplir les tables du gabarit. */
export function pagesDesTitres(doc, regles) {
  const copie = structuredClone(doc);
  const S = analyserStructure(copie, regles);
  return pagination(copie, S, regles.parametres.caracteres_par_page).parBloc;
}

function pagination(doc, S, parPage) {
  const lignes = [];
  const parBloc = new Map();
  let page = 1;
  let rempli = 0; // lignes déjà occupées sur la page courante
  let vide = true; // page commencée sans texte (paragraphes vides seulement) : pas de nouveau saut
  const nouvellePage = () => { if (!vide) { page++; rempli = 0; vide = true; } };
  const ajouter = (n) => {
    rempli += n;
    while (rempli > LIGNES_PAGE) { page++; rempli -= LIGNES_PAGE; }
  };
  const premiereApresGarde = doc.blocs.findIndex((b) => !b.garde);
  // Tableaux : une rangée occupe autant de lignes que sa cellule la plus longue (colonnes côte à côte).
  const colonnes = new Map();
  for (const b of doc.blocs) if (b.tableau !== null) colonnes.set(b.tableau, Math.max(colonnes.get(b.tableau) || 1, (b.colonne ?? 0) + 1));
  let rangee = null;
  for (const b of doc.blocs) {
    if (b.i === premiereApresGarde) nouvellePage();
    if ((b.sautAvant || (b.titre && b.niveau === 1)) && !b.garde) nouvellePage();
    if (b.titre) parBloc.set(b.i, page);
    if (b.titre && b.niveau <= 2 && !b.garde && !S.dans('plan-det', b.i) && !S.dans('plan-synth', b.i)) {
      const car = caracteres(doc.blocs.slice(b.i, b.fin));
      lignes.push({ i: b.i, intitule: b.intitule, niveau: b.niveau, page, pages: car / parPage });
    }
    if (b.tableau !== null) {
      const largeur = Math.max(10, Math.floor(CAR_LIGNE / colonnes.get(b.tableau)));
      const cle = `${b.tableau}|${b.ligne}`;
      const n = Math.max(1, Math.ceil(b.intitule.length / largeur));
      vide = false;
      if (!rangee || rangee.cle !== cle) { if (rangee) ajouter(rangee.n); rangee = { cle, n: 0, cellules: new Map() }; }
      rangee.cellules.set(b.colonne, (rangee.cellules.get(b.colonne) || 0) + n);
      rangee.n = Math.max(...rangee.cellules.values());
      continue;
    }
    if (rangee) { ajouter(rangee.n + 0.5); rangee = null; }
    const n = Math.max(1, Math.ceil(b.intitule.length / CAR_LIGNE));
    if (b.texte.trim() || b.garde) vide = false;
    ajouter(n + (b.titre ? 1 : b.texte.trim() ? 0.3 : 0));
  }
  if (rangee) ajouter(rangee.n);
  lignes.total = page;
  lignes.parBloc = parBloc;
  return lignes;
}
