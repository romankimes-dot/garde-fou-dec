// Lanceur : applique les règles du profil et du mode, calcule le feu et la pagination estimée.

import { analyserStructure, caracteres } from './structure.js';
import { CONTROLES } from './controles.js';
import { normaliser } from './texte.js';

export const cleOccurrence = (regle, o) => `${regle}|${normaliser(o.extrait).slice(0, 80)}|${o.message.slice(0, 50)}`;

/**
 * @param doc      modèle lu (docx.js ou office.js)
 * @param regles   contenu de regles.json, avec `consignes` (consignes.json) le cas échéant
 * @param options  { profil: 'notice'|'memoire', mode: 'redaction'|'depot', termes: [], ignores: Set }
 */
export function controler(doc, regles, options) {
  const { profil, mode } = options;
  const ignores = options.ignores || new Set();
  const p = regles.parametres;
  const S = analyserStructure(doc, regles);
  const ctx = { p, profil, mode, termes: options.termes || [] };

  const resultats = [];
  const infos = [];
  let bloquants = 0;
  let avertissements = 0;
  let ignoresCompte = 0;

  for (const r of regles.regles) {
    if (!r.profils.includes(profil)) continue;
    const niveau = r.niveau[mode];
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
    for (const o of sortie.occurrences) {
      o.niveau = niveau === 'bloquant' && o.niveau === 'avertissement' ? 'avertissement' : niveau;
      o.cle = cleOccurrence(r.id, o);
      if (ignores.has(o.cle)) { ignoresCompte++; continue; }
      if (o.niveau === 'bloquant') bloquants++; else avertissements++;
      occurrences.push(o);
    }
    resultats.push({ regle: r, niveau, occurrences });
  }

  if (profil === 'memoire' && S.zone('notice') && !S.corps) infos.unshift({ regle: '-', texte: 'Ce document ressemble à une notice (titre « Notice explicative », ni introduction ni conclusion) : choisir « Notice » en haut du volet.' });
  if (profil === 'notice' && !S.notice && S.corps) infos.unshift({ regle: '-', texte: 'Ce document ressemble à un mémoire (introduction et conclusion, pas de notice explicative) : choisir « Mémoire » en haut du volet.' });

  return {
    profil,
    mode,
    source: doc.source,
    feu: bloquants ? 'rouge' : avertissements ? 'orange' : 'vert',
    comptes: { bloquants, avertissements, ignores: ignoresCompte },
    resultats,
    infos,
    pagination: pagination(doc, S, p.caracteres_par_page),
    diagnostic: doc.diagnostic,
    titres: S.titres.map((t) => ({ i: t.i, role: t.role, niveau: t.niveau, zone: t.zone || null, section: t.sectionNotice || null })),
  };
}

/** Pagination estimée : 2 400 caractères par page, nouvelle page à chaque titre de niveau 1. */
function pagination(doc, S, parPage) {
  const lignes = [];
  let page = 1;
  let rempli = 0;
  const debut = doc.blocs.findIndex((b) => !b.garde);
  if (debut < 0) return lignes;
  for (const b of doc.blocs.slice(debut)) {
    if ((b.sautAvant || (b.titre && b.niveau === 1)) && rempli > 0) { page++; rempli = 0; }
    if (b.titre && b.niveau <= 2 && !S.dans('plan-det', b.i) && !S.dans('plan-synth', b.i)) {
      const car = caracteres(doc.blocs.slice(b.i, b.fin));
      lignes.push({ i: b.i, intitule: b.intitule, niveau: b.niveau, page, pages: car / parPage });
    }
    if (b.consigne) continue;
    rempli += b.intitule.length;
    while (rempli > parPage) { page++; rempli -= parPage; }
  }
  return lignes;
}
