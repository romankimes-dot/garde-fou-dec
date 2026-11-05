#!/usr/bin/env node
// Vérification de la mise en forme au gabarit (onglet Gabarit du complément), sans Word.
//
//   node outils/verifier-mise-en-forme.mjs SOURCE.docx notice|memoire [RESULTAT.docx] [--pdf] [--sortie DOSSIER]
//
// SOURCE : le texte d'origine (document où le texte a été collé). RESULTAT : le document créé par le complément
// (Fichier, Enregistrer sous, Télécharger une copie). Sans RESULTAT, il est composé ici par le même code que le volet.
// Contrôles : texte conservé dans l'ordre, styles attendus, tableaux, page de garde, restes du gabarit, alertes de
// Garde-fou apparues avec la mise en forme ; --pdf : rendu LibreOffice et nombre de pages.
// Sortie : rapport Markdown ; code 1 si un défaut (texte perdu, style, tableau, page de garde, reste du gabarit) ;
// les alertes Garde-fou plus nombreuses après mise en forme sont listées « à examiner ».

import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { join, basename } from 'node:path';
import { lireDocx } from '../sources/docx.js';
import { planifier, composer } from '../sources/composition.js';
import { controler } from '../moteur/moteur.js';

const args = process.argv.slice(2);
const options = new Set(args.filter((a) => a.startsWith('--')));
const valeur = (nom) => { const k = args.indexOf(nom); return k >= 0 ? args[k + 1] : null; };
const positionnels = args.filter((a, k) => !a.startsWith('--') && args[k - 1] !== '--sortie');
const [source, profil, resultat] = positionnels;
if (!source || !['notice', 'memoire'].includes(profil)) {
  console.error('Usage : node outils/verifier-mise-en-forme.mjs SOURCE.docx notice|memoire [RESULTAT.docx] [--pdf] [--sortie DOSSIER]');
  process.exit(2);
}
const ici = (f) => new URL(`../${f}`, import.meta.url);
const regles = JSON.parse(await readFile(ici('regles.json'), 'utf8'));
regles.consignes = JSON.parse(await readFile(ici('consignes.json'), 'utf8'));
const sortie = valeur('--sortie') || '/tmp/verifier-mise-en-forme';
await mkdir(sortie, { recursive: true });

const doc = await lireDocx(await readFile(source));
const plan = planifier(doc, regles, profil);
let octets;
if (resultat) octets = await readFile(resultat);
else {
  const gabarit = await readFile(ici(`gabarits/${profil === 'memoire' ? 'gabarit-memoire.docx' : 'gabarit-demande-agrement.docx'}`));
  octets = await composer(gabarit, plan, profil, {
    intitule: 'Intitulé de contrôle', auteur: 'Camille MARTIN', adresse: '75001 Paris', date: 'novembre 2026',
    session: "Session d'automne (novembre) 2027", tableaux: doc.tableauxXml || [],
  });
  await writeFile(join(sortie, `compose-${profil}.docx`), octets);
}
const res = await lireDocx(octets);

const defauts = []; // bloquants
const remarques = [];
const norme = (t) => (t || '').replace(/[\s  ]+/g, ' ').trim().toLowerCase();

// 1. Texte conservé, dans l'ordre : chaque ligne du plan (hors omis, tables Word et tableaux) se retrouve.
const corps = res.blocs.filter((b) => !b.garde && b.tableau === null).map((b) => norme(b.texte));
let k = 0;
const perdus = [];
for (const l of plan.lignes) {
  if (l.omis || l.champ || l.tableau !== undefined || !l.texte.trim()) continue;
  const t = norme(l.texte);
  const j = corps.findIndex((c, n) => n >= k && (c === t || c.endsWith(t)));
  if (j < 0) perdus.push(l); else k = j + 1;
}
if (perdus.length) defauts.push(`${perdus.length} paragraphe(s) du texte introuvable(s) ou déplacé(s) : ${perdus.slice(0, 5).map((l) => `¶ ${l.i + 1}`).join(', ')}${perdus.length > 5 ? '…' : ''}.`);
const mots = (blocs) => blocs.filter((b) => !b.garde && !b.consigne).reduce((n, b) => n + (b.texte.match(/\p{L}+/gu) || []).length, 0);
const omis = new Set(plan.lignes.filter((l) => l.omis).map((l) => l.i));
const motsSource = mots(doc.blocs.filter((b) => !omis.has(b.i)));
const motsResultat = mots(res.blocs);
remarques.push(`Mots hors page de garde : ${motsSource} dans la source (hors tables refaites par Word), ${motsResultat} dans le résultat (tables et titres du gabarit compris).`);
if (motsResultat < motsSource * 0.98) defauts.push(`Résultat plus court que la source : ${motsSource - motsResultat} mot(s) en moins.`);

// 2. Styles : chaque titre prévu porte le style attendu.
const NOMS = { Titrehorsplan: 'titre hors plan', Titrenotice: 'titre notice', 'Titrecatégorie': 'titre catégorie', Titreannexe: 'titre annexe',
  Heading1: 'heading 1', Heading2: 'heading 2', Heading3: 'heading 3', Heading4: 'heading 4', Normal: 'normal' };
const parTexte = new Map(res.blocs.filter((b) => !b.garde).map((b) => [norme(b.texte), b]));
const malStyles = plan.lignes.filter((l) => l.style && l.style !== 'Normal').filter((l) => {
  const b = parTexte.get(norme(l.texte)) || res.blocs.find((x) => norme(x.texte).endsWith(norme(l.texte)));
  return b && b.style !== NOMS[l.style];
});
if (malStyles.length) defauts.push(`${malStyles.length} titre(s) sans le style attendu : ${malStyles.slice(0, 5).map((l) => `¶ ${l.i + 1} (${NOMS[l.style]})`).join(', ')}.`);
const comptes = {};
for (const b of res.blocs) if (b.style && !b.garde) comptes[b.style] = (comptes[b.style] || 0) + 1;
remarques.push(`Styles du résultat : ${Object.entries(comptes).map(([s, n]) => `${s} ${n}`).join(' · ')}.`);

// 3. Tableaux.
const tableaux = (d) => new Set(d.blocs.filter((b) => b.tableau !== null && !b.garde).map((b) => b.tableau)).size;
if (tableaux(res) !== tableaux(doc)) defauts.push(`Tableaux : ${tableaux(doc)} dans la source, ${tableaux(res)} dans le résultat.`);
else remarques.push(`Tableaux repris : ${tableaux(res)}.`);

// 4. Page de garde et restes du gabarit.
const garde = res.blocs.filter((b) => b.garde).map((b) => b.texte).join(' | ');
remarques.push(`Page de garde lue : ${garde.replace(/\s+/g, ' ').slice(0, 300)}`);
if (/\[|\]|XXX|À COMPLÉTER|Intitulé du sujet|Prénom NOM|NOM Prénom/.test(garde)) defauts.push('Page de garde : un champ du gabarit semble non rempli.');
const consignes = res.blocs.filter((b) => b.consigne);
if (consignes.length) defauts.push(`${consignes.length} consigne(s) du gabarit restée(s) dans le résultat : ¶ ${consignes.slice(0, 5).map((b) => b.i + 1).join(', ')}.`);
for (const a of plan.avertissements) remarques.push(`Avertissement du plan : ${a}`);

// 5. Garde-fou : alertes de la source et du résultat ; celles apparues avec la mise en forme sont des régressions.
const lancer = (d) => controler(structuredClone(d), regles, { profil, termes: [], ignores: new Set() });
const avant = lancer(doc);
const apres = lancer(res);
// Tables du gabarit vides tant que Word n'a pas mis à jour les champs : attendu, pas une régression.
const AVANT_MAJ = /Mettre à jour les champs|Mettre à jour les cha/;
const parRegle = (r) => Object.fromEntries(r.resultats
  .map((x) => [x.regle.id, x.occurrences.filter((o) => !AVANT_MAJ.test(o.message)).length]).filter(([, n]) => n));
const pa = parRegle(avant);
const pb = parRegle(apres);
const nouvelles = Object.entries(pb).filter(([id, n]) => n > (pa[id] || 0));
const disparues = Object.entries(pa).filter(([id, n]) => n > (pb[id] || 0));
// À examiner (pas forcément un défaut : un sous-titre devenu texte en gras peut être lu autrement par une règle).
const examiner = nouvelles.map(([id, n]) => `${id} ${pa[id] || 0}→${n}`);
const detail = apres.resultats.filter((x) => nouvelles.some(([id]) => id === x.regle.id))
  .flatMap((x) => x.occurrences.slice(0, 4).map((o) => `  - ${x.regle.id} ¶ ${o.i === null ? '-' : o.i + 1} : ${o.message}`));

// 6. Rendu PDF (LibreOffice), facultatif.
let pdf = null;
if (options.has('--pdf')) {
  const fichier = resultat || join(sortie, `compose-${profil}.docx`);
  try {
    execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', sortie, fichier], { stdio: 'ignore', timeout: 120000 });
    const chemin = join(sortie, basename(fichier).replace(/\.docx$/i, '.pdf'));
    const pages = /Pages:\s+(\d+)/.exec(execFileSync('pdfinfo', [chemin]).toString())?.[1];
    pdf = `${chemin} (${pages} pages ; estimation Garde-fou : ${apres.pagination.total})`;
  } catch (e) { pdf = `rendu impossible : ${e.message}`; }
}

const lignes = [
  `# Vérification de la mise en forme – ${profil}`,
  `Source : ${basename(source)} · Résultat : ${resultat ? basename(resultat) : `composé ici (${join(sortie, `compose-${profil}.docx`)})`}`,
  '',
  `## Verdict : ${defauts.length ? `${defauts.length} défaut(s) à corriger` : 'aucun défaut détecté'}`,
  ...defauts.map((d) => `- ${d}`),
  '',
  `## À examiner : ${examiner.length ? `alertes Garde-fou plus nombreuses après mise en forme (${examiner.join(', ')})` : 'rien'}`,
  ...detail,
  '',
  '## Mesures',
  ...remarques.map((r) => `- ${r}`),
  `- Garde-fou : source ${avant.comptes.bloquants} bloquant(s) / ${avant.comptes.avertissements} à vérifier ; résultat ${apres.comptes.bloquants} / ${apres.comptes.avertissements}.`,
  disparues.length ? `- Alertes réglées par la mise en forme : ${disparues.map(([id, n]) => `${id} ${n}→${pb[id] || 0}`).join(', ')}.` : null,
  pdf ? `- PDF : ${pdf}` : null,
].filter((x) => x !== null);
console.log(lignes.join('\n'));
process.exit(defauts.length ? 1 : 0);
