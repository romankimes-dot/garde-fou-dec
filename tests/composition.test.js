// Mise en forme : texte d'un document → nouveau document au gabarit (page de garde, styles, tableaux repris).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { docx, lancer, regles } from './outils.js';
import { lireDocx } from '../sources/docx.js';
import { planifier, composer } from '../sources/composition.js';

const gabarit = (nom) => readFile(new URL(`../gabarits/${nom}`, import.meta.url));

test('notice fictive → document au gabarit : garde remplie, sections, tableaux, relu par le moteur', async () => {
  const source = await docx('fixtures/notice-fictive.docx');
  const plan = planifier(source, await regles(), 'notice');
  const octets = await composer(await gabarit('gabarit-demande-agrement.docx'), plan, 'notice',
    { intitule: 'Sujet fictif', auteur: 'Prénom NOM', adresse: '75001 Paris', date: 'novembre 2026', tableaux: source.tableauxXml });
  const sortie = await lireDocx(octets);
  const garde = sortie.blocs.filter((b) => b.garde).map((b) => b.texte).join(' | ');
  for (const t of ['Sujet fictif', 'Prénom NOM', '75001 Paris', 'novembre 2026', "Demande d'agrément du sujet de mémoire"]) assert.ok(garde.includes(t), `garde : ${t}`);
  assert.ok(!sortie.blocs.some((b) => b.consigne), 'aucune consigne du gabarit');
  assert.equal(sortie.tableaux, source.blocs.filter((b) => b.tableau !== null && !b.garde).reduce((s, b) => s.add(b.tableau), new Set()).size);
  const styles = new Set(sortie.blocs.map((b) => b.style));
  for (const st of ['titre hors plan', 'titre notice', 'titre catégorie']) assert.ok(styles.has(st), `style ${st}`);
  const r = await lancer(sortie, { profil: 'notice' });
  assert.ok(!r.ids.has('A19'), r.lignes.filter((l) => l.startsWith('A19')).join('\n'));
});

test('mémoire en texte simple : « Partie 1 – … » devient un titre numéroté automatiquement', async () => {
  const lignes = ['Introduction', 'Texte de l\'introduction assez long pour être un paragraphe ordinaire du mémoire.',
    'Partie 1 – Le diagnostic', 'Chapitre 1 – Les constats', 'Section 1 – Les faits', '§ 1 – Un point', 'Texte.', 'Conclusion', 'Fin.'];
  const doc = { source: 'word', blocs: lignes.map((texte, i) => ({ i, texte, label: '', titre: false, niveau: 0, role: null, style: null, tableau: null, liens: [], image: false, notes: 0, consigne: false, sautAvant: false, garde: false })) };
  const plan = planifier(doc, await regles(), 'memoire');
  const st = Object.fromEntries(plan.lignes.filter((l) => l.style).map((l) => [l.texte, l.style]));
  assert.equal(st['Le diagnostic'], 'Heading1');
  assert.equal(st['Les constats'], 'Heading2');
  assert.equal(st['Les faits'], 'Heading3');
  assert.equal(st['Un point'], 'Heading4');
  const octets = await composer(await gabarit('gabarit-memoire.docx'), plan, 'memoire',
    { intitule: 'Sujet fictif', auteur: 'NOM Prénom', adresse: '75001 Paris', session: "Session d'automne (novembre) 2027" });
  const sortie = await lireDocx(octets);
  const garde = sortie.blocs.filter((b) => b.garde).map((b) => b.texte).join(' | ');
  assert.ok(garde.includes('SUJET FICTIF') && garde.includes("Diplôme d'expertise comptable – Session d'automne (novembre) 2027") && garde.includes('MÉMOIRE'));
  const titres = sortie.blocs.filter((b) => /^heading [1-4]$/.test(b.style)).map((b) => b.label.trim());
  assert.deepEqual(titres, ['Partie 1 –', 'Chapitre 1 –', 'Section 1 –', '§ 1 –']);
});
