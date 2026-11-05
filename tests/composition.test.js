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
    { intitule: 'Sujet fictif', auteur: 'Prénom NOM', adresse: '75001 Paris', date: 'novembre 2026', tableaux: source.tableauxXml, regles: await regles() });
  const sortie = await lireDocx(octets);
  const garde = sortie.blocs.filter((b) => b.garde).map((b) => b.texte).join(' | ');
  for (const t of ['Sujet fictif', 'Prénom NOM', '75001 Paris', 'novembre 2026', "Demande d'agrément du sujet de mémoire"]) assert.ok(garde.includes(t), `garde : ${t}`);
  // Aucune consigne du gabarit ; seul le repère du tableau des outils (section 5 de la fixture sans tableau).
  assert.deepEqual(sortie.blocs.filter((b) => b.consigne).map((b) => b.texte.slice(0, 30)), ['[Tableau 2 à insérer ici : tab']);
  // Tableaux repris, plus un tableau par plan (synthétique, détaillé) et par tableau tapé en texte.
  const groupesPlan = plan.lignes.filter((l, k, t) => !l.omis && l.planLigne && !(t.slice(0, k).filter((x) => !x.omis).pop() || {}).planLigne).length;
  const tapes = plan.lignes.filter((l) => l.tableauTexte).length;
  assert.equal(sortie.tableaux, source.blocs.filter((b) => b.tableau !== null && !b.garde).reduce((s, b) => s.add(b.tableau), new Set()).size + groupesPlan + tapes);
  const styles = new Set(sortie.blocs.map((b) => b.style));
  for (const st of ['titre hors plan', 'titre notice', 'titre catégorie']) assert.ok(styles.has(st), `style ${st}`);
  // Table des matières pré-remplie (sections et sous-titres), sans sa propre entrée ; sous-titres stylés, sans « 1.1 ».
  const toc = sortie.blocs.filter((b) => /^toc \d$/.test(b.style)).map((b) => b.texte);
  assert.ok(toc.length >= 7, `entrées : ${toc.length}`);
  assert.ok(!toc.some((t) => /^table des mati/i.test(t)), 'la table ne se liste pas');
  assert.ok(sortie.blocs.some((b) => b.style === 'titre hors table' && /^Table des mati/.test(b.texte)));
  assert.ok(!sortie.blocs.some((b) => b.style === 'sous-titre notice' && /^\d+\.\d/.test(b.texte)));
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

test('notice collée en texte brut : tableaux « | », plans en tableau, légendes, sources, listes, bibliographie', async () => {
  const lignes = ["Demande d'agrément du sujet de mémoire présentée en vue de l'obtention du diplôme d'expertise comptable", 'Sujet', 'Prénom NOM',
    'Table des matières', 'Notice explicative', '1. Intérêt du sujet et problématique', 'Problématique : Comment faire ?',
    '- premier point de la liste', '- second point de la liste',
    "3. Différenciation avec les autres mémoires d'expertise comptable", 'Tableau 1 – Mémoires proches', 'Mémoire | Ce qu\'il traite | Différence',
    'Auteur 2010 | Objet | Écart', 'Auteur 2015 | Objet | Écart', 'Source : réalisé par l\'auteur.',
    '5. Apports du mémoire à la profession', 'Tableau 2 – Outils', "N° d'ordre | Outil | Nature | Objectif | Partie, chapitre, section",
    '1 | Outil un | Classeur | But | P1 C1 S1', 'Source : réalisé par l\'auteur.',
    'Plan synthétique', 'Introduction (4 pages)', 'Partie 1 – Le constat (48 pages)', 'Chapitre 1 – Les faits (24 pages)', 'Chapitre 2 – Les causes (24 pages)',
    'Partie 2 – La réponse (45 pages)', 'Chapitre 1 – La méthode (22,5 pages)', 'Chapitre 2 – La mise en œuvre (22,5 pages)', 'Conclusion (3 pages)',
    'Plan détaillé', 'Introduction (4 pages)', 'Partie 1 – Le constat (48 pages)', 'Chapitre 1 – Les faits (24 pages)',
    'Section 1 – Ce qui est reçu (12 pages)', '§ 1 – Les fichiers (6 pages) – outil n° 1 : table de passage', '§ 2 – Les délais (6 pages)',
    'Section 2 – Ce qui est fait (12 pages)', '§ 1 – Le traitement (6 pages)', '§ 2 – La livraison (6 pages)', 'Conclusion (3 pages)',
    'Bibliographie', 'Référence réglementaire', 'NOM Prénom, (2020), Titre, Éditeur, 100 pages.'];
  const doc = { source: 'word', blocs: lignes.map((texte, i) => ({ i, texte, label: '', titre: false, niveau: 0, role: null, style: null, tableau: null, liens: [], image: false, notes: 0, consigne: false, sautAvant: false, garde: false })) };
  const plan = planifier(doc, await regles(), 'notice');
  const st = (debut) => (plan.lignes.find((l) => l.texte.startsWith(debut)) || {}).style;
  assert.equal(st('Tableau 1'), 'Caption');
  assert.equal(st('Source'), 'Source');
  assert.equal(st('- premier'), 'Listetirets');
  assert.equal(st('NOM Prénom'), 'Referencebiblio');
  assert.equal(plan.lignes.filter((l) => l.tableauTexte).map((l) => l.tableauTexte.length).join(','), '3,2');
  const pl = plan.lignes.filter((l) => l.planLigne).map((l) => l.planLigne);
  assert.deepEqual(pl.find((x) => x.outil), { niveau: 4, titre: '§ 1 – Les fichiers', pages: '6', outil: 'outil n° 1 : table de passage' });
  const octets = await composer(await gabarit('gabarit-demande-agrement.docx'), plan, 'notice',
    { intitule: 'Sujet', auteur: 'Prénom NOM', adresse: '75001 Paris', date: 'novembre 2026', regles: await regles() });
  const sortie = await lireDocx(octets);
  assert.equal(sortie.tableaux, 4);
  assert.ok(sortie.blocs.some((b) => b.tableau !== null && b.texte === '22,5 pages'));
  // Relu par le moteur : tableaux présents (A10), plans lus dans leurs tableaux (A12, A13), aucun tableau de plan sans titre (B10).
  const r = await lancer(sortie, { profil: 'notice' });
  for (const id of ['A10', 'A13']) assert.ok(!r.ids.has(id), r.lignes.filter((l) => l.startsWith(id)).join('\n'));
  assert.ok(!r.lignes.some((l) => /^A12 .*(absent|sans nombre)/.test(l)), r.lignes.filter((l) => l.startsWith('A12')).join('\n'));
  assert.ok(!r.lignes.some((l) => /^B10 .*sans titre/.test(l)), r.lignes.filter((l) => l.startsWith('B10')).join('\n'));
});

test('repères « à insérer » et couleurs des titres', async () => {
  const lignes = ["Demande d'agrément du sujet de mémoire présentée en vue de l'obtention du diplôme d'expertise comptable", 'Sujet',
    'Notice explicative', '1. Intérêt du sujet et problématique', 'Problématique : Comment faire ?',
    "3. Différenciation avec les autres mémoires d'expertise comptable", 'Texte de la section trois, sans tableau, assez long pour être un paragraphe.',
    '4. Possibilités d\'atteindre les apports', 'Tableau 2 – Calendrier', 'Texte qui suit la légende sans tableau, assez long pour un paragraphe.',
    '5. Apports du mémoire à la profession', 'Tableau 3 – Outils', '1 | Outil | Classeur | But | P1 C1 S1', '2 | Outil | Classeur | But | P1 C1 S2',
    'Bibliographie', 'Référence réglementaire'];
  const doc = { source: 'word', blocs: lignes.map((texte, i) => ({ i, texte, label: '', titre: false, niveau: 0, role: null, style: null, tableau: null, liens: [], image: false, notes: 0, consigne: false, sautAvant: false, garde: false })) };
  const plan = planifier(doc, await regles(), 'notice');
  const reperes = plan.lignes.filter((l) => l.repere).map((l) => l.texte);
  assert.equal(reperes.length, 3, reperes.join('\n'));
  assert.match(reperes[0], /^\[Tableau 1 à insérer ici : tableau de différenciation/);
  assert.equal(reperes[1], '[Tableau 2 à insérer ici]');
  assert.match(reperes[2], /^\[Annexe 1 à insérer ici/);
  const octets = await composer(await gabarit('gabarit-demande-agrement.docx'), plan, 'notice', { intitule: 'Sujet', regles: await regles() });
  const sortie = await lireDocx(octets);
  assert.equal(sortie.blocs.filter((b) => b.consigne && /à insérer ici/.test(b.texte)).length, 3);
  const r = await lancer(sortie, { profil: 'notice' });
  assert.ok(r.lignes.some((l) => /^F01 .*Consigne/.test(l)), 'repères signalés comme restes de travail');
  assert.ok(!sortie.blocs.some((b) => b.rouge), 'aucun passage pris pour du rouge');
  const { lireZip } = await import('../sources/zip.js');
  const styles = await (await lireZip(octets)).texte('word/styles.xml');
  for (const [id, c] of [['Titrenotice', '1F3864'], ['Titrehorsplan', '1F3864'], ['Soustitrenotice', '2E5C8A'], ['Titrecatégorie', '2E5C8A']]) {
    assert.ok(new RegExp(`w:styleId="${id}"[\\s\\S]*?<w:color w:val="${c}"/>[\\s\\S]*?</w:style>`).test(styles) && !new RegExp(`w:styleId="${id}"(?:(?!</w:style>)[\\s\\S])*<w:color w:val="000000"/>`).test(styles), id);
  }
});
