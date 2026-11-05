// Tests du moteur sur les documents fictifs (python3 tests/fabriquer-fixtures.py) et sur les gabarits du dépôt.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { docx, lancer } from './outils.js';

const contient = (r, id, morceau) => assert.ok(
  r.lignes.some((l) => l.startsWith(id + ' ') && l.includes(morceau)),
  `${id} « ${morceau} » attendu ; obtenu :\n${r.lignes.join('\n')}`,
);

test('gabarit du mémoire : numéros automatiques lus depuis numbering.xml', async () => {
  const d = await docx('fixtures/gabarit-memoire.docx');
  const titres = d.blocs.filter((b) => b.style && b.style.startsWith('heading')).map((b) => b.label.trim());
  assert.deepEqual(titres.slice(0, 5), ['Partie 1 –', 'Chapitre 1 –', 'Section 1 –', '§ 1 –', '§ 2 –']);
  assert.equal(titres[5], 'Partie 2 –');
  assert.equal(titres[6], 'Chapitre 1 –');
});

test('gabarit de la demande : aucune erreur de structure en rédaction, hors contenu à écrire', async () => {
  const r = await lancer(await docx('fixtures/gabarit-demande-agrement.docx'), { profil: 'notice', mode: 'redaction' });
  for (const id of ['A06', 'A09', 'A14', 'A19']) assert.ok(!r.ids.has(id), `${id} inattendu :\n${r.lignes.join('\n')}`);
});

test('notice fictive en mode dépôt : chaque écart semé est trouvé', async () => {
  const r = await lancer(await docx('fixtures/notice-fictive.docx'), { profil: 'notice', mode: 'depot', termes: ['Fictif Durand'] });
  assert.equal(r.feu, 'rouge');
  contient(r, 'A06', 'proscrite');
  contient(r, 'A07', '« 2 » attendu');
  contient(r, 'A07', 'Partie 2 : 1 chapitre');
  contient(r, 'A08', 'seul sous « Section 2 »');
  contient(r, 'A09', 'Section 6');
  contient(r, 'A10', 'Section 5 sans tableau');
  contient(r, 'A11', 'Aucune phrase interrogative');
  contient(r, 'A13', 'que les parties et chapitres');
  contient(r, 'A13', 'pages absent');
  contient(r, 'A16', 'Fictif Durand');
  contient(r, 'A17', 'Tableau récapitulatif des usages absent');
  contient(r, 'A18', 'placée après « Ouvrage »');
  contient(r, 'A18', 'alphabétique');
  contient(r, 'B01', 'Mettre à jour');
  contient(r, 'B08', 'NPMQ');
  contient(r, 'F01', 'TODO');
  contient(r, 'F01', 'Consigne du gabarit');
  contient(r, 'F03', '( )');
  contient(r, 'F04', "pas d'espace avant « ; »");
  contient(r, 'F05', 'espaces doublés');
  contient(r, 'F06', 'Etat');
  contient(r, 'F06', 'A → À');
  contient(r, 'F07', '%');
  // Défini entre parenthèses : pas d'alerte.
  assert.ok(!r.lignes.some((l) => l.includes('« LAB »')));
  // Le plan est cohérent sur le total et les sommes.
  assert.ok(!r.lignes.some((l) => l.includes('somme des sous-niveaux')));
  assert.ok(!r.ids.has('A01') && !r.ids.has('A15'));
});

test('notice fictive en mode rédaction : restes de travail hors champ, typographie contrôlée, feu non vert', async () => {
  const r = await lancer(await docx('fixtures/notice-fictive.docx'), { profil: 'notice', mode: 'redaction', termes: ['Fictif Durand'] });
  for (const id of ['F01', 'F02']) assert.ok(!r.ids.has(id), `${id} hors du mode rédaction`);
  for (const id of ['F04', 'F05', 'F06', 'F07']) assert.ok(r.ids.has(id), `${id} attendu en rédaction aussi`);
  assert.ok(r.ids.has('A06') && r.ids.has('A16'));
});

test('ignorer une occurrence la retire du compte', async () => {
  const d1 = await docx('fixtures/notice-fictive.docx');
  const r1 = await lancer(d1, { profil: 'notice', mode: 'depot', termes: ['Fictif Durand'] });
  const o = r1.resultats.find((x) => x.regle.id === 'A16').occurrences[0];
  const r2 = await lancer(await docx('fixtures/notice-fictive.docx'), { profil: 'notice', mode: 'depot', termes: ['Fictif Durand'], ignores: new Set([o.cle]) });
  assert.equal(r2.comptes.ignores, 1);
  assert.equal(r2.comptes.bloquants, r1.comptes.bloquants - 1);
});

test('mémoire fictif : renvois, annexe IA, plan', async () => {
  const r = await lancer(await docx('fixtures/memoire-fictif.docx'), { profil: 'memoire', mode: 'depot' });
  contient(r, 'B06', "l'annexe 2, qui n'existe pas");
  contient(r, 'A17', "Usage de l'Intelligence Artificielle");
  contient(r, 'A07', 'Partie 1 : 1 chapitre');
  contient(r, 'B04', 'annonce du plan');
  contient(r, 'B05', 'conclusion partielle');
  contient(r, 'A16', 'vide');
  // Le renvoi « voir annexe 1 » du gabarit est un lien : pas d'alerte « non cliquable ».
  assert.ok(!r.lignes.some((l) => l.includes('non cliquable')));
  // Mise en page du gabarit conforme.
  assert.ok(!r.ids.has('A19'), r.lignes.filter((l) => l.startsWith('A19')).join('\n'));
});

test('annexe n° 1 de la charte IA laissée telle que le modèle officiel', async () => {
  const r = await lancer(await docx('fixtures/annexe-ia-modele.docx'), { profil: 'notice', mode: 'depot' });
  contient(r, 'A17', 'Exemple du modèle encore présent');
  contient(r, 'A17', 'Date non renseignée');
  contient(r, 'A17', 'Nom – prénom non renseignés');
  contient(r, 'A17', 'Signature non repérée');
});

test('plan : numérotation reprise à 1 et continue mélangées', async () => {
  const r = await lancer(await docx('fixtures/plan-numerotation-melangee.docx'), { profil: 'notice', mode: 'depot' });
  contient(r, 'A07', 'un seul usage pour tout le plan');
  assert.equal(r.lignes.filter((l) => l.includes('un seul usage')).length, 1);
});
