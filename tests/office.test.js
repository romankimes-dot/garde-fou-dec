// Lecture par Office.js, sur une imitation de l'interface Word construite à partir de la notice fictive.
// Vérifie que les deux sources (document ouvert et .docx) donnent les mêmes constats, hors contrôles propres au .docx.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { docx, lancer } from './outils.js';
import { lireWord, selectionnerParagraphe } from '../sources/office.js';

function installerFauxWord(modele) {
  const proxy = (o) => Object.assign(o, { load() { return o; } });
  const collection = (items) => proxy({ items });
  let selection = null;
  const garde = modele.blocs.filter((b) => b.garde).length;
  const paras = modele.blocs.map((b, i) => {
    const titre = /^heading (\d)$/.exec(b.style || '');
    const toc = /^toc (\d)$/.exec(b.style || '');
    return proxy({
      text: b.texte,
      style: titre ? `Titre ${titre[1]}` : toc ? `TM ${toc[1]}` : b.style === 'normal' ? 'Normal' : (b.style || 'Normal'),
      styleBuiltIn: titre ? `Heading${titre[1]}` : toc ? `Toc${toc[1]}` : b.style === 'normal' ? 'Normal' : 'Other',
      outlineLevel: 10,
      tableNestingLevel: b.tableau === null ? 0 : 1,
      isListItem: !!b.label,
      listItemOrNullObject: proxy({ isNullObject: !b.label, listString: b.label.trim() }),
      parentTableCellOrNullObject: proxy({ isNullObject: b.tableau === null, rowIndex: b.ligne, cellIndex: b.colonne }),
      select() { selection = i; },
    });
  });
  const vide = () => proxy({ text: '' });
  const section = (n) => proxy({ body: { paragraphs: collection(n) }, getHeader: vide, getFooter: vide });
  const document = {
    body: {
      paragraphs: collection(paras),
      getRange: () => ({ getHyperlinkRanges: () => collection(modele.blocs.flatMap((b) => b.liens).map((t) => ({ text: t }))) }),
      getComments: () => collection([]),
      getTrackedChanges: () => collection([]),
      footnotes: collection([]),
    },
    sections: collection(garde ? [section(paras.slice(0, garde)), section(paras.slice(garde))] : [section(paras)]),
    properties: proxy({ author: 'Prénom NOM' }),
  };
  globalThis.Office = { context: { requirements: { isSetSupported: () => true } } };
  globalThis.Word = { run: async (f) => f({ document, sync: async () => {} }) };
  return () => selection;
}

test('Office.js (imitation) : mêmes constats que sur le .docx', async () => {
  const options = { profil: 'notice', termes: ['Fictif Durand'] };
  const viaDocx = await lancer(await docx('fixtures/notice-fictive.docx'), options);
  const selection = installerFauxWord(await docx('fixtures/notice-fictive.docx'));
  const doc = await lireWord();
  assert.equal(doc.source, 'word');
  assert.equal(doc.tableaux, 1);
  const viaWord = await lancer(doc, options);
  const propres = (r) => r.lignes.filter((l) => !/^(A19|F02) /.test(l)).sort();
  assert.deepEqual(propres(viaWord), propres(viaDocx));
  assert.equal(await selectionnerParagraphe(20), true);
  assert.equal(selection(), 20);
});

test('Office.js (imitation) : mémoire, numéros des titres et renvois cliquables', async () => {
  const options = { profil: 'memoire' };
  const viaDocx = await lancer(await docx('fixtures/memoire-fictif.docx'), options);
  installerFauxWord(await docx('fixtures/memoire-fictif.docx'));
  const viaWord = await lancer(await lireWord(), options);
  const propres = (r) => r.lignes.filter((l) => !/^(A19|F02) /.test(l)).sort();
  assert.deepEqual(propres(viaWord), propres(viaDocx));
});
