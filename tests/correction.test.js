// Corrections dans Word : imitation minimale de Word.run et de paragraph.search (recherche littérale ou [0-9]).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { regles } from './outils.js';

function imiterWord(textes) {
  let attente = [];
  const items = textes.map((text) => ({ text }));
  for (const p of items) {
    p.load = () => {};
    p.insertText = (t) => attente.push({ p, debut: 0, fin: p.text.length, n: t });
    p.search = (q, o = {}) => {
      const motif = o.matchWildcards ? q : q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const re = new RegExp(motif, 'g');
      const res = { items: [], load() {} };
      let m;
      while ((m = re.exec(p.text))) {
        const [debut, fin] = [m.index, m.index + m[0].length];
        if (o.matchWholeWord && (/\p{L}/u.test(p.text[debut - 1] || '') || /\p{L}/u.test(p.text[fin] || ''))) continue;
        res.items.push({ text: m[0], insertText: (n) => attente.push({ p, debut, fin, n }) });
      }
      return res;
    };
  }
  globalThis.Word = {
    run: async (f) => f({
      document: { body: { paragraphs: { items, load() {} } } },
      sync: async () => {
        attente.sort((a, b) => b.debut - a.debut).forEach(({ p, debut, fin, n }) => { p.text = p.text.slice(0, debut) + n + p.text.slice(fin); });
        attente = [];
      },
    }),
  };
  return items;
}

test('corrections : insécables, espaces, unités, accents, numéros, guillemets, intitulés', async () => {
  const { corriger, apercu } = await import('../sources/correction.js');
  const textes = ['Une phrase : avec  deux espaces , et 28% puis 10 € « citation » ?', 'Etat et Ecole', 'Section 1.2 : Genèse',
    'Le "middle market" et "PME"', '1.1 Contexte en version 1.1', '5. Limites du sujet', 'Mesure de 12" et "x'];
  const items = imiterWord(textes);
  const doc = { blocs: textes.map((texte, i) => ({ i, texte })) };
  const liste = [{ type: 'insecables', i: 0 }, { type: 'espaces', i: 0 }, { type: 'unites', i: 0 }, { type: 'accents', i: 1 },
    { type: 'numero-simple', i: 2 }, { type: 'guillemets', i: 3 }, { type: 'sans-numero', i: 4 },
    { type: 'renommer', i: 5, texte: '7. Limites du mémoire' }, { type: 'guillemets', i: 6 }];
  assert.equal(apercu(liste[8], doc.blocs[6]), null, 'guillemets ambigus (pouces) : pas de correction');
  const { decales } = await corriger(liste, doc, (await regles()).parametres);
  assert.equal(decales, 0);
  const n = ' ';
  assert.deepEqual(items.map((p) => p.text), [
    `Une phrase${n}: avec deux espaces, et 28${n}% puis 10${n}€ «${n}citation${n}»${n}?`, 'État et École', 'Section 2 : Genèse',
    `Le «${n}middle market${n}» et «${n}PME${n}»`, 'Contexte en version 1.1', '7. Limites du mémoire', 'Mesure de 12" et "x']);
});

test('corrections : paragraphe modifié depuis le contrôle laissé tel quel', async () => {
  const { corriger } = await import('../sources/correction.js');
  const items = imiterWord(['Texte  modifié']);
  const r = await corriger([{ type: 'espaces', i: 0 }], { blocs: [{ i: 0, texte: 'Autre texte' }] }, (await regles()).parametres);
  assert.equal(r.decales, 1);
  assert.equal(items[0].text, 'Texte  modifié');
});
