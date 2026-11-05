// Corrections dans Word : imitation de Word où chaque caractère porte sa mise en forme, et où un texte inséré
// prend la police par défaut (Aptos 11), comme dans Word en ligne. Le complément doit réappliquer la mise en forme.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { regles } from './outils.js';

const DEFAUT = { name: 'Aptos', size: 11, italic: false, bold: false };
const PROPS = ['name', 'size', 'bold', 'italic', 'underline', 'color', 'highlightColor', 'superscript', 'subscript', 'strikeThrough'];

function imiterWord(paragraphes) {
  // paragraphes : [[texte, police, { debut, fin, police }?]]
  const items = paragraphes.map(([texte, police, zone]) => {
    const chars = [...texte].map((c, k) => ({ c, f: { ...DEFAUT, ...(zone && k >= zone.debut && k < zone.fin ? zone.police : police) } }));
    return { chars };
  });
  const plage = (p, a, b) => {
    const r = {
      get text() { return p.chars.slice(a, b).map((x) => x.c).join(''); },
      font: new Proxy({}, {
        get(_, k) {
          const v = p.chars.slice(a, b).map((x) => x.f[k]);
          return v.every((x) => x === v[0]) ? v[0] : null;
        },
        set(_, k, v) { p.chars.slice(a, b).forEach((x) => { x.f[k] = v; }); return true; },
      }),
      search(q) {
        const t = r.text;
        const res = { items: [], load() {} };
        for (let i = t.indexOf(q); i >= 0; i = t.indexOf(q, i + 1)) res.items.push(plage(p, a + i, a + i + q.length));
        return res;
      },
      insertText(n) {
        p.chars.splice(a, b - a, ...[...n].map((c) => ({ c, f: { ...DEFAUT } })));
        return plage(p, a, a + n.length);
      },
      load() {},
    };
    return r;
  };
  for (const p of items) {
    Object.defineProperty(p, 'text', { get: () => p.chars.map((x) => x.c).join('') });
    p.load = () => {};
    p.search = (q) => plage(p, 0, p.chars.length).search(q);
    p.insertText = (n) => plage(p, 0, p.chars.length).insertText(n);
    Object.defineProperty(p, 'font', { get: () => plage(p, 0, p.chars.length).font });
  }
  globalThis.Word = {
    run: async (f) => f({ document: { body: { paragraphs: { items, load() {} } } }, sync: async () => {} }),
  };
  return items;
}

const n = ' ';
const doc = (items) => ({ blocs: items.map((p, i) => ({ i, texte: p.text })) });

test('corrections : texte attendu pour chaque type', async () => {
  const { corriger } = await import('../sources/correction.js');
  const C = { name: 'Calibri', size: 12 };
  const items = imiterWord([
    ['Une phrase : avec  deux espaces , et 28% puis 10 € « citation » ?', C], ['Etat et Ecole', C], ['Section 1.2 : Genèse', C],
    ['Le "middle market" et "PME"', C], ['1.1 Contexte en version 1.1', C], ['5. Limites du sujet', C], ['Mesure de 12" et "x', C],
    ["L'analyse d'un cas", C], ['Sont volontairement exclus : ', C]]);
  const liste = [{ type: 'insecables', i: 0 }, { type: 'espaces', i: 0 }, { type: 'unites', i: 0 }, { type: 'accents', i: 1 },
    { type: 'numero-simple', i: 2 }, { type: 'guillemets', i: 3 }, { type: 'sans-numero', i: 4 },
    { type: 'renommer', i: 5, texte: '7. Limites du mémoire' }, { type: 'guillemets', i: 6 }, { type: 'apostrophes', i: 7 }, { type: 'insecables', i: 8 }];
  const r = await corriger(liste, doc(items), (await regles()).parametres);
  assert.equal(r.decales, 0);
  assert.deepEqual(items.map((p) => p.text), [
    `Une phrase${n}: avec deux espaces, et 28${n}% puis 10${n}€ «${n}citation${n}»${n}?`, 'État et École', 'Section 2 : Genèse',
    `Le «${n}middle market${n}» et «${n}PME${n}»`, 'Contexte en version 1.1', '7. Limites du mémoire', 'Mesure de 12" et "x',
    'L’analyse d’un cas', `Sont volontairement exclus${n}: `]);
});

test('corrections : la mise en forme d\'origine est conservée (police, taille, italique)', async () => {
  const { corriger } = await import('../sources/correction.js');
  const C = { name: 'Calibri', size: 12, italic: false };
  const texte = 'Le "Ré" et "middle market" : fin';
  const debut = texte.indexOf('Ré');
  const items = imiterWord([[texte, C, { debut, fin: debut + 2, police: { ...C, italic: true } }]]);
  await corriger([{ type: 'guillemets', i: 0 }, { type: 'insecables', i: 0 }], doc(items), (await regles()).parametres);
  const p = items[0];
  assert.equal(p.text, `Le «${n}Ré${n}» et «${n}middle market${n}»${n}: fin`);
  for (const x of p.chars) assert.deepEqual([x.f.name, x.f.size], ['Calibri', 12], `« ${x.c} » en ${x.f.name} ${x.f.size}`);
  const re = p.text.indexOf('Ré');
  assert.ok(p.chars[re].f.italic && p.chars[re + 1].f.italic, 'italique de « Ré » conservé');
  assert.ok(!p.chars[0].f.italic, 'pas d\'italique ailleurs');
});

test('corrections : annulation en un clic, texte et mise en forme', async () => {
  const { corriger, annuler } = await import('../sources/correction.js');
  const C = { name: 'Calibri', size: 12 };
  const textes = ['Sont exclus : les cas ; et 28% de "PME"', 'Etat  et Ecole'];
  const items = imiterWord(textes.map((t) => [t, C]));
  const r = await corriger([{ type: 'insecables', i: 0 }, { type: 'unites', i: 0 }, { type: 'guillemets', i: 0 }, { type: 'espaces', i: 1 },
    { type: 'accents', i: 1 }], doc(items), (await regles()).parametres);
  assert.ok(r.corrections >= 6);
  await annuler(r.journal);
  assert.deepEqual(items.map((p) => p.text), textes);
  for (const p of items) for (const x of p.chars) assert.deepEqual([x.f.name, x.f.size], ['Calibri', 12]);
});

test('corrections : paragraphe modifié depuis le contrôle laissé tel quel', async () => {
  const { corriger } = await import('../sources/correction.js');
  const items = imiterWord([['Texte  modifié', { name: 'Calibri', size: 12 }]]);
  const r = await corriger([{ type: 'espaces', i: 0 }], { blocs: [{ i: 0, texte: 'Autre texte' }] }, (await regles()).parametres);
  assert.equal(r.decales, 1);
  assert.equal(items[0].text, 'Texte  modifié');
});
