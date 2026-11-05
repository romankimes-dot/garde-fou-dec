// Analyseur XML minimal, sans dépendance : suffisant pour les parties d'un fichier .docx.
// Un nœud est { n: nom, a: attributs, c: enfants } ; un texte est une chaîne.

const ENTITES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decoder(s) {
  return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => {
    if (e[0] === '#') {
      const code = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return String.fromCodePoint(code);
    }
    return ENTITES[e] ?? m;
  });
}

const JETON = /<(\/?)([\w:.-]+)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<![^>]*>|([^<]+)/g;
const ATTRIBUT = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

export function analyserXml(xml) {
  const racine = { n: '#racine', a: {}, c: [] };
  const pile = [racine];
  JETON.lastIndex = 0;
  let m;
  while ((m = JETON.exec(xml))) {
    const haut = pile[pile.length - 1];
    if (m[6] !== undefined) { haut.c.push(decoder(m[6])); continue; }
    if (m[5] !== undefined) { haut.c.push(m[5]); continue; }
    if (!m[2]) continue;
    if (m[1]) { if (pile.length > 1) pile.pop(); continue; }
    const a = {};
    m[3].replace(ATTRIBUT, (_, k, v1, v2) => { a[k] = decoder(v1 ?? v2); return ''; });
    const el = { n: m[2], a, c: [] };
    haut.c.push(el);
    if (!m[4]) pile.push(el);
  }
  return racine;
}

export const enfants = (el, nom) => (el?.c || []).filter((x) => typeof x === 'object' && x.n === nom);
export const premier = (el, nom) => (el?.c || []).find((x) => typeof x === 'object' && x.n === nom) || null;
export const val = (el, a = 'w:val') => (el ? el.a[a] : undefined);

/** Premier descendant portant ce nom (profondeur d'abord). */
export function chercher(el, nom) {
  for (const x of el?.c || []) {
    if (typeof x !== 'object') continue;
    if (x.n === nom) return x;
    const r = chercher(x, nom);
    if (r) return r;
  }
  return null;
}

/** Tous les descendants portant ce nom. */
export function tous(el, nom, acc = []) {
  for (const x of el?.c || []) {
    if (typeof x !== 'object') continue;
    if (x.n === nom) acc.push(x);
    tous(x, nom, acc);
  }
  return acc;
}

/** Texte brut de tous les descendants (pour les petites parties : propriétés, en-têtes). */
export function texteBrut(el) {
  let s = '';
  for (const x of el?.c || []) s += typeof x === 'string' ? x : texteBrut(x);
  return s;
}
