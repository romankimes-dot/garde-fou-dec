// Corrections dans le document ouvert (Office.js), à la demande, paragraphe par paragraphe.
// Seules des corrections mécaniques sont proposées, avec aperçu quand le texte change de sens ;
// chacune s'annule par Ctrl+Z dans Word. Un paragraphe modifié depuis le contrôle est laissé tel quel.

const NBSP = '\u00a0';
const W = { matchWildcards: true };
const ACCENTS = { Ere: 'Ère' }; // sinon : E initial → É

export const LIBELLES = {
  insecables: 'espaces insécables avant : ; ! ? et dans « »',
  guillemets: 'guillemets droits " " remplacés par « »',
  espaces: 'espaces doublés et espace avant la virgule',
  unites: 'espace insécable avant % et €',
  accents: 'majuscules accentuées',
  'numero-simple': 'numéro simple (« Section 1.2 » → « Section 2 »)',
  'sans-numero': 'numéro 1.1 retiré du sous-titre',
  renommer: 'intitulé exact de la Note',
  trier: 'sélection pour le tri alphabétique de Word',
  'style-normal': 'style Normal en taille 12',
};

// Corrections dont le résultat est montré avant application.
export const AVEC_APERCU = new Set(['guillemets', 'renommer', 'sans-numero', 'numero-simple']);

/** Paires de guillemets droits d'un texte : null si ambiguës (nombre impair, pouces, espace intérieure). */
function pairesGuillemets(t) {
  const pos = [...t.matchAll(/"/g)].map((m) => m.index);
  if (!pos.length || pos.length % 2) return null;
  for (let k = 0; k < pos.length; k += 2) {
    const [a, b] = [pos[k], pos[k + 1]];
    if (/\d/.test(t[a - 1] || '') || /\s/.test(t[a + 1] || '') || /\s/.test(t[b - 1] || '') || b - a < 2) return null;
  }
  return pos;
}

/** Texte après correction (aperçu), ou null si la correction ne s'applique pas. */
export function apercu(c, bloc) {
  const t = bloc.texte;
  switch (c.type) {
    case 'renommer': return c.texte;
    case 'guillemets': {
      const pos = pairesGuillemets(t);
      if (!pos) return null;
      let k = 0;
      return t.replace(/"/g, () => (k++ % 2 ? `${NBSP}»` : `«${NBSP}`));
    }
    case 'sans-numero': {
      const m = /^\s*\d+(?:\.\d+)+\.?\s+/.exec(t);
      return m ? t.slice(m[0].length) : null;
    }
    case 'numero-simple': {
      const m = /(partie|chapitre|section|paragraphe|§)\s*(\d+(?:\.\d+)+)/i.exec(t);
      return m ? t.replace(m[0], `${m[1]} ${m[2].split('.').pop()}`) : null;
    }
    default: return null;
  }
}

/** Recherches et remplacements dans un paragraphe. */
function remplacements(c, bloc, parametres) {
  const t = bloc.texte;
  switch (c.type) {
    // Recherches en mode « caractères génériques » : l'espace est écrite [ ], jamais en tête ou en fin de chaîne.
    case 'insecables':
      return [
        ...[':', ';', '\\!', '\\?'].map((x) => ({ cherche: `[ ]${x}`, options: W, remplace: (t) => NBSP + t.slice(1) })),
        { cherche: '«[ ]', options: W, remplace: () => `«${NBSP}` },
        { cherche: '[ ]»', options: W, remplace: () => `${NBSP}»` },
      ];
    case 'guillemets': {
      const pos = pairesGuillemets(t);
      if (!pos) return [];
      // Word rapproche " de “ ” : on vérifie que les occurrences trouvées sont exactement les guillemets droits attendus.
      return [{ cherche: '"', attendu: pos.length, tous: (x, k) => (k % 2 ? `${NBSP}»` : `«${NBSP}`) }];
    }
    case 'espaces':
      return [{ cherche: '[ ]{2,}', options: W, remplace: () => ' ' }, { cherche: '[ ],', options: W, remplace: () => ',' }];
    case 'unites':
      return ['%', '€'].flatMap((u) => [
        { cherche: `[0-9]${u}`, options: W, remplace: (x) => x.slice(0, -1) + NBSP + u },
        { cherche: `[0-9][ ]${u}`, options: W, remplace: (x) => x[0] + NBSP + u },
      ]);
    case 'accents':
      return parametres.majuscules_accentuees.map((m) => ({
        cherche: m, options: { matchWholeWord: true }, remplace: () => ACCENTS[m] || `É${m.slice(1)}`,
      }));
    case 'numero-simple': {
      const m = /(partie|chapitre|section|paragraphe|§)\s*(\d+(?:\.\d+)+)/i.exec(t);
      return m ? [{ cherche: m[0], premier: true, remplace: () => `${m[1]} ${m[2].split('.').pop()}` }] : [];
    }
    case 'sans-numero': {
      const m = /^\s*(\d+(?:\.\d+)+\.?)\s+/.exec(t);
      return m ? [{ cherche: `${m[1]}[ ]`, options: W, premier: true, remplace: () => '' }] : [];
    }
    default:
      return [];
  }
}

// Comparaison tolérante (espaces, insécables, traits d'union conditionnels, caractères de contrôle) : sert à repérer
// un paragraphe réellement réécrit depuis le contrôle, pas une différence de lecture.
const comparable = (t) => (t || '').replace(/[\s\u00a0\u202f\u00ad\u200b\u0007\u000b\r]/g, '').slice(0, 50).toLowerCase();

/**
 * Applique des corrections [{ type, i, texte?, style? }] (i : numéro de body.paragraphs).
 * Rend { corrections, decales, ignorees }.
 */
export async function corriger(liste, doc, parametres) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load('items/text');
    await context.sync();
    let decales = 0;
    let ignorees = 0;
    let corrections = 0;
    let trouves = 0;
    const exemples = [];
    const cibles = [];
    const vus = new Set();
    for (const c of liste) {
      const cle = `${c.type}|${c.i}`;
      if (vus.has(cle)) continue;
      vus.add(cle);
      if (c.type === 'style-normal') { cibles.push({ c }); continue; } // style du document, pas un paragraphe
      const para = paras.items[c.i];
      const bloc = doc.blocs[c.i];
      if (!para || !bloc || comparable(para.text) !== comparable(bloc.texte)) {
        decales++;
        if (exemples.length < 3) exemples.push(`¶ ${c.i + 1} : ${para ? `${para.text.length} caractères lus contre ${bloc ? bloc.texte.length : '?'}` : 'introuvable'}`);
        continue;
      }
      cibles.push({ c, para, bloc });
    }

    // Paragraphe entier remplacé (intitulés) : le style de titre de référence est appliqué s'il est connu.
    for (const { c, para } of cibles.filter((x) => x.c.type === 'renommer')) {
      para.insertText(c.texte, 'Replace');
      if (c.style) para.styleBuiltIn = c.style;
      corrections++;
    }
    if (cibles.some((x) => x.c.type === 'style-normal')) {
      const normal = context.document.getStyles().getByNameOrNullObject('Normal');
      normal.load('isNullObject');
      await context.sync();
      if (!normal.isNullObject) { normal.font.size = parametres.taille_texte; corrections++; }
    }

    // Recherches dans les paragraphes, en plusieurs passes pour les remplacements répétés.
    const recherches = cibles.filter((x) => !['renommer', 'style-normal', 'trier'].includes(x.c.type));
    for (let passe = 0; passe < 4; passe++) {
      const lots = [];
      for (const { c, para, bloc } of recherches) {
        for (const r of remplacements(c, bloc, parametres)) {
          if (passe > 0 && !r.repeter) continue;
          const res = para.search(r.cherche, { matchCase: true, ...(r.options || {}) });
          res.load('items/text');
          lots.push({ res, r });
        }
      }
      if (!lots.length) break;
      await context.sync();
      let n = 0;
      for (const { res } of lots) trouves += res.items.length;
      for (const { res, r } of lots) {
        if (r.attendu !== undefined && (res.items.length !== r.attendu || res.items.some((x) => x.text !== '"'))) { ignorees++; continue; }
        const items = r.premier ? res.items.slice(0, 1) : res.items;
        items.forEach((x, k) => { x.insertText(r.tous ? r.tous(x.text, k) : r.remplace(x.text), 'Replace'); n++; });
      }
      if (!n) break;
      await context.sync();
      corrections += n;
    }
    await context.sync();
    return { corrections, decales, ignorees, trouves, paragraphes: cibles.length, exemples };
  });
}

/** Sélectionne les paragraphes i..j dans Word (tri par Accueil, Trier, qui conserve la mise en forme). */
export async function selectionnerPlage(i, j) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load('items');
    await context.sync();
    const a = paras.items[i];
    const b = paras.items[j];
    if (!a || !b) return false;
    a.getRange('Start').expandTo(b.getRange('End')).select();
    await context.sync();
    return true;
  });
}
