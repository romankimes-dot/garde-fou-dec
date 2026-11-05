// Corrections automatiques dans le document ouvert (Office.js), à la demande, paragraphe par paragraphe.
// Seules des corrections sûres et mécaniques sont proposées ; chacune s'annule par Ctrl+Z dans Word.

const NBSP = '\u00a0';
const ACCENTS = { Ere: 'Ère' }; // sinon : E initial → É

/** Recherches et remplacements d'une règle pour un paragraphe. */
function remplacements(regle, bloc, parametres) {
  switch (regle) {
    case 'F04':
      return [
        ...[':', ';', '!', '?'].map((c) => ({ cherche: ` ${c}`, remplace: () => NBSP + c })),
        { cherche: '« ', remplace: () => `«${NBSP}` },
        { cherche: ' »', remplace: () => `${NBSP}»` },
      ];
    case 'F05':
      return [
        { cherche: '  ', remplace: () => ' ', repeter: true },
        { cherche: ' ,', remplace: () => ',' },
      ];
    case 'F07':
      return ['%', '€'].flatMap((u) => [
        { cherche: `[0-9]${u}`, options: { matchWildcards: true }, remplace: (t) => t.slice(0, -1) + NBSP + u },
        { cherche: `[0-9] ${u}`, options: { matchWildcards: true }, remplace: (t) => t[0] + NBSP + u },
      ]);
    case 'F06':
      return parametres.majuscules_accentuees.map((m) => ({
        cherche: m, options: { matchWholeWord: true }, remplace: () => ACCENTS[m] || `É${m.slice(1)}`,
      }));
    case 'A06': {
      // Numéro tapé à la main seulement : un numéro automatique se corrige dans la définition de la liste.
      const m = /(partie|chapitre|section|paragraphe|§)\s*(\d+(?:\.\d+)+)/i.exec(bloc.texte);
      if (!m) return [];
      return [{ cherche: m[0], remplace: () => `${m[1]} ${m[2].split('.').pop()}` }];
    }
    default:
      return [];
  }
}

export const CORRIGEABLES = new Set(['A06', 'F04', 'F05', 'F06', 'F07']);

const comparable = (t) => (t || '').replace(/[\r\u0007\u000b]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);

/**
 * Applique les corrections d'une règle aux paragraphes `indices` (numéros de body.paragraphs).
 * Un paragraphe dont le texte a changé depuis la dernière lecture est laissé tel quel (décalage).
 * Rend { corrections, decales }.
 */
export async function corriger(regle, indices, doc, parametres) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load('items/text');
    await context.sync();
    const cibles = [];
    let decales = 0;
    for (const i of new Set(indices)) {
      const para = paras.items[i];
      const bloc = doc.blocs[i];
      if (!para || !bloc || comparable(para.text) !== comparable(bloc.texte)) { decales++; continue; }
      cibles.push({ para, bloc });
    }
    let corrections = 0;
    for (let passe = 0; passe < 4; passe++) {
      const lots = [];
      for (const { para, bloc } of cibles) {
        for (const r of remplacements(regle, bloc, parametres)) {
          if (passe > 0 && !r.repeter) continue;
          const res = para.search(r.cherche, { matchCase: true, ...(r.options || {}) });
          res.load('items/text');
          lots.push({ res, r });
        }
      }
      if (!lots.length) break;
      await context.sync();
      let n = 0;
      for (const { res, r } of lots) {
        for (const x of res.items) { x.insertText(r.remplace(x.text), 'Replace'); n++; }
      }
      if (!n) break;
      await context.sync();
      corrections += n;
    }
    return { corrections, decales };
  });
}
