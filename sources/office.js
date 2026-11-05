// Lecture du document ouvert dans Word (en ligne ou sur ordinateur) par l'interface officielle (Office.js),
// vers le modèle commun du moteur. Lecture seule : rien n'est écrit dans le document.

import { STYLES_TITRES, STYLES_CONSIGNE } from './docx.js';

const BUILTIN_TITRE = /^Heading([1-9])$/;
const BUILTIN_TOC = /^Toc([1-9])$/;
const PUCE_SYMBOLE = /^[-•▪●◦–-]$/;

const api = (version) => {
  try { return Office.context.requirements.isSetSupported('WordApi', version); } catch { return false; }
};

/** Lit le document ouvert. `Word` et `Office` sont fournis par office.js. */
export async function lireWord() {
  return Word.run(async (context) => {
    const document = context.document;
    const paras = document.body.paragraphs;
    paras.load('items/text,items/style,items/styleBuiltIn,items/outlineLevel,items/tableNestingLevel,items/isListItem');
    const sections = document.sections;
    sections.load('items');
    const props = document.properties;
    props.load('author,lastAuthor,company,title,subject,keywords,comments,manager');
    await context.sync();

    const items = paras.items;
    const listes = items.map((p) => (p.isListItem ? p.listItemOrNullObject : null));
    listes.forEach((l) => l && l.load('listString'));
    const cellules = items.map((p) => (p.tableNestingLevel > 0 ? p.parentTableCellOrNullObject : null));
    cellules.forEach((c) => c && c.load('rowIndex,cellIndex'));
    let premiereSection = null;
    if (sections.items.length > 1) {
      premiereSection = sections.items[0].body.paragraphs;
      premiereSection.load('items/tableNestingLevel');
    }
    const entetes = [];
    for (const s of sections.items) {
      for (const type of ['Primary', 'FirstPage', 'EvenPages']) {
        for (const corps of [s.getHeader(type), s.getFooter(type)]) { corps.load('text'); entetes.push(corps); }
      }
    }
    await context.sync();

    // Facultatif selon la version de Word : liens, commentaires, révisions, notes.
    const lire = async (version, f) => {
      if (!api(version)) return null;
      try { return await f(); } catch { return null; }
    };
    const liens = await lire('1.3', async () => {
      const r = document.body.getRange('Whole').getHyperlinkRanges();
      r.load('items/text');
      await context.sync();
      return r.items.map((x) => x.text);
    }) || [];
    const commentaires = await lire('1.4', async () => {
      const c = document.body.getComments();
      c.load('items/content');
      await context.sync();
      return c.items.map((x) => x.content);
    });
    const revisions = await lire('1.6', async () => {
      const t = document.body.getTrackedChanges();
      t.load('items');
      await context.sync();
      return t.items.length;
    });
    const notes = await lire('1.5', async () => {
      const n = document.body.footnotes;
      n.load('items');
      await context.sync();
      const corps = n.items.map((x) => { x.body.load('text'); return x.body; });
      await context.sync();
      return corps.map((b) => b.text).filter((t) => t.trim());
    });

    // Tableaux : un nouveau tableau commence en sortant du texte courant, ou à la cellule (0, 0).
    let tableaux = 0;
    let dansTableau = false;
    let marqueurs = 0;
    const blocs = items.map((p, i) => {
      const nom = (p.style || '').toLowerCase();
      const builtin = p.styleBuiltIn || '';
      const cellule = cellules[i] && !cellules[i].isNullObject ? cellules[i] : null;
      let tableau = null;
      if (p.tableNestingLevel > 0) {
        if (!dansTableau || (cellule && cellule.rowIndex === 0 && cellule.cellIndex === 0 && i > 0
          && !(cellules[i - 1] && !cellules[i - 1].isNullObject && cellules[i - 1].rowIndex === 0 && cellules[i - 1].cellIndex === 0))) {
          tableaux++;
        }
        dansTableau = true;
        tableau = tableaux - 1;
      } else dansTableau = false;

      let role = null;
      let niveau = 0;
      const titre = BUILTIN_TITRE.exec(builtin);
      if (STYLES_TITRES[nom]) [role, niveau] = STYLES_TITRES[nom];
      else if (titre) { role = 'plan'; niveau = Number(titre[1]); }
      else if (BUILTIN_TOC.test(builtin) || /^(toc|tm) ?\d$/.test(nom)) role = 'toc';
      else if (p.outlineLevel >= 1 && p.outlineLevel <= 9) { role = 'autre'; niveau = p.outlineLevel; }

      let label = '';
      const liste = listes[i];
      if (liste && !liste.isNullObject && liste.listString) {
        label = (PUCE_SYMBOLE.test(liste.listString.trim()) ? '•' : liste.listString.trim()) + ' ';
        marqueurs++;
      }
      return {
        i,
        texte: (p.text || '').replace(/\r/g, ' ').replace(/[\u0007\u000b]/g, ' '),
        label,
        titre: niveau > 0 && tableau === null,
        niveau: tableau === null ? niveau : 0,
        role: tableau === null || role === 'toc' ? role : null,
        style: nom || null,
        tableau,
        ligne: cellule ? cellule.rowIndex : null,
        colonne: cellule ? cellule.cellIndex : null,
        liens: [],
        image: false,
        notes: 0,
        consigne: STYLES_CONSIGNE.has(nom),
        sautAvant: false,
        garde: false,
        direct: null,
      };
    });

    if (premiereSection) for (let k = 0; k < premiereSection.items.length && k < blocs.length; k++) blocs[k].garde = true;

    // Chaque lien est rattaché au premier paragraphe, à partir du précédent, qui contient son texte.
    let j = 0;
    for (const l of liens) {
      const t = l.trim();
      if (!t) continue;
      let k = j;
      while (k < blocs.length && !blocs[k].texte.includes(t)) k++;
      if (k < blocs.length) { blocs[k].liens.push(l); j = k; }
    }

    const proprietes = {};
    for (const [cle, libelle] of [['author', 'Auteur'], ['lastAuthor', 'Modifié par'], ['company', 'Société'],
      ['manager', 'Responsable'], ['title', 'Titre'], ['subject', 'Objet'], ['keywords', 'Mots clés'], ['comments', 'Commentaires']]) {
      if (props[cle]) proprietes[libelle] = props[cle];
    }

    return {
      source: 'word',
      blocs,
      tableaux,
      notes: notes || [],
      commentaires: commentaires || [],
      revisions: revisions ?? null,
      entetes: entetes.map((e) => e.text).filter((t) => t && t.trim()),
      proprietes,
      labelsLus: true,
      miseEnPage: null,
      diagnostic: {
        paragraphes: blocs.length,
        titres: blocs.filter((b) => b.titre).length,
        marqueurs,
        tableaux,
        liens: liens.length,
        sections: sections.items.length,
        commentaires: commentaires === null ? 'non lus' : commentaires.length,
        revisions: revisions === null ? 'non lues' : revisions,
        notes: notes === null ? 'non lues' : notes.length,
        versions: ['1.3', '1.4', '1.5', '1.6'].filter(api),
      },
    };
  });
}

/** Sélectionne le paragraphe n° i dans Word (même ordre que body.paragraphs). */
export async function selectionnerParagraphe(i) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load('items/tableNestingLevel');
    await context.sync();
    const p = paras.items[i];
    if (!p) return false;
    p.select('Select');
    await context.sync();
    return true;
  });
}
