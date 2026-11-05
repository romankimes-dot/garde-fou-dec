// Corrections dans le document ouvert (Office.js), à la demande, avec aperçu ligne à ligne.
//
// Principes, tirés des essais dans Word en ligne :
// - un passage est d'abord retrouvé par son contexte (« lus : » plutôt que « : »), sans espace en tête ni en fin
//   de chaîne et sans caractères génériques (refusés par Word en ligne) ;
// - seul le plus petit passage est remplacé (le guillemet, l'espace et ses deux voisins), et sa mise en forme
//   d'origine (police, taille, gras, italique, couleur…) est réappliquée : Word donne sinon au texte inséré la
//   police par défaut ;
// - un seul remplacement par paragraphe et par passe, puis relecture du texte réel (Word en ligne décale les
//   remplacements groupés d'un même paragraphe) ;
// - chaque remplacement est consigné dans un journal, pour « Annuler la dernière correction » en un clic
//   (Ctrl+Z, dans Word, annule paragraphe par paragraphe).

const NBSP = ' ';
const ACCENTS = { Ere: 'Ère' }; // sinon : E initial → É
const PROPRIETES = ['name', 'size', 'bold', 'italic', 'underline', 'color', 'highlightColor', 'superscript', 'subscript', 'strikeThrough'];

export const LIBELLES = {
  insecables: 'espaces insécables avant : ; ! ? et dans « »',
  guillemets: 'guillemets droits " " remplacés par « »',
  apostrophes: "apostrophes droites ' remplacées par ’",
  espaces: 'espaces doublés et espace avant la virgule',
  unites: 'espace insécable avant % et €',
  accents: 'majuscules accentuées',
  'numero-simple': 'numéro simple (« Section 1.2 » → « Section 2 »)',
  'sans-numero': 'numéro 1.1 retiré du sous-titre',
  renommer: 'intitulé exact de la Note',
  'style-normal': 'style Normal en taille 12',
  tables: 'mise à jour des tables (sommaire, table des matières, liste des annexes)',
};

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

/**
 * Remplacement décrit par son contexte : `cherche` (texte à retrouver), `debut` et `longueur` du passage changé
 * dans ce contexte, `nouveau` (texte qui le remplace). Le contexte n'a ni espace en tête ni en fin.
 */
function contexte(t, debut, longueur, nouveau) {
  let a = debut;
  let b = debut + longueur;
  for (let k = 0; k < 3 && a > 0; k++) a--;
  while (a > 0 && /\s/.test(t[a])) a--;
  for (let k = 0; k < 3 && b < t.length; k++) b++;
  while (b < t.length && /\s/.test(t[b - 1])) b++;
  while (a < debut && /\s/.test(t[a])) a++;
  while (b > debut + longueur && /\s/.test(t[b - 1])) b--;
  const cherche = t.slice(a, b);
  if (/^\s|\s$|\^/.test(cherche) || cherche.length > 250 || cherche.length < 2) return null;
  const r = { cherche, debut: debut - a, longueur, nouveau };
  return apres(r) === cherche ? null : r;
}
const apres = (r) => r.cherche.slice(0, r.debut) + r.nouveau + r.cherche.slice(r.debut + r.longueur);

function parContexte(t, motif, transformer) {
  const liste = [];
  for (const m of t.matchAll(motif)) {
    const r = contexte(t, m.index, m[0].length, transformer(m));
    if (r) liste.push(r);
  }
  return liste;
}

/** Remplacements à faire dans un texte, pour une correction ; le premier est appliqué, puis le texte est relu. */
function remplacements(c, t, parametres) {
  switch (c.type) {
    case 'insecables':
      return [
        ...parContexte(t, / (?=[:;!?])/g, () => NBSP),
        ...parContexte(t, /(?<=«) /g, () => NBSP),
        ...parContexte(t, / (?=»)/g, () => NBSP),
      ];
    case 'guillemets': {
      // Le premier guillemet droit restant : fermant si un « est encore ouvert avant lui, ouvrant sinon.
      const i = t.indexOf('"');
      if (i < 0) return [];
      const avant = t.slice(0, i);
      const ouvert = (avant.match(/«/g) || []).length > (avant.match(/»/g) || []).length;
      const r = contexte(t, i, 1, ouvert ? `${NBSP}»` : `«${NBSP}`);
      return r ? [r] : [];
    }
    case 'apostrophes':
      return parContexte(t, /(?<=\p{L})'(?=\p{L})/gu, () => '’');
    case 'espaces':
      return [...parContexte(t, /(?<= ) +/g, () => ''), ...parContexte(t, / (?=,)/g, () => '')];
    case 'unites':
      return [...parContexte(t, /(?<=\d) (?=[%€])/g, () => NBSP), ...parContexte(t, /(?<=\d)(?=[%€])/g, () => NBSP)];
    case 'accents': {
      const mots = parametres?.majuscules_accentuees || [];
      if (!mots.length) return [];
      const re = new RegExp(`(?<!\\p{L})(${mots.join('|')})(?!\\p{L})`, 'gu');
      return [...t.matchAll(re)].map((m) => contexte(t, m.index, 1, (ACCENTS[m[1]] || 'É').slice(0, 1))).filter(Boolean);
    }
    case 'numero-simple': {
      const m = /(partie|chapitre|section|paragraphe|§)\s*(\d+(?:\.\d+)+)/i.exec(t);
      if (!m) return [];
      const debut = m.index + m[0].length - m[2].length;
      const r = contexte(t, debut, m[2].length - m[2].split('.').pop().length, '');
      return r ? [r] : [];
    }
    case 'sans-numero': {
      const m = /^(\s*)(\d+(?:\.\d+)+\.?\s+)/.exec(t);
      if (!m) return [];
      const r = contexte(t, m[1].length, m[2].length, '');
      return r ? [r] : [];
    }
    default:
      return [];
  }
}

/** Texte après correction (aperçu), ou null si rien ne change ou si le cas est ambigu. */
export function apercu(c, bloc, parametres = null) {
  const t = bloc.texte;
  if (c.type === 'renommer') return c.texte;
  if (c.type === 'style-normal') return null;
  if (c.type === 'guillemets' && !pairesGuillemets(t)) return null;
  let courant = t;
  for (let k = 0; k < 200; k++) {
    const r = remplacements(c, courant, parametres)[0];
    if (!r) break;
    const i = courant.indexOf(r.cherche);
    if (i < 0) break;
    courant = courant.slice(0, i) + apres(r) + courant.slice(i + r.cherche.length);
  }
  return courant === t ? null : courant;
}

// Comparaison tolérante (espaces, insécables, traits d'union conditionnels, caractères de contrôle) : sert à repérer
// un paragraphe réellement réécrit depuis le contrôle, pas une différence de lecture.
// Texte entier comparé : un mot ajouté n'importe où dans le paragraphe suffit à le laisser.
const comparable = (t) => (t || '').replace(/[\s  ­​\u0007\u000b\r]/g, '');
const nettoyer = (t) => (t || '').replace(/[\r\u0007\u000b]/g, ' ');

/** Plus petit passage à remplacer dans le contexte : le caractère changé, ou l'espace et ses deux voisins. */
function petitPassage(r) {
  const change = r.cherche.slice(r.debut, r.debut + r.longueur);
  if (change.length === 1 && /\S/.test(change)) return { texte: change, avant: r.cherche.slice(0, r.debut).split(change).length - 1, nouveau: r.nouveau };
  let a = r.debut;
  let b = r.debut + r.longueur;
  if (a > 0) a--;
  if (b < r.cherche.length) b++;
  while (a > 0 && /\s/.test(r.cherche[a])) a--;
  while (b < r.cherche.length && /\s/.test(r.cherche[b - 1])) b++;
  const texte = r.cherche.slice(a, b);
  return { texte, avant: r.cherche.slice(0, a).split(texte).length - 1, nouveau: texte.slice(0, r.debut - a) + r.nouveau + texte.slice(r.debut - a + r.longueur) };
}

/** Passage laissé tel quel, pour le retrouver depuis le volet. */
function noter(compte, t, motif) {
  if (!compte.lieux) return;
  const p = t.p || petitPassage(t.r);
  compte.lieux.push({ i: t.i, motif, passage: [{ texte: t.r.cherche, rang: 0 }, { texte: p.texte, rang: p.avant }] });
}

/**
 * Applique un remplacement par paragraphe, en une passe (trois allers-retours avec Word pour tous les paragraphes).
 * travaux : [{ para, r, i }]. Rend la liste des remplacements faits, à consigner au journal.
 */
async function unePasse(context, travaux, compte) {
  const etape1 = travaux.map((t) => {
    const res = t.para.search(t.r.cherche, { matchCase: true });
    res.load('items');
    return { ...t, res };
  });
  await context.sync();
  const etape2 = [];
  for (const t of etape1) {
    if (!t.res.items.length) { compte.introuvables++; noter(compte, t, 'introuvable'); continue; }
    const p = petitPassage(t.r);
    const sous = t.res.items[0].search(p.texte, { matchCase: true });
    sous.load(['items/text', ...PROPRIETES.map((x) => `items/font/${x}`)].join(','));
    etape2.push({ ...t, p, sous });
  }
  await context.sync();
  const faits = [];
  for (const t of etape2) {
    // Word rapproche « " » de « “ » et « ' » de « ’ » : seules les occurrences exactes comptent.
    const exacts = t.sous.items.filter((x) => x.text === t.p.texte);
    const cible = exacts[t.p.avant] || exacts[0];
    if (!cible) { compte.introuvables++; noter(compte, t, 'introuvable'); continue; }
    const police = {};
    for (const x of PROPRIETES) police[x] = cible.font[x];
    // Passage de plus d'un caractère en mise en forme mixte (une partie en italique…) : laissé tel quel.
    if (t.p.texte.length > 1 && ['name', 'size', 'italic', 'bold'].some((x) => police[x] === null || police[x] === undefined)) {
      compte.mixtes++;
      noter(compte, t, 'mise en forme mixte');
      continue;
    }
    const nouveau = cible.insertText(t.p.nouveau || '', 'Replace');
    if (t.p.nouveau) for (const x of PROPRIETES) if (police[x] !== null && police[x] !== undefined) nouveau.font[x] = police[x];
    compte.corrections++;
    faits.push({ i: t.i, r: t.r, police });
  }
  await context.sync();
  return faits;
}

/**
 * Applique des corrections [{ type, i, texte?, style? }] (i : numéro de body.paragraphs).
 * Rend { corrections, paragraphes, decales, dejaCorriges, mixtes, introuvables, ignorees, erreurs, journal }.
 */
export async function corriger(liste, doc, parametres) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load('items/text');
    await context.sync();
    const compte = { corrections: 0, decales: 0, dejaCorriges: 0, mixtes: 0, introuvables: 0, ignorees: 0, paragraphes: 0, lieux: [] };
    const exemples = [];
    const erreurs = [];
    const journal = [];
    const cibles = [];
    const vus = new Set();
    for (const c of liste) {
      const cle = `${c.type}|${c.i}`;
      if (vus.has(cle)) continue;
      vus.add(cle);
      if (c.type === 'style-normal' || c.type === 'tables') { cibles.push({ c }); continue; }
      const para = paras.items[c.i];
      const bloc = doc.blocs[c.i];
      if (!para || !bloc || comparable(para.text) !== comparable(bloc.texte)) {
        const attendu = bloc ? apercu(c, bloc, parametres) : null;
        if (para && attendu !== null && comparable(para.text) === comparable(attendu)) { compte.dejaCorriges++; continue; }
        compte.decales++;
        compte.lieux.push({ i: c.i, motif: 'modifié depuis le contrôle', passage: null });
        if (exemples.length < 3) exemples.push(`¶ ${c.i + 1} : ${para ? `${para.text.length} caractères lus contre ${bloc ? bloc.texte.length : '?'}` : 'introuvable'}`);
        continue;
      }
      if (c.type === 'guillemets' && !pairesGuillemets(bloc.texte)) { compte.ignorees++; continue; }
      cibles.push({ c, para, bloc });
    }
    compte.paragraphes = new Set(cibles.filter((x) => x.para).map((x) => x.c.i)).size;

    // Intitulés renommés : paragraphe réécrit, mise en forme du paragraphe réappliquée.
    try {
      const entiers = cibles.filter((x) => x.c.type === 'renommer');
      entiers.forEach((x) => x.para.load(PROPRIETES.map((p) => `font/${p}`).join(',')));
      if (entiers.length) await context.sync();
      for (const { c, para, bloc } of entiers) {
        const police = {};
        for (const x of PROPRIETES) police[x] = para.font[x];
        const nouveau = para.insertText(c.texte, 'Replace');
        for (const x of PROPRIETES) if (police[x] !== null && police[x] !== undefined) nouveau.font[x] = police[x];
        if (c.style) para.styleBuiltIn = c.style;
        journal.push({ i: c.i, entier: { avant: bloc.texte, police } });
        compte.corrections++;
      }
      await context.sync();
    } catch (e) {
      erreurs.push(`renommer : ${e.message}`);
    }

    if (cibles.some((x) => x.c.type === 'style-normal')) {
      try {
        await styleNormal(context, parametres.taille_texte, journal, compte);
      } catch (e) {
        erreurs.push(`style Normal : ${e.message}`);
      }
    }

    if (cibles.some((x) => x.c.type === 'tables')) {
      try {
        compte.tables = await mettreAJourTables(context);
        compte.corrections += compte.tables;
      } catch (e) {
        erreurs.push(`tables : ${e.message}`);
      }
    }

    // Remplacements dans le texte, type par type, un par paragraphe et par passe.
    const recherches = cibles.filter((x) => x.para && x.c.type !== 'renommer');
    for (const type of new Set(recherches.map((x) => x.c.type))) {
      try {
        let actives = recherches.filter((x) => x.c.type === type);
        for (let passe = 0; passe < 120 && actives.length; passe++) {
          // Texte relu à chaque passe : un type précédent a pu modifier le même paragraphe.
          actives.forEach((x) => x.para.load('text'));
          await context.sync();
          const travaux = [];
          for (const x of actives) {
            const texte = nettoyer(x.para.text);
            const r = remplacements(x.c, texte, parametres)[0];
            if (r) travaux.push({ ...x, i: x.c.i, r });
          }
          if (!travaux.length) break;
          const faits = await unePasse(context, travaux, compte);
          journal.push(...faits);
          const encore = new Set(faits.map((f) => f.i));
          actives = actives.filter((x) => encore.has(x.c.i));
        }
      } catch (e) {
        erreurs.push(`${type} : ${e.message}${e.debugInfo?.errorLocation ? ` (${e.debugInfo.errorLocation})` : ''}`);
      }
    }
    return { ...compte, exemples, erreurs, journal };
  });
}

/**
 * Style Normal à la taille du texte. Word en ligne ne donne pas toujours la taille d'un style (« propriété non
 * disponible ») : la taille d'origine est lue sur les paragraphes. Si le style ne peut pas être changé, les
 * paragraphes de style Normal sont mis à la taille un par un.
 */
async function styleNormal(context, taille, journal, compte) {
  const paras = context.document.body.paragraphs;
  paras.load('items/styleBuiltIn,items/style,items/font/size');
  await context.sync();
  const normaux = paras.items.map((p, i) => ({ p, i })).filter(({ p }) => p.styleBuiltIn === 'Normal' || /^(normal|standard)$/i.test(p.style || ''));
  const tailles = normaux.map(({ p }) => p.font.size).filter((x) => typeof x === 'number' && x !== taille);
  const frequence = (x) => tailles.filter((y) => y === x).length;
  const ancienne = tailles.length ? tailles.reduce((a, b) => (frequence(b) > frequence(a) ? b : a)) : null;
  try {
    const normal = context.document.getStyles().getByNameOrNullObject('Normal');
    normal.load('isNullObject');
    await context.sync();
    if (normal.isNullObject) throw new Error('style Normal introuvable');
    normal.font.size = taille;
    await context.sync();
    journal.push({ style: { nom: 'Normal', taille: ancienne } });
    compte.corrections++;
  } catch {
    // Repli : taille posée sur chaque paragraphe de style Normal qui n'est pas à la bonne taille.
    for (const { p, i } of normaux) {
      if (typeof p.font.size === 'number' && p.font.size !== taille) {
        journal.push({ i, taille: p.font.size });
        p.font.size = taille;
        compte.corrections++;
      }
    }
    await context.sync();
  }
}

/**
 * Met à jour les tables (champs TOC) : numéros de page et entrées recalculés par Word (WordApi 1.5).
 * Rend le nombre de tables mises à jour.
 */
async function mettreAJourTables(context) {
  const champs = context.document.body.fields;
  if (!champs || typeof Office === 'undefined' || !Office.context.requirements.isSetSupported('WordApi', '1.5')) {
    throw new Error('non disponible dans cette version de Word : clic droit sur la table, « Mettre à jour la table »');
  }
  champs.load('items/code');
  await context.sync();
  const tables = champs.items.filter((f) => /^\s*TOC\b/.test(f.code || ''));
  tables.forEach((f) => f.updateResult());
  await context.sync();
  return tables.length;
}

/** Annule les remplacements d'un journal, du dernier au premier, mise en forme comprise. */
export async function annuler(journal) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load('items/text');
    await context.sync();
    const compte = { corrections: 0, introuvables: 0, mixtes: 0 };
    const erreurs = [];
    for (const f of [...journal].reverse()) {
      try {
        if (f.style) {
          const st = context.document.getStyles().getByNameOrNullObject(f.style.nom);
          st.load('isNullObject');
          await context.sync();
          if (!st.isNullObject && f.style.taille) { st.font.size = f.style.taille; await context.sync(); compte.corrections++; }
          continue;
        }
        const para = paras.items[f.i];
        if (!para) { compte.introuvables++; continue; }
        if (f.taille) {
          para.font.size = f.taille;
          await context.sync();
          compte.corrections++;
          continue;
        }
        if (f.entier) {
          const nouveau = para.insertText(f.entier.avant, 'Replace');
          for (const x of PROPRIETES) if (f.entier.police[x] !== null && f.entier.police[x] !== undefined) nouveau.font[x] = f.entier.police[x];
          await context.sync();
          compte.corrections++;
          continue;
        }
        // Remplacement inverse : retrouver le contexte corrigé et y remettre le passage d'origine.
        const inverse = { cherche: apres(f.r), debut: f.r.debut, longueur: f.r.nouveau.length, nouveau: f.r.cherche.slice(f.r.debut, f.r.debut + f.r.longueur) };
        if (!inverse.cherche.trim()) { compte.introuvables++; continue; }
        await unePasse(context, [{ para, i: f.i, r: inverse }], compte);
      } catch (e) {
        erreurs.push(e.message);
      }
    }
    return { ...compte, erreurs };
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

/** Empreinte du texte du document : sert à savoir s'il a changé depuis le dernier contrôle. */
export async function empreinte() {
  return Word.run(async (context) => {
    const corps = context.document.body;
    corps.load('text');
    await context.sync();
    return signature(corps.text);
  });
}

export function signature(texte) {
  let h = 0;
  const t = (texte || '').replace(/\s+/g, ' ');
  for (let k = 0; k < t.length; k++) h = (h * 31 + t.charCodeAt(k)) | 0;
  return `${t.length}:${h}`;
}
