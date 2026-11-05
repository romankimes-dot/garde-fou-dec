// Volet du complément Word : lit le document ouvert (Office.js) ou un .docx, lance les contrôles, affiche le résultat.

import { lireDocx } from './sources/docx.js';
import { lireWord, selectionnerParagraphe } from './sources/office.js';
import { corriger, apercu, selectionnerPlage, LIBELLES, AVEC_APERCU } from './sources/correction.js';
import { controler } from './moteur/moteur.js';
import { pages } from './moteur/texte.js';

const $ = (id) => document.getElementById(id);
const CATEGORIES = { A: 'Conformité à la Note du jury', B: 'Structure', F: 'Propreté' };
const VERSION = '0.4.1';

const etat = {
  profil: 'notice', mode: 'redaction', termes: [], ignores: {},
  regles: null, doc: null, cleDoc: null, dansWord: false, nomDoc: '', ouvertes: new Set(),
};

// ---------------------------------------------------------------- stockage local (ce navigateur seulement)

const CLE = 'garde-fou-dec';
const stockage = {
  lire() {
    try { return JSON.parse(localStorage.getItem(CLE)) || {}; } catch { return {}; }
  },
  ecrire(valeurs) {
    try { localStorage.setItem(CLE, JSON.stringify({ ...stockage.lire(), ...valeurs })); } catch { /* non conservé */ }
  },
};

async function demarrer() {
  const info = typeof Office !== 'undefined' ? await Office.onReady() : { host: null };
  etat.dansWord = info.host === 'Word';
  const [regles, consignes] = await Promise.all([
    fetch('regles.json').then((r) => r.json()),
    fetch('consignes.json').then((r) => r.json()).catch(() => []),
  ]);
  regles.consignes = consignes;
  etat.regles = regles;
  const s = stockage.lire();
  Object.assign(etat, {
    profil: s.profil || 'notice', mode: s.mode || 'redaction', termes: s.termes || [], ignores: s.ignores || {},
  });
  for (const nom of ['profil', 'mode']) {
    document.querySelector(`input[name="${nom}"][value="${etat[nom]}"]`).checked = true;
    document.querySelectorAll(`input[name="${nom}"]`).forEach((r) => r.addEventListener('change', () => {
      etat[nom] = r.value;
      stockage.ecrire({ [nom]: r.value });
      if (etat.doc) afficher(); else afficherExigences();
    }));
  }
  $('termes').value = etat.termes.join('\n');
  afficherExigences();
  $('pied').textContent = `Garde-fou ${VERSION}. Règles du ${regles.version}. ${regles.reference}.`;

  $('controler-ouvert').hidden = !etat.dansWord;
  if (!etat.dansWord) statut('Hors de Word : seul le contrôle d\'un fichier .docx est possible.');
  $('controler-ouvert').addEventListener('click', lireOuvert);
  $('fichier').addEventListener('change', (e) => e.target.files[0] && lireFichier(e.target.files[0]));
  document.body.addEventListener('dragover', (e) => { e.preventDefault(); document.body.classList.add('depot-survol'); });
  document.body.addEventListener('dragleave', () => document.body.classList.remove('depot-survol'));
  document.body.addEventListener('drop', (e) => {
    e.preventDefault();
    document.body.classList.remove('depot-survol');
    const f = e.dataTransfer.files[0];
    if (f) lireFichier(f);
  });
  $('enregistrer-termes').addEventListener('click', () => {
    etat.termes = $('termes').value.split('\n').map((t) => t.trim()).filter(Boolean);
    stockage.ecrire({ termes: etat.termes });
    statut(`${etat.termes.length} terme(s) enregistré(s) dans ce navigateur.`);
    if (etat.doc) afficher();
  });
  $('oublier-ignores').addEventListener('click', () => {
    if (!etat.cleDoc) return;
    delete etat.ignores[etat.cleDoc];
    stockage.ecrire({ ignores: etat.ignores });
    if (etat.doc) afficher();
  });
  $('copier-alertes').addEventListener('click', async () => {
    const texte = alertesPourCalibrage();
    if (!texte) { statut('Lancer d\'abord un contrôle.'); return; }
    $('diagnostic').textContent = texte;
    if (await copier(texte)) { statut('Alertes copiées, sans extrait du document : les relire avant de les transmettre.'); return; }
    $('diagnostic').hidden = false;
    $('reglages').open = true;
    selectionner($('diagnostic'));
    statut('Copie refusée par Word : alertes affichées et sélectionnées dans Réglages, faire Ctrl+C (Cmd+C).', true);
  });
  $('copier-diagnostic').addEventListener('click', async () => {
    const d = { complement: VERSION, source: etat.doc?.source || null, ...(etat.doc?.diagnostic || {}), erreur: etat.erreur || null };
    const texte = JSON.stringify(d, null, 1);
    $('diagnostic').textContent = texte;
    $('diagnostic').hidden = false;
    if (await copier(texte)) { statut('Diagnostic copié (sans texte du document).'); return; }
    selectionner($('diagnostic'));
    statut('Copie refusée par Word : diagnostic sélectionné ci-dessous, faire Ctrl+C (Cmd+C).', true);
  });
}

/**
 * Presse-papiers. Dans le volet Word, l'API Clipboard est souvent refusée (ou ne répond jamais) :
 * copie d'abord par execCommand, synchrone dans le clic, puis par l'API avec un délai maximal.
 */
async function copier(texte) {
  const zone = document.createElement('textarea');
  zone.value = texte;
  zone.setAttribute('readonly', '');
  zone.style.cssText = 'position:fixed;top:0;left:0;width:1px;height:1px;opacity:0;';
  document.body.append(zone);
  zone.focus();
  zone.select();
  let ok = false;
  try { ok = document.execCommand('copy'); } catch { ok = false; }
  zone.remove();
  if (ok) return true;
  if (!navigator.clipboard?.writeText) return false;
  try {
    await Promise.race([navigator.clipboard.writeText(texte), new Promise((_, non) => setTimeout(() => non(new Error('délai')), 1500))]);
    return true;
  } catch { return false; }
}

function selectionner(element) {
  const plage = document.createRange();
  plage.selectNodeContents(element);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(plage);
  element.scrollIntoView({ block: 'nearest' });
}

/**
 * Alertes pour calibrage : règle, message, et forme des passages en cause (mots remplacés par « … », numéros,
 * ponctuation et mots du plan conservés). Messages identiques regroupés. Termes du cabinet et restes de travail masqués.
 */
const MOTS_GARDES = new Set(['partie', 'chapitre', 'section', 'paragraphe', 'introduction', 'conclusion', 'annexe', 'p', 'page', 'pages',
  'sommaire', 'notice', 'plan', 'bibliographie', 'synthese', 'outil', 'out']);
const SANS_FORME = new Set(['A15', 'A16', 'B08', 'F01']);

/** Vocabulaire de la Note (titres des sections de la notice et des catégories de la bibliographie) : laissé visible. */
function motsDeLaNote() {
  if (motsDeLaNote.cache) return motsDeLaNote.cache;
  const p = etat.regles?.parametres || {};
  const titres = [...(p.notice_sections || []), ...(p.bibliographie_categories || [])].map((x) => x.titre).join(' ');
  motsDeLaNote.cache = new Set((titres.match(/\p{L}{3,}/gu) || []).map(sansAccent));
  return motsDeLaNote.cache;
}
const sansAccent = (w) => w.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

function forme(i) {
  const b = i !== null && i !== undefined ? etat.doc.blocs[i] : null;
  if (!b) return '';
  if (b.garde) return '(page de garde)';
  const t = ((b.label || '') + b.texte).replace(/\s+/g, ' ').trim()
    .replace(/\p{L}+/gu, (w) => (w.length === 1 || /^[IVXLC]+$/.test(w) || MOTS_GARDES.has(sansAccent(w)) || motsDeLaNote().has(sansAccent(w)) ? w : '…'))
    .replace(/…(?:[\s'’-]*…)+/g, '…')
    .replace(/\d{4,}|\d[\d\s.]{4,}\d/g, '#');
  return t.length > 70 ? `${t.slice(0, 45)} […] ${t.slice(-20)}` : t;
}

function alertesPourCalibrage() {
  if (!etat.dernier) return null;
  const r = etat.dernier;
  const lignes = [`Garde-fou ${VERSION} · ${r.profil} · ${r.mode} · ${r.source} · ${etat.doc.diagnostic?.paragraphes} paragraphes · ${r.comptes.bloquants} bloquant(s), ${r.comptes.avertissements} à vérifier`];
  for (const x of r.resultats) {
    if (!x.occurrences.length) continue;
    lignes.push('', `${x.regle.id} ${x.regle.titre} (${x.occurrences.length})`);
    const groupes = new Map();
    for (const o of x.occurrences) {
      let m = o.message;
      if (x.regle.id === 'A16' || x.regle.id === 'F01') m = m.replace(/«[^»]*»/g, '« … »');
      const cle = `${o.niveau === 'bloquant' ? '■' : '▲'} ${m}`;
      if (!groupes.has(cle)) groupes.set(cle, []);
      groupes.get(cle).push(o);
    }
    for (const [cle, liste] of groupes) {
      lignes.push(`  ${cle}${liste.length > 1 ? ` (×${liste.length})` : ''}`);
      if (SANS_FORME.has(x.regle.id)) continue;
      for (const o of liste.slice(0, 3)) {
        const f = forme(o.i);
        if (f) lignes.push(`      ¶ ${o.i + 1} · ${f}`);
      }
      if (liste.length > 3) lignes.push('      …');
    }
  }
  if (r.infos.length) {
    lignes.push('', 'Informations');
    for (const x of r.infos) lignes.push(`  ${x.regle} · ${x.regle === 'A16' ? x.texte.replace(/«[^»]*»/g, '« … »') : x.texte}`);
  }
  if (r.titres?.length) {
    lignes.push('', `Structure lue (${r.titres.length} titres) : ¶ · niveau · rôle · pièce · forme`);
    for (const t of r.titres.slice(0, 120)) {
      lignes.push(`  ¶ ${t.i + 1} · n${t.niveau} · ${t.role || '-'}${t.section ? ` ${t.section}` : ''} · ${t.zone || '-'} · ${forme(t.i)}`);
    }
    if (r.titres.length > 120) lignes.push('  …');
  }
  return lignes.join('\n');
}

function statut(texte, erreur = false) {
  $('statut').textContent = texte;
  $('statut').classList.toggle('erreur', erreur);
}

function nomDuDocument() {
  try {
    const url = Office.context.document.url || '';
    const nom = decodeURIComponent(url.split(/[\\/]/).pop().split('?')[0]);
    return nom || 'Document Word';
  } catch { return 'Document Word'; }
}

// ---------------------------------------------------------------- lecture

async function lireOuvert() {
  statut('Lecture du document…');
  const debut = performance.now();
  try {
    etat.doc = await lireWord();
    etat.erreur = null;
    etat.nomDoc = nomDuDocument();
    if (etat.cleDoc !== 'w:' + etat.nomDoc) etat.ouvertes.clear();
    etat.cleDoc = 'w:' + etat.nomDoc;
    afficher(performance.now() - debut);
  } catch (e) {
    etat.erreur = `${e.name || 'Erreur'} : ${e.message}${e.code ? ` (${e.code})` : ''}`;
    statut(`Lecture impossible : ${etat.erreur}.${etat.doc ? ' Les résultats ci-dessous sont ceux de la lecture précédente.' : ''}`, true);
  }
}

async function lireFichier(f) {
  statut(`Lecture de ${f.name}…`);
  const debut = performance.now();
  try {
    etat.doc = await lireDocx(await f.arrayBuffer());
    etat.erreur = null;
    etat.nomDoc = f.name;
    if (etat.cleDoc !== 'f:' + f.name) etat.ouvertes.clear();
    etat.cleDoc = 'f:' + f.name;
    afficher(performance.now() - debut);
  } catch (e) {
    etat.erreur = e.message;
    statut(e.message, true);
  } finally {
    $('fichier').value = '';
  }
}

// ---------------------------------------------------------------- affichage

function el(balise, attrs = {}, ...enfants) {
  const e = document.createElement(balise);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v);
  }
  for (const x of enfants) if (x !== null && x !== undefined) e.append(x);
  return e;
}

function afficher(duree) {
  const doc = structuredClone(etat.doc);
  const ignores = new Set(etat.ignores[etat.cleDoc] || []);
  const r = controler(doc, etat.regles, { profil: etat.profil, mode: etat.mode, termes: etat.termes, ignores });
  etat.dernier = r;

  const d = doc.diagnostic || {};
  const temps = duree ? ` en ${(duree / 1000).toLocaleString('fr-FR', { maximumFractionDigits: 1 })} s` : '';
  statut(`${etat.nomDoc} · ${doc.source === 'docx' ? 'fichier .docx' : 'document ouvert'} · ${d.paragraphes} paragraphes lus${temps} · ${new Date().toLocaleTimeString('fr-FR')}.`);

  const feux = { rouge: 'Points bloquants', orange: 'Points à vérifier', vert: 'Aucun écart trouvé' };
  $('synthese').hidden = false;
  $('feu').className = `feu ${r.feu}`;
  $('feu-texte').textContent = feux[r.feu];
  $('comptes').textContent = `${r.comptes.bloquants} bloquant(s) · ${r.comptes.avertissements} à vérifier · ${r.comptes.ignores} ignoré(s) · ${etat.profil === 'notice' ? 'notice' : 'mémoire'}, mode ${etat.mode === 'depot' ? 'dépôt' : 'rédaction'}`;

  $('infos').hidden = !r.infos.length;
  const choix = r.infos.filter((x) => x.regle === '-');
  $('infos').replaceChildren(...choix.map((x) => el('p', { class: 'alerte-profil' }, x.texte)),
    el('details', {}, el('summary', {}, `Mesures et informations (${r.infos.length - choix.length})`),
      ...r.infos.filter((x) => x.regle !== '-').map((x) => el('p', {}, `${x.regle} · ${x.texte}`))));

  const conformes = [];
  const blocs = [];
  for (const [cat, nomCat] of Object.entries(CATEGORIES)) {
    const res = r.resultats.filter((x) => x.regle.categorie === cat);
    const avec = res.filter((x) => x.occurrences.length);
    res.filter((x) => !x.occurrences.length).forEach((x) => conformes.push(x));
    if (!avec.length) continue;
    blocs.push(el('h2', { class: 'categorie' }, `${cat}. ${nomCat}`));
    for (const x of avec) blocs.push(carteRegle(x));
  }
  if (conformes.length) {
    blocs.push(el('details', { class: 'conformes' },
      el('summary', {}, `Contrôles sans écart ou hors de portée (${conformes.length})`),
      el('ul', {}, ...conformes.map((x) => el('li', {}, `${x.regle.id} ${x.regle.titre}${x.horsSource ? ' : sur le .docx seulement' : ''}`)))));
  }
  $('resultats').replaceChildren(...blocs);

  afficherExigences();
  const aVerifier = etat.regles.exigences.filter((e) => e.profils.includes(etat.profil) && (!e.regles.length || e.manuel)).length;
  // (exigences non contrôlables ou contrôlées en partie)
  $('rappel').textContent = `${aVerifier} exigence(s) de la Note à vérifier aussi vous-même : voir « Exigences de la Note » plus bas.`;

  $('pagination-bloc').hidden = !r.pagination.length;
  $('pagination').replaceChildren(...r.pagination.map((l) => el('tr', { class: `n${l.niveau}` },
    el('td', {}, l.intitule), el('td', {}, `p. ${l.page}`), el('td', {}, pages(l.pages)))));
}

function carteRegle(x) {
  const bloquants = x.occurrences.filter((o) => o.niveau === 'bloquant').length;
  const niveau = bloquants ? 'bloquant' : 'avertissement';
  const n = x.occurrences.length;
  const libelle = niveau === 'bloquant' ? `${n} bloquant${n > 1 ? 's' : ''}` : `${n} à vérifier`;
  const carte = el('details', { class: 'regle' },
    el('summary', {},
      el('span', { class: 'nom' }, `${x.regle.id} ${x.regle.titre}`),
      el('span', { class: `badge ${niveau}` }, libelle)),
    el('p', { class: 'base' }, `${x.regle.nature === 'exigence' ? 'Note du jury' : 'Préférence'} : ${x.regle.base}`),
    x.regle.corriger ? el('p', { class: 'conseil' }, el('strong', {}, 'Comment corriger : '), x.regle.corriger) : null,
    blocAuto(x),
    el('ul', { class: 'occurrences' }, ...x.occurrences.map((o) => ligneOccurrence(o, x))));
  // Repliée par défaut ; l'état ouvert survit au réaffichage (Ignorer, changement de mode).
  carte.open = etat.ouvertes.has(x.regle.id);
  carte.addEventListener('toggle', () => {
    if (carte.open) etat.ouvertes.add(x.regle.id); else etat.ouvertes.delete(x.regle.id);
  });
  return carte;
}

const enWord = () => etat.doc?.source === 'word' && etat.dansWord;

/** Aperçu « avant → après » d'une liste de corrections, puis application sur confirmation. */
function apercuCorrections(liste, conteneur) {
  const lignes = liste.map((c) => {
    const b = etat.doc.blocs[c.i];
    const apres = b ? apercu(c, b) : null;
    return el('li', {}, el('span', { class: 'avant' }, b ? b.texte.trim() : ''),
      apres === null ? el('span', { class: 'discret' }, 'laissé tel quel (cas ambigu)') : el('span', { class: 'apres' }, `→ ${apres.trim()}`));
  });
  const applicables = liste.filter((c) => etat.doc.blocs[c.i] && apercu(c, etat.doc.blocs[c.i]) !== null);
  conteneur.replaceChildren(el('div', { class: 'apercu' },
    el('p', { class: 'discret' }, 'Aperçu : vérifier chaque ligne avant d\'appliquer (Ctrl+Z dans Word pour annuler ensuite).'),
    el('ul', {}, ...lignes),
    el('span', { class: 'gestes' },
      el('button', { type: 'button', class: 'auto arme', onclick: () => appliquer(applicables) }, `Appliquer (${applicables.length})`),
      el('button', { type: 'button', class: 'ignorer', onclick: () => conteneur.replaceChildren() }, 'Annuler'))));
}

/** Bouton de correction : aperçu pour les changements de texte, confirmation en deux temps sinon. */
function boutonCorriger(liste, libelle, conteneur) {
  const types = new Set(liste.map((c) => c.type));
  const bouton = el('button', { class: 'auto', type: 'button' }, libelle);
  let arme = false;
  bouton.addEventListener('click', () => {
    if ([...types].some((t) => AVEC_APERCU.has(t))) { apercuCorrections(liste, conteneur); return; }
    if (!arme) { arme = true; bouton.textContent = 'Confirmer : modifier le document (Ctrl+Z pour annuler)'; bouton.classList.add('arme'); return; }
    appliquer(liste);
  });
  return bouton;
}

/** Corrections d'une règle, regroupées par type. Hors de Word : simple indication. */
function blocAuto(x) {
  const toutes = x.occurrences.flatMap((o) => o.corrections || []);
  if (!toutes.length) return null;
  const parType = new Map();
  for (const c of toutes) parType.set(c.type, [...(parType.get(c.type) || []), c]);
  if (!enWord()) {
    return el('p', { class: 'discret auto' }, `Correction possible quand le document est ouvert dans Word : ${[...parType.keys()].map((t) => LIBELLES[t]).join(' ; ')}.`);
  }
  const apercuZone = el('div', {});
  const boutons = [...parType].map(([type, liste]) => boutonCorriger(liste,
    `Corriger : ${LIBELLES[type]}${type === 'style-normal' ? '' : ` (${liste.length})`}`, apercuZone));
  return el('div', { class: 'bloc-auto' }, ...boutons, apercuZone);
}

async function appliquer(liste) {
  statut('Correction en cours…');
  try {
    const { corrections, decales, ignorees } = await corriger(liste, etat.doc, etat.regles.parametres);
    const msg = `${corrections} correction(s) appliquée(s) dans le document (Ctrl+Z dans Word pour annuler)`
      + `${decales ? ` ; ${decales} paragraphe(s) modifié(s) depuis le contrôle, laissé(s) tel(s) quel(s)` : ''}`
      + `${ignorees ? ` ; ${ignorees} cas ambigu(s) laissé(s) tel(s) quel(s)` : ''}. Contrôle relancé.`;
    await lireOuvert();
    statut(msg);
  } catch (e) {
    statut(`Correction impossible : ${e.message}. Document protégé, en lecture seule, ou fonction absente de cette version de Word ?`, true);
  }
}

async function selectionnerPourTri(plage) {
  try {
    if (await selectionnerPlage(plage.debut, plage.fin)) {
      statut('Références sélectionnées : Accueil, Trier (bouton A→Z), Trier par Paragraphes, Texte, Croissant, OK ; puis relancer le contrôle.');
    } else statut('Sélection impossible : relancer le contrôle.', true);
  } catch (e) {
    statut(`Sélection impossible : ${e.message}`, true);
  }
}

function afficherExigences() {
  const r = etat.dernier && etat.dernier.profil === etat.profil ? etat.dernier : null;
  const groupes = new Map();
  for (const e of etat.regles.exigences.filter((x) => x.profils.includes(etat.profil))) {
    let etatEx = 'manuel';
    let libelle = 'À vérifier vous-même';
    if (e.regles.length) {
      const res = r ? r.resultats.filter((x) => e.regles.includes(x.regle.id)) : [];
      const n = res.reduce((k, x) => k + x.occurrences.length, 0);
      const ids = `Règle${e.regles.length > 1 ? 's' : ''} ${e.regles.join(', ')}`;
      const nonControle = r && r.infos.some((x) => e.regles.includes(x.regle) && /non contr[ôo]l/.test(x.texte));
      if (!r) { etatEx = 'auto'; libelle = `Contrôlé (${ids})`; }
      else if (n) { etatEx = 'alerte'; libelle = `${ids} : ${n} alerte(s) au total`; }
      else if (nonControle) { etatEx = 'manuel'; libelle = `${ids} : non contrôlé (pièce introuvable), à vérifier vous-même`; }
      else if (res.length && res.every((x) => x.horsSource)) { etatEx = 'manuel'; libelle = 'Sur le .docx seulement'; }
      else { etatEx = 'ok'; libelle = `${ids} : sans écart`; }
      if (e.manuel && etatEx === 'ok') libelle += ' ; reste à vérifier vous-même';
    }
    if (!groupes.has(e.groupe)) groupes.set(e.groupe, []);
    groupes.get(e.groupe).push(el('li', { class: `exigence ${etatEx}` },
      el('span', { class: 'statut-ex' }, libelle),
      el('span', {}, `${e.texte} (${e.page}).`),
      e.manuel ? el('span', { class: 'discret' }, e.manuel) : null));
  }
  const codes = etat.regles.parametres.codes_sujet;
  const blocs = [];
  for (const [g, items] of groupes) {
    blocs.push(el('li', { class: 'groupe-ex' }, g), ...items);
    if (g === 'Dépôt' && etat.profil === 'notice' && codes) {
      blocs.push(el('li', {}, el('details', {}, el('summary', {}, 'Codes du sujet (2 à 6 au document annexe)'),
        el('p', { class: 'discret' }, 'Disciplines'), el('ul', { class: 'codes' }, ...codes.disciplines.map((c) => el('li', {}, c))),
        el('p', { class: 'discret' }, 'Secteurs et branches'), el('ul', { class: 'codes' }, ...codes.secteurs.map((c) => el('li', {}, c))))));
    }
  }
  $('exigences').replaceChildren(...blocs);
}

function ligneOccurrence(o, x) {
  const allerPossible = etat.doc.source === 'word' && o.i !== null;
  const extrait = o.extrait
    ? (allerPossible
      ? el('button', { class: 'extrait', type: 'button', title: 'Sélectionner le passage dans Word', onclick: () => aller(o.i) }, o.extrait)
      : el('span', { class: 'extrait' }, o.i !== null ? `¶ ${o.i + 1} · ${o.extrait}` : o.extrait))
    : null;
  const zone = el('div', {});
  return el('li', { class: `occurrence ${o.niveau}` },
    el('p', { class: 'message' }, o.message),
    extrait,
    el('span', { class: 'gestes' },
      o.corrections?.length && enWord() ? (() => { const b = boutonCorriger(o.corrections, 'Corriger', zone); b.className = 'ignorer'; return b; })() : null,
      o.plage && enWord() ? el('button', { class: 'ignorer', type: 'button', title: 'Sélectionner les références de la catégorie pour le tri de Word', onclick: () => selectionnerPourTri(o.plage) }, 'Sélectionner pour trier') : null,
      el('button', { class: 'ignorer', type: 'button', title: 'Ne plus signaler cette occurrence pour ce document', onclick: () => ignorer(o.cle) }, 'Ignorer')),
    zone);
}

async function aller(i) {
  try {
    if (!(await selectionnerParagraphe(i))) statut('Passage introuvable : relancer le contrôle.', true);
  } catch (e) {
    statut(`Passage non sélectionné : ${e.message}`, true);
  }
}

function ignorer(cle) {
  const liste = new Set(etat.ignores[etat.cleDoc] || []);
  liste.add(cle);
  etat.ignores[etat.cleDoc] = [...liste];
  stockage.ecrire({ ignores: etat.ignores });
  afficher();
}

demarrer().catch((e) => statut(`Démarrage impossible : ${e.message}`, true));
