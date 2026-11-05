// Volet du complément Word : lit le document ouvert (Office.js) ou un .docx, lance les contrôles, affiche le résultat.

import { lireDocx } from './sources/docx.js';
import { lireWord, selectionnerParagraphe } from './sources/office.js';
import { controler } from './moteur/moteur.js';
import { pages } from './moteur/texte.js';

const $ = (id) => document.getElementById(id);
const CATEGORIES = { A: 'Conformité à la Note du jury', B: 'Structure', F: 'Propreté' };
const VERSION = '0.2.0';

const etat = {
  profil: 'notice', mode: 'redaction', termes: [], ignores: {},
  regles: null, doc: null, cleDoc: null, dansWord: false, nomDoc: '',
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
      if (etat.doc) afficher();
    }));
  }
  $('termes').value = etat.termes.join('\n');
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
  $('copier-diagnostic').addEventListener('click', async () => {
    const d = { complement: VERSION, source: etat.doc?.source || null, ...(etat.doc?.diagnostic || {}), erreur: etat.erreur || null };
    const texte = JSON.stringify(d, null, 1);
    $('diagnostic').textContent = texte;
    $('diagnostic').hidden = false;
    try { await navigator.clipboard.writeText(texte); statut('Diagnostic copié (sans texte du document).'); } catch {
      statut('Diagnostic affiché ci-dessous : le sélectionner et le copier.');
    }
  });
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

  const d = doc.diagnostic || {};
  const temps = duree ? ` en ${(duree / 1000).toLocaleString('fr-FR', { maximumFractionDigits: 1 })} s` : '';
  statut(`${etat.nomDoc} · ${doc.source === 'docx' ? 'fichier .docx' : 'document ouvert'} · ${d.paragraphes} paragraphes lus${temps} · ${new Date().toLocaleTimeString('fr-FR')}.`);

  const feux = { rouge: 'Points bloquants', orange: 'Points à vérifier', vert: 'Aucun écart trouvé' };
  $('synthese').hidden = false;
  $('feu').className = `feu ${r.feu}`;
  $('feu-texte').textContent = feux[r.feu];
  $('comptes').textContent = `${r.comptes.bloquants} bloquant(s) · ${r.comptes.avertissements} à vérifier · ${r.comptes.ignores} ignoré(s) · ${etat.profil === 'notice' ? 'notice' : 'mémoire'}, mode ${etat.mode === 'depot' ? 'dépôt' : 'rédaction'}`;

  $('infos').hidden = !r.infos.length;
  $('infos').replaceChildren(...r.infos.map((x) => el('p', {}, x.texte)));

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

  $('pagination-bloc').hidden = !r.pagination.length;
  $('pagination').replaceChildren(...r.pagination.map((l) => el('tr', { class: `n${l.niveau}` },
    el('td', {}, l.intitule), el('td', {}, `p. ${l.page}`), el('td', {}, pages(l.pages)))));
}

function carteRegle(x) {
  const bloquants = x.occurrences.filter((o) => o.niveau === 'bloquant').length;
  const niveau = bloquants ? 'bloquant' : 'avertissement';
  const n = x.occurrences.length;
  const libelle = niveau === 'bloquant' ? `${n} bloquant${n > 1 ? 's' : ''}` : `${n} à vérifier`;
  return el('details', { class: 'regle', open: '' },
    el('summary', {},
      el('span', { class: 'nom' }, `${x.regle.id} ${x.regle.titre}`),
      el('span', { class: `badge ${niveau}` }, libelle)),
    el('p', { class: 'base' }, `${x.regle.nature === 'exigence' ? 'Note du jury' : 'Préférence'} : ${x.regle.base}`),
    el('ul', { class: 'occurrences' }, ...x.occurrences.map((o) => ligneOccurrence(o))));
}

function ligneOccurrence(o) {
  const allerPossible = etat.doc.source === 'word' && o.i !== null;
  const extrait = o.extrait
    ? (allerPossible
      ? el('button', { class: 'extrait', type: 'button', title: 'Sélectionner le passage dans Word', onclick: () => aller(o.i) }, o.extrait)
      : el('span', { class: 'extrait' }, o.i !== null ? `¶ ${o.i + 1} · ${o.extrait}` : o.extrait))
    : null;
  return el('li', { class: `occurrence ${o.niveau}` },
    el('p', { class: 'message' }, o.message),
    extrait,
    el('button', { class: 'ignorer', type: 'button', title: 'Ne plus signaler cette occurrence pour ce document', onclick: () => ignorer(o.cle) }, 'Ignorer'));
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
