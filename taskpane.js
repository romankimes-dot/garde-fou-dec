// Volet du complément Word : lit le document ouvert (Office.js) ou un .docx, lance les contrôles, affiche le résultat.

import { lireDocx } from './sources/docx.js';
import { lireWord, selectionnerParagraphe, ouvrirNouveauDocument } from './sources/office.js';
import { planifier, composer, enBase64 } from './sources/composition.js';
import { corriger, annuler, apercu, selectionnerPlage, empreinte, LIBELLES } from './sources/correction.js';
import { controler } from './moteur/moteur.js';
import { pages } from './moteur/texte.js';

const $ = (id) => document.getElementById(id);
const CATEGORIES = { A: 'Conformité à la Note du jury', B: 'Structure', F: 'Typographie', R: 'Restes de travail' };
const VERSION = '0.10.0';

const etat = {
  profil: 'notice', termes: [], ignores: {}, verifiees: {},
  regles: null, doc: null, cleDoc: null, dansWord: false, nomDoc: '', ouvertes: new Set(), filtre: 'tous', deplies: new Set(),
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
    fetch('regles.json', { cache: 'no-cache' }).then((r) => r.json()),
    fetch('consignes.json', { cache: 'no-cache' }).then((r) => r.json()).catch(() => []),
  ]);
  regles.consignes = consignes;
  etat.regles = regles;
  const s = stockage.lire();
  Object.assign(etat, {
    profil: s.profil || 'notice', termes: s.termes || [], ignores: s.ignores || {}, verifiees: s.verifiees || {},
  });
  for (const nom of ['profil']) {
    const choisi = document.querySelector(`input[name="${nom}"][value="${etat[nom]}"]`)
      || document.querySelector(`input[name="${nom}"]`);
    if (choisi) { choisi.checked = true; etat[nom] = choisi.value; }
    document.querySelectorAll(`input[name="${nom}"]`).forEach((r) => r.addEventListener('change', () => {
      etat[nom] = r.value;
      stockage.ecrire({ [nom]: r.value });
      document.body.className = document.body.className.replace(/\bprofil-\S+/g, '').trim() + ` profil-${etat.profil}`;
      $('resultat-forme').replaceChildren();
      $('creer-forme').hidden = true;
      if (etat.doc) afficher(); else afficherExigences();
    }));
  }
  $('termes').value = etat.termes.join('\n');
  installerMiseEnForme(s.champs || {});
  afficherExigences();
  for (const b of document.querySelectorAll('.onglets [role="tab"]')) b.addEventListener('click', () => ouvrirOnglet(b.id.replace('onglet-', '')));
  for (const b of document.querySelectorAll('[data-filtre]')) {
    b.addEventListener('click', () => {
      etat.filtre = etat.filtre === b.dataset.filtre && b.classList.contains('tuile') ? 'tous' : b.dataset.filtre;
      if (etat.dernier) afficherAlertes(etat.dernier);
    });
  }
  // Retour dans le volet après une modification du document (frappe, Ctrl+Z…) : signaler des alertes périmées.
  const verifierFraicheur = async () => {
    if (!etat.dansWord || etat.doc?.source !== 'word' || !etat.empreinte || etat.enCours) return;
    try {
      const e = await empreinte();
      $('perime').hidden = e === etat.empreinte;
      if (e !== etat.empreinte) etat.journal = null; // annulation par journal impossible sur un texte modifié
    } catch { /* lecture impossible : rien à signaler */ }
  };
  window.addEventListener('focus', verifierFraicheur);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) verifierFraicheur(); });
  $('relancer').addEventListener('click', () => { etat.retour = null; lireOuvert(); });
  $('tout-replier').addEventListener('click', () => { etat.ouvertes.clear(); document.querySelectorAll('.regle[open]').forEach((d) => { d.open = false; }); });
  $('pied').textContent = `Garde-fou ${VERSION}. Règles du ${regles.version}. ${regles.reference}.`;
  $('version').textContent = `v${VERSION}`;
  $('version').title = `Garde-fou ${VERSION} · règles du ${regles.version}`;

  $('controler-ouvert').hidden = !etat.dansWord;
  if (!etat.dansWord) statut('Hors de Word : seul le contrôle d\'un fichier .docx est possible.');
  $('controler-ouvert').addEventListener('click', () => { etat.retour = null; lireOuvert(); });
  $('fichier').addEventListener('change', (e) => { etat.retour = null; if (e.target.files[0]) lireFichier(e.target.files[0]); });
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
  $('voir-ignorees').addEventListener('click', () => { etat.filtre = 'ignorees'; ouvrirOnglet('alertes'); if (etat.dernier) afficherAlertes(etat.dernier); });
  $('oublier-ignores').addEventListener('click', () => {
    if (!etat.cleDoc) return;
    delete etat.ignores[etat.cleDoc];
    stockage.ecrire({ ignores: etat.ignores });
    if (etat.filtre === 'ignorees') etat.filtre = 'tous';
    if (etat.doc) afficher();
  });
  $('copier-alertes').addEventListener('click', async () => {
    const texte = alertesPourCalibrage();
    if (!texte) { statut('Lancer d\'abord un contrôle.'); return; }
    if (await copier(texte)) { statut('Alertes copiées, sans extrait du document : les relire avant de les transmettre.'); return; }
    $('diagnostic').textContent = texte;
    $('diagnostic').hidden = false;
    ouvrirOnglet('reglages');
    selectionner($('diagnostic'));
    statut('Copie refusée par Word : alertes affichées et sélectionnées dans Réglages, faire Ctrl+C (Cmd+C).', true);
  });
  $('copier-diagnostic').addEventListener('click', async () => {
    const d = { complement: VERSION, source: etat.doc?.source || null, ...(etat.doc?.diagnostic || {}), erreur: etat.erreur || null,
      derniereCorrection: etat.dernierRapport || null };
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
// Règles dont les messages citent le document (termes, sigles, restes, titres) : citations masquées.
const CITATIONS_MASQUEES = new Set(['A16', 'F01', 'B08']);
const masquerCitations = (t) => t.replace(/«[^»]*»/g, '« … »');

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
    .replace(/\d{3,}|\d[\d\s.]{4,}\d/g, '#');
  return t.length > 70 ? `${t.slice(0, 45)} […] ${t.slice(-20)}` : t;
}

function alertesPourCalibrage() {
  if (!etat.dernier) return null;
  const r = etat.dernier;
  const lignes = [`Garde-fou ${VERSION} · ${r.profil} · ${r.source} · ${etat.doc.diagnostic?.paragraphes} paragraphes · ${r.comptes.bloquants} bloquant(s), ${r.comptes.avertissements} à vérifier, ${r.comptes.restes} reste(s) de travail`];
  for (const x of r.resultats) {
    if (!x.occurrences.length) continue;
    lignes.push('', `${x.regle.id} ${x.regle.titre} (${x.occurrences.length})`);
    const groupes = new Map();
    for (const o of x.occurrences) {
      let m = o.message;
      if (CITATIONS_MASQUEES.has(x.regle.id)) m = masquerCitations(m);
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
    for (const x of r.infos) lignes.push(`  ${x.regle} · ${x.regle === 'A16' || /titre approchant/.test(x.texte) ? masquerCitations(x.texte) : x.texte}`);
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

function occupe(oui) {
  $('controler-ouvert').classList.toggle('occupe', oui);
  $('controler-ouvert').querySelector('.libelle').textContent = oui ? 'Contrôle en cours…' : 'Contrôler ce document';
}

function ouvrirOnglet(nom) {
  for (const b of document.querySelectorAll('.onglets [role="tab"]')) {
    const actif = b.id === `onglet-${nom}`;
    b.setAttribute('aria-selected', String(actif));
    $(b.getAttribute('aria-controls')).hidden = !actif;
  }
}

async function lireOuvert() {
  statut('Lecture du document…');
  occupe(true);
  const debut = performance.now();
  try {
    etat.doc = await lireWord();
    etat.erreur = null;
    etat.nomDoc = nomDuDocument();
    if (etat.cleDoc !== 'w:' + etat.nomDoc) etat.ouvertes.clear();
    etat.cleDoc = 'w:' + etat.nomDoc;
    afficher(performance.now() - debut);
    $('perime').hidden = true;
    try { etat.empreinte = await empreinte(); } catch { etat.empreinte = null; }
  } catch (e) {
    etat.erreur = `${e.name || 'Erreur'} : ${e.message}${e.code ? ` (${e.code})` : ''}`;
    statut(`Lecture impossible : ${etat.erreur}.${etat.doc ? ' Les résultats affichés sont ceux de la lecture précédente.' : ''}`, true);
    if (etat.retour && !etat.retour.erreur) {
      etat.retour = { ...etat.retour, details: `${etat.retour.details || ''} · Relecture du document impossible : relancer le contrôle avant toute autre correction.` };
      if (etat.dernier) afficherAlertes(etat.dernier);
    }
  } finally {
    occupe(false);
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
  const r = controler(doc, etat.regles, { profil: etat.profil, termes: etat.termes, ignores });
  etat.dernier = r;

  const d = doc.diagnostic || {};
  const temps = duree ? ` en ${(duree / 1000).toLocaleString('fr-FR', { maximumFractionDigits: 1 })} s` : '';
  statut(`${etat.nomDoc} · ${doc.source === 'docx' ? 'fichier .docx' : 'document ouvert'} · ${d.paragraphes} paragraphes${temps} · ${new Date().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' })}`);

  const feux = { rouge: 'Points bloquants', orange: r.comptes.bloquants || r.comptes.avertissements ? 'Points à vérifier' : 'Restes de travail à nettoyer', vert: 'Prêt pour l\'export PDF' };
  $('accueil').hidden = true;
  $('synthese').hidden = false;
  $('feu').className = `feu ${r.feu}`;
  $('feu-texte').textContent = feux[r.feu];
  $('contexte').textContent = etat.profil === 'notice' ? 'Demande d\'agrément' : 'Mémoire';
  $('rappel-restes').hidden = !r.comptes.restes;
  $('rappel-restes').textContent = `À nettoyer avant l'export PDF : ${r.comptes.restes} reste(s) de travail (consignes du gabarit, repères, TODO, commentaires, modifications suivies). Voir →`;
  $('filtre-restes').hidden = !r.comptes.restes;
  $('n-bloquants').textContent = r.comptes.bloquants;
  $('n-avertissements').textContent = r.comptes.avertissements;
  $('n-corrigeables').textContent = r.resultats.reduce((n, x) => n + x.occurrences.filter((o) => o.corrections?.length).length, 0);
  $('bandeau-profil').replaceChildren(...r.infos.filter((x) => x.regle === '-').map((x) => el('p', { class: 'bandeau' }, x.texte)));
  $('compte-ignores').textContent = r.comptes.ignores ? `${r.comptes.ignores} occurrence(s) ignorée(s)` : '';
  $('pied-alertes').hidden = false;
  $('filtres').hidden = false;
  const n = r.comptes.bloquants + r.comptes.avertissements + r.comptes.restes;
  $('compte-alertes').textContent = n || '';
  $('compte-alertes').className = `compteur ${r.comptes.bloquants ? 'bloquant' : n ? 'avertissement' : ''}`;
  afficherAlertes(r);

  // Mesures : informations des contrôles (volumes, bibliographie, mise en page) et pagination estimée.
  const mesures = r.infos.filter((x) => x.regle !== '-');
  $('mesures-vide').hidden = mesures.length > 0;
  $('infos').replaceChildren(...mesures.map((x) => el('li', { class: 'mesure' }, el('span', { class: 'id' }, x.regle), el('span', {}, x.texte))));
  $('pagination-bloc').hidden = !r.pagination.length;
  $('pagination-total').textContent = r.pagination.total
    ? `Environ ${r.pagination.total} pages à l'affichage, page de garde comprise (estimation en lignes ; le PDF fait foi). La colonne « 2 400 car. » donne la mesure de la Note pour les limites de volume.`
    : '';
  $('pagination').replaceChildren(...r.pagination.map((l) => el('tr', { class: `n${l.niveau}` },
    el('td', {}, l.intitule), el('td', {}, `p. ${l.page}`), el('td', {}, pages(l.pages)))));

  afficherExigences();
}

/** Garde une occurrence selon le filtre courant. */
function garder(o) {
  switch (etat.filtre) {
    case 'bloquant': return o.niveau === 'bloquant' && !o.reste;
    case 'avertissement': return o.niveau !== 'bloquant' && !o.reste;
    case 'corrigeable': return !!o.corrections?.length;
    case 'restes': return o.reste;
    case 'ignorees': return true;
    default: return true;
  }
}

function afficherAlertes(r) {
  for (const b of document.querySelectorAll('[data-filtre]')) b.setAttribute('aria-pressed', String(b.dataset.filtre === etat.filtre));
  $('filtre-ignorees').textContent = `Ignorées${r.comptes.ignores ? ` (${r.comptes.ignores})` : ''}`;
  $('filtre-ignorees').hidden = !r.comptes.ignores && etat.filtre !== 'ignorees';
  const retour = etat.retour;
  $('retour').hidden = !retour;
  if (retour) {
    $('retour').className = `retour${retour.erreur ? ' erreur' : ''}`;
    const liens = (retour.lieux || []).length && enWord()
      ? el('span', { class: 'lieux' }, 'Passages laissés : ', ...retour.lieux.slice(0, 12).map((l) => el('button', { type: 'button', class: 'geste', title: l.motif, onclick: () => aller(l.i, l.passage) }, `¶ ${l.i + 1}`)))
      : null;
    $('retour').replaceChildren(...[el('strong', {}, retour.texte), retour.details ? el('span', { class: 'details' }, retour.details) : null, liens,
      retour.annulable && etat.journal?.length
        ? el('button', { type: 'button', class: 'geste fort annuler', onclick: annulerDerniere }, 'Annuler la dernière correction')
        : null].filter(Boolean));
  }
  const conformes = [];
  const blocs = [];
  for (const [cat, nomCat] of Object.entries(CATEGORIES)) {
    const res = r.resultats.filter((x) => x.regle.categorie === cat);
    res.filter((x) => !x.occurrences.length).forEach((x) => conformes.push(x));
    const source = (x) => (etat.filtre === 'ignorees' ? (x.ignorees || []).map((o) => ({ ...o, ignoree: true })) : x.occurrences);
    const avec = res.map((x) => ({ ...x, occurrences: source(x).map((o) => ({ ...o, reste: cat === 'R' })).filter(garder) })).filter((x) => x.occurrences.length);
    if (!avec.length) continue;
    const total = avec.reduce((k, x) => k + x.occurrences.length, 0);
    blocs.push(el('h2', { class: 'categorie' }, el('span', {}, nomCat), el('span', {}, `${total}`)));
    // Bloquants d'abord, puis par nombre d'occurrences.
    avec.sort((x, y) => (y.occurrences.some((o) => o.niveau === 'bloquant') - x.occurrences.some((o) => o.niveau === 'bloquant'))
      || x.regle.id.localeCompare(y.regle.id));
    for (const x of avec) blocs.push(carteRegle(x));
  }
  if (!blocs.length) {
    blocs.push(el('div', { class: 'vide' }, el('p', { class: 'vide-titre' }, etat.filtre === 'tous' ? 'Aucun écart trouvé' : etat.filtre === 'ignorees' ? 'Aucune occurrence ignorée' : 'Rien dans ce filtre'),
      el('p', {}, etat.filtre === 'tous' ? 'Voir aussi l\'onglet Exigences : certaines se vérifient à la main.' : 'Choisir « Toutes » pour tout revoir.')));
  }
  if (conformes.length && etat.filtre === 'tous') {
    blocs.push(el('details', { class: 'conformes' },
      el('summary', {}, `Contrôles sans écart ou hors de portée (${conformes.length})`),
      el('ul', {}, ...conformes.map((x) => el('li', {}, `${x.regle.id} ${x.regle.titre}${x.horsSource ? ' : sur le .docx seulement' : ''}`)))));
  }
  $('resultats').replaceChildren(...blocs);
}

const CHEVRON = () => {
  const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  s.setAttribute('viewBox', '0 0 16 16');
  s.setAttribute('class', 'chevron');
  s.setAttribute('aria-hidden', 'true');
  s.innerHTML = '<path d="m6 3 5 5-5 5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>';
  return s;
};
const VISIBLES = 8;

function carteRegle(x) {
  const bloquants = x.occurrences.filter((o) => o.niveau === 'bloquant').length;
  const niveau = bloquants ? 'bloquant' : 'avertissement';
  const n = x.occurrences.length;
  const corrigeables = x.occurrences.filter((o) => o.corrections?.length).length;
  const occurrences = etat.deplies.has(x.regle.id) ? x.occurrences : x.occurrences.slice(0, VISIBLES);
  const liste = el('ul', { class: 'occurrences' }, ...occurrences.map((o) => ligneOccurrence(o, x)));
  const plus = n > occurrences.length
    ? el('button', { type: 'button', class: 'plus', onclick: () => { etat.deplies.add(x.regle.id); afficherAlertes(etat.dernier); } }, `Afficher les ${n - occurrences.length} autres`)
    : null;
  const carte = el('details', { class: `regle ${x.regle.categorie === 'R' ? 'reste' : niveau}` },
    el('summary', {},
      CHEVRON(),
      el('span', { class: 'id' }, x.regle.id),
      el('span', { class: 'nom' }, x.regle.titre),
      corrigeables ? el('span', { class: 'badge auto', title: 'Corrigeables dans Word' }, `✓ ${corrigeables}`) : null,
      (() => {
        const premiere = x.occurrences.find((o) => o.i !== null && o.i >= 0);
        if (!premiere || !enWord()) return null;
        const b = el('button', { type: 'button', class: 'aller', title: 'Aller à la première occurrence dans le document', 'aria-label': 'Aller à la première occurrence' }, '↗');
        b.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); aller(premiere.i, premiere.passage); });
        return b;
      })(),
      el('span', { class: `badge ${x.regle.categorie === 'R' ? 'avertissement' : niveau}`, title: x.regle.categorie === 'R' ? 'À nettoyer avant l\'export' : niveau === 'bloquant' ? 'Bloquants' : 'À vérifier' }, `${n}`)),
    el('div', { class: 'corps' },
      el('p', { class: 'base' },
        el('span', { class: `etiquette ${x.regle.nature === 'exigence' ? '' : 'preference'}` }, x.regle.nature === 'exigence' ? 'Note du jury' : 'Préférence'),
        x.regle.base),
      x.regle.corriger ? el('div', { class: 'conseil' }, el('strong', {}, 'Comment corriger'), x.regle.corriger) : null,
      blocAuto(x),
      liste, plus));
  // Repliée par défaut ; l'état ouvert survit au réaffichage (Ignorer, changement de mode, filtre).
  carte.open = etat.ouvertes.has(x.regle.id);
  carte.addEventListener('toggle', () => {
    if (carte.open) etat.ouvertes.add(x.regle.id); else etat.ouvertes.delete(x.regle.id);
  });
  return carte;
}

const enWord = () => etat.doc?.source === 'word' && etat.dansWord;

/** Aperçu « avant → après » d'une liste de corrections, puis application sur confirmation. */
/** Aperçu « avant → après » de chaque paragraphe, puis application sur confirmation. */
function apercuCorrections(liste, conteneur) {
  const p = etat.regles.parametres;
  if (liste.every((c) => c.type === 'style-normal')) {
    conteneur.replaceChildren(el('div', { class: 'apercu' },
      el('p', {}, `Le style Normal passera en taille ${p.taille_texte} : tous les paragraphes sans mise en forme directe changent d'un coup.`),
      el('span', { class: 'gestes' },
        el('button', { type: 'button', class: 'principal petit', onclick: () => appliquer(liste) }, 'Appliquer'),
        el('button', { type: 'button', class: 'geste', onclick: () => conteneur.replaceChildren() }, 'Annuler'))));
    return;
  }
  const vues = liste.map((c) => ({ c, b: etat.doc.blocs[c.i] })).map((x) => ({ ...x, apres: x.b ? apercu(x.c, x.b, p) : null }));
  const applicables = vues.filter((x) => x.apres !== null).map((x) => x.c);
  const nb = new Set(applicables.map((c) => c.i)).size;
  conteneur.replaceChildren(el('div', { class: 'apercu' },
    el('p', { class: 'discret' }, `${nb} paragraphe(s) seront modifiés. Vérifier les lignes ; la mise en forme d'origine est conservée, et « Annuler la dernière correction » rétablit tout en un clic.`),
    el('ul', {}, ...vues.map((x) => el('li', {}, el('span', { class: 'avant' }, x.b ? x.b.texte.trim() : ''),
      x.apres === null ? el('span', { class: 'discret' }, 'laissé tel quel (rien à changer ou cas ambigu)') : el('span', { class: 'apres' }, `→ ${x.apres.trim()}`)))),
    el('span', { class: 'gestes' },
      el('button', { type: 'button', class: 'principal petit', onclick: () => appliquer(applicables) }, `Appliquer à ${nb} paragraphe(s)`),
      el('button', { type: 'button', class: 'geste', onclick: () => conteneur.replaceChildren() }, 'Annuler'))));
}

/** Bouton de correction : ouvre toujours l'aperçu avant toute modification. */
function boutonCorriger(liste, libelle, conteneur) {
  const bouton = el('button', { class: 'auto', type: 'button' }, libelle);
  bouton.addEventListener('click', () => apercuCorrections(liste, conteneur));
  return bouton;
}

/** Corrections d'une règle, regroupées par type. Hors de Word : simple indication. */
function blocAuto(x) {
  const toutes = x.occurrences.filter((o) => !o.ignoree).flatMap((o) => o.corrections || []);
  if (!toutes.length) return null;
  const parType = new Map();
  for (const c of toutes) parType.set(c.type, [...(parType.get(c.type) || []), c]);
  if (!enWord()) {
    return el('div', { class: 'bloc-auto' }, el('p', { class: 'indication-auto' }, `Correction en un clic quand le document est ouvert dans Word : ${[...parType.keys()].map((t) => LIBELLES[t]).join(' ; ')}.`));
  }
  const apercuZone = el('div', {});
  const boutons = [...parType].map(([type, liste]) => {
    const nb = new Set(liste.map((c) => c.i)).size;
    return boutonCorriger(liste, `Corriger : ${LIBELLES[type]}${type === 'style-normal' ? '' : ` (${nb} paragraphe${nb > 1 ? 's' : ''})`}`, apercuZone);
  });
  return el('div', { class: 'bloc-auto' }, ...boutons, apercuZone);
}

async function appliquer(liste) {
  // Un seul traitement à la fois : un second clic pendant l'application est ignoré.
  if (etat.enCours) return;
  etat.enCours = true;
  document.body.classList.add('correction-en-cours');
  statut('Correction en cours…');
  try {
    // Document modifié depuis le contrôle (frappe, Ctrl+Z) : aucune correction sur des alertes périmées.
    if (etat.empreinte) {
      let e = null;
      try { e = await empreinte(); } catch { e = null; }
      if (e && e !== etat.empreinte) {
        $('perime').hidden = false;
        etat.journal = null;
        etat.retour = { erreur: true, texte: 'Aucune correction appliquée : le document a changé depuis le contrôle.', details: 'Relancer le contrôle, puis corriger.' };
        statut('Document modifié depuis le contrôle : relancer le contrôle.', true);
        afficherAlertes(etat.dernier);
        return;
      }
    }
    const r = await corriger(liste, etat.doc, etat.regles.parametres);
    const details = [];
    if (r.dejaCorriges) details.push(`${r.dejaCorriges} paragraphe(s) déjà corrigé(s)`);
    if (r.decales) details.push(`${r.decales} paragraphe(s) laissé(s) : texte modifié depuis le contrôle (${r.exemples.join(' ; ')}), relancer le contrôle`);
    if (r.mixtes) details.push(`${r.mixtes} passage(s) en mise en forme mixte (italique ou gras sur une partie) laissé(s) tels quels : à corriger à la main`);
    if (r.introuvables) details.push(`${r.introuvables} passage(s) introuvable(s)`);
    if (r.ignorees) details.push(`${r.ignorees} cas ambigu(s) laissé(s)`);
    if (r.erreurs?.length) details.push(`refusé par Word : ${r.erreurs.join(' ; ')}`);
    if (r.journal?.length) etat.journal = r.journal;
    etat.retour = r.corrections || (r.dejaCorriges && !r.decales)
      ? { texte: !r.corrections ? 'Déjà corrigé : rien à changer.' : (r.paragraphes ? `${r.corrections} remplacement(s) dans ${r.paragraphes} paragraphe(s), mise en forme conservée.` : `Taille ${etat.regles.parametres.taille_texte} appliquée (${r.corrections} modification(s) de style).`), details: details.join(' · '), annulable: !!r.journal?.length, lieux: r.lieux }
      : { erreur: true, texte: 'Aucune correction appliquée.', details: `${details.join(' · ') || 'rien trouvé'}. Relancer le contrôle ; si le problème persiste, copier le diagnostic (Réglages).`, lieux: r.lieux };
    etat.dernierRapport = { liste: liste.length, types: [...new Set(liste.map((c) => c.type))], ...r, journal: r.journal?.length, lieux: r.lieux?.map((l) => `${l.i + 1} ${l.motif}`) };
    await lireOuvert();
    // Montrer la correction dans le document : premier paragraphe corrigé sélectionné.
    const premier = r.journal?.find((f) => f.i !== undefined);
    if (premier) await aller(premier.i);
  } catch (e) {
    etat.retour = { erreur: true, texte: `Correction impossible : ${e.message}`, details: 'Document protégé, en lecture seule, ou fonction absente de cette version de Word. Copier le diagnostic (Réglages).' };
    etat.dernierRapport = { erreur: `${e.name} : ${e.message}${e.code ? ` (${e.code})` : ''}`, debug: e.debugInfo ? JSON.stringify(e.debugInfo).slice(0, 400) : null };
    if (etat.dernier) afficherAlertes(etat.dernier);
  } finally {
    etat.enCours = false;
    document.body.classList.remove('correction-en-cours');
  }
  ouvrirOnglet('alertes');
  $('retour').scrollIntoView({ block: 'nearest' });
}

/** Annule en un clic la dernière correction (tous ses remplacements), mise en forme comprise. */
async function annulerDerniere() {
  if (etat.enCours || !etat.journal?.length) return;
  etat.enCours = true;
  document.body.classList.add('correction-en-cours');
  statut('Annulation en cours…');
  try {
    const r = await annuler(etat.journal);
    etat.journal = null;
    etat.dernierRapport = { ...(etat.dernierRapport || {}), annulation: { retablis: r.corrections, introuvables: r.introuvables, erreurs: r.erreurs } };
    etat.retour = r.erreurs.length || r.introuvables
      ? { erreur: true, texte: `Annulation partielle : ${r.corrections} remplacement(s) rétabli(s).`, details: `${r.introuvables} introuvable(s)${r.erreurs.length ? ` ; ${r.erreurs.join(' ; ')}` : ''} : finir avec Ctrl+Z dans Word.` }
      : { texte: `Dernière correction annulée : ${r.corrections} remplacement(s) rétabli(s).`, details: 'Ne pas faire Ctrl+Z ensuite : Word rejouerait la correction annulée.' };
    await lireOuvert();
  } catch (e) {
    etat.retour = { erreur: true, texte: `Annulation impossible : ${e.message}`, details: 'Utiliser Ctrl+Z dans Word.' };
    if (etat.dernier) afficherAlertes(etat.dernier);
  } finally {
    etat.enCours = false;
    document.body.classList.remove('correction-en-cours');
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

const ICONES = { ok: '✓', alerte: '▲', manuel: '☐', auto: '•', na: '–', verifie: '✓' };

const motifDe = (e, id) => {
  const m = e.motifs?.[id] ?? e.motif;
  return m ? new RegExp(m, 'i') : null;
};

/**
 * Rattache chaque alerte à une seule exigence : celle dont le motif reconnaît le message, sinon celle où la règle
 * vient en premier, sinon la première (aucune alerte n'est perdue ni comptée deux fois).
 */
function rattacher(exigences, r) {
  const parExigence = new Map(exigences.map((e) => [e, 0]));
  if (!r) return parExigence;
  for (const x of r.resultats) {
    const candidates = exigences.filter((e) => e.regles.includes(x.regle.id));
    if (!candidates.length) continue;
    for (const o of x.occurrences) {
      const note = (e) => {
        const m = motifDe(e, x.regle.id);
        if (m) return m.test(o.message) ? 0 : null;
        return 1 + e.regles.indexOf(x.regle.id);
      };
      const notees = candidates.map((e) => ({ e, n: note(e) })).filter((c) => c.n !== null);
      const choix = notees.length ? notees.reduce((a, c) => (c.n < a.n ? c : a)).e : candidates[0];
      parExigence.set(choix, parExigence.get(choix) + 1);
    }
  }
  return parExigence;
}

/** Exigences cochées « vérifié » par le candidat, pour ce document. */
const cleVerif = () => `${etat.cleDoc || 'aucun'}|${etat.profil}`;
function basculerVerifie(texte) {
  const liste = new Set(etat.verifiees[cleVerif()] || []);
  if (liste.has(texte)) liste.delete(texte); else liste.add(texte);
  etat.verifiees[cleVerif()] = [...liste];
  stockage.ecrire({ verifiees: etat.verifiees });
  afficherExigences();
}

function afficherExigences() {
  const r = etat.dernier && etat.dernier.profil === etat.profil ? etat.dernier : null;
  const exigences = etat.regles.exigences.filter((x) => x.profils.includes(etat.profil));
  const alertes = rattacher(exigences, r);
  const verifiees = new Set(etat.verifiees[cleVerif()] || []);
  const groupes = new Map();
  const totaux = { ok: 0, alerte: 0, manuel: 0, auto: 0, na: 0, verifie: 0 };
  for (const e of exigences) {
    let etatEx = 'manuel';
    let libelle = 'À vérifier vous-même';
    if (e.regles.length) {
      const n = alertes.get(e);
      const ids = e.regles.join(', ');
      const propres = (x) => e.regles.includes(x.regle) && (!e.motif || new RegExp(e.motif, 'i').test(x.texte));
      const nonApplicable = r && !n && r.infos.some((x) => propres(x) && /^Non applicable/.test(x.texte));
      const nonControle = r && !n && r.infos.some((x) => propres(x) && /non contr[ôo]l/.test(x.texte));
      const res = r ? r.resultats.filter((x) => e.regles.includes(x.regle.id)) : [];
      if (!r) { etatEx = 'auto'; libelle = `Contrôlé · ${ids}`; }
      else if (n) { etatEx = 'alerte'; libelle = `${n} alerte(s) · ${ids}`; }
      else if (nonApplicable) { etatEx = 'na'; libelle = 'Non applicable à ce document'; }
      else if (nonControle) { etatEx = 'manuel'; libelle = `Non contrôlé (pièce introuvable) · ${ids}`; }
      else if (res.length && res.every((x) => x.horsSource)) { etatEx = 'manuel'; libelle = 'Sur le .docx seulement'; }
      else if (e.manuel) { etatEx = 'manuel'; libelle = `Sans écart · ${ids} · reste à vérifier vous-même`; }
      else { etatEx = 'ok'; libelle = `Sans écart · ${ids}`; }
    }
    // Partie manuelle cochée par le candidat (seulement sans alerte).
    const cochable = r && etatEx === 'manuel';
    if (cochable && verifiees.has(e.texte)) { etatEx = 'verifie'; libelle = `${libelle.replace(/ · reste à vérifier vous-même$|^À vérifier vous-même$/, '') || 'Manuel'} · vérifié par vous`; }
    totaux[etatEx]++;
    if (!groupes.has(e.groupe)) groupes.set(e.groupe, { items: [], alertes: 0, manuels: 0 });
    const g = groupes.get(e.groupe);
    if (etatEx === 'alerte') g.alertes++;
    if (etatEx === 'manuel') g.manuels++;
    g.items.push(el('li', { class: `exigence ${etatEx}` },
      el('span', { class: 'icone', 'aria-hidden': 'true' }, ICONES[etatEx]),
      el('span', { class: 'statut-ex' }, libelle),
      el('span', {}, `${e.texte} (${e.page}).`),
      e.manuel ? el('span', { class: 'discret' }, e.manuel) : null,
      cochable || etatEx === 'verifie'
        ? el('button', { type: 'button', class: 'geste verif', onclick: () => basculerVerifie(e.texte) }, etatEx === 'verifie' ? 'Décocher' : 'Marquer vérifié')
        : null));
  }
  const codes = etat.regles.parametres.codes_sujet;
  const blocs = [];
  for (const [nom, g] of groupes) {
    if (nom === 'Dépôt' && etat.profil === 'notice' && codes) {
      g.items.push(el('li', { class: 'exigence auto' }, el('span', { class: 'icone' }, 'i'), el('details', {},
        el('summary', {}, 'Codes du sujet (2 à 6 au document annexe)'),
        el('p', { class: 'discret' }, 'Disciplines'), el('ul', { class: 'codes' }, ...codes.disciplines.map((c) => el('li', {}, c))),
        el('p', { class: 'discret' }, 'Secteurs et branches'), el('ul', { class: 'codes' }, ...codes.secteurs.map((c) => el('li', {}, c))))));
    }
    const resume = [g.alertes ? `${g.alertes} en alerte` : '', g.manuels ? `${g.manuels} à vérifier` : ''].filter(Boolean).join(' · ') || 'traité';
    const groupe = el('details', { class: 'groupe' }, el('summary', {}, CHEVRON(), el('span', {}, nom), el('span', { class: 'resume' }, resume)), el('ul', {}, ...g.items));
    groupe.open = g.alertes > 0 || !r;
    blocs.push(groupe);
  }
  $('exigences').replaceChildren(...blocs);
  const total = Object.values(totaux).reduce((a, b) => a + b, 0);
  $('progression').hidden = !r;
  if (r) {
    $('progression').querySelector('.ok').style.width = `${(100 * (totaux.ok + totaux.verifie + totaux.na)) / total}%`;
    $('progression').querySelector('.alerte').style.width = `${(100 * totaux.alerte) / total}%`;
    $('progression-texte').textContent = `${totaux.alerte + totaux.manuel} à traiter (${totaux.alerte} en alerte, ${totaux.manuel} à vérifier vous-même) · ${totaux.ok} sans écart${totaux.verifie ? ` · ${totaux.verifie} vérifiée(s) par vous` : ''}${totaux.na ? ` · ${totaux.na} non applicable(s)` : ''} · ${total} exigences`;
  }
  $('compte-exigences').textContent = r ? `${totaux.alerte + totaux.manuel}` : '';
  $('compte-exigences').title = 'Exigences à traiter : en alerte ou à vérifier vous-même';
}

function ligneOccurrence(o, x) {
  const allerPossible = enWord() && o.i !== null;
  const extrait = o.extrait
    ? (allerPossible
      ? el('button', { class: 'extrait', type: 'button', title: 'Aller au passage dans le document', onclick: () => aller(o.i, o.passage) }, o.extrait)
      : el('span', { class: 'extrait' }, o.i !== null ? `¶ ${o.i + 1} · ${o.extrait}` : o.extrait))
    : null;
  // Alerte regroupée (« 35 lignes… ») : parcours des paragraphes concernés un à un.
  const cibles = [...new Set((o.cibles || []).filter((i) => i !== null && i >= 0))];
  let rang = 0;
  const suivant = allerPossible && cibles.length > 1
    ? el('button', { class: 'geste', type: 'button', title: 'Aller au paragraphe concerné suivant' }, `Suivant › 1/${cibles.length}`)
    : null;
  if (suivant) {
    suivant.addEventListener('click', () => {
      rang = (rang + 1) % cibles.length;
      suivant.textContent = `Suivant › ${rang + 1}/${cibles.length}`;
      aller(cibles[rang]);
    });
  }
  const zone = el('div', {});
  const message = el('p', { class: `message${allerPossible ? ' cliquable' : ''}`, title: allerPossible ? 'Aller au passage dans le document' : '' }, o.message);
  if (allerPossible) message.addEventListener('click', () => aller(o.i, o.passage));
  return el('li', { class: `occurrence ${o.reste ? 'avertissement' : o.niveau}${o.ignoree ? ' ignoree' : ''}` },
    message,
    extrait,
    el('span', { class: 'gestes' },
      o.corrections?.length && enWord() && !o.ignoree ? (() => { const b = boutonCorriger(o.corrections, 'Corriger', zone); b.className = 'geste fort'; return b; })() : null,
      o.plage && enWord() ? el('button', { class: 'geste fort', type: 'button', title: 'Sélectionner les références de la catégorie pour le tri de Word', onclick: () => selectionnerPourTri(o.plage) }, 'Sélectionner pour trier') : null,
      o.ignoree
        ? el('button', { class: 'geste fort', type: 'button', title: 'Signaler de nouveau cette occurrence', onclick: () => retablir(o.cle) }, 'Rétablir')
        : el('button', { class: 'geste', type: 'button', title: 'Ne plus signaler cette occurrence pour ce document', onclick: () => ignorer(o.cle) }, 'Ignorer'),
      suivant),
    zone);
}

async function aller(i, passage = null) {
  if (i === null || i === undefined || i < 0 || !enWord()) return;
  try {
    if (!(await selectionnerParagraphe(i, passage))) statut('Passage introuvable : relancer le contrôle.', true);
  } catch (e) {
    statut(`Passage non sélectionné : ${e.message}`, true);
  }
}

function retablir(cle) {
  const liste = new Set(etat.ignores[etat.cleDoc] || []);
  liste.delete(cle);
  etat.ignores[etat.cleDoc] = [...liste];
  stockage.ecrire({ ignores: etat.ignores });
  if (!liste.size && etat.filtre === 'ignorees') etat.filtre = 'tous';
  afficher();
}

function ignorer(cle) {
  const liste = new Set(etat.ignores[etat.cleDoc] || []);
  liste.add(cle);
  etat.ignores[etat.cleDoc] = [...liste];
  stockage.ecrire({ ignores: etat.ignores });
  afficher();
}

// ---------------------------------------------------------------- mise en forme (texte simple → gabarit)

const NOMS_STYLES = {
  Titrehorsplan: 'Titre de pièce', Titrenotice: 'Section notice', 'Titrecatégorie': 'Catégorie', Titreannexe: 'Annexe',
  Heading1: 'Partie', Heading2: 'Chapitre', Heading3: 'Section', Heading4: '§', Normal: 'Texte',
};

function installerMiseEnForme(champs) {
  document.body.classList.add(`profil-${etat.profil}`);
  const ids = ['intitule', 'auteur', 'adresse', 'date', 'session'];
  const maintenant = new Date().toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' });
  for (const id of ids) {
    const champ = $(`champ-${id}`);
    champ.value = champs[id] ?? (id === 'date' ? maintenant : '');
    champ.addEventListener('change', () => stockage.ecrire({ champs: lireChamps() }));
  }
  $('analyser-forme').addEventListener('click', analyserForme);
  $('creer-forme').addEventListener('click', creerForme);
}

function lireChamps() {
  const c = {};
  for (const id of ['intitule', 'auteur', 'adresse', 'date', 'session']) c[id] = $(`champ-${id}`).value.trim();
  return c;
}

/** Lit le document (ou reprend la dernière lecture) et montre le style attribué à chaque paragraphe. */
async function analyserForme() {
  if (etat.dansWord && etat.doc?.source !== 'word') await lireOuvert();
  if (!etat.doc) { statut('Mise en forme : ouvrir le document dans Word, ou contrôler d\'abord un .docx.', true); return; }
  if (!etat.doc) return;
  const plan = planifier(etat.doc, etat.regles, etat.profil);
  etat.planForme = plan;
  const compte = {};
  for (const l of plan.lignes) {
    const cle = l.omis ? null : l.champ ? 'Table (champ Word)' : l.tableau !== undefined ? 'Tableau' : NOMS_STYLES[l.style] + (l.gras ? ' (gras)' : '');
    if (cle) compte[cle] = (compte[cle] || 0) + 1;
  }
  const visibles = plan.lignes.filter((l) => !l.omis && (l.champ || l.tableau !== undefined || l.style !== 'Normal' || l.gras));
  $('resultat-forme').replaceChildren(el('div', { class: 'carte-simple' },
    el('h2', { class: 'titre-section' }, 'Ce qui sera mis en forme'),
    el('div', { class: 'styles-forme' }, ...Object.entries(compte).map(([k, n]) => el('span', {}, `${k} : ${n}`))),
    ...plan.avertissements.map((a) => el('p', { class: 'bandeau' }, a)),
    el('p', { class: 'discret' }, `Page de garde et entrées de table du texte d'origine remplacées par celles du gabarit (${plan.lignes.filter((l) => l.omis).length} paragraphe(s)). Titres et tableaux ci-dessous ; le reste en texte courant.`),
    el('ul', { class: 'liste-forme' }, ...visibles.map((l) => el('li', {},
      el('span', { class: 'style' }, l.champ ? 'Table' : l.tableau !== undefined ? 'Tableau' : NOMS_STYLES[l.style] || l.style),
      el('span', { class: l.gras ? 'gras' : '' }, l.tableau !== undefined ? `Tableau ${l.tableau + 1} (repris tel quel)` : l.texte.slice(0, 140)))))));
  $('creer-forme').hidden = false;
}

function telecharger(octets, nom) {
  const url = URL.createObjectURL(new Blob([octets], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
  const a = el('a', { href: url, download: nom, class: 'bouton-lien' }, `Télécharger ${nom}`);
  return a;
}

/** Compose le nouveau document au gabarit et l'ouvre dans Word (sinon, le propose au téléchargement). */
async function creerForme() {
  if (!etat.planForme) return;
  const champs = lireChamps();
  stockage.ecrire({ champs });
  statut('Création du document au gabarit…');
  try {
    const nomGabarit = etat.profil === 'memoire' ? 'gabarit-memoire.docx' : 'gabarit-demande-agrement.docx';
    const gabarit = await fetch(`gabarits/${nomGabarit}?v=${VERSION}`).then((r) => { if (!r.ok) throw new Error(`gabarit introuvable (${r.status})`); return r.arrayBuffer(); });
    const octets = await composer(gabarit, etat.planForme, etat.profil, { ...champs, tableaux: etat.doc.tableauxXml || [] });
    const nom = etat.profil === 'memoire' ? 'memoire-mis-en-forme.docx' : 'demande-agrement-mise-en-forme.docx';
    let ouvert = false;
    if (etat.dansWord) {
      try { ouvert = await ouvrirNouveauDocument(enBase64(octets)); } catch (e) { ouvert = false; statut(`Ouverture refusée par Word : ${e.message}`, true); }
    }
    const lien = telecharger(octets, nom);
    $('resultat-forme').prepend(el('div', { class: 'retour' },
      el('strong', {}, ouvert ? 'Nouveau document ouvert dans Word.' : 'Document prêt.'),
      el('span', { class: 'details' }, ouvert
        ? 'Dans le nouveau document : clic droit sur la table des matières, « Mettre à jour les champs », puis lancer Garde-fou.'
        : 'Word n\'a pas pu l\'ouvrir directement : le télécharger, puis l\'ouvrir dans Word (en ligne : le déposer dans OneDrive).'),
      el('div', { class: 'ligne-boutons' }, lien)));
    statut(ouvert ? 'Document au gabarit créé et ouvert.' : 'Document au gabarit prêt au téléchargement.');
  } catch (e) {
    statut(`Mise en forme impossible : ${e.message}`, true);
  }
}

demarrer().catch((e) => statut(`Démarrage impossible : ${e.message}`, true));
