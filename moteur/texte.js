// Outils de texte communs aux contrôles.

/** Minuscules, sans accents, apostrophes et espaces unifiées. */
export function normaliser(s) {
  return (s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[’`]/g, "'")
    .replace(/[  \t\s]+/g, ' ')
    .replace(/[–—]/g, '-')
    .trim();
}

/** Clé d'un titre : sans numéro (« Partie 1 – », « 1. », « I. »…) ni ponctuation finale. */
export function cleTitre(s) {
  return normaliser(s)
    .replace(/^[•▪◦]\s*/, '')
    .replace(/^(partie|chapitre|section|§|paragraphe|annexe)\s*(n°\s*)?\d+\s*[-:.]?\s*/, '')
    .replace(/^([ivxlc]+|\d+(\.\d+)*|[a-z])\s*[.)-]\s+/, '')
    .replace(/[\s.:;,]+$/, '')
    .trim();
}

/** Extrait lisible autour d'une position. */
export function extrait(texte, pos = 0, longueur = 90) {
  const t = (texte || '').replace(/[\t\n]+/g, ' ');
  const debut = Math.max(0, pos - 30);
  const fin = Math.min(t.length, debut + longueur);
  return (debut > 0 ? '…' : '') + t.slice(debut, fin).trim() + (fin < t.length ? '…' : '');
}

/** Remplace les URL et adresses électroniques par des espaces (positions conservées). */
export function masquerUrl(t) {
  return t.replace(/\b(?:https?:\/\/|www\.)\S+|[\w.+-]+@[\w-]+\.[\w.-]+/g, (m) => ' '.repeat(m.length));
}

export const nombre = (n) => n.toLocaleString('fr-FR');
export const pages = (n) => n.toLocaleString('fr-FR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
