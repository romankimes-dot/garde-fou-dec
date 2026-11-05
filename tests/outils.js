import { readFile } from 'node:fs/promises';
import { lireDocx } from '../sources/docx.js';
import { controler } from '../moteur/moteur.js';

export async function regles() {
  const r = JSON.parse(await readFile(new URL('../regles.json', import.meta.url), 'utf8'));
  r.consignes = JSON.parse(await readFile(new URL('../consignes.json', import.meta.url), 'utf8'));
  return r;
}

export async function docx(nom) {
  return lireDocx(await readFile(new URL(nom, import.meta.url)));
}

/** Contrôle et rend la liste « ID message » des occurrences. */
export async function lancer(doc, options) {
  const r = controler(doc, await regles(), options);
  r.lignes = r.resultats.flatMap((x) => x.occurrences.map((o) => `${x.regle.id} ${o.message}`));
  r.ids = new Set(r.resultats.filter((x) => x.occurrences.length).map((x) => x.regle.id));
  return r;
}
