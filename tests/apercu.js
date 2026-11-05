// Aperçu en console : node tests/apercu.js FICHIER.docx notice|memoire [termes,séparés,par,virgules]
import { readFile } from 'node:fs/promises';
import { lireDocx } from '../sources/docx.js';
import { controler } from '../moteur/moteur.js';

const [fichier, profil = 'notice', termes = ''] = process.argv.slice(2);
const regles = JSON.parse(await readFile(new URL('../regles.json', import.meta.url), 'utf8'));
regles.consignes = JSON.parse(await readFile(new URL('../consignes.json', import.meta.url), 'utf8'));
const doc = await lireDocx(await readFile(fichier));
const r = controler(doc, regles, { profil, termes: termes ? termes.split(',') : [] });
console.log(`Feu ${r.feu} : ${r.comptes.bloquants} bloquant(s), ${r.comptes.avertissements} avertissement(s)`);
for (const x of r.infos) console.log(`  i ${x.regle} ${x.texte}`);
for (const res of r.resultats) {
  for (const o of res.occurrences) console.log(`${o.niveau === 'bloquant' ? '■' : '▲'} ${res.regle.id} ${o.message}  [${o.i ?? '-'}] ${o.extrait}`);
}
if (process.env.STRUCTURE) for (const b of doc.blocs) if (b.titre) console.log(`  T${b.niveau} ${b.role} ${b.zone || ''} | ${b.intitule}`);
