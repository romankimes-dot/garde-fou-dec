// Même numéro de version partout (outils/version.py) : volet, manifeste, paramètre ?v= des fichiers chargés.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';

const lire = (f) => readFile(new URL(`../${f}`, import.meta.url), 'utf8');

test('version cohérente et tous les modules versionnés', async () => {
  const v = /const VERSION = '([^']+)'/.exec(await lire('taskpane.js'))[1];
  assert.match(await lire('manifest.xml'), new RegExp(`<Version>${v.replace(/\./g, '\\.')}\\.0</Version>`));
  const html = await lire('taskpane.html');
  assert.ok(html.includes(`taskpane.js?v=${v}`) && html.includes(`taskpane.css?v=${v}`));
  const carte = JSON.parse(/<script type="importmap">([\s\S]*?)<\/script>/.exec(html)[1]).imports;
  for (const dossier of ['sources', 'moteur']) {
    for (const f of (await readdir(new URL(`../${dossier}/`, import.meta.url))).filter((x) => x.endsWith('.js'))) {
      assert.equal(carte[`./${dossier}/${f}`], `./${dossier}/${f}?v=${v}`, `${dossier}/${f} absent de l'import map`);
    }
  }
});
