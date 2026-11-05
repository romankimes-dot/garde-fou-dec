// Lecture d'une archive ZIP (un .docx en est une), sans dépendance.
// Décompression par DecompressionStream('deflate-raw'), disponible dans Chrome et dans Node 18 et plus.

export async function lireZip(donnees) {
  const u8 = donnees instanceof Uint8Array ? donnees : new Uint8Array(donnees);
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let fin = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { fin = i; break; }
  }
  if (fin < 0) throw new Error("Ce fichier n'est pas un .docx (archive ZIP introuvable).");
  const nombre = dv.getUint16(fin + 10, true);
  let p = dv.getUint32(fin + 16, true);
  const entrees = new Map();
  const utf8 = new TextDecoder('utf-8');
  for (let k = 0; k < nombre; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Archive ZIP endommagée.');
    const methode = dv.getUint16(p + 10, true);
    const taille = dv.getUint32(p + 20, true);
    const lnom = dv.getUint16(p + 28, true);
    const lextra = dv.getUint16(p + 30, true);
    const lcom = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    entrees.set(utf8.decode(u8.subarray(p + 46, p + 46 + lnom)), { methode, taille, local });
    p += 46 + lnom + lextra + lcom;
  }

  async function octets(nom) {
    const e = entrees.get(nom);
    if (!e) return null;
    const debut = e.local + 30 + dv.getUint16(e.local + 26, true) + dv.getUint16(e.local + 28, true);
    const brut = u8.subarray(debut, debut + e.taille);
    if (e.methode === 0) return brut;
    if (e.methode !== 8) throw new Error(`Compression ZIP non prise en charge (${e.methode}).`);
    const flux = new Blob([brut]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(flux).arrayBuffer());
  }

  return {
    noms: [...entrees.keys()],
    octets,
    async texte(nom) {
      const o = await octets(nom);
      return o ? utf8.decode(o) : null;
    },
  };
}

// ---------------------------------------------------------------- écriture (sans compression)

const TABLE_CRC = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(u8) {
  let c = 0xffffffff;
  for (let i = 0; i < u8.length; i++) c = TABLE_CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** Archive ZIP « stockée » (sans compression) : [{ nom, donnees: Uint8Array }] → Uint8Array. */
export function ecrireZip(fichiers) {
  const utf8 = new TextEncoder();
  const morceaux = [];
  const central = [];
  let decalage = 0;
  for (const { nom, donnees } of fichiers) {
    const n = utf8.encode(nom);
    const crc = crc32(donnees);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true); // noms en UTF-8
    local.setUint16(8, 0, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, donnees.length, true);
    local.setUint32(22, donnees.length, true);
    local.setUint16(26, n.length, true);
    morceaux.push(new Uint8Array(local.buffer), n, donnees);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true);
    c.setUint16(4, 20, true);
    c.setUint16(6, 20, true);
    c.setUint16(8, 0x0800, true);
    c.setUint32(16, crc, true);
    c.setUint32(20, donnees.length, true);
    c.setUint32(24, donnees.length, true);
    c.setUint16(28, n.length, true);
    c.setUint32(42, decalage, true);
    central.push(new Uint8Array(c.buffer), n);
    decalage += 30 + n.length + donnees.length;
  }
  const tailleCentral = central.reduce((k, x) => k + x.length, 0);
  const fin = new DataView(new ArrayBuffer(22));
  fin.setUint32(0, 0x06054b50, true);
  fin.setUint16(8, fichiers.length, true);
  fin.setUint16(10, fichiers.length, true);
  fin.setUint32(12, tailleCentral, true);
  fin.setUint32(16, decalage, true);
  const tout = [...morceaux, ...central, new Uint8Array(fin.buffer)];
  const sortie = new Uint8Array(tout.reduce((k, x) => k + x.length, 0));
  let p = 0;
  for (const x of tout) { sortie.set(x, p); p += x.length; }
  return sortie;
}
