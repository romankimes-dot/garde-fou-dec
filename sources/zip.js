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
    async texte(nom) {
      const o = await octets(nom);
      return o ? utf8.decode(o) : null;
    },
  };
}
