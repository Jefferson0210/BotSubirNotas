/* Genera claves de licencia validas para la extension (barrera basica).
   Debe usar EXACTAMENTE el mismo checksum que validarLicencia() en content.js.
   Uso:
     node tools/generar-licencia.cjs            -> una clave aleatoria-deterministica
     node tools/generar-licencia.cjs ABCD 1234  -> clave a partir de 2 grupos dados
*/
function checksumLic(cuerpo) {
  let h = 2166136261 >>> 0;
  for (const ch of cuerpo) { h ^= ch.charCodeAt(0); h = Math.imul(h, 16777619) >>> 0; }
  const al = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let out = "";
  for (let i = 0; i < 4; i++) out += al[(h >>> (i * 5)) & 31];
  return out;
}
function grupoDesde(n) {
  const al = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let s = "";
  for (let i = 0; i < 4; i++) { s += al[n & 31]; n = Math.floor(n / 32); }
  return s;
}
const norm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 4).padEnd(4, "A");

let g1, g2;
if (process.argv[2] && process.argv[3]) {
  g1 = norm(process.argv[2]); g2 = norm(process.argv[3]);
} else {
  // deterministico a partir de la hora (sin depender de Math.random)
  const t = Date.now();
  g1 = grupoDesde(t & 0xfffff); g2 = grupoDesde((t >> 7) & 0xfffff);
}
const clave = `SN-${g1}-${g2}-${checksumLic(g1 + g2)}`;
console.log(clave);
