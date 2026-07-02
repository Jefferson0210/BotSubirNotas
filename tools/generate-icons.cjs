/* Genera iconos PNG placeholder (badge azul con check blanco) para la extension.
   Uso: node tools/generate-icons.cjs   -> crea icons/icon16|48|128.png
   Reemplaza luego por tu logo definitivo (mismos nombres/tamanos). */
const zlib = require("zlib");
const fs = require("fs");
const path = require("path");

const BG = [21, 101, 192];   // #1565c0 (BRAND.colorPrimario)
const FG = [255, 255, 255];

function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length, 0);
  const td = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td), 0);
  return Buffer.concat([len, td, crc]);
}
function png(size) {
  const px = (x, y) => {
    const cx = size / 2, cy = size / 2;
    // fondo circular tipo "badge"
    const dist = Math.hypot(x - cx, y - cy);
    if (dist > size * 0.48) return [0, 0, 0, 0];           // fuera del circulo: transparente
    // check mark (dos segmentos)
    const s = size;
    const p1 = [0.30 * s, 0.54 * s], p2 = [0.45 * s, 0.68 * s], p3 = [0.72 * s, 0.34 * s];
    const grosor = Math.max(1.4, s * 0.09);
    if (distSeg(x, y, p1, p2) < grosor || distSeg(x, y, p2, p3) < grosor) return [...FG, 255];
    return [...BG, 255];
  };
  const raw = Buffer.alloc(size * (size * 4 + 1));
  let o = 0;
  for (let y = 0; y < size; y++) {
    raw[o++] = 0; // filtro none
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = px(x + 0.5, y + 0.5);
      raw[o++] = r; raw[o++] = g; raw[o++] = b; raw[o++] = a;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw, { level: 9 })),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}
function distSeg(px, py, a, b) {
  const dx = b[0] - a[0], dy = b[1] - a[1];
  const l2 = dx * dx + dy * dy;
  let t = l2 ? ((px - a[0]) * dx + (py - a[1]) * dy) / l2 : 0;
  t = Math.max(0, Math.min(1, t));
  return Math.hypot(px - (a[0] + t * dx), py - (a[1] + t * dy));
}

const dir = path.join(__dirname, "..", "icons");
fs.mkdirSync(dir, { recursive: true });
for (const s of [16, 48, 128]) {
  fs.writeFileSync(path.join(dir, `icon${s}.png`), png(s));
  console.log("icons/icon" + s + ".png");
}
console.log("Listo.");
