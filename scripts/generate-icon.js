// Generates build/icon.ico from assets tray PNGs so installer/shortcut icon
// matches the system tray icon artwork. Packs multiple sizes when available
// for crisper icons at small sizes.

const fs = require('fs');
const path = require('path');

function main(){
  try {
    const projectRoot = __dirname ? path.join(__dirname, '..') : process.cwd();
    const outDir = path.join(projectRoot, 'build');
    const outIco = path.join(outDir, 'icon.ico');

    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const sizes = [16, 24, 32, 48, 64, 128, 256];
    const entries = [];
    for (const sz of sizes){
      const p = path.join(projectRoot, 'assets', `tray-${sz}.png`);
      if (fs.existsSync(p)){
        const buf = fs.readFileSync(p);
        entries.push({ sz, buf });
      }
    }

    // Fallback: at least pack 256 if present; otherwise bail
    if (entries.length === 0){
      const p256 = path.join(projectRoot, 'assets', 'tray-256.png');
      if (!fs.existsSync(p256)){
        console.error('Missing tray PNGs; expected at least assets/tray-256.png');
        process.exit(1);
      }
      entries.push({ sz: 256, buf: fs.readFileSync(p256) });
    }

    const count = entries.length;
    const header = Buffer.alloc(6);
    header.writeUInt16LE(0, 0); // reserved
    header.writeUInt16LE(1, 2); // type: icon
    header.writeUInt16LE(count, 4); // number of images

    const dirs = [];
    let offset = 6 + 16 * count;
    for (const e of entries){
      const dir = Buffer.alloc(16);
      dir.writeUInt8(e.sz === 256 ? 0 : e.sz, 0); // width (0 => 256)
      dir.writeUInt8(e.sz === 256 ? 0 : e.sz, 1); // height (0 => 256)
      dir.writeUInt8(0, 2); // colors in palette
      dir.writeUInt8(0, 3); // reserved
      dir.writeUInt16LE(1, 4); // planes
      dir.writeUInt16LE(32, 6); // bpp
      dir.writeUInt32LE(e.buf.length, 8); // size
      dir.writeUInt32LE(offset, 12); // offset
      dirs.push(dir);
      offset += e.buf.length;
    }

    const out = Buffer.concat([header, ...dirs, ...entries.map(e => e.buf)]);
    fs.writeFileSync(outIco, out);
    console.log('Wrote', path.relative(projectRoot, outIco), `(sizes: ${entries.map(e=>e.sz).join(', ')})`);
  } catch (e) {
    console.error('Failed to generate build/icon.ico:', e && e.message || e);
    process.exit(1);
  }
}

if (require.main === module) main();
