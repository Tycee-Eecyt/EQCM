// Generates build/icon.ico from assets/tray-256.png so installer/shortcut icon
// matches the system tray icon artwork.
// Creates a minimal ICO (single 256x256 PNG entry) which Windows will scale.

const fs = require('fs');
const path = require('path');

function main(){
  try {
    const projectRoot = __dirname ? path.join(__dirname, '..') : process.cwd();
    const trayPng = path.join(projectRoot, 'assets', 'tray-256.png');
    const outDir = path.join(projectRoot, 'build');
    const outIco = path.join(outDir, 'icon.ico');

    if (!fs.existsSync(trayPng)) {
      console.error('Missing assets/tray-256.png; cannot generate icon.ico');
      process.exit(1);
    }
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const png = fs.readFileSync(trayPng);

    // ICO header: 6 bytes
    const header = Buffer.alloc(6);
    header.writeUInt16LE(0, 0); // reserved
    header.writeUInt16LE(1, 2); // type: icon
    header.writeUInt16LE(1, 4); // number of images

    // Directory entry: 16 bytes
    const dir = Buffer.alloc(16);
    dir.writeUInt8(0, 0); // width (0 means 256)
    dir.writeUInt8(0, 1); // height (0 means 256)
    dir.writeUInt8(0, 2); // color palette
    dir.writeUInt8(0, 3); // reserved
    dir.writeUInt16LE(1, 4); // color planes
    dir.writeUInt16LE(32, 6); // bits per pixel
    dir.writeUInt32LE(png.length, 8); // size of image data
    dir.writeUInt32LE(6 + 16, 12); // offset of image data

    const out = Buffer.concat([header, dir, png]);
    fs.writeFileSync(outIco, out);
    console.log('Wrote', path.relative(projectRoot, outIco));
  } catch (e) {
    console.error('Failed to generate build/icon.ico:', e && e.message || e);
    process.exit(1);
  }
}

if (require.main === module) main();

