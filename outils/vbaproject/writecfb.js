// usage: node writecfb.js manifest.json out.bin   manifest: [[path, file], ...]
// Écrit le conteneur OLE (vbaProject.bin) avec la bibliothèque SheetJS « cfb »,
// sans le flux témoin « \u0001Sh33tJ5 » qu'elle ajoute d'habitude (inutile pour Excel).
const fs = require('fs'), path = require('path'), vm = require('vm');
const src = fs.readFileSync(require.resolve('cfb'), 'utf8')
  .replace('function seed_cfb(cfb) {', 'function seed_cfb(cfb) { return;');
if (src.indexOf('function seed_cfb(cfb) { return;') < 0) throw new Error('cfb : fonction seed_cfb introuvable');
const mod = { exports: {} };
vm.runInThisContext('(function(module, exports, require){' + src + '\n})')(mod, mod.exports, require);
const CFB = mod.exports;
const [,, man, out] = process.argv;
const items = JSON.parse(fs.readFileSync(man, 'utf8'));
const cfb = CFB.utils.cfb_new();
for (const [p, f] of items) CFB.utils.cfb_add(cfb, p, fs.readFileSync(f));
CFB.utils.cfb_gc(cfb);
fs.writeFileSync(out, CFB.write(cfb, { type: 'buffer', fileType: 'cfb' }));
