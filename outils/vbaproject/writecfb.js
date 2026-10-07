// usage: node writecfb.js manifest.json out.bin   manifest: [[path, file], ...]
const CFB = require('cfb'); const fs = require('fs');
const [,, man, out] = process.argv;
const items = JSON.parse(fs.readFileSync(man, 'utf8'));
const cfb = CFB.utils.cfb_new();
for (const [p, f] of items) CFB.utils.cfb_add(cfb, p, fs.readFileSync(f));
CFB.utils.cfb_gc(cfb);
fs.writeFileSync(out, CFB.write(cfb, { type: 'buffer', fileType: 'cfb' }));
