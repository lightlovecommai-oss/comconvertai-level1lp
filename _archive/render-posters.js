// 海報產生器：poster_lv{1,2,3}.svg → 同名 .png
// 用法：node render-posters.js   （改 SVG 後一定要重跑，PNG 才不會停在舊詞）
const sharp = require('sharp');
const fs = require('fs');
(async () => {
  for (const n of [1, 2, 3]) {
    const svg = `poster_lv${n}.svg`;
    if (!fs.existsSync(svg)) continue;
    await sharp(Buffer.from(fs.readFileSync(svg))).png().toFile(`poster_lv${n}.png`);
    console.log(`✓ poster_lv${n}.png`);
  }
})();
