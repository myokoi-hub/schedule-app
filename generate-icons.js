// アイコン生成スクリプト（カレンダー風の青いアイコン）
const { PNG } = require('pngjs');
const fs = require('fs');
const path = require('path');

function createIcon(size) {
  const png = new PNG({ width: size, height: size });
  const bg   = { r: 74,  g: 144, b: 217 }; // #4A90D9 青
  const dark = { r: 42,  g: 110, b: 190 }; // 濃い青（ヘッダー部分）
  const white = { r: 255, g: 255, b: 255 };

  const pad    = Math.round(size * 0.08);
  const radius = Math.round(size * 0.15);
  const headerH = Math.round(size * 0.28);

  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const idx = (size * y + x) * 4;
      const inRect = x >= pad && x < size - pad && y >= pad && y < size - pad;

      // 角丸判定
      const corners = [
        [pad + radius, pad + radius],
        [size - pad - radius - 1, pad + radius],
        [pad + radius, size - pad - radius - 1],
        [size - pad - radius - 1, size - pad - radius - 1],
      ];
      let inCornerCutout = false;
      if (x < pad + radius && y < pad + radius)
        inCornerCutout = Math.hypot(x - corners[0][0], y - corners[0][1]) > radius;
      if (x >= size - pad - radius && y < pad + radius)
        inCornerCutout = Math.hypot(x - corners[1][0], y - corners[1][1]) > radius;
      if (x < pad + radius && y >= size - pad - radius)
        inCornerCutout = Math.hypot(x - corners[2][0], y - corners[2][1]) > radius;
      if (x >= size - pad - radius && y >= size - pad - radius)
        inCornerCutout = Math.hypot(x - corners[3][0], y - corners[3][1]) > radius;

      const inside = inRect && !inCornerCutout;
      const inHeader = inside && y < pad + headerH;
      const inBody   = inside && y >= pad + headerH;

      // グリッド描画（本体部分）
      const bodyX = x - pad;
      const bodyY = y - (pad + headerH);
      const cellW = (size - pad * 2) / 7;
      const cellH = (size - pad * 2 - headerH) / 5;
      const gridLine = bodyX % Math.round(cellW) < 1 || bodyY % Math.round(cellH) < 1;

      let color = { r: 0, g: 0, b: 0, a: 0 };
      if (inHeader) {
        color = { ...dark, a: 255 };
        // ヘッダーに白いドット（飾り）
        const dotY = pad + Math.round(headerH * 0.5);
        const dotsX = [
          pad + Math.round((size - pad * 2) * 0.25),
          pad + Math.round((size - pad * 2) * 0.75),
        ];
        for (const dx of dotsX) {
          if (Math.hypot(x - dx, y - dotY) < size * 0.04) {
            color = { ...white, a: 255 };
          }
        }
        // 「31」文字の代わりに白い横線3本
        const lineY = [0.3, 0.55, 0.78].map(r => pad + Math.round(headerH * r));
        if (lineY.some(ly => Math.abs(y - ly) < Math.max(1, size * 0.02))
            && x > pad + (size - pad * 2) * 0.35
            && x < pad + (size - pad * 2) * 0.65) {
          color = { ...white, a: 255 };
        }
      } else if (inBody) {
        color = { ...bg, a: 255 };
        if (gridLine) color = { r: 60, g: 120, b: 200, a: 255 };
        // 数字風の白い点
        const dotCellX = Math.floor(bodyX / cellW);
        const dotCellY = Math.floor(bodyY / cellH);
        const dotCX = pad + Math.round((dotCellX + 0.5) * cellW);
        const dotCY = pad + headerH + Math.round((dotCellY + 0.5) * cellH);
        if (Math.hypot(x - dotCX, y - dotCY) < Math.max(1.5, size * 0.025)) {
          color = { ...white, a: 220 };
        }
      }

      png.data[idx]     = color.r;
      png.data[idx + 1] = color.g;
      png.data[idx + 2] = color.b;
      png.data[idx + 3] = color.a ?? 0;
    }
  }
  return png;
}

const outDir = path.join(__dirname, 'public');
for (const size of [192, 512]) {
  const png = createIcon(size);
  const buf = PNG.sync.write(png);
  fs.writeFileSync(path.join(outDir, `icon-${size}.png`), buf);
  console.log(`icon-${size}.png を作成しました`);
}
