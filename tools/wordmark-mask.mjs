import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
const b = await chromium.launch();
const p = await b.newPage({ viewport: { width: 1300, height: 900 } });
await p.setContent(`<html><head><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Squada+One&display=swap">
<style>body{margin:0;background:#000;width:1300px;height:900px;display:grid;place-items:center}
.b{display:grid;gap:0;justify-items:center;transform:scaleY(1.32)}
.l{font:200px/0.8 'Squada One';color:#fff;white-space:nowrap;transform-origin:center}</style></head>
<body><div class="b"><div class="l" id="a">GOLRANG</div><div class="l" id="c">FUTURE</div><div class="l" id="d">STARS</div></div></body></html>`);
await p.waitForTimeout(1500);
// every line exactly the same width, like the original block; STARS by spreading its letters
await p.evaluate(() => {
  const W = document.getElementById('a').getBoundingClientRect().width;
  for (const id of ['c']) { const el = document.getElementById(id); el.style.transform = 'scaleX(' + (W / el.getBoundingClientRect().width) + ')'; }
  const s = document.getElementById('d'); const n = s.textContent.length;
  let ls = 0; for (let i = 0; i < 40; i++) { s.style.letterSpacing = ls + 'px'; s.style.marginInlineStart = ls + 'px'; const w = s.getBoundingClientRect().width - ls; if (Math.abs(w - W) < 2) break; ls += (W - w) / n; }
});
await p.waitForTimeout(200);
await p.screenshot({ path: 'logomask3.png' });
await b.close();
