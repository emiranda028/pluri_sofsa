// Genera la presentación en PDF (vectorial, A4 apaisado) a partir de index.html.
// Uso: node tools/build_pdf.js [salida.pdf] [carpeta_fuentes]
//   carpeta_fuentes (opcional): copia local de fonts.css + woff2 de Google Fonts,
//   útil en entornos sin acceso directo a fonts.googleapis.com.
const path = require('path'), fs = require('fs');
let pw; try { pw = require('playwright'); } catch (e) { pw = require(require('child_process').execSync('npm root -g').toString().trim() + '/playwright'); }
const ROOT = path.resolve(__dirname, '..');
const OUT = path.resolve(process.argv[2] || path.join(ROOT, 'Presupuesto_2027_SOFSA_presentacion.pdf'));
const FONTS = process.argv[3];
(async () => {
  const b = await pw.chromium.launch();
  const p = await b.newPage({ viewport: { width: 1440, height: 900 } });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  if (FONTS) {
    await p.route('https://fonts.googleapis.com/**', r => r.fulfill({ contentType: 'text/css',
      body: fs.readFileSync(path.join(FONTS, 'fonts.css'), 'utf8').replace(/https:\/\/fonts\.gstatic\.com\/([^)]*)/g, (_, f) => 'https://fonts.gstatic.com/' + f) }));
    await p.route('https://fonts.gstatic.com/**', r => { const f = new URL(r.request().url()).pathname.slice(1).replace(/\//g, '_');
      r.fulfill({ contentType: f.endsWith('.ttf') ? 'font/ttf' : 'font/woff2', body: fs.readFileSync(path.join(FONTS, f)) }); });
  }
  await p.goto('file://' + path.join(ROOT, 'index.html'), { waitUntil: 'networkidle' });
  await p.evaluate(() => document.fonts.ready);
  await p.click('.tab[data-tab="pres"]'); await p.waitForTimeout(300);
  await p.evaluate(() => { window.print = () => {}; presPrint(); });
  await p.waitForTimeout(900);
  await p.evaluate(() => document.fonts.ready);
  const n = await p.$$eval('#pres-print .pslide', x => x.length);
  await p.emulateMedia({ media: 'print' });
  await p.pdf({ path: OUT, printBackground: true, preferCSSPageSize: true });
  console.log('PDF', OUT, 'páginas', n, errs.length ? 'ERRORES ' + JSON.stringify(errs) : '');
  await b.close(); if (errs.length) process.exit(1);
})();
