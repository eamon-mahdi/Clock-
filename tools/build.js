// Builds:
//   public/            the website Cloudflare serves (index.html, js/, _headers, robots.txt)
//   dist/haematology-rota.html  one self-contained page (demo mode, saves in the browser)
//   templates/*.xlsx   the import templates
// Run: node tools/build.js
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
global.RotaRules = require('../app/js/rules.js');
global.RotaXlsx = require('../app/js/xlsx.js');
global.RotaStore = require('../app/js/store.js');
const E = require('../app/js/excel.js');

const body = fs.readFileSync(path.join(root, 'app/index.html'), 'utf8');
const page = (inner) => `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 30 30'%3E%3Crect x='3' y='5' width='24' height='22' rx='4' fill='%2315202b'/%3E%3Crect x='7' y='14' width='5' height='6' rx='1' fill='%23f6b4ae'/%3E%3Crect x='13' y='14' width='5' height='6' rx='1' fill='%23a3cdf3'/%3E%3Crect x='19' y='14' width='5' height='6' rx='1' fill='%23b7dd9a'/%3E%3C/svg%3E">
${inner.replace(/\n(<header)/, '\n</head>\n<body>\n$1')}
</body>
</html>
`;

// public/ for Cloudflare
const pub = path.join(root, 'public');
fs.rmSync(pub, { recursive: true, force: true });
fs.mkdirSync(path.join(pub, 'js'), { recursive: true });
for (const f of ['xlsx.js', 'rules.js', 'store.js', 'excel.js', 'ui.js']) fs.copyFileSync(path.join(root, 'app/js', f), path.join(pub, 'js', f));
fs.writeFileSync(path.join(pub, 'index.html'), page(body));
fs.writeFileSync(path.join(pub, 'robots.txt'), 'User-agent: *\nDisallow: /\n');
fs.writeFileSync(path.join(pub, '_headers'), `/*
  X-Content-Type-Options: nosniff
  X-Frame-Options: DENY
  Referrer-Policy: no-referrer
  Permissions-Policy: camera=(), microphone=(), geolocation=()
  X-Robots-Tag: noindex, nofollow
  Content-Security-Policy: default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; font-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'
/js/*
  Cache-Control: no-cache
/
  Cache-Control: no-cache
`);

// dist/ single file
let single = body.replace(/<script src="(js\/[\w.]+)"><\/script>/g, (_, src) => {
  const code = fs.readFileSync(path.join(root, 'app', src), 'utf8');
  if (/<\/script/i.test(code)) throw new Error(`${src} contains a closing script tag`);
  return `<script>\n${code}</script>`;
});
fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
fs.writeFileSync(path.join(root, 'dist/haematology-rota.html'), page(single));

const today = new Date().toISOString().slice(0, 10);
const demo = global.RotaStore.demoState(today);
fs.mkdirSync(path.join(root, 'templates'), { recursive: true });
fs.writeFileSync(path.join(root, 'templates/oncall-import-template.xlsx'), E.oncallTemplate(today));
fs.writeFileSync(path.join(root, 'templates/weekly-plan-template.xlsx'), E.planTemplate(demo));
fs.writeFileSync(path.join(root, 'templates/staff-list-template.xlsx'), E.staffTemplate(demo));
console.log('Built public/ (Cloudflare), dist/haematology-rota.html (' + Math.round(single.length / 1024) + ' KB) and 3 templates');
