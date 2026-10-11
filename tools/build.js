// Builds dist/haematology-rota.html (one self-contained page, scripts inlined)
// and the import templates in templates/. Run: node tools/build.js
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
global.RotaRules = require('../app/js/rules.js');
global.RotaXlsx = require('../app/js/xlsx.js');
global.RotaStore = require('../app/js/store.js');
const E = require('../app/js/excel.js');

let html = fs.readFileSync(path.join(root, 'app/index.html'), 'utf8');
html = html.replace(/<script src="(js\/[\w.]+)"><\/script>/g, (_, src) => {
  const code = fs.readFileSync(path.join(root, 'app', src), 'utf8');
  if (/<\/script/i.test(code)) throw new Error(`${src} contains a closing script tag`);
  return `<script>\n${code}</script>`;
});
fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
fs.writeFileSync(path.join(root, 'dist/haematology-rota.html'), html);

const today = new Date().toISOString().slice(0, 10);
const demo = global.RotaStore.demoState(today);
fs.mkdirSync(path.join(root, 'templates'), { recursive: true });
fs.writeFileSync(path.join(root, 'templates/oncall-import-template.xlsx'), E.oncallTemplate(today));
fs.writeFileSync(path.join(root, 'templates/weekly-plan-template.xlsx'), E.planTemplate(demo));
fs.writeFileSync(path.join(root, 'templates/staff-list-template.xlsx'), E.staffTemplate(demo));
console.log('Built dist/haematology-rota.html (' + Math.round(html.length / 1024) + ' KB) and 3 templates');
