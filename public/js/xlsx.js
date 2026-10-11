/* Minimal .xlsx reader and writer with no dependencies.
 * Writer: stored (uncompressed) zip, shared styles (fills, fonts, borders, date format).
 * Reader: handles deflate via DecompressionStream, shared and inline strings, date-formatted cells.
 * Works in browsers and in Node 18+ (for tests and build scripts). */
(function (root) {
  'use strict';

  // ---------- zip ----------
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }
  const enc = (s) => new TextEncoder().encode(s);
  const dec = (b) => new TextDecoder('utf-8').decode(b);

  function zip(files) {
    // files: [{name, data: Uint8Array}] -> Uint8Array (method 0, stored)
    const parts = [];
    const central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc(f.name);
      const crc = crc32(f.data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0x0800, true); // UTF-8 names
      local.setUint16(8, 0, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, f.data.length, true);
      local.setUint32(22, f.data.length, true);
      local.setUint16(26, name.length, true);
      parts.push(new Uint8Array(local.buffer), name, f.data);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true);
      cd.setUint16(4, 20, true);
      cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x0800, true);
      cd.setUint32(16, crc, true);
      cd.setUint32(20, f.data.length, true);
      cd.setUint32(24, f.data.length, true);
      cd.setUint16(28, name.length, true);
      cd.setUint32(42, offset, true);
      central.push(new Uint8Array(cd.buffer), name);
      offset += 30 + name.length + f.data.length;
    }
    const cdSize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    const all = parts.concat(central, [new Uint8Array(end.buffer)]);
    const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
    let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(buf) {
    const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('This file is not a valid .xlsx (zip) file.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const out = {};
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Damaged .xlsx file (central directory).');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true);
      const elen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const lofs = dv.getUint32(p + 42, true);
      const name = dec(bytes.subarray(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(lofs + 26, true);
      const lelen = dv.getUint16(lofs + 28, true);
      const start = lofs + 30 + lnlen + lelen;
      const raw = bytes.subarray(start, start + csize);
      if (method === 0) out[name] = raw;
      else if (method === 8) out[name] = await inflateRaw(raw);
      else throw new Error('Unsupported compression in .xlsx file.');
      p += 46 + nlen + elen + clen;
    }
    return out;
  }

  // ---------- helpers ----------
  const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const xmlUnesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d)).replace(/&amp;/g, '&');
  function colName(i) { // 0 -> A
    let s = '';
    i += 1;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  function colIndex(ref) {
    const m = /^([A-Z]+)/.exec(ref);
    let n = 0;
    for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
  function isoToSerial(iso) {
    const [y, m, d] = iso.split('-').map(Number);
    return (Date.UTC(y, m - 1, d) - EXCEL_EPOCH) / 86400000;
  }
  function serialToIso(n) {
    const d = new Date(EXCEL_EPOCH + Math.round(n) * 86400000);
    return d.toISOString().slice(0, 10);
  }

  // ---------- writer ----------
  // sheet: {name, rows: [[cell]], widths: [n], freeze: {row, col}, heights: {rowIndex: pt}}
  // cell: null | string | number | {v, date: true, style: {...}}
  // style: {fill: 'RRGGBB', bold, color: 'RRGGBB', size, center, wrap, border}
  function write(sheets) {
    const styleKeys = new Map(); // json -> index
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const fonts = ['<font><sz val="10"/><name val="Calibri"/></font>'];
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>',
      '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>'];
    const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    const fillIdx = new Map(); const fontIdx = new Map();
    function styleIndex(style, isDate) {
      if (!style && !isDate) return 0;
      const s = style || {};
      const key = JSON.stringify([s, !!isDate]);
      if (styleKeys.has(key)) return styleKeys.get(key);
      let fillId = 0;
      if (s.fill) {
        if (!fillIdx.has(s.fill)) { fillIdx.set(s.fill, fills.length); fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${s.fill}"/><bgColor indexed="64"/></patternFill></fill>`); }
        fillId = fillIdx.get(s.fill);
      }
      let fontId = 0;
      if (s.bold || s.color || s.size) {
        const fk = [s.bold, s.color, s.size].join('|');
        if (!fontIdx.has(fk)) {
          fontIdx.set(fk, fonts.length);
          fonts.push(`<font>${s.bold ? '<b/>' : ''}<sz val="${s.size || 10}"/>${s.color ? `<color rgb="FF${s.color}"/>` : ''}<name val="Calibri"/></font>`);
        }
        fontId = fontIdx.get(fk);
      }
      const borderId = s.border ? 1 : 0;
      const align = s.center || s.wrap ? `<alignment${s.center ? ' horizontal="center" vertical="center"' : ''}${s.wrap ? ' wrapText="1"' : ''}/>` : '';
      const numFmt = isDate ? 164 : 0;
      xfs.push(`<xf numFmtId="${numFmt}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${fillId ? ' applyFill="1"' : ''}${fontId ? ' applyFont="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${isDate ? ' applyNumberFormat="1"' : ''}${align ? ' applyAlignment="1">' + align + '</xf>' : '/>'}`);
      styleKeys.set(key, xfs.length - 1);
      return xfs.length - 1;
    }

    const sheetXml = sheets.map((sh) => {
      const rowsXml = sh.rows.map((row, r) => {
        const cells = (row || []).map((cell, c) => {
          if (cell === null || cell === undefined || cell === '') return '';
          const obj = typeof cell === 'object' ? cell : { v: cell };
          const ref = colName(c) + (r + 1);
          const s = styleIndex(obj.style, obj.date);
          const sAttr = s ? ` s="${s}"` : '';
          if (obj.v === null || obj.v === undefined || obj.v === '') return `<c r="${ref}"${sAttr}/>`;
          if (obj.date) return `<c r="${ref}"${sAttr}><v>${isoToSerial(obj.v)}</v></c>`;
          if (typeof obj.v === 'number') return `<c r="${ref}"${sAttr}><v>${obj.v}</v></c>`;
          return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(obj.v)}</t></is></c>`;
        }).join('');
        const ht = sh.heights && sh.heights[r] ? ` ht="${sh.heights[r]}" customHeight="1"` : '';
        return `<row r="${r + 1}"${ht}>${cells}</row>`;
      }).join('');
      const cols = sh.widths ? `<cols>${sh.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '';
      let pane = '';
      if (sh.freeze) {
        const { row = 0, col = 0 } = sh.freeze;
        const tl = colName(col) + (row + 1);
        pane = `<sheetViews><sheetView workbookViewId="0"><pane${col ? ` xSplit="${col}"` : ''}${row ? ` ySplit="${row}"` : ''} topLeftCell="${tl}" activePane="${row && col ? 'bottomRight' : row ? 'bottomLeft' : 'topRight'}" state="frozen"/></sheetView></sheetViews>`;
      }
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${pane}${cols}<sheetData>${rowsXml}</sheetData></worksheet>`;
    });

    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts><fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xmlEsc(s.name.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`;
    const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;

    const files = [
      { name: '[Content_Types].xml', data: enc(types) },
      { name: '_rels/.rels', data: enc(rels) },
      { name: 'xl/workbook.xml', data: enc(workbook) },
      { name: 'xl/_rels/workbook.xml.rels', data: enc(wbRels) },
      { name: 'xl/styles.xml', data: enc(styles) }
    ].concat(sheetXml.map((x, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: enc(x) })));
    return zip(files);
  }

  // ---------- reader ----------
  const DATE_FMT_IDS = new Set([14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
  function textOf(fragment) {
    // concatenates every <t>…</t> inside a fragment
    let out = '';
    const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
    let m;
    while ((m = re.exec(fragment))) out += m[1] ? xmlUnesc(m[1]) : '';
    return out;
  }
  async function read(buf) {
    const files = await unzip(buf);
    const get = (n) => (files[n] ? dec(files[n]) : null);
    const wb = get('xl/workbook.xml');
    if (!wb) throw new Error('This file has no workbook inside. Is it an .xlsx file?');
    const rels = get('xl/_rels/workbook.xml.rels') || '';
    const relMap = {};
    rels.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const id = /Id="([^"]+)"/.exec(tag); const t = /Target="([^"]+)"/.exec(tag);
      if (id && t) relMap[id[1]] = t[1].replace(/^\/?xl\//, '').replace(/^\//, '');
      return '';
    });
    const shared = [];
    const ss = get('xl/sharedStrings.xml');
    if (ss) { const re = /<si>([\s\S]*?)<\/si>/g; let m; while ((m = re.exec(ss))) shared.push(textOf(m[1])); }
    // date styles
    const dateXf = [];
    const st = get('xl/styles.xml');
    if (st) {
      const custom = {};
      st.replace(/<numFmt\b[^>]*>/g, (tag) => {
        const id = /numFmtId="(\d+)"/.exec(tag); const code = /formatCode="([^"]*)"/.exec(tag);
        if (id && code) custom[+id[1]] = xmlUnesc(code[1]);
        return '';
      });
      const xfBlock = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(st);
      if (xfBlock) {
        const re = /<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g; let m;
        while ((m = re.exec(xfBlock[1]))) {
          const id = +((/numFmtId="(\d+)"/.exec(m[0]) || [0, 0])[1]);
          const code = custom[id] || '';
          dateXf.push(DATE_FMT_IDS.has(id) || (/[dy]/i.test(code.replace(/"[^"]*"/g, '')) && !/^\[?h/i.test(code)));
        }
      }
    }
    const sheets = [];
    const sre = /<sheet\b[^>]*>/g; let sm;
    while ((sm = sre.exec(wb))) {
      const name = xmlUnesc((/name="([^"]*)"/.exec(sm[0]) || [0, ''])[1]);
      const rid = (/r:id="([^"]+)"/.exec(sm[0]) || [0, ''])[1];
      const path = 'xl/' + (relMap[rid] || '');
      const xml = get(path);
      if (!xml) continue;
      const rows = [];
      const rre = /<row\b([^>]*)>([\s\S]*?)<\/row>|<row\b([^>]*)\/>/g; let rm;
      while ((rm = rre.exec(xml))) {
        const rAttr = /\br="(\d+)"/.exec(rm[1] || rm[3] || '');
        const rIdx = rAttr ? +rAttr[1] - 1 : rows.length;
        const row = [];
        const cre = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g; let cm;
        let auto = 0;
        while ((cm = cre.exec(rm[2] || ''))) {
          const attrs = cm[1];
          const ref = /\br="([A-Z]+\d+)"/.exec(attrs);
          const c = ref ? colIndex(ref[1]) : auto;
          auto = c + 1;
          const t = (/\bt="(\w+)"/.exec(attrs) || [0, 'n'])[1];
          const s = +((/\bs="(\d+)"/.exec(attrs) || [0, 0])[1]);
          const body = cm[2] || '';
          const v = /<v>([\s\S]*?)<\/v>/.exec(body);
          let val = null;
          if (t === 's' && v) val = shared[+v[1]] ?? '';
          else if (t === 'inlineStr') val = textOf(body);
          else if (t === 'str' && v) val = xmlUnesc(v[1]);
          else if (t === 'b' && v) val = v[1] === '1';
          else if (v) {
            const n = Number(v[1]);
            val = dateXf[s] && isFinite(n) ? { date: serialToIso(n) } : n;
          }
          row[c] = val;
        }
        rows[rIdx] = row;
      }
      for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
      sheets.push({ name, rows });
    }
    return sheets;
  }

  const api = { write, read, zip, unzip, crc32, isoToSerial, serialToIso, colName };
  root.RotaXlsx = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
