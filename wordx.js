'use strict';
// Xuất Word (.docx): ghép NGUYÊN VĂN đoạn Word gốc của từng câu (data/word_parts.json, sinh bởi tools/tao_word_parts.py)
// vào khung (kiểu chữ, đánh số, hình...) của chính file Word nguồn. Không cần cài gói: tự đọc/ghi zip bằng zlib.
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

let P = null, SRC = null;   // dữ liệu nạp khi cần lần đầu

// ---------- zip ----------
function readZip(buf) {
  let e = buf.length - 22;
  while (e >= 0 && buf.readUInt32LE(e) !== 0x06054b50) e--;
  if (e < 0) throw new Error('zip hỏng');
  const n = buf.readUInt16LE(e + 10);
  let p = buf.readUInt32LE(e + 16);
  const out = new Map();
  for (let i = 0; i < n; i++) {
    const method = buf.readUInt16LE(p + 10), csize = buf.readUInt32LE(p + 20);
    const nl = buf.readUInt16LE(p + 28), xl = buf.readUInt16LE(p + 30), cl = buf.readUInt16LE(p + 32), lo = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nl);
    const start = lo + 30 + buf.readUInt16LE(lo + 26) + buf.readUInt16LE(lo + 28);
    out.set(name, { method, raw: buf.subarray(start, start + csize) });
    p += 46 + nl + xl + cl;
  }
  return out;
}
function entry(name) {
  const x = SRC.get(name);
  if (!x) return null;
  return x.method === 8 ? zlib.inflateRawSync(x.raw) : Buffer.from(x.raw);
}
const CRC = (() => { const t = new Int32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c; } return t; })();
function crc32(b) { let c = -1; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ -1) >>> 0; }
function writeZip(files) {   // files: [{name, data: Buffer}]
  const parts = [], cen = [];
  let off = 0;
  for (const f of files) {
    const name = Buffer.from(f.name, 'utf8'), comp = zlib.deflateRawSync(f.data), crc = crc32(f.data);
    const h = Buffer.alloc(30);
    h.writeUInt32LE(0x04034b50, 0); h.writeUInt16LE(20, 4); h.writeUInt16LE(0x0800, 6); h.writeUInt16LE(8, 8);
    h.writeUInt32LE(0, 10); h.writeUInt32LE(crc, 14); h.writeUInt32LE(comp.length, 18); h.writeUInt32LE(f.data.length, 22);
    h.writeUInt16LE(name.length, 26); h.writeUInt16LE(0, 28);
    parts.push(h, name, comp);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(0x0800, 8); c.writeUInt16LE(8, 10);
    c.writeUInt32LE(0, 12); c.writeUInt32LE(crc, 16); c.writeUInt32LE(comp.length, 20); c.writeUInt32LE(f.data.length, 24);
    c.writeUInt16LE(name.length, 28); c.writeUInt32LE(off, 42);
    cen.push(c, name);
    off += 30 + name.length + comp.length;
  }
  const cd = Buffer.concat(cen), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(off, 16);
  return Buffer.concat(parts.concat([cd, end]));
}

// ---------- đoạn văn tự tạo (tiêu đề, phần, dòng họ tên) ----------
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function para(txt, o = {}) {
  const rpr = (o.b ? '<w:b/>' : '') + (o.i ? '<w:i/>' : '') + '<w:sz w:val="' + (o.sz || 23) + '"/><w:szCs w:val="' + (o.sz || 23) + '"/>' + (o.color ? '<w:color w:val="' + o.color + '"/>' : '');
  const ppr = (o.center ? '<w:jc w:val="center"/>' : '') + '<w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after == null ? 60 : o.after) + '"/>' + (o.keep ? '<w:keepNext/>' : '') + (o.pb ? '<w:pageBreakBefore/>' : '');
  return '<w:p><w:pPr>' + ppr + '</w:pPr><w:r>' + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(txt) + '</w:t></w:r></w:p>';
}
const PH = { mc: 'PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn', tf: 'PHẦN II. Câu trắc nghiệm đúng sai', short: 'PHẦN III. Câu trắc nghiệm trả lời ngắn', view: 'Tự luận' };

// File gốc bật "giữ với đoạn sau" (keepNext) cho MỌI đoạn: ghép nhiều câu thì cả chuỗi dính liền, Word đẩy sang trang mới
// để lại trang trắng. Giữ keepNext bên trong một câu (câu không bị cắt ngang trang) nhưng bỏ ở phần tử cuối của câu/lời giải.
function lastTopStart(x) {
  const re = /<(\/?)w:(p|tbl)\b[^>]*?(\/?)>/g;
  let m, depth = 0, last = 0;
  while ((m = re.exec(x))) {
    if (m[3]) { if (!depth) last = m.index; continue; }
    if (m[1]) depth--; else { if (!depth) last = m.index; depth++; }
  }
  return last;
}
function unchain(x) {
  if (!x) return x;
  x = x.replace(/<w:pageBreakBefore\/>/g, '');
  const i = lastTopStart(x);
  return x.slice(0, i) + x.slice(i).replace(/<w:keepNext\/>/g, '');
}

// ---------- sắp lại bố cục câu giống bản in giấy ----------
// Nhiều lựa chọn: A B C D trên một hàng (phương án ngắn) hoặc lưới 2 cột; đúng sai: bảng "Phát biểu | Đúng | Sai".
// Nội dung từng phương án vẫn là run Word gốc (công thức Equation, hình... giữ nguyên). Bỏ mã câu [B8-NB-I-21] khỏi đề.
// Cỡ chữ: chữ 13 pt của file nguồn -> 11.5 pt, các cỡ khác đổi theo cùng tỉ lệ (cỡ ghi theo thang nguồn rồi mới đổi).
const SZ = 23;                       // 11.5 pt (đơn vị nửa point)
const scaleSz = x => x && x.replace(/<w:(sz|szCs) w:val="(\d+)"\/>/g, (m, t, v) => '<w:' + t + ' w:val="' + Math.max(2, Math.round(v * SZ / 26)) + '"/>');
const TEXT_W = 11906 - 1134 * 2;     // bề rộng vùng chữ (A4, lề 2 cm của file nguồn)

function kidsOf(x) {                 // các phần tử cấp ngoài cùng của một đoạn XML
  const out = [], re = /<(\/?)([\w:]+)\b[^>]*?(\/?)>/g;
  let m, d = 0, s = 0;
  while ((m = re.exec(x))) {
    if (m[3]) { if (!d) out.push(x.slice(m.index, re.lastIndex)); continue; }
    if (m[1]) { d--; if (!d) out.push(x.slice(s, re.lastIndex)); } else { if (!d) s = m.index; d++; }
  }
  return out;
}
const nameOf = el => /^<([\w:]+)/.exec(el)[1];
const innerOf = el => /\/>$/.test(el) ? '' : el.slice(el.indexOf('>') + 1, el.lastIndexOf('</'));
const unesc = s => s.replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d))
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const DROP = /^w:(fldChar|instrText|bookmarkStart|bookmarkEnd|proofErr|lastRenderedPageBreak|permStart|permEnd|commentRangeStart|commentRangeEnd)$/;

// tách đoạn XML thành khối (đoạn/bảng), mỗi đoạn là dãy "nguyên tử" (chữ cắt được, tab, xuống dòng, công thức, hình...)
function parse(x) {
  const blocks = [];
  for (const el of kidsOf(x)) {
    if (nameOf(el) !== 'w:p') { blocks.push({ tbl: el }); continue; }
    const b = { ppr: '', atoms: [] };
    for (const c of kidsOf(innerOf(el))) {
      const cn = nameOf(c);
      if (cn === 'w:pPr') { b.ppr = c; continue; }
      if (DROP.test(cn)) continue;
      if (cn !== 'w:r') { b.atoms.push({ raw: c, c: /^(m:oMath|m:oMathPara|w:hyperlink|w:fldSimple|w:smartTag|w:ins|w:sdt)$/.test(cn) ? '￼' : '' }); continue; }
      let rpr = '';
      for (const rc of kidsOf(innerOf(c))) {
        const rn = nameOf(rc);
        if (rn === 'w:rPr') { rpr = rc; continue; }
        if (DROP.test(rn)) continue;
        if (rn === 'w:t') b.atoms.push({ rpr, t: unesc(innerOf(rc)) });
        else b.atoms.push({ rpr, x: rc, c: rn === 'w:tab' ? '\t' : (rn === 'w:br' || rn === 'w:cr') ? '\n' : rn === 'w:noBreakHyphen' ? '-' : /^w:(drawing|object|pict|sym)$/.test(rn) ? '￼' : '' });
      }
    }
    blocks.push(b);
  }
  // chuỗi ký tự tương ứng: mỗi đoạn kết thúc bằng '\n', bảng là '￼\n'
  let S = '';
  const ends = [];
  blocks.forEach((b, i) => {
    if (b.tbl) { b.at = S.length; S += '￼\n'; ends[i] = S.length - 1; return; }
    for (const a of b.atoms) { a.at = S.length; S += a.t != null ? a.t : a.c; }
    ends[i] = S.length; S += '\n';
  });
  return { blocks, S, ends };
}

// lấy phần nội dung thuộc các khoảng [s, e) -> dãy khối {ppr, xml} | {tbl}
function slice(D, ranges) {
  const inR = p => ranges.some(([s, e]) => p >= s && p < e);
  const out = [];
  D.blocks.forEach((b, i) => {
    if (b.tbl) { if (inR(b.at)) out.push({ tbl: b.tbl }); return; }
    let xml = '', any = inR(D.ends[i]);
    for (const a of b.atoms) {
      if (a.t != null) {
        let t = '';
        for (const [s, e] of ranges) { const lo = Math.max(s, a.at), hi = Math.min(e, a.at + a.t.length); if (hi > lo) t += a.t.slice(lo - a.at, hi - a.at); }
        if (t) { any = true; xml += '<w:r>' + a.rpr + '<w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>'; }
      } else if (inR(a.at)) { any = true; xml += a.raw != null ? a.raw : '<w:r>' + a.rpr + a.x + '</w:r>'; }
    }
    if (any) out.push({ ppr: b.ppr, xml });
  });
  return out;
}
const trim = (S, s, e) => { while (s < e && /[\s ]/.test(S[s])) s++; while (e > s && /[\s ]/.test(S[e - 1])) e--; return [s, e]; };
const blocksXml = bs => bs.map(b => b.tbl || '<w:p>' + b.ppr + b.xml + '</w:p>').join('');
// đoạn "Câu N." sát lề trái như bản in (bỏ thụt đầu dòng riêng của vài câu trong file nguồn)
const stemXml = bs => { if (bs[0] && !bs[0].tbl) bs[0].ppr = bs[0].ppr.replace(/<w:ind\b[^>]*\/>/, ''); return blocksXml(bs); };

// vị trí các nhãn phương án (A. B. C. D. hoặc a) b) c) d)), tìm lần lượt; nhãn đầu phải ở đầu dòng/sau tab
function markers(S, from, labs) {
  const out = [];
  let pos = from;
  for (let k = 0; k < labs.length; k++) {
    const L = labs[k].replace(/[.)]/g, '\\$&');
    const strict = new RegExp('(^|[\\n\\t])[ \\u00a0]*(' + L + ')', 'g'), loose = new RegExp('([\\s\\u00a0])(' + L + ')', 'g');
    let m = null;
    for (const re of k ? [strict, loose] : [strict]) { re.lastIndex = pos; m = re.exec(S); if (m) break; }
    if (!m) return null;
    const at = m.index + m[0].length - m[2].length;
    out.push({ at, end: at + m[2].length });
    pos = at + m[2].length;
  }
  return out;
}

const cellPPr = (keep, jc) => '<w:pPr>' + (keep ? '<w:keepNext/>' : '') + '<w:spacing w:before="0" w:after="0"/><w:ind w:left="0" w:right="0" w:firstLine="0"/>' + (jc ? '<w:jc w:val="' + jc + '"/>' : '') + '</w:pPr>';
const run = (t, rpr) => '<w:r>' + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
const RPR = (o = {}) => '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' + (o.b ? '<w:b/><w:bCs/>' : '') + '<w:sz w:val="' + (o.sz || 26) + '"/><w:szCs w:val="' + (o.sz || 26) + '"/>';
const boldRpr = rpr => (rpr || '').replace(/<\/?w:rPr>/g, '').replace(/<w:b\/>|<w:bCs\/>|<w:b w:val="[^"]*"\/>|<w:bCs w:val="[^"]*"\/>/g, '').replace(/^(<w:rStyle[^>]*\/>)?(<w:rFonts[^>]*\/>)?/, '$1$2<w:b/><w:bCs/>');
// nội dung một ô: nhãn in đậm (kiểu chữ của chính nhãn gốc) + các đoạn của phương án
function cellBody(D, lab, rng, keep, labRpr) {
  const bs = rng ? slice(D, [rng]) : [];
  const head = lab ? '<w:r><w:rPr>' + boldRpr(labRpr) + '</w:rPr><w:t xml:space="preserve">' + esc(lab) + ' </w:t></w:r>' : '';
  if (!bs.length) return '<w:p>' + cellPPr(keep) + head + '</w:p>';
  return bs.map((b, i) => '<w:p>' + cellPPr(keep) + (i ? '' : head) + b.xml + '</w:p>').join('');
}
const tc = (w, body, extra) => '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + (extra || '') + '</w:tcPr>' + body + '</w:tc>';
function table(widths, rows, o) {
  const B = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => '<w:' + s + ' w:val="' + (o.border ? 'single' : 'nil') + '" w:sz="4" w:space="0" w:color="000000"/>').join('') + '</w:tblBorders>';
  return '<w:tbl><w:tblPr><w:tblW w:w="' + widths.reduce((a, b) => a + b, 0) + '" w:type="dxa"/><w:tblInd w:w="' + (o.ind || 0) + '" w:type="dxa"/>' + B +
    '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="' + (o.padV || 0) + '" w:type="dxa"/><w:left w:w="' + (o.padH || 0) + '" w:type="dxa"/><w:bottom w:w="' + (o.padV || 0) + '" w:type="dxa"/><w:right w:w="' + (o.padH || 0) + '" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr>' +
    '<w:tblGrid>' + widths.map(w => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>' +
    rows.map(r => '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + r + '</w:tr>').join('') + '</w:tbl>';
}
const GAP = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>';

function relayout(stem, q) {
  const D = parse(stem), S = D.S;
  // bỏ mã câu "[B8-NB-I-21 (*)] " ngay sau "Câu ⟦N⟧."
  const cm = /\[[A-Z]\d+-[A-Z]+-[IVX]+-\d+[^\]\n]{0,8}\][  ]*/.exec(S.slice(0, 60));
  const cEnd = cm ? cm.index + cm[0].length : 0;
  const keepR = e => cm ? [[0, cm.index], [cEnd, e]] : [[0, e]];
  const whole = () => { const x = stemXml(slice(D, keepR(S.length))); return /<\/w:tbl>$/.test(x) ? x + GAP : x; };
  const labs = q.type === 'mc' ? ['A.', 'B.', 'C.', 'D.'] : q.type === 'tf' ? ['a)', 'b)', 'c)', 'd)'] : null;
  if (!labs || (q.options || []).length !== 4) return whole();
  const mk = markers(S, cEnd, labs);
  if (!mk) return whole();
  if (q.type === 'tf') {      // một số câu có thêm ý e) f) g)... mỗi ý ở đầu một dòng: thêm hàng
    for (const L of 'efghijklmnop') {
      const m = new RegExp('[\\n\\t][ \\u00a0]*(' + L + '\\))', 'g');
      m.lastIndex = mk[mk.length - 1].end;
      const r = m.exec(S);
      if (!r || !D.ends.some(e => e > mk[mk.length - 1].at && e <= r.index)) break;
      const at = r.index + r[0].length - 2;
      mk.push({ at, end: at + 2 }); labs.push(L + ')');
    }
  }
  const N = mk.length;
  const lastEnd = D.ends.find(e => e > mk[N - 1].at);          // hết đoạn chứa phương án cuối
  // có bảng nằm trong phần phương án thì giữ nguyên bố cục gốc
  if (D.blocks.some(b => b.tbl && b.at > mk[0].at && b.at < lastEnd)) return whole();
  const opts = mk.map((m, k) => trim(S, m.end, k < N - 1 ? mk[k + 1].at : lastEnd));
  const stemR = keepR(mk[0].at);
  stemR[stemR.length - 1] = trim(S, stemR[stemR.length - 1][0], stemR[stemR.length - 1][1]);
  const afterR = trim(S, lastEnd + 1, S.length);
  const after = afterR[1] > afterR[0] ? slice(D, [afterR]) : [];
  const atoms = D.blocks.flatMap(b => b.atoms || []);
  const labRpr = k => { const a = atoms.find(x => x.t != null && mk[k].at >= x.at && mk[k].at < x.at + x.t.length); return a ? a.rpr : ''; };
  const has = (k, re) => atoms.some(a => a.x && re.test(a.x) && a.at >= opts[k][0] && a.at < opts[k][1]);
  const body = k => opts[k][1] > opts[k][0] ? opts[k] : null;
  const keepLast = after.length > 0;
  let tbl;
  if (q.type === 'mc') {
    const two = q.options.some(o => String(o).length > 38 || /img|fig/.test(o)) || [0, 1, 2, 3].some(k => has(k, /^<w:(drawing|object|pict)/));
    const IND = 210, n = two ? 2 : 4, w = Math.floor((TEXT_W - IND) / n), ws = Array(n).fill(w);
    const cells = mk.map((m, k) => tc(w, cellBody(D, labs[k], body(k), two ? (k < 2 || keepLast) : keepLast, labRpr(k))));
    tbl = table(ws, two ? [cells[0] + cells[1], cells[2] + cells[3]] : [cells.join('')], { ind: IND, padH: 40 });
  } else {
    const wL = 520, wB = 820, wS = TEXT_W - wL - 2 * wB, ws = [wL, wS, wB, wB];
    const box = keep => '<w:p>' + cellPPr(keep, 'center') + run('☐', RPR({ sz: 29 })) + '</w:p>';
    const th = (w, t) => tc(w, '<w:p>' + cellPPr(1, 'center') + (t ? run(t, RPR({ b: 1 })) : '') + '</w:p>');
    const rows = [th(wL, '') + th(wS, 'Phát biểu') + th(wB, 'Đúng') + th(wB, 'Sai')].concat(mk.map((m, k) => {
      const keep = k < N - 1 || keepLast;
      return tc(wL, '<w:p>' + cellPPr(keep) + '<w:r><w:rPr>' + boldRpr(labRpr(k)) + '</w:rPr><w:t>' + labs[k] + '</w:t></w:r></w:p>') +
        tc(wS, cellBody(D, '', body(k), keep, '')) + tc(wB, box(keep)) + tc(wB, box(keep));
    }));
    tbl = table(ws, rows, { border: 1, padH: 90, padV: 15 });
  }
  return stemXml(slice(D, stemR)) + tbl + (after.length ? blocksXml(after) : GAP);
}

// ids: mã câu theo đúng thứ tự cần xuất; byId: mã -> câu (cfg); opt: {title, sub, ans, groupName(q)}
function build(root, ids, byId, opt) {
  if (!P) {
    P = JSON.parse(fs.readFileSync(path.join(root, 'data', 'word_parts.json'), 'utf8'));
    for (const [id, it] of Object.entries(P.items)) {
      it.stem = unchain(it.stem); it.ans = unchain(it.ans);
      if (byId[id]) it.stem = relayout(it.stem, byId[id]);
      it.stem = scaleSz(it.stem); it.ans = scaleSz(it.ans);
      for (const k of ['intro1', 'intro2']) if (it[k]) it[k] = scaleSz(it[k].replace(/<w:pageBreakBefore\/>/g, ''));
    }
  }
  if (!SRC) SRC = readZip(fs.readFileSync(path.join(root, P.src)));
  const num = {};
  ids = ids.filter(id => P.items[id]);   // câu không có đoạn Word gốc (ví dụ câu ghép từ trang Ôn tập) thì bỏ, đánh số liền mạch
  ids.forEach((id, i) => { num[id] = i + 1; });
  const fill = (x, n, a, b) => x.split('⟦N⟧').join(n).split('⟦A⟧').join(a).split('⟦B⟧').join(b);
  let body = para(opt.title || '', { b: 1, sz: 30, center: 1, after: 40 });
  if (opt.sub) body += para(opt.sub, { center: 1, after: 40 });
  body += para(ids.length + ' câu' + (opt.ans ? ' · BẢN CÓ ĐÁP ÁN (giáo viên)' : ''), { center: 1, i: 1, after: 120 });
  if (!opt.ans) body += para('Họ và tên: .................................................................... Lớp: ...............', { after: 160 });
  let last = null;
  const done = new Set();
  for (const id of ids) {
    const q = byId[id], it = P.items[id];
    if (!q || !it) continue;
    const key = q.type + (opt.groupName ? '|' + q.group : '');
    if (key !== last) {
      body += para((opt.groupName ? opt.groupName(q) + ' · ' : '') + (PH[q.type] || ''), { b: 1, before: 160, after: 60, keep: 1 });
      last = key;
    }
    if (it.intro2 && !done.has(it.pair)) {   // câu chùm: dữ kiện chung đặt trước câu đầu tiên của cặp
      const pn = num[it.pair];
      body += pn && pn === num[id] + 1 ? fill(it.intro2, '', num[id], pn) : fill(it.intro1, '', num[id], '');
    }
    body += fill(it.stem, num[id], '', '');
    if (q.type === 'short' && !opt.ans) body += para('Đáp số: ....................................', { after: 120 });
    done.add(id);
  }
  if (opt.ans) {
    body += para('ĐÁP ÁN VÀ HƯỚNG DẪN', { b: 1, sz: 26, center: 1, pb: 1, after: 120 });
    for (const id of ids) {
      const it = P.items[id];
      if (!it) continue;
      body += para('Câu ' + num[id] + ' [' + id + ']', { b: 1, color: '555555', before: 80, after: 20, keep: 1 });
      body += it.ans ? fill(it.ans, num[id], '', '') : para('(Nguồn không có đáp án)', { i: 1 });
    }
  }
  const docXml = P.head + body + P.tail;
  // chỉ giữ các quan hệ (hình, đối tượng nhúng...) mà nội dung thật sự dùng
  const used = new Set();
  docXml.replace(/r:(?:embed|id|link|pict)="(rId\d+)"/g, (m, r) => used.add(r));
  const relsXml = entry('word/_rels/document.xml.rels').toString('utf8');
  const keepT = new Set();
  const rels = relsXml.replace(/<Relationship [^>]*\/>/g, r => {
    const id = /Id="([^"]+)"/.exec(r)[1], t = /Target="([^"]+)"/.exec(r)[1];
    const isMedia = /^(media|embeddings)\//.test(t);
    if (isMedia && !used.has(id)) return '';
    if (isMedia) keepT.add('word/' + t);
    return r;
  });
  const files = [];
  for (const name of SRC.keys()) {
    if (name === 'word/document.xml') files.push({ name, data: Buffer.from(docXml, 'utf8') });
    else if (name === 'word/_rels/document.xml.rels') files.push({ name, data: Buffer.from(rels, 'utf8') });
    else if (name === 'word/styles.xml') files.push({ name, data: Buffer.from(entry(name).toString('utf8').replace(/(<w:style w:type="paragraph" w:default="1"[\s\S]*?<\/w:style>)/, s => s.replace(/<w:(sz|szCs) w:val="\d+"\/>/g, '<w:$1 w:val="23"/>')), 'utf8') });
    else if (/^word\/(media|embeddings)\//.test(name)) { if (keepT.has(name)) files.push({ name, data: entry(name) }); }
    else files.push({ name, data: entry(name) });
  }
  return writeZip(files);
}

module.exports = { build };
