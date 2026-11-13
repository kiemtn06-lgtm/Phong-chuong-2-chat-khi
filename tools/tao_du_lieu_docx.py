# -*- coding: utf-8 -*-
"""Chuyển 2 file Word của Chương 2 - Chất khí sang dữ liệu phòng làm bài.

  nguon-goc/CHUONG2_KHI_LI_TUONG_He_thong_theo_bai_va_muc_do.docx   -> ngân hàng tổng (câu, đáp án, lời giải theo mã câu)
  nguon-goc/CHUONG2_KHI_LI_TUONG_Bo_on_theo_buoi.docx               -> chia câu thành 6 buổi (và đối chiếu đáp án)

Ra: data/cfg.json, data/key.json, public/img/*, tools/bao_cao_chuyen_doi.txt

Chạy:  python tools/tao_du_lieu_docx.py [--nguon <thư mục chứa 2 file .docx>]
Cần:   pip install python-docx
Nguyên tắc: giữ NGUYÊN VĂN câu hỏi, phương án, đáp án, lời giải. Công thức Word (OMML) được đổi sang LaTeX,
không viết lại chữ nào. Gặp cấu trúc lạ thì dừng và báo lỗi thay vì đoán.
"""
import sys, os, re, json, hashlib, argparse
from docx import Document
from docx.oxml.ns import qn

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOI = []   # các điểm cần xem lại
try:
    with open(os.path.join(ROOT, 'tools', 'mo_phong.json'), encoding='utf-8') as _f:
        MO_PHONG = {k: v for k, v in json.load(_f).items() if not k.startswith('_')}
except FileNotFoundError:
    MO_PHONG = {}


def loi(msg):
    LOI.append(msg)


# ---------------------------------------------------------------- công thức OMML -> LaTeX
SYM = {'⋅': r'\cdot ', '∙': r'\cdot ', '·': r'\cdot ', '×': r'\times ', '≈': r'\approx ', '≤': r'\le ', '≥': r'\ge ',
       '≠': r'\ne ', '±': r'\pm ', '→': r'\to ', '⇒': r'\Rightarrow ', '⇔': r'\Leftrightarrow ', '∞': r'\infty ',
       '−': '-', '–': '-', '—': '-', '°': r'^{\circ}', '∘': r'^{\circ}', '∝': r'\propto ', '÷': r'\div ', '∆': r'\Delta ',
       'Δ': r'\Delta ', 'α': r'\alpha ', 'β': r'\beta ', 'γ': r'\gamma ', 'δ': r'\delta ', 'ε': r'\varepsilon ',
       'θ': r'\theta ', 'λ': r'\lambda ', 'μ': r'\mu ', 'π': r'\pi ', 'ρ': r'\rho ', 'σ': r'\sigma ', 'τ': r'\tau ',
       'φ': r'\varphi ', 'ω': r'\omega ', 'Ω': r'\Omega ', 'η': r'\eta ', 'ν': r'\nu ', 'κ': r'\kappa ', 'ξ': r'\xi ',
       'ψ': r'\psi ', 'Σ': r'\Sigma ', 'Φ': r'\Phi ', '∑': r'\sum ', '√': r'\sqrt{}', '∼': r'\sim ', '~': r'\sim ',
       '∶': ':', '⁄': '/', '′': "'", '″': "''", '∈': r'\in ', '⊂': r'\subset ', '∥': r'\parallel ', '⊥': r'\perp ',
       '↑': r'\uparrow ', '↓': r'\downarrow ', '←': r'\leftarrow ', '⟶': r'\to ', '⟹': r'\Rightarrow ', ' ': r'\ ',
       ' ': r'\ ', '{': r'\{', '}': r'\}', '%': r'\%', '#': r'\#', '&': r'\&', '_': r'\_', '$': r'\$'}
ACC = {'̅': r'\overline', '¯': r'\overline', '‾': r'\overline', '⃗': r'\vec', '→': r'\vec',
       '̂': r'\hat', '̃': r'\tilde', '̇': r'\dot', '̈': r'\ddot', 'ˉ': r'\overline'}


def lname(el):
    return el.tag.split('}')[1] if '}' in el.tag else el.tag


def m_text(s, plain):
    out = []
    for ch in s:
        out.append(SYM.get(ch, ch))
    t = ''.join(out)
    t = re.sub(r'(\d),(\d)', r'\1{,}\2', t)        # dấu phẩy thập phân
    if plain and re.search(r'[A-Za-zÀ-ỹ]', t):
        t2 = re.sub(r'\\$', '', t.strip())             # khoảng trắng cuối bị ánh xạ thành "\ " rồi bị cắt mất dấu cách
        return r'\mathrm{' + t2 + '}' if t2 else t
    return t


def m_children(el):
    return ''.join(m_node(c) for c in el)


def m_arg(el, name):
    for c in el:
        if lname(c) == name:
            return m_children(c)
    return ''


def wrap(s):
    return '{' + s + '}'


def m_node(el):
    n = lname(el)
    if n in ('ctrlPr', 'rPr', 'fPr', 'sSubPr', 'sSupPr', 'sSubSupPr', 'radPr', 'dPr', 'accPr', 'barPr', 'funcPr',
             'naryPr', 'limLowPr', 'limUppPr', 'eqArrPr'):
        return ''
    if n == 'r':
        txt = ''.join(t.text or '' for t in el if lname(t) == 't')
        plain = False
        rpr = [c for c in el if lname(c) == 'rPr']
        if rpr:
            for c in rpr[0]:
                if lname(c) == 'sty' and c.get(M + 'val') in ('p',):
                    plain = True
                if lname(c) == 'nor':
                    plain = True
        return m_text(txt, plain)
    if n == 'f':
        return r'\dfrac' + wrap(m_arg(el, 'num')) + wrap(m_arg(el, 'den'))
    if n == 'sSub':
        return wrap(m_arg(el, 'e')) + '_' + wrap(m_arg(el, 'sub'))
    if n == 'sSup':
        return wrap(m_arg(el, 'e')) + '^' + wrap(m_arg(el, 'sup'))
    if n == 'sSubSup':
        return wrap(m_arg(el, 'e')) + '_' + wrap(m_arg(el, 'sub')) + '^' + wrap(m_arg(el, 'sup'))
    if n == 'rad':
        deg = m_arg(el, 'deg')
        return r'\sqrt' + ('[' + deg + ']' if deg.strip() else '') + wrap(m_arg(el, 'e'))
    if n == 'd':
        beg, end, sep = '(', ')', '|'
        for c in el:
            if lname(c) == 'dPr':
                for p in c:
                    if lname(p) == 'begChr':
                        beg = p.get(M + 'val', '')
                    elif lname(p) == 'endChr':
                        end = p.get(M + 'val', '')
                    elif lname(p) == 'sepChr':
                        sep = p.get(M + 'val', '|')
        parts = [m_children(c) for c in el if lname(c) == 'e']
        esc = lambda ch: {'{': r'\{', '}': r'\}', '': '.'}.get(ch, ch)
        return r'\left' + esc(beg) + sep.join(parts) + r'\right' + esc(end)
    if n == 'acc':
        chr_ = '̂'
        for c in el:
            if lname(c) == 'accPr':
                for p in c:
                    if lname(p) == 'chr':
                        chr_ = p.get(M + 'val', chr_)
        if chr_ not in ACC:
            raise ValueError('dấu accent lạ: %r' % chr_)
        return ACC[chr_] + wrap(m_arg(el, 'e'))
    if n == 'bar':
        pos = 'top'
        for c in el:
            if lname(c) == 'barPr':
                for p in c:
                    if lname(p) == 'pos':
                        pos = p.get(M + 'val', 'top')
        return (r'\overline' if pos == 'top' else r'\underline') + wrap(m_arg(el, 'e'))
    if n == 'func':
        return r'\mathrm' + wrap(m_arg(el, 'fName')) + m_arg(el, 'e')
    if n in ('e', 'num', 'den', 'sub', 'sup', 'deg', 'fName', 'oMath', 'oMathPara'):
        return m_children(el)
    if n == 'nary':
        chr_ = '∫'
        for c in el:
            if lname(c) == 'naryPr':
                for p in c:
                    if lname(p) == 'chr':
                        chr_ = p.get(M + 'val', chr_)
        op = {'∑': r'\sum', '∫': r'\int', '∏': r'\prod'}.get(chr_)
        if not op:
            raise ValueError('toán tử nary lạ: %r' % chr_)
        s = m_arg(el, 'sub'); u = m_arg(el, 'sup')
        return op + ('_' + wrap(s) if s else '') + ('^' + wrap(u) if u else '') + wrap(m_arg(el, 'e'))
    raise ValueError('phần tử công thức chưa hỗ trợ: ' + n)


def omml(el):
    s = m_node(el)
    s = re.sub(r'\s+', ' ', s).strip()
    s = re.sub(r'(?:\\ )+(?=\}|$)', '', s)                     # khoảng cách thừa ở cuối nhóm
    s = re.sub(r'[À-ỹ]+', lambda m: '\\text{%s}' % m.group(0), s)   # chữ Việt có dấu trong công thức
    return s


# ---------------------------------------------------------------- đoạn văn -> chuỗi giàu định dạng
def esc_text(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def run_fmt(r):
    rpr = r.find(W + 'rPr')
    f = {'b': False, 'i': False, 'sup': False, 'sub': False}
    if rpr is not None:
        b = rpr.find(W + 'b')
        if b is not None and b.get(W + 'val', '1') not in ('0', 'false'):
            f['b'] = True
        i = rpr.find(W + 'i')
        if i is not None and i.get(W + 'val', '1') not in ('0', 'false'):
            f['i'] = True
        va = rpr.find(W + 'vertAlign')
        if va is not None:
            v = va.get(W + 'val')
            if v == 'superscript':
                f['sup'] = True
            elif v == 'subscript':
                f['sub'] = True
    return (f['b'], f['i'], f['sup'], f['sub'])


class Doc:
    def __init__(self, path, img_dir, tag):
        self.d = Document(path)
        self.path = path
        self.img_dir = img_dir
        self.tag = tag
        self.imgs = {}
        os.makedirs(img_dir, exist_ok=True)

    def image_file(self, rid):
        if rid in self.imgs:
            return self.imgs[rid]
        part = self.d.part.related_parts[rid]
        ct = part.content_type
        ext = {'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp'}.get(ct)
        if not ext:
            loi('Ảnh định dạng %s chưa hỗ trợ (rId %s)' % (ct, rid))
            self.imgs[rid] = None
            return None
        h = hashlib.sha1(part.blob).hexdigest()[:10]
        name = 'ch2_%s.%s' % (h, ext)
        with open(os.path.join(self.img_dir, name), 'wb') as f:
            f.write(part.blob)
        self.imgs[rid] = name
        return name

    def segs(self, p):
        """Đoạn văn -> danh sách phần tử ('t', text, fmt) | ('m', latex) | ('img', file) | ('flat', text) ."""
        out = []

        def walk(node):
            for c in node:
                n = lname(c)
                if n == 'r':
                    fm = run_fmt(c)
                    for ch in c:
                        cn = lname(ch)
                        if cn == 't':
                            if ch.text:
                                out.append(('t', ch.text, fm))
                        elif cn == 'tab':
                            out.append(('t', '\t', fm))
                        elif cn == 'br':
                            if ch.get(W + 'type') != 'page':
                                out.append(('t', '\n', fm))
                        elif cn == 'drawing':
                            for bl in ch.iter(A + 'blip'):
                                rid = bl.get(R_EMBED)
                                if rid:
                                    out.append(('img', self.image_file(rid)))
                        elif cn == 'object' or cn == 'pict':
                            loi('Đối tượng nhúng (object/pict) trong đoạn: %r' % ''.join(x[1] for x in out if x[0] == 't')[:60])
                elif n in ('oMath',):
                    out.append(('m', omml(c)))
                elif n == 'oMathPara':
                    for mm in c:
                        if lname(mm) == 'oMath':
                            out.append(('m', omml(mm)))
                elif n in ('hyperlink', 'ins', 'smartTag', 'sdt', 'sdtContent', 'fldSimple'):
                    walk(c)
        walk(p)
        return out


def segs_plain(segs):
    return ''.join(s[1] if s[0] == 't' else ('\u0001' if s[0] == 'm' else '') for s in segs)


def render(segs, inline=False):
    """Gộp các phần tử thành chuỗi: văn bản có thẻ <b> <i> <sup> <sub>, công thức trong $...$.
    inline=True: ảnh được chèn tại chỗ dạng [[img:tên]] (dùng cho lời giải); mặc định bỏ ảnh (ảnh của đề được gom riêng)."""
    # gộp văn bản cùng định dạng liền nhau
    merged = []
    for s in segs:
        if s[0] == 't' and merged and merged[-1][0] == 't' and merged[-1][2] == s[2]:
            merged[-1] = ('t', merged[-1][1] + s[1], s[2])
        else:
            merged.append(s)
    out = []
    for s in merged:
        if s[0] == 't':
            txt, (b, i, sup, sub) = s[1], s[2]
            if not txt.strip('\t\n ') :   # khoảng trắng thuần: không bọc thẻ
                out.append(esc_text(txt))
                continue
            core = txt.strip(' \t\n')
            lead = txt[:len(txt) - len(txt.lstrip(' \t\n'))]
            tail = txt[len(txt.rstrip(' \t\n')):]
            body = esc_text(core)
            for tag, on in (('sub', sub), ('sup', sup), ('i', i), ('b', b)):
                if on:
                    body = '<%s>%s</%s>' % (tag, body, tag)
            out.append(esc_text(lead) + body + esc_text(tail))
        elif s[0] == 'm':
            out.append('$' + s[1] + '$')
        elif s[0] == 'img' and inline and s[1]:
            out.append(' [[img:%s]] ' % s[1])
    return ''.join(out)


def blocks(doc):
    """Danh sách khối theo thứ tự: {'k':'p'|'t', 'style', 'segs'|'rows', 'imgs', 'text'}"""
    res = []
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    for ch in doc.d.element.body.iterchildren():
        if ch.tag == qn('w:p'):
            p = Paragraph(ch, doc.d)
            sg = doc.segs(ch)
            res.append({'k': 'p', 'style': p.style.name if p.style is not None else '', 'segs': sg,
                        'imgs': [s[1] for s in sg if s[0] == 'img' and s[1]], 'plain': segs_plain(sg).strip()})
        elif ch.tag == qn('w:tbl'):
            t = Table(ch, doc.d)
            rows, imgs = [], []
            for r in t.rows:
                cells = []
                seen = set()
                for c in r.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    paras = []
                    for cp in c._tc.iterchildren(qn('w:p')):
                        sg = doc.segs(cp)
                        imgs += [s[1] for s in sg if s[0] == 'img' and s[1]]
                        paras.append(sg)
                    cells.append(paras)
                rows.append(cells)
            res.append({'k': 't', 'style': 'tbl', 'rows': rows, 'imgs': imgs, 'plain': ''})
    return res


def cell_rich(paras, inline=False):
    return '\n'.join(render(p, inline) for p in paras).strip()


def cell_plain(paras):
    return '\n'.join(segs_plain(p) for p in paras).replace('\u0001', '').strip()


# ---------------------------------------------------------------- tách câu hỏi
QSTART = re.compile(r'^\s*Câu\s+(\d+)\s*\.\s*\[([A-Z]+\d+-[A-Z]+-[IV]+-\d+)(\s*\(\*\))?\]\s*', re.S)
CLUSTER = re.compile(r'^\s*Câu\s+\d+\s*[–-]\s*\d+\s*\(câu chùm\)\s*', re.S)
CODE = re.compile(r'^B(\d+)-(NB|TH|VD|VDC)-(I|II|III|IV)-(\d+)$')
LV = {'NB': 'nb', 'TH': 'th', 'VD': 'vd', 'VDC': 'vdc'}
TY = {'I': 'mc', 'II': 'tf', 'III': 'short', 'IV': 'view'}


def strip_prefix(segs, n):
    """Bỏ n ký tự văn bản đầu của danh sách phần tử (n tính trên segs_plain)."""
    res, left = [], n
    for s in segs:
        if s[0] == 't' and left > 0:
            t = s[1]
            if len(t) <= left:
                left -= len(t)
                continue
            res.append(('t', t[left:], s[2]))
            left = 0
        elif left > 0 and s[0] == 'm':
            raise ValueError('công thức nằm trong phần tiền tố')
        else:
            res.append(s)
    return res


def parse_master(bl):
    """Trả về danh sách câu (dict) theo thứ tự tài liệu."""
    units, cur = [], None
    heads = {'bai': '', 'muc': '', 'dang': ''}
    for i, b in enumerate(bl):
        if b['k'] == 'p' and b['style'].startswith('Heading'):
            if b['style'] == 'Heading 1':
                heads['bai'] = b['plain']
            elif b['style'] == 'Heading 2':
                heads['muc'] = b['plain']
            elif b['style'] == 'Heading 3':
                heads['dang'] = b['plain']
            cur = None
            continue
        pl = b['plain'] if b['k'] == 'p' else ''
        m = QSTART.match(pl) if b['k'] == 'p' else None
        mc = CLUSTER.match(pl) if b['k'] == 'p' else None
        if m:
            cur = {'code': m.group(2), 'n': int(m.group(1)), 'star': bool(m.group(3)), 'blocks': [b], 'prefix': len(m.group(0)),
                   'heads': dict(heads), 'idx': i, 'cluster': None}
            units.append(cur)
        elif mc:
            cur = {'cluster_header': b, 'prefix': len(mc.group(0)), 'blocks': [], 'heads': dict(heads), 'idx': i}
            units.append(cur)
        elif cur is not None:
            cur['blocks'].append(b)
    return units


def block_text_rich(b, drop=0, inline=False):
    segs = b['segs']
    if drop:
        segs = strip_prefix(segs, drop)
    return render(segs, inline).strip(' \n\t')


def split_markers(blocks_):
    """Chia nội dung một câu thành: thân (stem), phần đáp án, phần hướng dẫn, dòng nguồn."""
    stem, ans, guide, src = [], [], [], None
    mode = 'stem'
    for b in blocks_:
        pl = b['plain'] if b['k'] == 'p' else ''
        if b['k'] == 'p' and pl.startswith('Nguồn:'):
            src = b
            mode = 'end'
            continue
        if b['k'] == 'p' and re.match(r'^(?:Đáp án|Đáp số|ĐA)\s*:', pl) and mode == 'stem':
            ans.append(b)
            mode = 'afterans'
            continue
        if b['k'] == 'p' and re.match(r'^(Hướng dẫn giải|Lời giải)', pl) and mode in ('stem', 'afterans'):
            guide.append(b)
            mode = 'guide'
            continue
        {'stem': stem, 'afterans': guide, 'guide': guide, 'end': guide}[mode].append(b)
    return stem, ans, guide, src


def rich_lines(blks, first_drop=0):
    """Khối -> danh sách dòng chữ (đoạn văn là dòng; bảng được xử lý riêng)."""
    out = []
    for k, b in enumerate(blks):
        if b['k'] == 'p':
            t = block_text_rich(b, first_drop if k == 0 else 0)
            if t:
                out.append(t)
    return out


def table_html(rows):
    h = ['<table class="data">']
    for ri, r in enumerate(rows):
        h.append('<tr>' + ''.join(('<th>%s</th>' if ri == 0 else '<td>%s</td>') % cell_rich(c).replace('\n', '<br>') for c in r) + '</tr>')
    h.append('</table>')
    return ''.join(h)


OPT = re.compile(r'(?:(?<=\n)|(?<=\t)|^)[ \t]*([A-D])\.[ \t]*')


def unlabel(text):
    """Bỏ định dạng đậm quanh nhãn phương án (A. B. C. D.) và nhãn ý (a) b) c) d))."""
    text = re.sub(r'<b>\s*([A-D]\.)\s*</b>', r'\1', text)
    text = re.sub(r'<b>\s*([A-D])\s*</b>\s*\.', r'\1.', text)        # chữ đậm, dấu chấm ngoài thẻ
    text = re.sub(r'<b>\s*([a-d]\))\s*</b>', r'\1', text)
    text = re.sub(r'<b>\s*([a-d])\s*</b>\s*\)', r'\1)', text)
    text = re.sub(r'<i>\s*([A-D]\.|[a-d]\))\s*</i>', r'\1', text)
    return text


def split_options(text):
    """Tách phương án A-D theo thứ tự; trả (thân, [A,B,C,D]) hoặc None."""
    text = unlabel(text)
    pos = []
    want = 'A'
    for m in OPT.finditer(text):
        if m.group(1) == want:
            pos.append((m.start(), m.end()))
            want = chr(ord(want) + 1)
            if want == 'E':
                break
    if len(pos) != 4:
        return None
    stem = text[:pos[0][0]]
    opts = []
    for k in range(4):
        end = pos[k + 1][0] if k < 3 else len(text)
        opts.append(text[pos[k][1]:end].strip(' \n\t'))
    return stem.strip(' \n\t'), opts


def build_question(u):
    code = u['code']
    mm = CODE.match(code)
    if not mm:
        raise ValueError('mã câu lạ: ' + code)
    bai, lv, dang = int(mm.group(1)), mm.group(2), mm.group(3)
    ty = TY[dang]
    stem_b, ans_b, guide_b, src_b = split_markers(u['blocks'])
    # văn bản thân: đoạn đầu bỏ tiền tố "Câu n. [mã]"
    paras, tables, imgs, caps = [], [], [], []
    for k, b in enumerate(stem_b):
        imgs += b['imgs']
        if b['k'] == 'p':
            if b['style'] == 'Image Caption':
                caps.append(block_text_rich(b))
                continue
            t = block_text_rich(b, u['prefix'] if k == 0 else 0)
            if t:
                paras.append(t)
        else:
            tables.append(b)
    q = {'code': code, 'bai': bai, 'lv': lv, 'dang': dang, 'type': ty, 'star': u['star'], 'n': u['n'], 'heads': u['heads']}
    ans_text = '\n'.join(block_text_rich(b) for b in ans_b)
    guide_lines = []
    for b in guide_b:
        if b['k'] == 'p':
            t = block_text_rich(b, 0, True)
            if t:
                guide_lines.append(t)
        else:
            for r in b['rows']:
                cells = [cell_rich(c, True) for c in r]
                if any(cells):
                    guide_lines.append(' · '.join(x for x in cells if x))
    q['guide'] = guide_lines
    q['ans_text'] = ans_text
    q['src'] = block_text_rich(src_b).replace('<b>', '').replace('</b>', '') if src_b else ''
    q['imgs'] = imgs
    q['caps'] = caps
    # ---- theo loại
    if ty == 'mc':
        text = '\n'.join(paras)
        for t in tables:      # bảng dữ liệu nằm trong câu: giữ làm hình HTML
            q.setdefault('tables', []).append(table_html(t['rows']))
        sp = split_options(text)
        if not sp:
            raise ValueError('không tách được 4 phương án')
        q['text'], q['options'] = sp
        notag = lambda s: re.sub(r'<[^>]+>', '', s)
        mans = re.search(r'Đáp án\s*:\s*([A-D])\b', notag(ans_text)) or re.search(r'Đáp án\s*:\s*([A-D])\b', notag(' '.join(guide_lines)))
        q['answer'] = mans.group(1) if mans else None
    elif ty == 'tf':
        lines, stmts = [], []
        for k, b in enumerate(stem_b):
            if b['k'] == 'p':
                if b['style'] == 'Image Caption':
                    continue
                t = unlabel(block_text_rich(b, u['prefix'] if k == 0 else 0))
                for ln in t.split('\n'):
                    ln = ln.strip(' \t')
                    if not ln:
                        continue
                    if re.match(r'^[a-d]\)\s*', ln):
                        stmts.append(ln)
                    elif stmts:
                        stmts[-1] = stmts[-1] + ' ' + ln      # dòng nối tiếp của ý trước
                    else:
                        lines.append(ln)
            else:
                got = False
                for r in b['rows']:
                    c0 = unlabel(cell_rich(r[0])).strip() if r else ''
                    if re.match(r'^[a-d]\)', c0):
                        stmts.append(c0.replace('\n', ' '))
                        got = True
                if not got and b['rows'] and any(cell_rich(c) for r in b['rows'] for c in r):
                    q.setdefault('tables', []).append(table_html(b['rows']))
        q['text'] = '\n'.join(lines + caps_i(caps))
        q['options'] = [re.sub(r'^[a-d]\)\s*\.?\s*', '', s).strip() for s in stmts]
        notag = lambda s: re.sub(r'<[^>]+>', '', s)
        marks = re.findall(r'([a-d])\)\s*(Đ|S)\b', notag(ans_text))
        if len(marks) != 4:      # đáp án có thể nằm trong phần hướng dẫn: "Đáp án: a) S; b) S; c) Đ; d) Đ."
            for gl in guide_lines:
                g = notag(gl)
                if re.match(r'^\s*Đáp án\s*:', g):
                    marks = re.findall(r'([a-d])\)\s*(Đ|S)\b', g)
                    if len(marks) == 4:
                        break
        if len(marks) != 4:      # dòng bảng giải thích: "a) · Đ · Loại ý · ..."
            marks = []
            for gl in guide_lines:
                mt = re.match(r'^\s*([a-d])\)\s*·\s*(Đ|S)\b', notag(gl))
                if mt:
                    marks.append((mt.group(1), mt.group(2)))
        q['answer'] = [mk == 'Đ' for _, mk in marks] if len(marks) == 4 and [x for x, _ in marks] == list('abcd') else None
    else:   # short, view
        for t in tables:
            q.setdefault('tables', []).append(table_html(t['rows']))
        q['text'] = '\n'.join(paras + caps_i(caps))
        q['answer'] = None
    if ty != 'tf':
        q['text'] = q['text'] if ty == 'mc' else q['text']
    return q


def caps_i(caps):
    return ['<i>%s</i>' % c for c in caps if c]


def short_answer(q, extra=None):
    """Đáp số của câu trả lời ngắn: dòng 'Đáp án:' hoặc 'Đáp số:' hoặc 'Đáp án:' cuối lời giải."""
    cands = [re.sub(r'<[^>]+>', '', q['ans_text'])]
    for ln in (extra if extra is not None else q['guide']):
        cands.append(re.sub(r'<[^>]+>', '', re.sub(r'\$[^$]*\$', '', ln)))
    pat = re.compile(r'(?:Đáp (?:số|án)|ĐA)\s*:\s*([-−]?\s*\d[\d\.,]*)')
    for c in cands:
        found = pat.findall(c)
        if found:
            v = found[-1].replace('−', '-').replace(' ', '').rstrip('.,')
            return v
    return None


def decs(a):
    p = a.replace('.', ',').split(',')
    return len(p[1]) if len(p) > 1 else 0


def num(s):
    return float(s.replace(',', '.'))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--nguon', default=os.path.join(ROOT, 'nguon-goc'))
    ap.add_argument('--tray', action='store_true', help='chỉ phân tích và báo cáo, không ghi dữ liệu')
    args = ap.parse_args()
    f_master = os.path.join(args.nguon, 'CHUONG2_KHI_LI_TUONG_He_thong_theo_bai_va_muc_do.docx')
    f_buoi = os.path.join(args.nguon, 'CHUONG2_KHI_LI_TUONG_Bo_on_theo_buoi.docx')
    img_dir = os.path.join(ROOT, 'public', 'img')
    dm = Doc(f_master, img_dir, 'master')
    db = Doc(f_buoi, img_dir, 'buoi')
    bm, bb = blocks(dm), blocks(db)
    print('khoi:', len(bm), len(bb))
    um = parse_master(bm)
    # --- câu chùm
    questions, seen = [], set()
    i = 0
    while i < len(um):
        u = um[i]
        if 'cluster_header' in u:
            ctx = block_text_rich(u['cluster_header'], u['prefix'])
            subs = []
            j = i + 1
            while j < len(um) and 'code' in um[j] and len(subs) < 2:
                subs.append(um[j]); j += 1
            if len(subs) != 2:
                raise ValueError('câu chùm không đủ 2 câu')
            # lời giải chung nằm ở đơn vị cuối
            s2, a2, g2, src2 = split_markers(subs[1]['blocks'])
            for k, sb in enumerate(subs):
                sb2 = dict(sb)
                if k == 0:
                    sb2['blocks'] = sb['blocks']
                else:
                    sb2['blocks'] = s2
                qq = build_question(sb2)
                qq['cluster_ctx'] = ctx
                qq['cluster_imgs'] = [im for b in u['cluster_header'] and [u['cluster_header']] for im in b['imgs']]
                if k == 1:
                    qq['_g'] = (g2, src2)
                questions.append(qq)
            # chia hướng dẫn chung theo từng "Hướng dẫn:"
            g_lines = []
            for b in g2:
                if b['k'] == 'p':
                    t = block_text_rich(b, 0, True)
                    if t:
                        g_lines.append(t)
            nt = lambda s: re.sub(r'<[^>]+>', '', s).strip()
            hd = [x for x in g_lines if nt(x).startswith('Hướng dẫn:')]
            head = [x for x in g_lines if nt(x).startswith('Hướng dẫn giải')]
            if len(hd) == 2:
                questions[-2]['guide'] = head + [hd[0]]
                questions[-1]['guide'] = head + [hd[1]]
            else:
                loi('Câu chùm %s: không tách được hướng dẫn riêng, dùng chung' % subs[0]['code'])
                for qx in questions[-2:]:
                    qx['guide'] = g_lines
            srcline = block_text_rich(src2).replace('<b>', '').replace('</b>', '') if src2 else ''
            questions[-2]['src'] = questions[-1]['src'] = srcline
            i = j
            continue
        try:
            questions.append(build_question(u))
        except Exception as e:
            loi('LỖI %s: %s' % (u['code'], e))
        i += 1
    print('cau trong ngan hang tong:', len(questions))
    # --- đáp số trả lời ngắn
    for q in questions:
        if q['type'] == 'short':
            q['answer'] = short_answer(q)
            if q['answer'] is None:
                loi('Không tìm thấy đáp số: %s' % q['code'])
        if q['type'] == 'mc' and not q.get('answer'):
            loi('Thiếu đáp án trắc nghiệm: %s' % q['code'])
        if q['type'] == 'tf' and (not q.get('answer') or len(q['options']) != 4):
            loi('Đúng/sai chưa đủ 4 ý hoặc 4 đáp án: %s (ý=%d, đáp án=%s)' % (q['code'], len(q['options']), q.get('answer')))
    # --- trùng mã
    codes = {}
    for q in questions:
        codes.setdefault(q['code'], []).append(q)
    for c, l in codes.items():
        if len(l) > 1:
            loi('Mã câu trùng: %s (%d lần)' % (c, len(l)))
    # --- đối chiếu số câu theo tiêu đề "(n câu)"
    heads = {}
    for b in bm:
        if b['k'] == 'p' and b['style'] == 'Heading 3':
            m = re.search(r'\((\d+) câu\)', b['plain'])
            if m:
                heads[b['plain']] = int(m.group(1))
    cnt = {}
    for q in questions:
        cnt[q['heads']['dang']] = cnt.get(q['heads']['dang'], 0) + 1
    for h, n in heads.items():
        if cnt.get(h, 0) != n:
            loi('Số câu không khớp tiêu đề "%s": tiêu đề %d, đọc được %d' % (h, n, cnt.get(h, 0)))
    # --- bộ ôn theo buổi
    ub = parse_master(bb)
    buoi = {}       # code -> số buổi
    order_b = []    # (buổi, code)
    cur = 0
    for b in bb:
        pass
    # xác định buổi theo tiêu đề "Buổi k –" gần nhất trước mỗi câu
    cur = 0
    bidx = {}
    for i, b in enumerate(bb):
        if b['k'] == 'p' and b['style'] == 'Heading 1':
            mm = re.match(r'Buổi\s+(\d+)', b['plain'])
            if mm:
                cur = int(mm.group(1))
            else:
                cur = 0
        bidx[i] = cur
    for u in ub:
        if 'code' in u:
            bi = bidx[u['idx']]
            if bi:
                buoi.setdefault(u['code'], bi)
                order_b.append((bi, u['code']))
    print('cau trong bo on theo buoi:', len(order_b), '| ma khac nhau:', len(buoi))
    miss = [c for c in buoi if c not in codes]
    if miss:
        loi('Mã trong bộ ôn theo buổi nhưng không có trong ngân hàng tổng: %s' % ', '.join(miss[:20]))
    dup = len(order_b) - len(buoi)
    if dup:
        loi('Có %d câu xuất hiện ở nhiều buổi' % dup)
    return questions, codes, buoi, order_b, bb, ub, args


# ================================================================ xuất dữ liệu phòng
from collections import Counter

notag = lambda s: re.sub(r'<[^>]+>', '', s)


def tol_of(ans):
    a = abs(num(ans))
    return round(max(0.49 * 10 ** (-decs(ans)), 0.002 * a), 9)


def parse_buoi_answers(bb):
    """Mục 'Đáp án và hướng dẫn – Buổi k' của bộ ôn theo buổi: mã câu -> các dòng (đã render)."""
    res, cur, inans = {}, None, False
    for b in bb:
        if b['k'] != 'p':
            continue
        if b['style'].startswith('Heading'):
            inans = b['plain'].startswith('Đáp án và hướng dẫn')
            cur = None
            continue
        if not inans:
            continue
        m = re.match(r'^Câu\s+\d+\s*\[([A-Z]+\d+-[A-Z]+-[IV]+-\d+)', b['plain'])
        if m:
            cur = m.group(1)
            res[cur] = []
        elif cur:
            res[cur].append(block_text_rich(b, 0, True))
    return res


def main_xuat():
    questions, codes, buoi, order_b, bb, ub, args = main()
    by_code = {q['code']: q for q in questions}
    bao = []
    P = lambda *a: bao.append(' '.join(str(x) for x in a))

    # ---------- đối chiếu nội dung đề với bộ ôn theo buổi
    nd, tot = 0, 0
    i = 0
    while i < len(ub):
        u = ub[i]
        i += 1
        if 'code' not in u or u['code'] not in by_code:
            continue
        q = by_code[u['code']]
        if q.get('cluster_ctx'):
            continue
        tot += 1
        try:
            qb = build_question(u)
        except Exception as e:
            loi('Bộ ôn theo buổi: không đọc được %s: %s' % (u['code'], e))
            continue
        if qb['text'] != q['text'] or qb.get('options') != q.get('options'):
            nd += 1
            loi('Khác nội dung đề giữa 2 file: %s' % u['code'])
    P('Đối chiếu đề giữa 2 file Word: %d/%d câu đơn của bộ ôn theo buổi trùng khớp với ngân hàng tổng' % (tot - nd, tot))

    # ---------- đối chiếu đáp án
    banswers = parse_buoi_answers(bb)
    na = nchk = nmiss = 0
    for code, lines in banswers.items():
        q = by_code.get(code)
        if not q:
            continue
        t = notag(' '.join(lines))
        got = None
        if q['type'] == 'mc':
            m = re.search(r'Đáp án\s*:\s*([A-D])\b', t)
            got = m.group(1) if m else None
        elif q['type'] == 'tf':
            mk = re.findall(r'([a-d])\)\s*(Đ|S)\b', t)
            if len(mk) == 4:
                got = [x == 'Đ' for _, x in mk]
        else:
            got = short_answer({'ans_text': '', 'guide': lines})
        if got is None:
            nmiss += 1
            continue
        nchk += 1
        if got != q['answer']:
            na += 1
            loi('Đáp án khác nhau giữa 2 file: %s  tổng=%r  buổi=%r' % (code, q['answer'], got))
    P('Đối chiếu đáp án: %d câu có đáp án rõ ở cả hai file, %d câu khác nhau (%d câu bộ ôn không ghi đáp án dạng đọc được)' % (nchk, na, nmiss))

    # ---------- chọn câu đưa vào phòng
    loai_bo = [q for q in questions if not q.get('answer')]
    ok = [q for q in questions if q.get('answer')]
    okset = {q['code'] for q in ok}
    gnames = {}
    for b in bb:
        if b['k'] == 'p' and b['style'] == 'Heading 1':
            m = re.match(r'(Buổi\s+(\d+))\s*[–-]\s*(.*)', b['plain'])
            if m:
                nm = re.sub(r'^Bài\s+(\d+)\.\s*', r'Bài \1 · ', m.group(3)).strip()
                gnames[int(m.group(2))] = '%s · %s' % (m.group(1), nm)
    bank_names = {}
    for q in questions:
        m = re.match(r'Bài\s+(\d+)\.\s*(.*)', q['heads']['bai'])
        if m:
            bank_names[q['bai']] = 'Ngân hàng · Bài %s · %s' % (m.group(1), m.group(2))
    order = []
    for bi, code in order_b:
        if code in okset:
            order.append((bi, by_code[code]))
    for q in ok:
        if q['code'] not in buoi:
            order.append((10 + q['bai'], q))
    order.sort(key=lambda x: x[0])       # sort ổn định: giữ thứ tự trong nhóm
    LVN = {'nb': 'nb', 'th': 'th', 'vd': 'vd', 'vdc': 'vdc'}
    cfg_q, groups, key_a, key_e, key_n = [], {}, {}, {}, {}
    for g, q in order:
        groups[g] = gnames.get(g) or bank_names.get(g - 10) or ('Nhóm %d' % g)
        text = q['text']
        ctx = q.get('cluster_ctx')
        if ctx:
            text = ctx + '\n' + text
        imgs = []
        for f in (q.get('cluster_imgs') or []) + q['imgs']:
            if f and f not in imgs:
                imgs.append(f)
        o = {'id': q['code'], 'group': g, 'set': 'g', 'topic': q['bai'], 'level': LVN[q['lv'].lower()], 'type': q['type'], 'text': text}
        if q['type'] in ('mc', 'tf'):
            o['options'] = q['options']
        fig = {}
        if imgs:
            fig['img'] = imgs
            fig['alt'] = 'Hình của câu ' + q['code']
        if q.get('tables'):
            fig['html'] = ''.join(q['tables'])
        if fig:
            o['fig'] = fig
        if q['code'] in MO_PHONG:   # mô phỏng gắn theo câu (tools/mo_phong.json)
            o['sim'] = MO_PHONG[q['code']]
        cfg_q.append(o)
        key_a[q['code']] = q['answer'] if q['type'] != 'short' else {'v': q['answer'], 'tol': tol_of(q['answer'])}
        guide = [x for x in q['guide'] if not re.match(r'^(?:<[^>]+>)*Hướng dẫn giải \(theo nguồn\)\s*:?(?:</[^>]+>)*$', x.strip())]
        key_e[q['code']] = '\n'.join(guide) if guide else ''
        # Không hiện "Ghi chú cho giáo viên" (nguồn, ghi chú đáp án, mức độ tự xếp) trong phòng – theo yêu cầu GV 14/11/2026.
    sessions, total = [], 0
    for g in sorted(groups):
        n = sum(1 for x in cfg_q if x['group'] == g)
        sessions.append({'name': '%s (câu %d–%d)' % (groups[g], total + 1, total + n), 'end': total + n})
        total += n
    tabs = [{'id': 'mc', 'ten': 'Trắc nghiệm', 'types': ['mc']}, {'id': 'tf', 'ten': 'Đúng sai', 'types': ['tf']},
            {'id': 'short', 'ten': 'Trả lời ngắn', 'types': ['short']}]
    cfg = {'title': 'Luyện tập Vật lí 12 · Chương 2 · Chất khí', 'released': 0, 'closed': 0, 'tabs': tabs,
           'groups': {str(g): n for g, n in sorted(groups.items())}, 'sessions': sessions, 'questions': cfg_q}
    key = {'answers': key_a, 'explain': key_e, 'note': key_n}
    P('Đưa vào phòng: %d câu (%s) trong %d nhóm' % (len(cfg_q), ', '.join('%s %d' % (t, sum(1 for x in cfg_q if x['type'] == t)) for t in ('mc', 'tf', 'short')), len(groups)))
    for g in sorted(groups):
        P('  nhóm %2d | %3d câu | %s' % (g, sum(1 for x in cfg_q if x['group'] == g), groups[g]))
    P('Câu KHÔNG đưa vào phòng vì nguồn không có đáp án (%d):' % len(loai_bo))
    for q in loai_bo:
        P('  - %s | %s' % (q['code'], (notag(q['ans_text']).strip() or q['src'])[:90]))
    used = sorted({f for o in cfg_q for f in (o.get('fig', {}).get('img') or [])})
    P('Ảnh dùng trong đề: %d tệp' % len(used))
    P('Câu có bảng số liệu: %d | câu chùm: %d' % (sum(1 for o in cfg_q if o.get('fig', {}).get('html')), sum(1 for q in questions if q.get('cluster_ctx'))))
    P('Cảnh báo cần xem (%d):' % len(LOI))
    for x in LOI:
        P('  - ' + x)
    if args.tray:
        print('\n'.join(bao))
        print('(chạy thử, chưa ghi dữ liệu)')
        return
    with open(os.path.join(ROOT, 'data', 'cfg.json'), 'w', encoding='utf8') as f:
        json.dump(cfg, f, ensure_ascii=False)
    with open(os.path.join(ROOT, 'data', 'key.json'), 'w', encoding='utf8') as f:
        json.dump(key, f, ensure_ascii=False)
    P('Đã ghi data/cfg.json và data/key.json')
    with open(os.path.join(ROOT, 'tools', 'bao_cao_chuyen_doi.txt'), 'w', encoding='utf8') as f:
        f.write('\n'.join(bao) + '\n')
    print('\n'.join(bao))


if __name__ == '__main__':
    main_xuat()
