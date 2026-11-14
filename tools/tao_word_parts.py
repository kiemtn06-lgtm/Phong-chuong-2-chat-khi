# -*- coding: utf-8 -*-
"""Tách NGUYÊN VĂN đoạn Word (XML gốc: chữ, công thức Equation, hình, bảng) của từng câu trong phòng
từ file 523 câu, để máy chủ ghép thành file .docx khi giáo viên bấm "Xuất Word".

Ra: data/word_parts.json  = {src, head, tail, items: {mã câu: {stem, ans, intro2?, intro1?, pair?}}}
  - stem/ans/intro là chuỗi XML các đoạn Word; số câu thay bằng chỗ giữ ⟦N⟧ (câu), ⟦A⟧ ⟦B⟧ (2 câu chùm).
  - Bỏ dòng "Nguồn:" khỏi đáp án; bỏ ngắt trang.
Chạy (sau tao_du_lieu_523.py):  python -I -X utf8 tools/tao_word_parts.py
"""
import sys, os, re, json, copy, zipfile
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from lxml import etree
from tach_cau import tach, text, W

SRC_REL = 'nguon-goc/CHUONG2_KHI_LI_TUONG_tu_de_thi_tong_hop_DA_GOP_523_cau.docx'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
QS = re.compile(r'^\s*Câu\s+(\d+)\.\s*\[([A-Z0-9]+-[A-Z]+-[IVX]+-\d+)')


def sub_text(p, old, new):
    ts = list(p.iter(W + 't'))
    full = ''.join(t.text or '' for t in ts)
    i = full.find(old)
    if i < 0:
        return False
    j, pos, done = i + len(old), 0, False
    for t in ts:
        s = t.text or ''
        a, b = pos, pos + len(s)
        pos = b
        if b <= i or a >= j:
            continue
        lo, hi = max(i, a) - a, min(j, b) - a
        t.text = s[:lo] + (new if not done else '') + s[hi:]
        done = True
        t.set(XML_SPACE, 'preserve')
    return True


def clean(e):
    e = copy.deepcopy(e)
    for br in list(e.iter(W + 'br')):
        if br.get(W + 'type') == 'page':
            br.getparent().remove(br)
    for x in list(e.iter(W + 'lastRenderedPageBreak')):
        x.getparent().remove(x)
    return e


def ser(els):
    return ''.join(etree.tostring(e, encoding='unicode') for e in els)


def main():
    src = os.path.join(ROOT, SRC_REL)
    cfg = json.load(open(os.path.join(ROOT, 'data', 'cfg.json'), encoding='utf-8'))
    ids = {q['id'] for q in cfg['questions']}
    d, units = tach(src)
    items = {}
    for u in units:
        codes = [c.split(' ')[0] for c in u['codes']]
        stem = [clean(e) for e in u['stem']]
        ans = [clean(e) for e in u['ans'] if not text(e).strip().startswith('Nguồn:')]
        if not u['chum']:
            c, n = codes[0], u['cau'][0]
            if c not in ids:
                continue
            sub_text(stem[0], 'Câu %d.' % n, 'Câu ⟦N⟧.')
            items[c] = {'stem': ser(stem), 'ans': ser(ans)}
            continue
        # câu chùm: đoạn dữ kiện chung (+ hình/chú thích theo sau) rồi từng câu
        intro, per, cur = [stem[0]], [], None
        for e in stem[1:]:
            if e.tag == W + 'p' and QS.match(text(e)):
                cur = [e]
                per.append(cur)
            elif cur is None:
                intro.append(e)
            else:
                cur.append(e)
        a, b = u['cau']
        intro2, intro1 = [copy.deepcopy(x) for x in intro], [copy.deepcopy(x) for x in intro]
        sub_text(intro2[0], 'Câu %d và Câu %d' % (a, b), 'Câu ⟦A⟧ và Câu ⟦B⟧')
        sub_text(intro1[0], 'Câu %d và Câu %d' % (a, b), 'Câu ⟦A⟧')
        marked = any(e.tag == W + 'p' and re.match(r'\s*Đáp án câu\s+\d+', text(e)) for e in ans)
        ans_per, k, hd = [[] for _ in per], 0, -1
        for e in ans:
            t = text(e) if e.tag == W + 'p' else ''
            m = re.match(r'\s*Đáp án câu\s+(\d+)', t)
            if m and int(m.group(1)) in u['cau']:
                k = u['cau'].index(int(m.group(1)))
                sub_text(e, 'Đáp án câu %s' % m.group(1), 'Đáp án câu ⟦N⟧')
            elif not marked and re.match(r'\s*Hướng dẫn\s*:', t):
                hd = min(hd + 1, len(per) - 1)
                k = hd
            ans_per[k].append(e)
        for ci, (c, n) in enumerate(zip(codes, u['cau'])):
            if c not in ids:
                continue
            sub_text(per[ci][0], 'Câu %d.' % n, 'Câu ⟦N⟧.')
            items[c] = {'stem': ser(per[ci]), 'ans': ser(ans_per[ci]), 'intro2': ser(intro2), 'intro1': ser(intro1),
                        'pair': codes[1 - ci]}
    doc = zipfile.ZipFile(src).read('word/document.xml').decode('utf-8')
    head = doc[:doc.index('<w:body>') + len('<w:body>')]
    tail = doc[doc.rindex('<w:sectPr'):]
    miss = sorted(ids - set(items))
    out = {'src': SRC_REL, 'head': head, 'tail': tail, 'items': items}
    json.dump(out, open(os.path.join(ROOT, 'data', 'word_parts.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    print('Đoạn Word gốc: %d/%d câu | thiếu: %s | %.1f MB' % (len(items), len(ids), miss or 'không',
          os.path.getsize(os.path.join(ROOT, 'data', 'word_parts.json')) / 1048576))
    bad = [c for c, it in items.items() if 'Câu ⟦N⟧.' not in it['stem']]
    if bad:
        print('Không đặt được chỗ đánh số:', bad)


if __name__ == '__main__':
    main()
