# -*- coding: utf-8 -*-
"""Nạp bộ 523 câu Chương 2 (file "tu_de_thi_tong_hop_DA_GOP_523_cau.docx") vào phòng, chia 26 buổi
đúng như file Word "CHUONG2_KHI_LI_TUONG_Bo_on_theo_buoi_DU_523_cau.docx".

  nguon-goc/CHUONG2_KHI_LI_TUONG_tu_de_thi_tong_hop_DA_GOP_523_cau.docx -> data/cfg.json, data/key.json, public/img/*

Chạy:  python -I -X utf8 tools/tao_du_lieu_523.py [đường dẫn file .docx]      rồi  node tools/kiem_tra.js
Dùng lại bộ chuyển Word của tao_du_lieu_docx.py (chữ đậm/nghiêng, chỉ số, công thức OMML -> LaTeX, ảnh, bảng, sai số đáp án).
tach_cau.py: tách đơn vị câu; chia_buoi.py: chia buổi (mỗi buổi tối đa 18 TN, 4 Đ/S, 6 TLN; câu chùm đi cùng nhau).
Nguyên tắc: giữ NGUYÊN VĂN câu hỏi, phương án, đáp án, lời giải. Câu nguồn ghi "không có đáp án" không đưa vào phòng.
Mô phỏng (tools/mo_phong.json, ghi theo mã cũ) được gắn lại qua dòng "Mã cũ:" trong phần Nguồn của từng câu.
"""
import sys, os, re, json, glob
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
import tao_du_lieu_docx as T
from tach_cau import tach
from chia_buoi import chia, TEN_BAI

ROOT = os.path.dirname(HERE)
notag = lambda s: re.sub(r'<[^>]+>', '', s)
V_IMG = '{urn:schemas-microsoft-com:vml}imagedata'
R_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
# Hình vẽ bằng nét (VML) không có ảnh nhúng: ảnh chụp từ Word đặt sẵn ở tools/hinh_ve/<mã câu>.png
HINH_VE = {f[:-4] for f in os.listdir(os.path.join(HERE, 'hinh_ve'))} if os.path.isdir(os.path.join(HERE, 'hinh_ve')) else set()


def hinh_ve(doc, code):
    import shutil
    name = 'ch2_ve_%s.png' % code
    shutil.copyfile(os.path.join(HERE, 'hinh_ve', code + '.png'), os.path.join(doc.img_dir, name))
    return name


CHUM = re.compile(r'^\s*Sử dụng các thông tin sau cho Câu\s+\d+\s+và\s+Câu\s+\d+\s*:\s*')
MA_CU = re.compile(r'Mã cũ\s*:\s*((?:B\d+-[A-Z]+-[IV]+-\d+(?:\s*,\s*)?)+)')


def mk_block(doc, ch):
    """Một phần tử Word (đoạn hoặc bảng) -> khối cùng dạng với T.blocks()."""
    if ch.tag == qn('w:p'):
        p = Paragraph(ch, doc.d)
        sg = doc.segs(ch)
        imgs = [s[1] for s in sg if s[0] == 'img' and s[1]]
        for im in ch.iter(V_IMG):          # hình kiểu cũ của Word (VML: w:pict, w:object) có ảnh xem trước
            rid = im.get(R_ID)
            if rid and rid in doc.d.part.related_parts:
                name = doc.image_file(rid)
                if name and name not in imgs:
                    imgs.append(name)
        code = T.QSTART.match(T.segs_plain(sg).strip())
        if code and code.group(2) in HINH_VE:   # hình vẽ bằng nét của Word (không có ảnh): dùng ảnh đã chụp sẵn
            imgs.append(hinh_ve(doc, code.group(2)))
        return {'k': 'p', 'style': p.style.name if p.style is not None else '', 'segs': sg,
                'imgs': imgs, 'plain': T.segs_plain(sg).strip()}
    if ch.tag == qn('w:tbl'):
        t = Table(ch, doc.d)
        rows, imgs = [], []
        for r in t.rows:
            cells, seen = [], set()
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
        for im in ch.iter(V_IMG):          # ảnh xem trước của đối tượng nhúng nằm trong ô bảng
            rid = im.get(R_ID)
            if rid and rid in doc.d.part.related_parts:
                name = doc.image_file(rid)
                if name and name not in imgs:
                    imgs.append(name)
        return {'k': 't', 'style': 'tbl', 'rows': rows, 'imgs': imgs, 'plain': ''}
    return None


_split_goc = T.split_options


def split_options(text):
    """Như T.split_options; thêm trường hợp cả dòng phương án được in đậm: <b>A. ...\tB. ...</b>."""
    r = _split_goc(text)
    if r:
        return r
    return _split_goc(re.sub(r'(^|\n)<b>(\s*A\.[^\n]*)</b>(?=\n|$)', r'\1\2', text))


T.split_options = split_options
_omml_goc = T.omml


def omml(el):
    """Như T.omml; bỏ dấu cách cứng (\\ ) còn sót ở cuối công thức (gây lỗi KaTeX)."""
    s = re.sub(r'(?:\\\s*)+$', '', _omml_goc(el)).strip()
    # 2 công thức (B10-NB-I-01, B11-NB-I-01) còn chữ "hằng số" ở bảng mã cũ TCVN3 (.VnTime): "h»ng sè" -> giải mã, không đổi nội dung
    return s.replace('h»ng\\ s\\text{è}', '\\text{hằng số}')


T.omml = omml


def old_codes(src):
    m = MA_CU.search(notag(src or ''))
    return [c.strip() for c in m.group(1).split(',') if c.strip()] if m else []


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'nguon-goc', 'CHUONG2_KHI_LI_TUONG_tu_de_thi_tong_hop_DA_GOP_523_cau.docx')
    img_dir = os.path.join(ROOT, 'public', 'img')
    for f in glob.glob(os.path.join(img_dir, 'ch2_*')):
        os.remove(f)
    doc = T.Doc(src, img_dir, 'ch2')
    d, units = tach(src)
    doc.d = d
    bs = chia(units)
    bao = []
    P = bao.append
    cfg_q, groups, key_a, key_e, loai_bo, sim_used = [], {}, {}, {}, [], []
    sessions, total = [], 0

    def emit(g, q, ctx=None, ctx_imgs=None):
        """q: dict câu (kiểu T.build_question) đã có 'answer'. Trả về False nếu không đưa vào phòng."""
        if q['answer'] is None:
            loai_bo.append(q)
            return False
        text = q['text']
        if ctx:
            text = '<i>Dữ kiện dùng chung (câu chùm):</i> ' + ctx + '\n' + text
        imgs = []
        for f in (ctx_imgs or []) + q['imgs']:
            if f and f not in imgs:
                imgs.append(f)
        o = {'id': q['code'], 'group': g, 'set': 'g', 'topic': q['bai'], 'level': q['lv'].lower(), 'type': q['type'], 'text': text}
        if q['type'] in ('mc', 'tf'):
            o['options'] = q['options']
            if len(o['options']) != 4:
                T.loi('%s: số phương án/ý khác 4 (%d)' % (q['code'], len(o['options'])))
        fig = {}
        if imgs:
            fig['img'] = imgs
            fig['alt'] = 'Hình của câu ' + q['code']
        if q.get('tables'):
            fig['html'] = ''.join(q['tables'])
        if fig:
            o['fig'] = fig
        if q.get('old') and q['old'] in T.MO_PHONG:
            o['sim'] = T.MO_PHONG[q['old']]
            sim_used.append(q['old'])
        cfg_q.append(o)
        key_a[q['code']] = q['answer'] if q['type'] != 'short' else {'v': q['answer'], 'tol': T.tol_of(q['answer'])}
        guide = [x for x in q['guide'] if not re.match(r'^(?:<[^>]+>)*Hướng dẫn giải \(theo nguồn\)\s*:?(?:</[^>]+>)*$', x.strip())]
        key_e[q['code']] = '\n'.join(guide) if guide else ''
        return True

    for x in bs:
        g = x['so']
        groups[g] = 'Buổi %d · Bài %d · %s' % (g, x['bai'], TEN_BAI[x['bai']])
        n0 = len(cfg_q)
        for part in ('I', 'II', 'III'):
            for u in x[part]:
                sb = [b for b in (mk_block(doc, e) for e in u['stem']) if b]
                ab = [b for b in (mk_block(doc, e) for e in u['ans']) if b]
                if not u['chum']:
                    m = T.QSTART.match(sb[0]['plain'])
                    if not m:
                        raise ValueError('không nhận ra đầu câu: ' + sb[0]['plain'][:60])
                    q = T.build_question({'code': m.group(2), 'n': int(m.group(1)), 'star': bool(m.group(3)), 'blocks': sb + ab,
                                          'prefix': len(m.group(0)), 'heads': {}, 'idx': u['idx'], 'cluster': None})
                    if q['type'] == 'short':
                        q['answer'] = T.short_answer(q)
                    oc = old_codes(q['src'])
                    q['old'] = oc[0] if len(oc) == 1 else None
                    emit(g, q)
                    continue
                # ---- cụm câu chùm (Phần III): 1 đoạn dữ kiện chung + 2 câu + lời giải từng câu + nguồn
                mctx = CHUM.match(sb[0]['plain'])
                ctx = T.render(T.strip_prefix(sb[0]['segs'], len(mctx.group(0)))).strip()
                ctx_imgs = list(sb[0]['imgs'])
                ctx_tables = []
                per = []
                for b in sb[1:]:
                    if b['k'] == 'p' and T.QSTART.match(b['plain']):
                        per.append([b])
                    elif per:
                        per[-1].append(b)
                    else:   # hình, chú thích hình, bảng hoặc đoạn nối tiếp của dữ kiện chung
                        ctx_imgs += b['imgs']
                        if b['k'] == 'p':
                            t = T.block_text_rich(b)
                            if t:
                                ctx += '\n' + ('<i>%s</i>' % t if b['style'] == 'Image Caption' else t)
                        else:
                            ctx_tables.append(T.table_html(b['rows']))
                if len(per) != len(u['cau']):
                    raise ValueError('câu chùm không đủ câu: %s' % u['codes'])
                marked = any(b['k'] == 'p' and re.match(r'Đáp án câu\s+\d+', b['plain']) for b in ab)
                ans_per = [[] for _ in per]
                src_b, k, hd = None, 0, -1
                for b in ab:
                    pl = b['plain'] if b['k'] == 'p' else ''
                    if pl.startswith('Nguồn:'):
                        src_b = b
                        continue
                    mm = re.match(r'Đáp án câu\s+(\d+)', pl)
                    if mm and int(mm.group(1)) in u['cau']:
                        k = u['cau'].index(int(mm.group(1)))
                    elif not marked and re.match(r'Hướng dẫn\s*:', pl):
                        hd = min(hd + 1, len(per) - 1)
                        k = hd
                    ans_per[k].append(b)
                src_txt = T.block_text_rich(src_b) if src_b else ''
                oc = old_codes(src_txt)
                for ci, blks in enumerate(per):
                    m = T.QSTART.match(blks[0]['plain'])
                    code = m.group(2)
                    mm = T.CODE.match(code)
                    if T.TY[mm.group(3)] != 'short':
                        raise ValueError('câu chùm không phải trả lời ngắn: ' + code)
                    paras, tables, imgs = [], [], []
                    for j, b in enumerate(blks):
                        imgs += b['imgs']
                        if b['k'] == 'p':
                            t = T.block_text_rich(b, len(m.group(0)) if j == 0 else 0)
                            if t:
                                paras.append('<i>%s</i>' % t if b['style'] == 'Image Caption' else t)
                        else:
                            tables.append(T.table_html(b['rows']))
                    guide = []
                    for b in ans_per[ci]:
                        if b['k'] == 'p':
                            t = T.block_text_rich(b, 0, True)
                            if t:
                                guide.append(re.sub(r'Đáp án câu\s+\d+', 'Đáp án', t))
                        else:
                            for r in b['rows']:
                                cells = [T.cell_rich(c, True) for c in r]
                                if any(cells):
                                    guide.append(' · '.join(c for c in cells if c))
                    q = {'code': code, 'bai': int(mm.group(1)), 'lv': mm.group(2), 'type': 'short', 'text': '\n'.join(paras), 'imgs': imgs,
                         'guide': guide, 'ans_text': '', 'src': src_txt, 'old': oc[ci] if len(oc) == len(per) else None}
                    if ctx_tables or tables:
                        q['tables'] = ctx_tables + tables
                    q['answer'] = T.short_answer(q)
                    emit(g, q, ctx, ctx_imgs)
        n = len(cfg_q) - n0
        sessions.append({'name': '%s (câu %d–%d)' % (groups[g], total + 1, total + n), 'end': total + n})
        total += n

    tabs = [{'id': 'mc', 'ten': 'Trắc nghiệm', 'types': ['mc']}, {'id': 'tf', 'ten': 'Đúng sai', 'types': ['tf']},
            {'id': 'short', 'ten': 'Trả lời ngắn', 'types': ['short']}]
    cfg = {'title': 'Luyện tập Vật lí 12 · Chương 2 · Chất khí', 'released': 0, 'closed': 0, 'tabs': tabs,
           'groups': {str(g): n for g, n in sorted(groups.items())}, 'sessions': sessions, 'questions': cfg_q}
    key = {'answers': key_a, 'explain': key_e, 'note': {}}
    P('Nguồn: ' + os.path.basename(src))
    P('Đưa vào phòng: %d câu (%s) trong %d nhóm (buổi)' % (len(cfg_q), ', '.join('%s %d' % (t, sum(1 for q in cfg_q if q['type'] == t)) for t in ('mc', 'tf', 'short')), len(groups)))
    for g in sorted(groups):
        c = [q for q in cfg_q if q['group'] == g]
        P('  nhóm %2d | %2d câu (TN %2d, ĐS %d, TLN %d) | %s' % (g, len(c), sum(q['type'] == 'mc' for q in c), sum(q['type'] == 'tf' for q in c), sum(q['type'] == 'short' for q in c), groups[g]))
    P('Câu KHÔNG đưa vào phòng vì nguồn không có đáp án (%d):' % len(loai_bo))
    for q in loai_bo:
        P('  - %s | %s' % (q['code'], (notag(q.get('ans_text', '')).strip() or notag(q.get('src', '')))[:90]))
    used = sorted({f for o in cfg_q for f in (o.get('fig', {}).get('img') or [])})
    P('Ảnh dùng trong đề: %d tệp | câu có bảng số liệu: %d | câu chùm: %d' % (len(used), sum(1 for o in cfg_q if o.get('fig', {}).get('html')), sum(len(u['cau']) for x in bs for u in x['III'] if u['chum'])))
    P('Mô phỏng gắn lại theo mã cũ: %d/%d' % (len(set(sim_used)), len([k for k in T.MO_PHONG if not k.startswith('_')])))
    miss = sorted(k for k in T.MO_PHONG if not k.startswith('_') and k not in sim_used)
    if miss:
        P('  chưa gắn được (không thấy "Mã cũ" tương ứng): ' + ', '.join(miss))
    P('Cảnh báo cần xem (%d):' % len(T.LOI))
    for w in T.LOI:
        P('  - ' + w)
    with open(os.path.join(ROOT, 'data', 'cfg.json'), 'w', encoding='utf8') as f:
        json.dump(cfg, f, ensure_ascii=False)
    with open(os.path.join(ROOT, 'data', 'key.json'), 'w', encoding='utf8') as f:
        json.dump(key, f, ensure_ascii=False)
    P('Đã ghi data/cfg.json và data/key.json')
    with open(os.path.join(ROOT, 'tools', 'bao_cao_chuyen_doi_523.txt'), 'w', encoding='utf8') as f:
        f.write('\n'.join(bao) + '\n')
    print('\n'.join(bao))


if __name__ == '__main__':
    main()
