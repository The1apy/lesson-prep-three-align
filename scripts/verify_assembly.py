# -*- coding: utf-8 -*-
"""
装配结构保真校验 —— 技能自带脚本，禁止每轮现写。

用法:
    python verify_assembly.py <模板原件.docx> <成品.docx>

仅用标准库（zipfile + xml.etree），不依赖 python-docx，任何环境可直接运行。
任一检查项 FAIL 即判装配失败：不得交付成品，须重做装配。
"""
import sys, io, zipfile, re
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
RED_HINT = ('C00000',)          # 淡红说明文字；其它非黑颜色一并计入"带色"
BLACKS = (None, 'auto', '000000')


def load(path):
    return ET.fromstring(zipfile.ZipFile(path).read('word/document.xml')).find(W + 'body')


def run_color(r):
    rpr = r.find(W + 'rPr')
    if rpr is None:
        return None
    c = rpr.find(W + 'color')
    return c.get(W + 'val') if c is not None else None


def runs(el):
    """返回 [(color, text), ...]，跳过空文本。"""
    out = []
    for r in el.iter(W + 'r'):
        t = ''.join(x.text or '' for x in r.findall(W + 't'))
        if t.strip():
            out.append((run_color(r), t))
    return out


def norm(s):
    return re.sub(r'\s+', '', s or '')


def cell_text(tc):
    return ''.join(t for _, t in runs(tc))


def black_text(el):
    return ''.join(t for c, t in runs(el) if c in BLACKS)


def red_texts(el):
    return [t for c, t in runs(el) if c not in BLACKS]


def tables(body):
    return body.findall(W + 'tbl')


def cell_sig(tc):
    """该单元格的合并特征 (gridSpan, vMerge)。"""
    pr = tc.find(W + 'tcPr')
    span, vm = 1, ''
    if pr is not None:
        gs = pr.find(W + 'gridSpan')
        if gs is not None:
            span = int(gs.get(W + 'val', '1'))
        v = pr.find(W + 'vMerge')
        if v is not None:
            vm = v.get(W + 'val') or 'continue'
    return (span, vm)


def table_shape(tbl):
    grid = tbl.find(W + 'tblGrid')
    ncol = len(grid.findall(W + 'gridCol')) if grid is not None else 0
    rows = tbl.findall(W + 'tr')
    return ncol, [[cell_sig(tc) for tc in tr.findall(W + 'tc')] for tr in rows]


def body_paragraphs(body):
    """表格外的非空段落文本（按顺序）。"""
    out = []
    for ch in body:
        if ch.tag == W + 'p':
            t = ''.join(x.text or '' for x in ch.iter(W + 't'))
            if t.strip():
                out.append(t)
    return out


def all_cells(body):
    for tbl in tables(body):
        for tr in tbl.findall(W + 'tr'):
            for tc in tr.findall(W + 'tc'):
                yield tc


def main(tpl_path, out_path):
    tpl, out = load(tpl_path), load(out_path)
    results = []       # (名称, 通过?, 说明)

    def rec(name, ok, msg=''):
        results.append((name, ok, msg))

    # ---- 1. 来自模板副本（先决项）----
    colored = sum(1 for tc in all_cells(out) for c, _ in runs(tc) if c not in BLACKS)
    colored += sum(1 for ch in out if ch.tag == W + 'p'
                   for c, _ in runs(ch) if c not in BLACKS)
    ok1 = colored > 0
    rec('来自模板副本（带色 run 数 > 0）', ok1, '带色 run %d 处' % colored)
    if not ok1:
        rec('后续检查', False, '先决项未通过：成品是新建文档而非模板副本，其余检查略过')
        return report(results)

    # ---- 2. 表格数量 ----
    tt, ot = tables(tpl), tables(out)
    rec('表格数量', len(tt) == len(ot), '模板 %d / 成品 %d' % (len(tt), len(ot)))

    # ---- 3+4. 行列数与合并结构 ----
    for i, (a, b) in enumerate(zip(tt, ot)):
        na, sa = table_shape(a)
        nb, sb = table_shape(b)
        rec('表%d 行列数' % (i + 1), (len(sa) == len(sb) and na == nb),
            '模板 %d行×%d列网格 / 成品 %d行×%d列网格' % (len(sa), na, len(sb), nb))
        rec('表%d 合并结构' % (i + 1), sa == sb,
            '一致' if sa == sb else 'gridSpan/vMerge 分布不同')

    # ---- 5. 栏目名保留（原件黑字须原样存在于成品同格）----
    bad = []
    for idx, (a, b) in enumerate(zip(all_cells(tpl), all_cells(out))):
        need = norm(black_text(a))
        if need and need not in norm(cell_text(b)):
            bad.append('格#%d 缺「%s」' % (idx, black_text(a)[:20]))
    rec('栏目名逐格保留', not bad, '；'.join(bad[:4]) if bad else '全部保留')

    # ---- 6.「不改动」栏原样（逐字相等，不用 in）----
    locked, bad2 = 0, []
    for idx, (a, b) in enumerate(zip(all_cells(tpl), all_cells(out))):
        if any('不改动' in t for t in red_texts(a)):
            locked += 1
            if norm(cell_text(a)) != norm(cell_text(b)):
                bad2.append('格#%d 被改动' % idx)
    rec('「不改动」栏原样', not bad2,
        ('共 %d 格，均未改动' % locked) if not bad2 else '；'.join(bad2))

    # ---- 7. 说明文字已替换（「不改动」栏除外）----
    left = []
    out_all = norm(''.join(cell_text(tc) for tc in all_cells(out)) + ''.join(body_paragraphs(out)))
    for a in all_cells(tpl):
        if any('不改动' in t for t in red_texts(a)):
            continue
        for t in red_texts(a):
            if len(norm(t)) >= 6 and norm(t) in out_all:
                left.append(t[:22])
    rec('说明文字已替换', not left,
        '残留：' + '；'.join(dict.fromkeys(left[:4])) if left else '无残留')

    # ---- 8. 表格外正文骨架保留（双区装配的正文区）----
    tp, op = body_paragraphs(tpl), body_paragraphs(out)
    op_n = [norm(x) for x in op]
    miss, j = [], 0
    for t in tp:
        nt = norm(t)
        if len(nt) < 3:
            continue
        hit = next((k for k in range(j, len(op_n)) if nt in op_n[k]), None)
        if hit is None:
            miss.append(t[:18])
        else:
            j = hit + 1
    rec('正文骨架保留（表格外段落）', not miss,
        '模板 %d 段 / 成品 %d 段；缺：%s' % (len(tp), len(op), '；'.join(miss[:4]))
        if miss else '模板 %d 段全部按序保留' % len(tp))

    return report(results)


def report(results):
    width = max(len(n) for n, _, _ in results)
    print('=' * 72)
    print('装配结构保真校验')
    print('=' * 72)
    for name, ok, msg in results:
        print('  [%s] %-*s  %s' % ('PASS' if ok else 'FAIL', width, name, msg))
    failed = [n for n, ok, _ in results if not ok]
    print('-' * 72)
    if failed:
        print('结果：FAIL（%d 项未通过：%s）' % (len(failed), '、'.join(failed)))
        print('→ 不得交付成品，须重做装配。')
        return 1
    print('结果：PASS（全部 %d 项通过）→ 可交付。' % len(results))
    return 0


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1], sys.argv[2]))
