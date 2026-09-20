# -*- coding: utf-8 -*-
"""
模板填空位清单 —— 技能自带脚本，装配前先跑，禁止凭印象猜栏目。

用法:
    python list_slots.py <模板.docx>

输出每个位置的坐标、类型与处理方式，装配时照单执行：
    [填]   有淡红说明文字 → 把说明整段替换为本课内容（同格黑字原样保留）
    [空]   单元格为空     → 直接填入
    [锁]   说明含"不改动" → 跳过，一个字都不改
    [选]   黑字为打√选项  → 只在选项上做标记，不改写选项文字
    [正文] 表格外段落     → 双区装配的正文区，在原段落内替换文字，不得增删段落

仅用标准库，不依赖 python-docx。
"""
import sys, io, zipfile, re
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
BLACKS = (None, 'auto', '000000')


def run_color(r):
    rpr = r.find(W + 'rPr')
    if rpr is None:
        return None
    c = rpr.find(W + 'color')
    return c.get(W + 'val') if c is not None else None


def runs(el):
    out = []
    for r in el.iter(W + 'r'):
        t = ''.join(x.text or '' for x in r.findall(W + 't'))
        if t.strip():
            out.append((run_color(r), t))
    return out


def clip(s, n=52):
    s = re.sub(r'\s+', ' ', s or '').strip()
    return s[:n] + ('…' if len(s) > n else '')


def main(path):
    body = ET.fromstring(zipfile.ZipFile(path).read('word/document.xml')).find(W + 'body')
    fill = lock = empty = choice = 0

    for ti, tbl in enumerate(body.findall(W + 'tbl')):
        print('\n表 %d' % (ti + 1))
        print('-' * 70)
        for ri, tr in enumerate(tbl.findall(W + 'tr')):
            for ci, tc in enumerate(tr.findall(W + 'tc')):
                rs = runs(tc)
                blk = ''.join(t for c, t in rs if c in BLACKS)
                red = [t for c, t in rs if c not in BLACKS]
                pos = 't%d r%-2d c%d' % (ti, ri, ci)
                if not rs:
                    empty += 1
                    print('  %s  [空]   ' % pos)
                elif any('不改动' in t for t in red):
                    lock += 1
                    print('  %s  [锁]   %s' % (pos, clip(blk)))
                elif red:
                    fill += 1
                    tag = '[填]'
                    if '√' in blk or '打√' in blk:
                        tag, choice, fill = '[选]', choice + 1, fill - 1
                    print('  %s  %s   栏目「%s」' % (pos, tag, clip(blk, 24)))
                    print('  %s         说明→替换：%s' % (' ' * len(pos), clip(red[0])))
                elif '打√' in blk or '√' in blk:
                    choice += 1
                    print('  %s  [选]   %s' % (pos, clip(blk)))
                else:
                    print('  %s  [名]   %s' % (pos, clip(blk, 30)))

    paras = [(i, ''.join(x.text or '' for x in ch.iter(W + 't')))
             for i, ch in enumerate(body) if ch.tag == W + 'p']
    paras = [(i, t) for i, t in paras if t.strip()]
    if paras:
        print('\n表格外正文（双区装配的正文区，共 %d 段）' % len(paras))
        print('-' * 70)
        for i, t in paras:
            print('  p%-4d [正文] %s' % (i, clip(t, 58)))

    print('\n' + '=' * 70)
    print('合计：待填 %d 格 ｜ 空格 %d ｜ 打√选项 %d 格 ｜ 锁定(不改动) %d 格 ｜ 正文 %d 段'
          % (fill, empty, choice, lock, len(paras)))
    print('装配后必须运行 verify_assembly.py 校验，未通过不得交付。')


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    main(sys.argv[1])
