# -*- coding: utf-8 -*-
"""苹果财报风格骨架示例 —— 白底、超大留白、一页一个指标、guidance 收尾。

【脱敏说明】设计系统与函数写法完整保留；文案为占位词、数值为演示占位值。
页序（正式版 15 页，本骨架抽 9 页示范）：
  01 封面 → 02 一页看完 → 03 主角数字（电量）→ 04 分组柱图（分月）
  → 05 主角数字（限电率）→ 06 横条对比（结构）→ 07 折线（度电考核走势）
  → 08 苹果式数据表（成本）→ 09 guidance 收尾

三条铁律（不遵守就变成普通汇报）：
  1. 一页只讲一个数：巨数字 56—104pt，其余文字都降到 11—17pt 当配角
  2. 大数字后不要紧跟中文单位——66pt 数字 + 单位在 3.7in 栏内必换行；
     数字单独一行，单位与指标名合并成下一行小字
  3. 只用横线，不用竖线、不用单元格底色；数值列右对齐，非数值列加 0.24in 左内距
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deck_lib import new_deck, rgb, _fix_font
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'demo_apple.pptx')

SW, SH = 13.333, 7.5
LM = 1.02                      # 左边距（苹果式大留白）
CW = SW - 2 * LM

# 文字与中性（黑灰阶，承载版式）
INK   = '1D1D1F'
GRAY  = '6E6E73'
LGRAY = '86868B'
LINE  = 'D2D2D7'
FILL  = 'F5F5F7'
WHITE = 'FFFFFF'

# 企业标准色（取自公司色卡；换成自家 VI 即可）
C1 = '004A86'   # 主色 —— 强调、最大项
C2 = '006AAF'   # 主数据 —— 柱、条、折线
C3 = '00A2E9'   # 次数据
G1 = '03994C'   # 正向 / 达标
G2 = '83C034'   # 备用正向
G3 = 'DADA00'   # 中性提示
N1 = 'ED6C00'   # 负向 / 未达标（色卡无正红时用它替代）
N2 = 'F6AD02'   # 中间态
N3 = 'FED900'   # 高亮块

BLUE, DEEP, SKY, GREEN, RED, AMBER, LIME = C2, C1, C3, G1, N1, N2, G2
LBLUE = '9FD5F2'

CN = '等线'      # 比微软雅黑更接近苹方
EN = 'Arial'     # 最接近 SF Pro 的字形（数字 1 无底横）


def aslide(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl._anim_groups = []
    bg = sl.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = rgb(WHITE)
    return sl


def t2(sl, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
       ls=1.0, spc=0):
    """runs: [(text, size, color, bold, font)]；spc 单位为 1/100 pt（负值=收紧）"""
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = ls
    for text, size, color, bold, font in runs:
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = rgb(color)
        if spc:
            r._r.get_or_add_rPr().set('spc', str(int(spc)))
        _fix_font(r, font)
    return tb


def line(sl, x, y, w, color=LINE, pt=0.75):
    ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y),
                                 Inches(x + w), Inches(y))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(pt)
    return ln


def eyebrow(sl, text, color=LGRAY, y=0.90):
    """顶部小标：11pt 浅灰，字距放宽。y=0.90 起，给母版 logo 与横线让位。"""
    tb = t2(sl, LM, y, CW, 0.26, [(text, 11, color, False, CN)], spc=60)
    tb.name = 'EYEBROW'


def h1(sl, text, y=1.14, size=44, color=INK, w=None, ls=1.10):
    tb = t2(sl, LM, y, w or CW, size / 72.0 * 1.35 * max(1, text.count('\n') + 1),
            [(text, size, color, True, CN)], ls=ls, spc=-40)
    tb.name = 'STITLE'


def kicker(sl, text, y=6.80):
    """页脚：细线 + 数据口径。宽度收到 CW-0.90，避开母版右下角页码。"""
    ln = line(sl, LM, y - 0.14, CW - 0.90, LINE, 0.75)
    ln.name = 'KLINE'
    tb = t2(sl, LM, y, CW - 0.7, 0.24, [(text, 9, LGRAY, False, CN)], spc=10)
    tb.name = 'KICKER'


def big(sl, x, y, num, unit, size=104, color=INK, unit_size=19, w=9.0, sub_color=GRAY):
    """主角数字：数字单独一行，单位以 19pt 灰字跟在后面。"""
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(size / 72.0 * 1.12))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.line_spacing = 0.94
    r = p.add_run(); r.text = num
    r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = rgb(color)
    r._r.get_or_add_rPr().set('spc', '-60')
    _fix_font(r, EN)
    if unit:
        r2 = p.add_run(); r2.text = ' ' + unit
        r2.font.size = Pt(unit_size); r2.font.bold = False
        r2.font.color.rgb = rgb(sub_color)
        _fix_font(r2, CN)
    return tb


def sub(sl, x, y, text, size=17, color=GRAY, w=None, ls=1.35):
    t2(sl, x, y, w or CW, 0.46 * max(1, text.count('\n') + 1),
       [(text, size, color, False, CN)], ls=ls)


# ---------- 极简图表（无坐标轴、无网格） ----------
def _plain(sh):
    st = sh._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}style')
    if st is not None:
        sh._element.remove(st)
    sh.shadow.inherit = False
    return sh


def vbars(sl, x, y, w, h, items, maxv, bw=0.62, label_size=15, cat_size=12, fmt='%.2f'):
    """items: [(label, value, color)]——柱顶数值、柱底标签"""
    n, slot = len(items), w / max(1, len(items))
    for i, (lb, v, col) in enumerate(items):
        cx = x + slot * (i + 0.5)
        bh = max(v / maxv * h, 0.01)
        sh = _plain(sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(cx - bw / 2),
                                        Inches(y + h - bh), Inches(bw), Inches(bh)))
        sh.fill.solid(); sh.fill.fore_color.rgb = rgb(col)
        sh.line.fill.background()
        t2(sl, cx - slot / 2, y + h - bh - 0.34, slot, 0.30,
           [(fmt % v, label_size, INK, True, EN)], align=PP_ALIGN.CENTER)
        t2(sl, cx - slot / 2, y + h + 0.12, slot, 0.28,
           [(lb, cat_size, GRAY, False, CN)], align=PP_ALIGN.CENTER)


def group_bars(sl, x, y, w, h, cats, series, colors, maxv, bw=0.26, fmt='%.2f'):
    """series: [(name, [v...], )]——分组柱，用于目标 vs 实际"""
    n, m, slot = len(cats), len(series), w / max(1, len(cats))
    for i, c in enumerate(cats):
        cx = x + slot * (i + 0.5)
        for k, (nm, vs) in enumerate(series):
            v = vs[i]
            bh = max(v / maxv * h, 0.01)
            left = cx - (m * bw + (m - 1) * 0.06) / 2 + k * (bw + 0.06)
            sh = _plain(sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left),
                                            Inches(y + h - bh), Inches(bw), Inches(bh)))
            sh.fill.solid(); sh.fill.fore_color.rgb = rgb(colors[k])
            sh.line.fill.background()
            t2(sl, left - 0.16, y + h - bh - 0.30, bw + 0.32, 0.26,
               [(fmt % v, 12.5, INK, True, EN)], align=PP_ALIGN.CENTER)
        t2(sl, cx - slot / 2, y + h + 0.12, slot, 0.28, [(c, 12, GRAY, False, CN)],
           align=PP_ALIGN.CENTER)


def hbrow(sl, x, y, w, label, value, ratio, color=BLUE, lw=2.6, label_w=2.30,
          val_w=1.30):
    """细横条：只有一条轨道线 + 一条数值线"""
    t2(sl, x, y - 0.05, label_w - 0.16, 0.30, [(label, 13, GRAY, False, CN)])
    bw = w - label_w - val_w
    tr = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x + label_w),
                                 Inches(y + 0.14), Inches(x + label_w + bw), Inches(y + 0.14))
    tr.line.color.rgb = rgb('E8E8ED'); tr.line.width = Pt(lw)
    sh = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x + label_w),
                                 Inches(y + 0.14), Inches(x + label_w + bw * ratio),
                                 Inches(y + 0.14))
    sh.line.color.rgb = rgb(color); sh.line.width = Pt(lw)
    t2(sl, x + label_w + bw + 0.14, y - 0.05, val_w, 0.30,
       [(value, 14, INK, True, EN)])


def dline(sl, x, y, w, h, cats, series, colors, maxv, minv=0, fmt='%.1f'):
    """极简折线：无网格、只在首末点与主序列打点、末点标数值"""
    n, span = len(cats), maxv - minv
    px, pw = x, w - 0.55
    for k, (nm, vs) in enumerate(series):
        pts = [(px + pw * i / (n - 1), y + h * (1 - (v - minv) / span)) for i, v in enumerate(vs)]
        for a, b in zip(pts, pts[1:]):
            ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(a[0]), Inches(a[1]),
                                         Inches(b[0]), Inches(b[1]))
            ln.line.color.rgb = rgb(colors[k])
            ln.line.width = Pt(2.25)
        for i, (mx, my) in enumerate(pts):
            if i in (0, n - 1) or k == 0:
                d = 0.058
                sh = _plain(sl.shapes.add_shape(MSO_SHAPE.OVAL, Inches(mx - d / 2),
                                                Inches(my - d / 2), Inches(d), Inches(d)))
                sh.fill.solid(); sh.fill.fore_color.rgb = rgb(colors[k])
                sh.line.fill.background()
        last = pts[-1]
        t2(sl, last[0] + 0.10, last[1] - 0.17, 1.05, 0.26,
           [(fmt % vs[-1], 12.5, INK, True, EN)])
    for i, c in enumerate(cats):
        if i % 2 == 0 or i == n - 1:
            t2(sl, px + pw * i / (n - 1) - 0.28, y + h + 0.14, 0.56, 0.26,
               [(c, 11.5, LGRAY, False, CN)], align=PP_ALIGN.CENTER)


def rowlines(sl, x, y, w, cols, widths, rows, row_h=0.50, fs=13, head_fs=11,
             num_cols=(1, 2, 3), top=True):
    """苹果式数据表：只有横线，无竖线无填充"""
    total = sum(widths)
    xpos = [x]
    for ww in widths[:-1]:
        xpos.append(xpos[-1] + ww * w / total)
    if top:
        line(sl, x, y - 0.30, w, INK, 1.0)
    for j, c in enumerate(cols):
        t2(sl, xpos[j] + (0.24 if j > 0 and j not in num_cols else 0), y - 0.24,
           widths[j] * w / total - 0.08, 0.26, [(c, head_fs, LGRAY, False, CN)],
           align=PP_ALIGN.RIGHT if j in num_cols else PP_ALIGN.LEFT, spc=20)
    line(sl, x, y, w, LINE, 0.75)
    for i, row in enumerate(rows):
        ry = y + 0.10 + i * row_h
        for j, cell in enumerate(row):
            t2(sl, xpos[j] + (0.24 if j > 0 and j not in num_cols else 0), ry,
               widths[j] * w / total - 0.08, 0.30,
               [(str(cell), fs, INK if j == 0 else GRAY, j == 0,
                 EN if j in num_cols else CN)],
               align=PP_ALIGN.RIGHT if j in num_cols else PP_ALIGN.LEFT)
        if i == len(rows) - 1:
            line(sl, x, ry + 0.36, w, LINE, 0.75)


# ======================== 01 封面 ========================
def p01(prs):
    sl = aslide(prs)
    eyebrow(sl, '示例能源公司　·　20XX 年第三季度', y=2.20)
    t2(sl, LM, 2.52, CW, 1.30,
       [('生产运营', 60, INK, True, CN)], ls=1.06, spc=-40)
    t2(sl, LM, 3.42, CW, 0.80,
       [('一个季度，四个数。', 26, GRAY, False, CN)], ls=1.20)
    line(sl, LM, 6.10, CW - 0.90, LINE, 0.75)
    t2(sl, LM, 6.24, CW - 0.7, 0.26,
       [('生产管理部　·　20XX 年 10 月', 9, LGRAY, False, CN)], spc=10)


# ======================== 02 一页看完 ========================
def p02(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '本季度一览')
    h1(sl, '四个数，一句话。', y=1.14, size=40)
    sub(sl, LM, 2.00, '电量欠发，考核优于对标，成本在预算内——问题集中在限电。')

    # 四宫格：单位写进标签，巨数字不带单位——52pt 数字 + 单位一定会挤到邻列
    items = [('30.00', '上网电量（亿千瓦时）', '完成率 93.8%', DEEP),
             ('12.0%', '限电率', '优于对标 1.0pp', RED),
             ('0.0400', '度电考核（元/千瓦时）', '低于对标 5.0%', GREEN),
             ('10,000', '运维成本（万元）', '支出率 77%', BLUE)]
    for i, (num, lab, note, col) in enumerate(items):
        cx = LM + i * (CW / 4)
        t2(sl, cx, 2.86, CW / 4 - 0.30, 0.28, [(lab, 11.5, LGRAY, False, CN)], spc=20)
        big(sl, cx, 3.16, num, '', size=46, color=col, w=CW / 4 - 0.30)
        t2(sl, cx, 4.20, CW / 4 - 0.30, 0.28, [(note, 12, GRAY, False, CN)])
    kicker(sl, '口径：本季度累计，与上季度同口径；度电考核＝考核电量÷上网电量。')


# ======================== 03 主角数字：电量 ========================
def p03(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '电量')
    h1(sl, '少发 2.00 亿千瓦时。', y=1.14, size=40)
    big(sl, LM, 2.30, '30.00', '亿千瓦时', size=104, color=INK, unit_size=22)
    sub(sl, LM, 3.86, '计划 32.00 亿千瓦时，完成率 93.8%。\n缺口七成在存量项目，三成在新增装机。',
        size=17)
    line(sl, LM, 5.10, CW - 0.90, LINE, 0.75)
    for i, (k, v) in enumerate([('计划电量', '32.00'), ('缺口', '−2.00'),
                                ('去年同期', '32.60'), ('同比', '−8.0%')]):
        cx = LM + i * (CW / 4.6)
        t2(sl, cx, 5.28, 2.20, 0.26, [(k, 11, LGRAY, False, CN)], spc=30)
        t2(sl, cx, 5.56, 2.20, 0.40, [(v, 20, INK, True, EN)])
    kicker(sl, '口径：上网电量（非发电量），含试运行电量；数据截至 9 月 30 日。')


# ======================== 04 分组柱图：分月电量 ========================
def p04(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '电量　·　月度')
    h1(sl, '7 月是低点，8—9 月连续回升。', y=1.14, size=34)

    group_bars(sl, LM, 2.20, CW - 0.60, 3.30,
               ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'],
               [('计划', [3.4, 3.2, 4.0, 3.9, 3.8, 3.6, 3.5, 3.5, 3.1]),
                ('实际', [3.1, 3.0, 3.8, 3.7, 3.4, 3.2, 3.0, 3.3, 3.5])],
               [LBLUE, BLUE], 5.0, bw=0.30, fmt='%.1f')
    # 图例
    for i, (nm, col) in enumerate([('计划', LBLUE), ('实际', BLUE)]):
        lx = LM + i * 1.30
        sh = _plain(sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(lx), Inches(6.08),
                                        Inches(0.16), Inches(0.16)))
        sh.fill.solid(); sh.fill.fore_color.rgb = rgb(col)
        sh.line.fill.background()
        t2(sl, lx + 0.24, 6.02, 0.90, 0.28, [(nm, 12, GRAY, False, CN)])
    kicker(sl, '口径：单位亿千瓦时；计划为年初分解值，不含季度调整。')


# ======================== 05 主角数字：限电率 ========================
def p05(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '限电')
    h1(sl, '限电率 12.0%，多丢了 5.80 亿千瓦时。', y=1.14, size=34)
    big(sl, LM, 2.24, '12.0', '%', size=104, color=RED, unit_size=34)
    sub(sl, LM, 3.76, '同比上升 5 个百分点，限电量由 2.80 亿增至 5.80 亿千瓦时。\n'
                      '断面约束、外送通道受限是主因，短期内无解。', size=17)
    kicker(sl, '口径：限电量＝可发电量−实际发电量；限电率＝限电量÷可发电量。')


# ======================== 06 横条对比：缺口结构 ========================
def p06(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '电量　·　缺口结构')
    h1(sl, '一半是限电，四成是风。', y=1.14, size=34)

    rows = [('限电上升', '3.00', 1.00, RED), ('风资源下降', '2.25', 0.75, AMBER),
            ('大部件损坏', '0.30', 0.10, BLUE), ('主动弃电', '0.15', 0.05, SKY),
            ('项目达产滞后', '0.10', 0.03, LBLUE)]
    for i, (lab, val, r, col) in enumerate(rows):
        hbrow(sl, LM, 2.28 + i * 0.66, CW - 0.60, lab, val, r, color=col,
              label_w=2.60, val_w=1.10)
    line(sl, LM, 5.92, CW - 0.90, LINE, 0.75)
    t2(sl, LM, 6.04, CW - 0.9, 0.30,
       [('合计 −5.80 亿千瓦时', 13, INK, True, CN)])
    kicker(sl, '口径：单位亿千瓦时；按影响电量降序，五项互不重叠。')


# ======================== 07 折线：度电考核走势 ========================
def p07(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '考核　·　度电考核金额')
    h1(sl, '度电考核一路下行，9 月 0.0360 元。', y=1.14, size=32)
    sub(sl, LM, 1.92, '低于对标均值 5.0%，同业排名第三。', size=16)
    dline(sl, LM, 2.60, CW - 0.60, 2.90,
          ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'],
          [('度电考核', [0.046, 0.045, 0.043, 0.042, 0.041, 0.040, 0.039, 0.037, 0.036])],
          [BLUE], 0.050, minv=0.030, fmt='%.4f')
    kicker(sl, '口径：元/千瓦时；度电考核＝当月考核电量÷当月上网电量。')


# ======================== 08 苹果式数据表：成本 ========================
def p08(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '成本')
    h1(sl, '比预算省 3,000 万元。', y=1.14, size=40)
    sub(sl, LM, 1.94, '材料费涨得最快，其余科目基本持平。', size=16)

    rowlines(sl, LM, 2.72, CW - 0.60,
             ['成本项', '本年', '上年', '同比'],
             [2.4, 1.2, 1.2, 1.2],
             [['材料费', '3,200', '3,000', '+6.7%'],
              ['修理费', '2,800', '2,750', '+1.8%'],
              ['人工费', '2,200', '2,150', '+2.3%'],
              ['其他费用', '1,800', '1,760', '+2.3%'],
              ['合计', '10,000', '9,660', '+3.5%']],
             row_h=0.56, fs=14)
    kicker(sl, '口径：万元，累计数；不含折旧与财务费用。')


# ======================== 09 guidance 收尾 ========================
def p09(prs, n):
    sl = aslide(prs)
    eyebrow(sl, '全年指引')
    h1(sl, '四季度补 11.5 亿千瓦时，全年 41.5 亿。', y=1.14, size=34)
    sub(sl, LM, 2.10, '完成率 98.8%，缺口 0.5 亿——可以接受，但不能再降。', size=17)

    for i, (k, v, col) in enumerate([('四季度目标', '11.5', DEEP), ('全年预计', '41.5', DEEP),
                                     ('完成率', '98.8%', GREEN)]):
        cx = LM + i * (CW / 3)
        t2(sl, cx, 3.10, 2.60, 0.28, [(k, 12, LGRAY, False, CN)], spc=30)
        big(sl, cx, 3.42, v, '亿千瓦时' if i < 2 else '', size=56, color=col,
            unit_size=14, w=3.40)

    line(sl, LM, 5.30, CW - 0.90, LINE, 0.75)
    sub(sl, LM, 5.46, '前提：风资源回到常年水平、限电率不高于 12%、新增 0.3GW 按期并网。', size=15)
    kicker(sl, '口径：亿千瓦时；四季度为预计值，按 9 月末在运容量测算。')


def main():
    prs = new_deck()
    p01(prs)
    p02(prs, 1); p03(prs, 2); p04(prs, 3); p05(prs, 4)
    p06(prs, 5); p07(prs, 6); p08(prs, 7); p09(prs, 8)
    # 本样式不加任何转场与元素动画（2026-10-23 定），故此处无 set_morph 调用
    prs.save(OUT)
    print('saved:', OUT)


if __name__ == '__main__':
    main()
