# -*- coding: utf-8 -*-
"""第一性原理风格（深色）骨架示例 —— 深底、巨字、一页一个结论、下判断。

【脱敏说明】设计系统与函数写法完整保留；文案为占位词、数值为演示占位值。
页序（正式版 14 页，本骨架抽 8 页示范）：
  01 封面 → 02 第一性原理（物理上限）→ 03 主角数字（缺口）
  → 04 横条图（因素拆解）→ 05 三个小数字（结构）→ 06 判断页（要么…要么…）
  → 07 目标 vs 实际 → 08 收尾一句话

四条铁律：
  1. 一页只讲一件事：19pt 灰标题 + 巨数字（110—150pt）+ 三个 32—36pt 小数字 + 26pt 判断句
  2. 语气用陈述句，缺口直接写「差 29%」「要么提速，要么改预期」，不写「进一步/持续/有效」
  3. 深底上 653 深蓝只能做大面积色块，不能做小字或细线（看不清）；主强调用最亮的次色
  4. 不加任何动画与转场（翻页即硬切）；用左侧竖向进度脊提供静态推进感
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deck_lib import new_deck, rect, txt, rgb, _fix_font
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'demo_dark.pptx')

# ---------- 深底色板 ----------
BG   = '0A0D12'      # 页面底：近黑
BG2  = '141A22'      # 面板
TRK  = '1A222C'      # 轨道
LN   = '232C37'      # 分隔线
FG   = 'F2F5F8'      # 主文字
FG2  = 'A3AEBB'      # 次级文字
FG3  = '5E6A78'      # 三级文字 / 页脚

# 企业标准色（深底适配：取色卡中最亮的一档做强调，最深一档只做大色块）
C1 = '004A86'        # 最深 —— 封面底纹、大面积色块（不可做小字）
C2 = '006AAF'        # 次深 —— 次级数据、对比条
C3 = '00A2E9'        # 最亮 —— 主强调（深底最清晰）
G1 = '03994C'        # 正向（深底偏暗，少用）
G2 = '83C034'        # 正向（深底首选）
G3 = 'DADA00'        # 中性提示
N1 = 'ED6C00'        # 负向 / 缺口
N2 = 'F6AD02'        # 中间态
N3 = 'FED900'        # 高亮

ACC, RED, AMB, GRN = C3, N1, N2, G2

SW, M = 13.333, 0.85
CW = SW - 2 * M
NUM = 'Bahnschrift'   # Win10+ 自带，DIN 观感
YH  = '微软雅黑'


def dslide(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl._anim_groups = []
    bg = sl.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = rgb(BG)
    return sl


def mix(sl, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, ls=1.0):
    """runs: [(text, size, color, bold, font)]"""
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
        _fix_font(r, font)
    return tb


def label(sl, text, color=ACC):
    """顶部细标签：0.30in 短横线 + 10.5pt 灰字"""
    rect(sl, M, 0.60, 0.30, 0.030, fill=color)
    txt(sl, M + 0.42, 0.47, CW - 0.42, 0.28, text, size=10.5, color=FG3)


def pn(sl, n, total=8):
    """页码 + 左侧竖向进度脊（静态推进感，不用动画）"""
    txt(sl, SW - M - 0.6, 6.94, 0.6, 0.24, '%02d' % n, size=10, color=FG3,
        align=PP_ALIGN.RIGHT, font=NUM)
    x, y0, h = 0.32, 1.34, 4.30
    rect(sl, x, y0, 0.042, h, fill=LN)
    rect(sl, x, y0, 0.042, h * n / total, fill=ACC)


def foot(sl, text, y=6.94):
    txt(sl, M, y, CW - 1.0, 0.24, text, size=9.5, color=FG3)


def hero(sl, x, y, num, unit, color=FG, size=128, unit_size=22, w=9.5):
    """巨大数字 + 单位：同一段内，底部对齐"""
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(size / 72.0 * 1.10))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.line_spacing = 0.92
    r = p.add_run(); r.text = num
    r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = rgb(color)
    _fix_font(r, NUM)
    if unit:
        r2 = p.add_run(); r2.text = '  ' + unit
        r2.font.size = Pt(unit_size); r2.font.bold = False
        r2.font.color.rgb = rgb(FG2)
        _fix_font(r2, YH)
    return tb


def stat(sl, x, y, num, unit, cap, color=FG, size=34, w=2.9):
    """小号数字 + 两行说明，三项并排用"""
    mix(sl, x, y, w, size / 72.0 * 1.16,
        [(num, size, color, True, NUM), (' ' + unit, 12, FG2, False, YH)],
        anchor=MSO_ANCHOR.BOTTOM)
    txt(sl, x, y + size / 72.0 * 1.20, w, 0.52, cap, size=12, color=FG3, ls=1.30)


def judge(sl, text, y=4.92, color=FG, size=26, w=None):
    """判断句：左侧 0.055in 竖条 + 26pt 白字"""
    rect(sl, M, y + 0.04, 0.055, 0.62, fill=ACC)
    txt(sl, M + 0.30, y, w or (CW - 0.30), 0.70, text, size=size, color=color, bold=True,
        ls=1.16)


def hbars(sl, x, y, w, row_h, items, maxv, label_w, vfmt='%.2f', color_default=AMB):
    """横条图：右对齐标签 + 全宽轨道 + 条尾数值"""
    bw = w - label_w - 1.35
    for i, (name, val, col) in enumerate(items):
        cy = y + i * row_h
        txt(sl, x, cy, label_w - 0.24, row_h, name, size=15, color=FG2,
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        rect(sl, x + label_w, cy + row_h * 0.31, bw, row_h * 0.38, fill=TRK)
        L = max(val / maxv * bw, 0.02)
        rect(sl, x + label_w, cy + row_h * 0.31, L, row_h * 0.38, fill=col or color_default)
        txt(sl, x + label_w + L + 0.14, cy, 1.2, row_h, vfmt % val, size=17, color=FG,
            anchor=MSO_ANCHOR.MIDDLE, font=NUM)


# ======================== 01 封面 ========================
def p01(prs):
    sl = dslide(prs)
    rect(sl, 0, 0, 0.34, 7.5, fill=C1)      # 左侧深蓝色块（653 只做大面积）
    txt(sl, M, 1.16, CW, 0.28, '示例能源公司　·　生产管理部', size=11, color=FG3)
    mix(sl, M, 1.92, 11.4, 2.30, [
        ('一个季度，\n', 58, FG, True, YH),
        ('只该记住三个数。', 58, ACC, True, YH),
    ], ls=1.12)
    txt(sl, M, 4.72, 11.0, 0.34, '20XX 年第三季度 · 生产运营', size=16, color=FG2)
    foot(sl, '本稿只讲生产运维，不含电价与售电。')


# ======================== 02 第一性原理：物理上限 ========================
def p02(prs, n):
    sl = dslide(prs)
    label(sl, '第一性原理 / 从物理上限倒推')
    txt(sl, M, 0.92, CW, 0.40, '装机 3,000MW × 8,760 小时 = 262.8 亿千瓦时', size=19, color=FG2)
    hero(sl, M, 1.50, '13.7', '%', size=110, color=ACC)
    txt(sl, M, 3.34, 9.0, 0.30, '容量的实际利用率（容量系数）', size=14, color=FG3)

    for i, (num, unit, cap, col) in enumerate([('262.8', '亿千瓦时', '物理上限产量', FG2),
                                               ('36.0', '亿千瓦时', '全年实际产量', FG),
                                               ('8,760', '小时', '理论满发时长', FG3)]):
        stat(sl, M + i * 3.95, 3.96, num, unit, cap, color=col, size=34, w=3.55)

    judge(sl, '9 成的产量，卡在风、卡在限电——不是卡在设备。')
    foot(sl, '自算：容量系数＝实际产量÷（装机×8760）；装机 3,000MW、全年预计 36.0 亿千瓦时。')
    pn(sl, n)


# ======================== 03 主角数字：缺口 ========================
def p03(prs, n):
    sl = dslide(prs)
    label(sl, '电量 / 缺口')
    txt(sl, M, 0.92, CW, 0.40, '前三季度，少发了', size=19, color=FG2)
    hero(sl, M, 1.34, '2.00', '亿千瓦时', size=128, color=RED)
    txt(sl, M, 3.30, 10.5, 0.62, '相当于 12 天的全公司产量，且四季度补不回来。',
        size=22, color=FG2, ls=1.24)

    for i, (num, unit, cap, col) in enumerate([('93.8', '%', '完成率', FG),
                                               ('32.00', '亿千瓦时', '计划电量', FG2),
                                               ('−8.0', '%', '同比', RED)]):
        stat(sl, M + i * 3.95, 4.24, num, unit, cap, color=col, size=34, w=3.55)
    foot(sl, '口径：上网电量，含试运行；同比对比去年同期 32.60 亿千瓦时。')
    pn(sl, n)


# ======================== 04 横条图：因素拆解 ========================
def p04(prs, n):
    sl = dslide(prs)
    label(sl, '缺口 / 拆到不能再拆')
    txt(sl, M, 0.90, CW, 0.40, '五项因素，只有一项能自己解决。', size=19, color=FG2)
    hbars(sl, M, 1.58, CW, 0.80,
          [('限电上升', 3.00, RED), ('风资源下降', 2.25, AMB),
           ('大部件损坏', 0.30, ACC), ('主动弃电', 0.15, C2),
           ('项目达产滞后', 0.10, C2)],
          3.00, 3.10, vfmt='−%.2f')
    judge(sl, '限电与风占 9 成——靠努力补不回来，只能改预期。', y=5.72)
    foot(sl, '口径：亿千瓦时；按影响电量降序，五项互不重叠。')
    pn(sl, n)


# ======================== 05 三个小数字：结构 ========================
def p05(prs, n):
    sl = dslide(prs)
    label(sl, '考核 / 结构决定对策')
    txt(sl, M, 0.92, CW, 0.40, '考核电量里，七成来自一件事。', size=19, color=FG2)
    hero(sl, M, 1.44, '76', '%', size=110, color=ACC)
    txt(sl, M, 3.22, 9.0, 0.30, '来自功率预测偏差', size=14, color=FG3)
    for i, (num, unit, cap, col) in enumerate([('3.40', '亿千瓦时', '总考核电量', FG2),
                                               ('0.0400', '元/千瓦时', '度电考核', GRN),
                                               ('−5.0', '%', '低于对标', GRN)]):
        stat(sl, M + i * 3.95, 3.86, num, unit, cap, color=col, size=34, w=3.55)
    judge(sl, '预测提 1 个点，少交 0.20 亿千瓦时。')
    foot(sl, '口径：度电考核＝考核电量÷上网电量；对标为区域 9 家集团均值。')
    pn(sl, n)


# ======================== 06 判断页：要么…要么… ========================
def p06(prs, n):
    sl = dslide(prs)
    label(sl, '限电 / 无解的题怎么答')
    txt(sl, M, 0.92, CW, 0.40, '限电率 12.0%，同比涨了 5 个点。', size=19, color=FG2)
    hero(sl, M, 1.40, '5.80', '亿千瓦时', size=118, color=RED)
    txt(sl, M, 3.30, 10.5, 0.34, '全年丢掉的可发电量', size=14, color=FG3)

    for i, (num, unit, cap, col) in enumerate([('2.80', '亿千瓦时', '去年同期限电', FG2),
                                               ('2.07', '倍', '同比放大', RED),
                                               ('12.0', '%', '本年限电率', AMB)]):
        stat(sl, M + i * 3.95, 3.94, num, unit, cap, color=col, size=34, w=3.55)
    judge(sl, '断面约束两年内不会解——要么抢通道，要么把电卖在本地。')
    foot(sl, '口径：限电量＝可发电量−实际发电量；倍数为 5.80÷2.80。')
    pn(sl, n)


# ======================== 07 目标 vs 实际 ========================
def p07(prs, n):
    sl = dslide(prs)
    label(sl, '全年 / 预期要诚实')
    txt(sl, M, 0.92, CW, 0.40, '年初说 42.0，现在只能给 41.5。', size=19, color=FG2)
    hero(sl, M, 1.44, '98.8', '%', size=110, color=AMB)
    txt(sl, M, 3.22, 9.0, 0.30, '全年完成率', size=14, color=FG3)

    for i, (num, unit, cap, col) in enumerate([('42.0', '亿千瓦时', '年初目标', FG2),
                                               ('41.5', '亿千瓦时', '最新预计', FG),
                                               ('0.5', '亿千瓦时', '缺口', RED)]):
        stat(sl, M + i * 3.95, 3.86, num, unit, cap, color=col, size=34, w=3.55)
    judge(sl, '0.5 亿的缺口可以接受，但不能再降——四季度每月必须跑满计划。')
    foot(sl, '自算：全年＝前三季度 30.00＋四季度预计 11.5；完成率＝41.5÷42.0。')
    pn(sl, n)


# ======================== 08 收尾 ========================
def p08(prs, n):
    sl = dslide(prs)
    label(sl, '一句话')
    mix(sl, M, 1.90, 11.6, 2.60, [
        ('电发少了，\n', 52, FG, True, YH),
        ('不是设备的问题。', 52, ACC, True, YH),
    ], ls=1.16)
    txt(sl, M, 4.40, 11.0, 0.80,
        '设备可用率 98.2%，人员没减，检修没拖——少的是风和通道。', size=20, color=FG2, ls=1.30)
    foot(sl, '数据来源：生产管理系统月报；派生值页脚已标自算。')
    pn(sl, n)


def main():
    prs = new_deck()
    p01(prs)
    p02(prs, 1); p03(prs, 2); p04(prs, 3); p05(prs, 4)
    p06(prs, 5); p07(prs, 6); p08(prs, 7)
    # 本样式不加任何动画与转场（2026-10-23 定），推进感靠 pn() 的静态进度脊
    prs.save(OUT)
    print('saved:', OUT)


if __name__ == '__main__':
    main()
