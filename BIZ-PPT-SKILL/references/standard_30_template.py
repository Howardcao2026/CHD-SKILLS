# -*- coding: utf-8 -*-
"""常规版（详细版）骨架示例 —— 白底、正文 14pt 起、结论句 20pt、明细进表格。

【脱敏说明】
本文件是版式骨架，不是某家公司的真实报告：
  · 页型（每页用哪几个版式函数、放在什么坐标）完整保留，可直接照抄；
  · 文案一律用「指标一 / 场站A / 季度要点」这类占位词，数值为演示占位值；
  · 页数压缩到 15 页，覆盖详细版全部页型；正式使用按下方「详细版页序清单」扩展。

详细版页序清单（每页一个函数，函数名 = 页码）：
  01 封面 / 02 目录 / 03 分隔页
  04 核心指标总览（4 大卡 + 明细表）
  05 电量：完成情况（4 卡 + 分组柱图 + 右侧要点）
  06 电量：偏差归因（横条图 + 右侧说明）
  07 场站完成率（明细表 + 达标分布条）
  08 小时数：总体（折线图 + 要点）
  09 小时数：场站分布（横条图）
  10 限电：走势（柱图 + 结论）
  11 对标：与区域/同行比较（表格 + 色阶）
  12 考核：两张巨数字大卡 + 结构条
  13 考核：新政与对策（左右对比）
  14 储能：运行表现（3 卡 + 分站表）
  15 储能：配套运行（充放电量表）
  16 成本：总量与结构（结构条 + 表）
  17 降本专项（进度条行）
  18 资本化（明细表）
  19 设备：风险与跳闸（2×3 风险卡）
  20 设备：无人值守（完成率进度 + 表）
  21 增量项目（里程碑表）
  22 全年目标预计（目标 vs 实际柱图）
  23 风险与难点（风险卡）
  24 攻坚方向（三列）
  25 任务清单与次年安排（3 列表格）
  26 总结：一页看懂（大数字块）
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deck_lib import *
from deck_lib import _fix_font
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'demo_standard.pptx')

F14, F18, F13, F12 = 14, 18, 13, 12


# ---------- 本版专用小部件 ----------
def H(sl, text):
    """一句话结论：20pt，句末标点必须完整。"""
    headline(sl, text, size=20)


def kpi_big(sl, x, y, w, h, label, value, note, accent=BLUE, vcolor=NAVY, vsize=30,
            unit='', ncolor=ORANGE, fill=LIGHT2):
    """大号 KPI 卡：顶部色条 + 标签 + 巨数字 + 橙色注脚。"""
    small = h < 1.55
    rect(sl, x, y, w, h, fill=fill, line=LINE, lw=0.75)
    rect(sl, x, y, w, 0.075, fill=accent)
    txt(sl, x + 0.18, y + (0.13 if small else 0.20), w - 0.28, 0.28, label,
        size=(11.5 if small else 12), color=GRAY2)
    vs = min(vsize, 22) if small else vsize
    tb = sl.shapes.add_textbox(Inches(x + 0.16), Inches(y + (0.40 if small else 0.56)),
                               Inches(w - 0.22), Inches(0.46 if small else 0.60))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = value
    r.font.size = Pt(vs); r.font.bold = True; r.font.color.rgb = rgb(vcolor)
    _fix_font(r)
    if unit:
        r2 = p.add_run(); r2.text = unit
        r2.font.size = Pt(13); r2.font.color.rgb = rgb(GRAY2)
        _fix_font(r2)
    txt(sl, x + 0.18, y + h - (0.36 if small else 0.46), w - 0.26, 0.34, note,
        size=(11 if small else 12), color=ncolor, ls=1.15)


def body(sl, x, y, w, h, items, size=F14, gap=10):
    """要点区：('引导词：', '正文') 或纯字符串。14pt 下每行约 30 字。"""
    arr = []
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            lead, rest = it
            arr.append({'runs': [(lead, size, NAVY, True), (rest, size, GRAY, False)],
                        'ls': 1.34, 'before': 0 if i == 0 else gap})
        else:
            arr.append({'runs': [('· ', size, BLUE, True), (it, size, GRAY, False)],
                        'ls': 1.34, 'before': 0 if i == 0 else gap, 'hang': True,
                        'marL': 0.24, 'indent': -0.24})
    paras(sl, x, y, w, h, arr)


def bar_row(sl, x, y, w, label, value, ratio, color=BLUE, vsize=14):
    """细进度行：标签 + 轨道 + 比例条 + 右侧数值。"""
    txt(sl, x, y, 2.60, 0.32, label, size=vsize, color=GRAY)
    rect(sl, x + 2.70, y + 0.08, w - 4.10, 0.16, fill='E3EDF6')
    rect(sl, x + 2.70, y + 0.08, max((w - 4.10) * ratio, 0.03), 0.16, fill=color)
    txt(sl, x + w - 1.30, y, 1.30, 0.32, value, size=vsize, color=NAVY, bold=True,
        align=PP_ALIGN.RIGHT)


def note_panel(sl, x, y, w, h, text, accent=BLUE, size=F14):
    """左竖条备注块。"""
    panel(sl, x, y, w, h, fill=LIGHT2)
    rect(sl, x, y, 0.05, h, fill=accent)
    txt(sl, x + 0.26, y + 0.13, w - 0.52, h - 0.26, text, size=size, color=GRAY, ls=1.26)


# ================= 01 封面 =================
def p01(prs):
    sl = blank(prs)
    txt(sl, 0, 2.10, SW, 0.32, '示例能源公司', size=12, color=BLUE, bold=True,
        align=PP_ALIGN.CENTER)
    txt(sl, 0, 2.52, SW, 0.60, '20XX年第三季度经营活动分析会', size=26, color=NAVY, bold=True,
        align=PP_ALIGN.CENTER)
    txt(sl, 0, 3.14, SW, 0.80, '生产运营报告', size=40, color=NAVY, bold=True,
        align=PP_ALIGN.CENTER)
    rect(sl, SW / 2 - 0.45, 4.05, 0.9, 0.032, fill=BLUE)
    txt(sl, 0, 4.28, SW, 0.30, '生产管理部　|　20XX年10月', size=14, color=GRAY2,
        align=PP_ALIGN.CENTER)
    # 左下角淡色线稿（风机 + 光伏阵列），不影响文字
    rect(sl, 0.70, 6.74, 3.95, 0.012, fill=PALE)
    turbine(sl, 1.62, 5.32, 1.42, 1.02, color=PALE, lw=2.2, rot=18)
    turbine(sl, 2.86, 6.10, 0.64, 0.40, color=PALE, lw=1.5, rot=74)
    gx, gy, gw, gh = 3.42, 6.36, 1.02, 0.38
    rect(sl, gx, gy, gw, gh, fill=None, line=PALE, lw=1.2)
    for i in range(1, 3):
        rect(sl, gx, gy + gh * i / 3, gw, 0.007, fill=PALE)
    for i in range(1, 4):
        rect(sl, gx + gw * i / 4, gy, 0.007, gh, fill=PALE)
    rect(sl, gx - 0.05, gy + gh, 0.007, 0.12, fill=PALE)
    rect(sl, gx + gw + 0.04, gy + gh, 0.007, 0.12, fill=PALE)


# ================= 02 目录 =================
def p02(prs, num):
    sl = blank(prs)
    page_title(sl, '目　录', num)
    items = ['前三季度指标完成与偏差分析', '全年目标预计及风险难点',
             '四季度攻坚及次年安排', '总结：一页看懂']
    for i, t in enumerate(items):
        y = 1.66 + i * 1.34
        rect(sl, 2.90, y, 0.80, 0.80, fill=WHITE, line=BLUE, lw=1.2)
        txt(sl, 2.90, y + 0.16, 0.80, 0.48, str(i + 1), size=24, color=BLUE, bold=True,
            align=PP_ALIGN.CENTER)
        rect(sl, 3.86, y + 0.05, 6.60, 0.70, fill=BLUE if i < 3 else NAVY)
        txt(sl, 3.86, y + 0.18, 6.60, 0.44, t, size=20, color=WHITE, bold=True,
            align=PP_ALIGN.CENTER)


# ================= 分隔页 =================
def part(prs, num, no, title):
    sl = blank(prs)
    page_title(sl, '', num, rule=False, label=False)
    rect(sl, 1.62, 2.34, 1.55, 1.55, fill=BLUE)
    rect(sl, 1.86, 2.56, 1.55, 1.55, fill=WHITE, line=BLUE, lw=1.4)
    txt(sl, 1.86, 2.90, 1.55, 0.80, no, size=42, color=NAVY, align=PP_ALIGN.CENTER)
    txt(sl, 1.86, 4.28, 1.55, 0.28, 'PART', size=12, color=BLUE, bold=True,
        align=PP_ALIGN.CENTER)
    txt(sl, 3.72, 2.98, 8.6, 0.55, title, size=24, color=NAVY, bold=True)
    arrow_right(sl, 3.75, 3.72, 5.95, color=BLUE, lw=1.4)


# ================= 04 核心指标总览（4 大卡 + 明细表） =================
def p04(prs, num):
    sl = blank(prs)
    page_title(sl, '核心指标总览', num)
    H(sl, '电量完成率93.8%、欠发2.00亿千瓦时，考核与成本优于对标。')

    y, h, w, g = 1.72, 1.52, 3.015, 0.145
    kpi_big(sl, M, y, w, h, '上网电量（亿千瓦时）', '30.00', '完成率93.8%，欠发2.00亿', BLUE, NAVY, 30)
    kpi_big(sl, M + (w + g), y, w, h, '限电率（%）', '12.0', '同比+5pp，优于对标1.0pp', BLUE, NAVY, 30)
    kpi_big(sl, M + 2 * (w + g), y, w, h, '度电考核（元/千瓦时）', '0.0400', '低于对标5.0%，同业第三',
            GREEN, GREEN, 30, ncolor=GREEN)
    kpi_big(sl, M + 3 * (w + g), y, w, h, '运维成本（万元）', '10,000', '支出率77%，全年13,000',
            ORANGE, ORANGE, 30)

    chip(sl, M, 3.44, '其他关键生产指标', size=15)
    rows = [['指标', '完成', '对标 / 偏差'],
            ['风电利用小时', '1,200小时', '对标+5.0%，相对+4.5pp'],
            ['光伏利用小时', '1,050小时', '对标+25%，相对-6.0pp'],
            ['无人值守完成率', '90.0%', '1,800MW／2,000MW'],
            ['集电线路跳闸', '20次', '同比减少15次'],
            ['大部件下塔维修', '15次', '单次平均14天']]
    table(sl, M, 3.86, CW, [2.90, 3.60, 5.99], rows, row_h=0.48, header_h=0.42, fs=F13, hfs=F13,
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT], bold_first_col=True)


# ================= 05 电量完成情况（4 卡 + 分组柱图 + 要点） =================
def p05(prs, num):
    sl = blank(prs)
    page_title(sl, '一  发电量：完成情况', num)
    H(sl, '完成率93.8%、欠发2.00亿千瓦时，存量项目占缺口七成。')

    y, h, w, g = 1.72, 1.16, 3.015, 0.145
    kpi_big(sl, M, y, w, h, '计划电量（亿千瓦时）', '32.00', '年度计划分解', BLUE, NAVY, 24,
            ncolor=GRAY2)
    kpi_big(sl, M + (w + g), y, w, h, '实际电量（亿千瓦时）', '30.00', '1—9月累计', BLUE, NAVY, 24,
            ncolor=GRAY2)
    kpi_big(sl, M + 2 * (w + g), y, w, h, '完成率（%）', '93.8', '低于攻坚值', ORANGE, ORANGE, 24)
    kpi_big(sl, M + 3 * (w + g), y, w, h, '欠发电量（亿千瓦时）', '2.00', '四季度须补回',
            RED, RED, 24, ncolor=RED)

    band(sl, M, 3.02, 7.20, 0.34, '分项目类型完成情况（亿千瓦时）', fill=BLUE, size=13)
    schart(sl, M - 0.05, 3.44, 7.30, 2.95,
           ['存量项目', '上年增量', '本年增量'],
           [('计划', [20.00, 10.00, 2.00]), ('实际', [19.00, 9.80, 1.20])],
           kind='column', colors=[PALE, BLUE], gap=85, fmt='0.00', label_size=13,
           legend=True, vsize=11, show_vaxis=False, vmax=25, label_color=NAVY, cat_size=13)

    x = 7.85
    chip(sl, x, 3.02, '缺口构成', size=15)
    body(sl, x, 3.46, 5.06, 2.90, [
        ('按类型：', '存量占七成、上年增量占两成、本年增量占一成。'),
        ('按场站：', '30座场站完成率均值93.6%，5座低于90%、3座超过100%。'),
        ('回升：', '9月单月3.20亿千瓦时，欠发收窄至0.10亿，连续两月回升。'),
    ], size=F14, gap=12)


# ================= 06 偏差归因（横条图 + 说明） =================
def p06(prs, num):
    sl = blank(prs)
    page_title(sl, '一  发电量：偏差归因', num)
    H(sl, '五项因素合计影响5.80亿千瓦时，风资源与限电占九成。')

    band(sl, M, 1.72, 6.60, 0.34, '各因素影响电量（亿千瓦时）', fill=BLUE, size=13)
    cats = ['主动弃电', '项目达产滞后', '大部件损坏', '限电率上升', '风资源下降']
    vals = [0.15, 0.10, 0.30, 3.00, 2.25]
    schart(sl, M - 0.05, 2.12, 6.70, 3.76, cats, [('影响电量', vals)], kind='bar',
           colors=[RED], gap=115, fmt='"-"0.00', label_size=13, label_color=NAVY,
           cat_size=13, show_vaxis=False, vmax=3.6, vmin=0)
    txt(sl, M, 6.06, 6.60, 0.32, '合计 -5.80 亿千瓦时，同比下降约16.0%', size=F14,
        color=RED, bold=True)

    x = 7.30
    chip(sl, x, 1.72, '归因说明', size=15)
    body(sl, x, 2.16, 5.62, 4.50, [
        ('风资源：', '年均风速同比下降12%，三个场站降幅超15%，影响2.25亿。'),
        ('限电：', '限电量由2.80亿增至5.80亿千瓦时，断面制约，影响3.00亿。'),
        ('大部件：', '叶片、齿轮箱损坏致长停，年内大部件下塔15次。'),
        ('达产与弃电：', '新项目达产滞后影响0.10亿；低价时段主动弃电0.15亿。'),
    ], size=F14, gap=14)


# ================= 07 场站完成率（表 + 达标分布） =================
def p07(prs, num):
    sl = blank(prs)
    page_title(sl, '一  发电量：场站完成率分布', num)
    H(sl, '30座场站中5座完成率低于90%，缺口集中在4座。')

    rows = [['场站', '计划', '实际', '完成率', '缺口', '评级'],
            ['场站A', '4.00', '3.20', '80.0%', '-0.80', '未达标'],
            ['场站B', '3.50', '3.05', '87.1%', '-0.45', '未达标'],
            ['场站C', '3.00', '2.70', '90.0%', '-0.30', '接近'],
            ['场站D', '2.50', '2.45', '98.0%', '-0.05', '达标'],
            ['场站E', '2.00', '2.10', '105.0%', '+0.10', '超额'],
            ['其他25座', '17.00', '16.50', '97.1%', '-0.50', '达标'],
            ['合计', '32.00', '30.00', '93.8%', '-2.00', '—']]
    cf = {(1, 5): 'FBEAE0', (2, 5): 'FBEAE0', (3, 5): 'FFF7E6', (7, 5): LIGHT}
    cc = {(1, 5): RED, (2, 5): RED, (3, 5): ORANGE, (1, 3): RED, (2, 3): RED,
          (5, 3): GREEN, (7, 3): NAVY}
    table(sl, M, 1.76, 7.90, [1.90, 1.10, 1.10, 1.30, 1.10, 1.40], rows, row_h=0.44,
          header_h=0.42, fs=F13, hfs=F13, bold_first_col=True,
          aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 5, cell_fill=cf, cell_color=cc)

    x = 8.55
    chip(sl, x, 1.76, '完成率分档', size=15)
    for i, (lab, n, r, c) in enumerate([
            ('100% 以上（超额）', 3, 0.10, GREEN),
            ('95%—100%（达标）', 15, 0.50, BLUE),
            ('90%—95%（接近）', 7, 0.23, SKY),
            ('90% 以下（未达标）', 5, 0.17, RED)]):
        yy = 2.28 + i * 0.98
        txt(sl, x, yy, 4.30, 0.30, lab, size=F13, color=GRAY)
        rect(sl, x, yy + 0.36, 4.30, 0.22, fill='E3EDF6')
        rect(sl, x, yy + 0.36, 4.30 * r, 0.22, fill=c)
        txt(sl, x + 4.30 * r + 0.08, yy + 0.32, 0.90, 0.30, f'{n}座', size=F13,
            color=NAVY, bold=True)
    note_panel(sl, x, 6.02, 4.30, 0.62, '未达标5座合计缺口1.55亿，占全部缺口78%。', accent=RED, size=F13)


# ================= 08 小时数总体（折线图） =================
def p08(prs, num):
    sl = blank(prs)
    page_title(sl, '二  利用小时：月度走势', num)
    H(sl, '风电利用小时1,200小时，同比低80小时，9月差距收窄。')

    band(sl, M, 1.72, 8.30, 0.34, '月度利用小时（小时）', fill=BLUE, size=13)
    schart(sl, M - 0.05, 2.16, 8.40, 4.10,
           ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'],
           [('本年', [130, 125, 160, 155, 140, 120, 105, 125, 140]),
            ('上年', [150, 140, 175, 165, 150, 130, 115, 135, 160])],
           kind='line_marker', colors=[BLUE, PALE], fmt='0', label_size=12,
           legend=True, show_vaxis=True, gridlines=True, vsize=11, cat_size=F13,
           vmax=200, vmin=80)

    x = 8.95
    chip(sl, x, 1.72, '关键读数', size=15)
    body(sl, x, 2.16, 3.96, 4.30, [
        ('累计：', '1,200小时，同比减少80小时，降幅6.3%。'),
        ('对标：', '比区域均值高5.0%，相对领先4.5pp。'),
        ('拐点：', '7月最低105小时，8—9月连续回升。'),
        ('判断：', '四季度需月均150小时才能补回缺口。'),
    ], size=F14, gap=13)


# ================= 09 场站分布（横条图） =================
def p09(prs, num):
    sl = blank(prs)
    page_title(sl, '二  利用小时：场站分布', num)
    H(sl, '场站间差异达420小时，风资源与机型是主因。')

    band(sl, M, 1.72, 7.60, 0.34, '各场站累计利用小时（小时）', fill=BLUE, size=13)
    cats = ['场站A', '场站B', '场站C', '场站D', '场站E', '场站F']
    vals = [1450, 1320, 1260, 1180, 1090, 1030]
    schart(sl, M - 0.05, 2.14, 7.70, 4.30, cats, [('利用小时', vals)], kind='bar',
           colors=[BLUE], gap=120, fmt='0', label_size=13, label_color=NAVY,
           cat_size=F13, show_vaxis=False, vmax=1700, vmin=0)

    x = 8.25
    chip(sl, x, 1.72, '差异拆解', size=15)
    body(sl, x, 2.16, 4.66, 4.30, [
        ('最高：', '场站A 1,450小时，风速高于均值8%。'),
        ('最低：', '场站F 1,030小时，机型轮毂高度偏低。'),
        ('共性：', '6座场站中4座低于均值，集中在同一资源区。'),
        ('对策：', '低风速机组技改列入次年计划。'),
    ], size=F14, gap=13)


# ================= 12 考核：两张巨数字大卡 =================
def p12(prs, num):
    sl = blank(prs)
    page_title(sl, '三  并网考核：总考核电量与度电考核', num)
    H(sl, '总考核电量3.40亿千瓦时，度电考核0.0400元、低于对标5.0%。')

    y = 1.72
    kpi_big(sl, M, y, 5.86, 1.30, '总考核电量（亿千瓦时）', '3.40', '占上网电量 11.3%',
            BLUE, NAVY, 40, ncolor=GRAY2)
    kpi_big(sl, M + 6.04, y, 5.86, 1.30, '度电考核金额（元/千瓦时）', '0.0400',
            '低于对标均值 5.0%，同业第三', GREEN, GREEN, 40, ncolor=GREEN)
    chip(sl, M, 3.26, '考核结构（占月度考核电量比重）', size=15)
    bar_row(sl, M, 3.74, 7.40, '功率预测偏差', '52%—78%', 0.78, BLUE)
    bar_row(sl, M, 4.34, 7.40, '发电计划曲线', '12%—46%', 0.46, BLUE)
    bar_row(sl, M, 4.94, 7.40, '功率预测（AGC）', '8%—19%', 0.19, SKY)

    x = 8.20
    chip(sl, x, 3.26, '三点结论', size=15)
    body(sl, x, 3.70, 4.70, 2.40, [
        ('结构：', '预测偏差占七成以上，是治理重点。'),
        ('水平：', '度电考核低于对标，处于同业前列。'),
        ('空间：', '预测准确率每提升1pp，可减考核约0.20亿千瓦时。'),
    ], size=F14, gap=13)
    note_panel(sl, M, 5.72, CW, 0.78,
               '口径：总考核电量＝各月考核电量合计；度电考核金额＝考核电量÷上网电量。'
               '两项均为本季度累计值，不含跨期调整。', size=F12)


# ================= 16 成本总量与结构 =================
def p16(prs, num):
    sl = blank(prs)
    page_title(sl, '四  运维成本：总量与结构', num)
    H(sl, '运维成本10,000万元、支出率77%，同比增3.5%。')

    y = 1.74
    kpi_big(sl, M, y, 3.015, 1.16, '累计成本（万元）', '10,000', '预算13,000', BLUE, NAVY, 24,
            ncolor=GRAY2)
    kpi_big(sl, M + 3.16, y, 3.015, 1.16, '支出率（%）', '77.0', '时序进度78%', BLUE, NAVY, 24,
            ncolor=GRAY2)
    kpi_big(sl, M + 6.32, y, 3.015, 1.16, '同比增幅（%）', '+3.5', '材料与服务涨价', ORANGE, ORANGE, 24)
    kpi_big(sl, M + 9.48, y, 3.015, 1.16, '单位成本（元/千瓦时）', '0.033', '同比持平', GREEN, GREEN, 24,
            ncolor=GREEN)

    band(sl, M, 3.06, 7.40, 0.34, '成本构成（万元）', fill=BLUE, size=13)
    rows = [['成本项', '本年', '上年', '同比'],
            ['材料费', '3,200', '3,000', '+6.7%'],
            ['修理费', '2,800', '2,750', '+1.8%'],
            ['人工费', '2,200', '2,150', '+2.3%'],
            ['其他费用', '1,800', '1,760', '+2.3%'],
            ['合计', '10,000', '9,660', '+3.5%']]
    table(sl, M, 3.46, 7.40, [2.40, 1.70, 1.70, 1.60], rows, row_h=0.48, header_h=0.40,
          fs=F13, hfs=F13, bold_first_col=True,
          aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 3,
          cell_color={(5, 0): NAVY, (5, 1): NAVY, (5, 2): NAVY})

    x = 8.05
    chip(sl, x, 3.06, '成本管控三条', size=15)
    body(sl, x, 3.50, 4.86, 3.00, [
        ('材料费：', '涨幅最高，集中在大部件备件。'),
        ('修理费：', '基本持平，检修计划完成率96%。'),
        ('人工费：', '随人员结构优化小幅上升。'),
    ], size=F14, gap=12)


# ================= 17 降本专项（进度条行） =================
def p17(prs, num):
    sl = blank(prs)
    page_title(sl, '四  降本专项：八项措施进度', num)
    H(sl, '八项措施累计降本1,200万元，完成年度目标的80%。')

    chip(sl, M, 1.74, '各措施完成进度（万元 / 完成率）', size=15)
    data = [('备件联合储备', 320, 0.91, BLUE), ('检修自主化', 260, 0.87, BLUE),
            ('以大代小置换', 180, 0.72, BLUE), ('功率预测优化', 150, 0.83, SKY),
            ('无人值守减员', 120, 0.60, SKY), ('备件国产化', 100, 0.50, SKY),
            ('车辆与差旅', 45, 0.90, PALE), ('其他', 25, 0.42, PALE)]
    for i, (lab, amt, r, c) in enumerate(data):
        yy = 2.24 + i * 0.56
        txt(sl, M, yy, 2.00, 0.30, lab, size=F14, color=GRAY)
        rect(sl, M + 2.10, yy + 0.07, 3.50, 0.18, fill='E3EDF6')
        rect(sl, M + 2.10, yy + 0.07, 3.50 * r, 0.18, fill=c)
        txt(sl, M + 5.75, yy, 0.95, 0.30, f'{amt}', size=F14, color=NAVY, bold=True,
            align=PP_ALIGN.RIGHT)
        txt(sl, M + 6.85, yy, 0.78, 0.30, f'{r*100:.0f}%', size=F14, color=GRAY, bold=True,
            align=PP_ALIGN.RIGHT)

    note_panel(sl, 8.30, 1.74, 4.61, 1.14,
               '合计降本1,200万元，年度目标1,500万元，完成率80%。', accent=GREEN)
    chip(sl, 8.30, 3.10, '薄弱环节', size=15)
    body(sl, 8.30, 3.54, 4.61, 2.60, [
        ('备件国产化：', '进度50%，认证周期长，须并行送检。'),
        ('无人值守：', '进度60%，通信改造滞后一个月。'),
    ], size=F14, gap=13)


# ================= 19 设备风险（2×3 风险卡） =================
def p19(prs, num):
    sl = blank(prs)
    page_title(sl, '五  设备治理：风险与跳闸', num)
    H(sl, '集电线路跳闸20次、同比减15次，大部件仍是最大风险。')

    chip(sl, M, 1.74, '六类风险与处置', size=15)
    cards = [('01', '大部件损坏', '叶片、齿轮箱长停，单次平均14天。', '处置：备用部件前置、48小时复役', RED),
             ('02', '集电线路跳闸', '20次，同比减少15次，电缆头为主因。', '处置：电缆头批次更换', ORANGE),
             ('03', '变桨系统', '故障率0.8次/台年，集中在上年机型。', '处置：专项技改列入次年', ORANGE),
             ('04', '功率预测偏差', '准确率82%，低于对标3pp。', '处置：预测模型迭代+人工复核', BLUE),
             ('05', '无人值守改造', '完成率90%，尚有200MW未投运。', '处置：通信改造与主机扩容', BLUE),
             ('06', '备件供应', '关键备件到货周期延长至60天。', '处置：联合储备、提前锁量', SKY)]
    for i, (no, title, desc, act, col) in enumerate(cards):
        cx = M + (i % 3) * 4.16
        cy = 2.20 + (i // 3) * 2.10
        rect(sl, cx, cy, 3.96, 1.88, fill=LIGHT2, line=LINE, lw=0.75)
        rect(sl, cx, cy, 3.96, 0.07, fill=col)
        txt(sl, cx + 0.20, cy + 0.24, 0.70, 0.40, no, size=20, color=col, bold=True)
        txt(sl, cx + 0.95, cy + 0.28, 2.85, 0.34, title, size=15, color=NAVY, bold=True)
        txt(sl, cx + 0.20, cy + 0.80, 3.56, 0.62, desc, size=F13, color=GRAY, ls=1.28)
        rect(sl, cx + 0.20, cy + 1.48, 3.56, 0.008, fill=LINE)
        txt(sl, cx + 0.20, cy + 1.56, 3.56, 0.26, act, size=11.5, color=col)


# ================= 22 全年目标预计（目标 vs 实际柱图） =================
def p22(prs, num):
    sl = blank(prs)
    page_title(sl, '六  全年目标预计', num)
    H(sl, '全年预计完成41.5亿千瓦时，缺口0.5亿、完成率98.8%。')

    band(sl, M, 1.72, 7.60, 0.34, '全年电量：目标 / 预计（亿千瓦时）', fill=BLUE, size=13)
    schart(sl, M - 0.05, 2.16, 7.70, 3.80,
           ['一季度', '二季度', '三季度', '四季度E'],
           [('目标', [10.0, 11.5, 10.5, 9.5]), ('预计', [9.8, 11.0, 9.2, 9.5])],
           kind='column', colors=[PALE, BLUE], gap=90, fmt='0.0', label_size=13,
           legend=True, vsize=11, show_vaxis=False, vmax=14, label_color=NAVY, cat_size=F13)
    txt(sl, M, 6.06, 7.60, 0.30, '全年缺口 0.5 亿千瓦时，四季度须与计划持平。', size=F14,
        color=RED, bold=True)

    x = 8.25
    chip(sl, x, 1.72, '判断与前提', size=15)
    body(sl, x, 2.16, 4.66, 4.20, [
        ('基准情形：', '风资源恢复到常年水平，限电率不高于12%。'),
        ('风险情形：', '风资源再降5%，全年缺口扩大至1.5亿。'),
        ('对冲：', '四季度新增装机0.3GW贡献电量0.4亿。'),
    ], size=F14, gap=14)


# ================= 24 攻坚方向（三列） =================
def p24(prs, num):
    sl = blank(prs)
    page_title(sl, '七  四季度攻坚方向', num)
    H(sl, '三条主线：补电量、降考核、控成本。')

    cols = [('补电量', BLUE, ['抢发高价时段电量', '新增装机按期并网', '大部件48小时复役']),
            ('降考核', SKY, ['预测准确率提至85%', 'AGC联控优化', '计划曲线申报精细化']),
            ('控成本', GREEN, ['备件国产化提速', '检修自主化扩面', '无人值守全投运'])]
    for i, (title, col, items) in enumerate(cols):
        x = M + i * 4.16
        rect(sl, x, 1.80, 3.96, 4.60, fill=LIGHT2, line=LINE, lw=0.75)
        rect(sl, x, 1.80, 3.96, 0.52, fill=col)
        txt(sl, x, 1.90, 3.96, 0.34, title, size=18, color=WHITE, bold=True,
            align=PP_ALIGN.CENTER)
        for j, t in enumerate(items):
            yy = 2.62 + j * 1.16
            txt(sl, x + 0.26, yy, 3.44, 0.30, f'0{j+1}', size=13, color=col, bold=True)
            txt(sl, x + 0.26, yy + 0.34, 3.44, 0.72, t, size=F14, color=GRAY, ls=1.25)


# ================= 25 任务清单（3 列表格，无模糊表述） =================
def p25(prs, num):
    sl = blank(prs)
    page_title(sl, '七  四季度重点任务清单', num)
    H(sl, '十项任务全部明确交付物与时点，按月跟踪。')

    rows = [['序号', '任务', '交付物 / 完成标志', '时点'],
            ['1', '欠发场站专项提升', '5座场站逐站方案批复', '10月底'],
            ['2', '大部件长停复役', '在修15台全部并网', '11月底'],
            ['3', '新增装机并网', '0.3GW 全容量并网', '12月中'],
            ['4', '预测准确率提升', '月度准确率≥85%', '12月底'],
            ['5', 'AGC联控优化', '三座储能站参数下发', '11月中'],
            ['6', '无人值守投运', '剩余200MW投运', '12月底'],
            ['7', '备件国产化认证', '3类备件完成认证', '12月底'],
            ['8', '检修自主化扩面', '自主检修占比≥60%', '12月底'],
            ['9', '成本预算收口', '全年支出率≤100%', '12月底'],
            ['10', '次年计划编制', '计划文本上报', '11月底']]
    table(sl, M, 1.76, CW, [0.90, 4.20, 5.30, 1.51], rows, row_h=0.44, header_h=0.40,
          fs=F13, hfs=F13, aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT,
                                   PP_ALIGN.CENTER])


# ================= 26 总结：一页看懂（大数字块） =================
def p26(prs, num):
    sl = blank(prs)
    page_title(sl, '总结　一页看懂', num)
    H(sl, '三句话：电量差2.0亿、考核优5%、成本控在预算内。')

    tiles = [('30.00', '亿千瓦时', '上网电量', '完成率 93.8%', BLUE),
             ('2.00', '亿千瓦时', '电量缺口', '存量占七成', RED),
             ('12.0%', '', '限电率', '优于对标 1.0pp', ORANGE),
             ('0.0400', '元/千瓦时', '度电考核', '低于对标 5.0%', GREEN),
             ('90.0%', '', '无人值守', '1,800MW 已投运', SKY),
             ('10,000', '万元', '运维成本', '支出率 77%', BLUE)]
    for i, (num, unit, lab, note, col) in enumerate(tiles):
        cx = M + (i % 3) * 4.16
        cy = 1.80 + (i // 3) * 2.14
        rect(sl, cx, cy, 3.96, 1.92, fill=WHITE, line=LINE, lw=0.9)
        rect(sl, cx, cy, 0.06, 1.92, fill=col)
        txt(sl, cx + 0.26, cy + 0.20, 3.40, 0.30, lab, size=F13, color=GRAY2)
        txt(sl, cx + 0.26, cy + 0.56, 3.40, 0.62, num, size=34, color=col, bold=True)
        txt(sl, cx + 0.26, cy + 1.22, 3.40, 0.26, unit, size=11, color=GRAY3)
        txt(sl, cx + 0.26, cy + 1.50, 3.40, 0.30, note, size=F13, color=GRAY)


def p27(prs, num):
    sl = blank(prs)
    page_title(sl, '', num, rule=False, label=False)
    txt(sl, 0, 2.62, SW, 0.50, '汇报完毕　请指正', size=30, color=NAVY, bold=True,
        align=PP_ALIGN.CENTER)
    rect(sl, SW / 2 - 0.45, 3.28, 0.9, 0.030, fill=BLUE)
    txt(sl, 0, 3.52, SW, 0.30, '生产管理部　|　20XX年10月', size=13, color=GRAY2,
        align=PP_ALIGN.CENTER)


def main():
    prs = new_deck()
    p01(prs)
    p02(prs, 2)
    part(prs, 3, '01', '前三季度指标完成与偏差分析')
    p04(prs, 4); p05(prs, 5); p06(prs, 6); p07(prs, 7)
    p08(prs, 8); p09(prs, 9); p12(prs, 10); p16(prs, 11)
    p17(prs, 12); p19(prs, 13); p22(prs, 14); p24(prs, 15)
    p25(prs, 16); p26(prs, 17); p27(prs, 18)
    apply_chart_anim(prs)   # ANIM_CHARTS=False，等于空操作
    prs.save(OUT)
    print('saved:', OUT)


if __name__ == '__main__':
    main()
