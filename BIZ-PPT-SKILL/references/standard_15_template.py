# -*- coding: utf-8 -*-
"""简略版骨架示例 —— 15 页以内，白底、正文 14pt、结论句 20pt。

【脱敏说明】页型与函数写法完整保留，文案为占位词、数值为演示占位值。

简略版的唯一任务是「合并」：详细版 30 页压到 15 页，靠的是把同类证据并到一页，
而不是把字号调小。已验证的五条合并规则：
  1. 发电量 + 利用小时 + 限电 → 一页（左图：计划/实际；右表：小时数与限电）
  2. 并网考核 + 资本化 → 一页（两个大卡 + 结构条）
  3. 成本 + 降本专项 → 一页（成本结构表 + 降本进度条）
  4. 设备治理 + 无人值守 → 一页（风险卡 2×3）
  5. 全年目标 + 风险难点 → 一页（目标对比图 + 风险清单）
每页保留：一句话结论（20pt，标点完整）+ 一张图或一张表 + 3—5 条要点，不做第二个主轴。
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from deck_lib import *
from deck_lib import _fix_font
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'demo_brief.pptx')

F14, F13, F12 = 14, 13, 12


def H(sl, text):
    headline(sl, text, size=20)


def kpi_big(sl, x, y, w, h, label, value, note, accent=BLUE, vcolor=NAVY, vsize=28,
            ncolor=ORANGE):
    small = h < 1.55
    rect(sl, x, y, w, h, fill=LIGHT2, line=LINE, lw=0.75)
    rect(sl, x, y, w, 0.075, fill=accent)
    txt(sl, x + 0.18, y + (0.13 if small else 0.20), w - 0.28, 0.28, label,
        size=(11.5 if small else 12), color=GRAY2)
    vs = min(vsize, 22) if small else vsize
    tb = sl.shapes.add_textbox(Inches(x + 0.16), Inches(y + (0.40 if small else 0.52)),
                               Inches(w - 0.22), Inches(0.50))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = value
    r.font.size = Pt(vs); r.font.bold = True; r.font.color.rgb = rgb(vcolor)
    _fix_font(r)
    txt(sl, x + 0.18, y + h - (0.36 if small else 0.42), w - 0.26, 0.32, note,
        size=(11 if small else 12), color=ncolor, ls=1.12)


def body(sl, x, y, w, h, items, size=F14, gap=9):
    arr = []
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            lead, rest = it
            arr.append({'runs': [(lead, size, NAVY, True), (rest, size, GRAY, False)],
                        'ls': 1.30, 'before': 0 if i == 0 else gap})
        else:
            arr.append({'runs': [('· ', size, BLUE, True), (it, size, GRAY, False)],
                        'ls': 1.30, 'before': 0 if i == 0 else gap, 'hang': True,
                        'marL': 0.24, 'indent': -0.24})
    paras(sl, x, y, w, h, arr)


def bar_row(sl, x, y, w, label, value, ratio, color=BLUE, vsize=F14):
    txt(sl, x, y, 2.20, 0.30, label, size=vsize, color=GRAY)
    rect(sl, x + 2.30, y + 0.07, w - 3.60, 0.16, fill='E3EDF6')
    rect(sl, x + 2.30, y + 0.07, max((w - 3.60) * ratio, 0.03), 0.16, fill=color)
    txt(sl, x + w - 1.20, y, 1.20, 0.30, value, size=vsize, color=NAVY, bold=True,
        align=PP_ALIGN.RIGHT)


def note_panel(sl, x, y, w, h, text, accent=BLUE, size=F14):
    panel(sl, x, y, w, h, fill=LIGHT2)
    rect(sl, x, y, 0.05, h, fill=accent)
    txt(sl, x + 0.26, y + 0.13, w - 0.52, h - 0.26, text, size=size, color=GRAY, ls=1.24)


# ================= 02 总览：一页讲完 =================
def p02(prs, num):
    sl = blank(prs)
    page_title(sl, '一  前三季度总体：四个数', num)
    H(sl, '电量欠发2.00亿千瓦时，考核优于对标、成本控在预算内。')

    y, h, w, g = 1.72, 1.44, 3.015, 0.145
    kpi_big(sl, M, y, w, h, '上网电量（亿千瓦时）', '30.00', '完成率93.8%，欠发2.00亿')
    kpi_big(sl, M + (w + g), y, w, h, '限电率（%）', '12.0', '同比+5pp，优于对标1pp')
    kpi_big(sl, M + 2 * (w + g), y, w, h, '度电考核（元/千瓦时）', '0.0400', '低于对标5.0%',
            GREEN, GREEN, 28, ncolor=GREEN)
    kpi_big(sl, M + 3 * (w + g), y, w, h, '运维成本（万元）', '10,000', '支出率77%',
            ORANGE, ORANGE, 28)

    chip(sl, M, 3.40, '三件事讲完这季度', size=15)
    rows = [['维度', '结论', '支撑'],
            ['电量', '欠发2.00亿，存量占七成', '限电与风资源各占一半'],
            ['考核', '度电考核低于对标5.0%', '预测偏差占结构七成'],
            ['成本', '支出率77%，同比+3.5%', '材料费涨幅最高']]
    table(sl, M, 3.84, CW, [1.80, 5.10, 5.59], rows, row_h=0.72, header_h=0.42,
          fs=F14, hfs=F13, bold_first_col=True,
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT])


# ================= 03 电量 + 小时数 + 限电（合并页） =================
def p03(prs, num):
    sl = blank(prs)
    page_title(sl, '二  电量、利用小时与限电（合并）', num)
    H(sl, '电量完成93.8%，利用小时同比低80小时，限电率升至12.0%。')

    band(sl, M, 1.72, 7.30, 0.34, '分月上网电量（亿千瓦时）', fill=BLUE, size=13)
    schart(sl, M - 0.05, 2.14, 7.40, 3.20,
           ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'],
           [('计划', [3.4, 3.2, 4.0, 3.9, 3.8, 3.6, 3.5, 3.5, 3.1]),
            ('实际', [3.1, 3.0, 3.8, 3.7, 3.4, 3.2, 3.0, 3.3, 3.5])],
           kind='column', colors=[PALE, BLUE], gap=60, fmt='0.0', label_size=11,
           legend=True, vsize=10, show_vaxis=False, vmax=5.0, label_color=NAVY, cat_size=F12)

    x = 8.00
    chip(sl, x, 1.72, '三个关键读数', size=15)
    body(sl, x, 2.16, 4.91, 1.60, [
        ('利用小时：', '1,200小时，同比-80小时。'),
        ('限电率：', '12.0%，同比+5pp。'),
        ('限电量：', '5.80亿千瓦时，同比翻倍。'),
    ], size=F14, gap=8)

    chip(sl, x, 3.92, '限电与欠发的关系', size=15)
    rows = [['因素', '影响电量', '占比'],
            ['限电上升', '3.00', '52%'],
            ['风资源下降', '2.25', '39%'],
            ['大部件损坏', '0.30', '5%'],
            ['达产与弃电', '0.25', '4%']]
    table(sl, x, 4.34, 4.91, [2.11, 1.60, 1.20], rows, row_h=0.40, header_h=0.36,
          fs=F13, hfs=F13, bold_first_col=True,
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER])
    note_panel(sl, M, 5.62, 7.30, 0.86,
               '结论：电量缺口的一半来自限电、四成来自风资源，两者都难在四季度补回，'
               '计划应转向"少欠"而不是"补平"。', accent=RED, size=F14)


# ================= 04 考核 + 资本化（合并页） =================
def p04(prs, num):
    sl = blank(prs)
    page_title(sl, '三  并网考核与资本化（合并）', num)
    H(sl, '度电考核0.0400元、低于对标5.0%；资本化支出完成80%。')

    y = 1.72
    kpi_big(sl, M, y, 5.86, 1.20, '总考核电量（亿千瓦时）', '3.40', '占上网电量 11.3%',
            BLUE, NAVY, 36, ncolor=GRAY2)
    kpi_big(sl, M + 6.04, y, 5.86, 1.20, '资本化支出（万元）', '4,000', '预算 5,000，完成 80%',
            AMBER, NAVY, 36, ncolor=GRAY2)

    chip(sl, M, 3.16, '考核结构（占月度考核电量比重）', size=15)
    bar_row(sl, M, 3.62, 6.30, '功率预测偏差', '52%—78%', 0.78, BLUE)
    bar_row(sl, M, 4.18, 6.30, '发电计划曲线', '12%—46%', 0.46, BLUE)
    bar_row(sl, M, 4.74, 6.30, '功率预测（AGC）', '8%—19%', 0.19, SKY)

    chip(sl, 7.20, 3.16, '资本化项目进度', size=15)
    rows = [['项目', '预算', '完成', '进度'],
            ['技改类', '2,400', '2,050', '85%'],
            ['数字化', '1,400', '1,000', '71%'],
            ['安全类', '1,200', '950', '79%']]
    table(sl, 7.20, 3.60, 5.71, [1.80, 1.30, 1.30, 1.31], rows, row_h=0.46, header_h=0.38,
          fs=F13, hfs=F13, bold_first_col=True,
          aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 3)
    note_panel(sl, M, 5.62, 6.30, 0.86,
               '口径：总考核电量与度电考核均为累计值，度电考核＝考核电量÷上网电量。',
               accent=BLUE, size=F13)
    note_panel(sl, 7.20, 5.62, 5.71, 0.86,
               '数字化项目进度最慢，主因是通信改造与主机到货。', accent=AMBER, size=F13)


# ================= 05 储能运行 =================
def p05(prs, num):
    sl = blank(prs)
    page_title(sl, '四  储能：运行表现', num)
    H(sl, '在运300MW／600MWh，综合效率90.8%，k3均值1.51。')

    y, h, w, g = 1.72, 1.30, 3.015, 0.145
    kpi_big(sl, M, y, w, h, '在运规模（MW）', '300', '三座站合计', BLUE, NAVY, 30, ncolor=GRAY2)
    kpi_big(sl, M + (w + g), y, w, h, '综合效率（%）', '90.8', '目标 90.0', GREEN, GREEN, 30,
            ncolor=GREEN)
    kpi_big(sl, M + 2 * (w + g), y, w, h, 'k3 均值', '1.51', '目标 1.40', GREEN, GREEN, 30,
            ncolor=GREEN)
    kpi_big(sl, M + 3 * (w + g), y, w, h, '充放电量（MWh）', '45,000', '1—9月累计', BLUE, NAVY, 30,
            ncolor=GRAY2)

    chip(sl, M, 3.26, '分站运行指标', size=15)
    rows = [['场站', '规模', 'k3', '综合效率', '可用系数', '评价'],
            ['储能站A', '100MW', '1.58', '91.2%', '96.0%', '优'],
            ['储能站B', '100MW', '1.49', '90.5%', '94.5%', '良'],
            ['储能站C', '100MW', '1.46', '90.7%', '93.0%', '良']]
    table(sl, M, 3.72, 8.20, [1.90, 1.20, 1.10, 1.40, 1.40, 1.20], rows, row_h=0.52,
          header_h=0.40, fs=F14, hfs=F13, bold_first_col=True,
          aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 5,
          cell_color={(1, 5): GREEN, (2, 5): NAVY, (3, 5): NAVY})

    x = 8.80
    chip(sl, x, 3.26, '两条判断', size=15)
    body(sl, x, 3.70, 4.11, 1.60, [
        ('效率：', '三站均达目标，B站偏低主因是温控。'),
        ('策略：', '高价时段下网电费上升，报价与充放不匹配。'),
    ], size=F14, gap=12)
    note_panel(sl, M, 5.86, CW, 0.78,
               '口径：k3＝统计期放电量÷充电量；综合效率含辅助用电；可用系数＝可用小时÷统计小时。',
               size=F14)


# ================= 06 成本 + 降本（合并页） =================
def p06(prs, num):
    sl = blank(prs)
    page_title(sl, '五  运维成本与降本专项（合并）', num)
    H(sl, '成本10,000万元、支出率77%，八项措施累计降本1,200万元。')

    band(sl, M, 1.72, 6.60, 0.34, '成本构成（万元）', fill=BLUE, size=13)
    rows = [['成本项', '本年', '上年', '同比'],
            ['材料费', '3,200', '3,000', '+6.7%'],
            ['修理费', '2,800', '2,750', '+1.8%'],
            ['人工费', '2,200', '2,150', '+2.3%'],
            ['其他费用', '1,800', '1,760', '+2.3%'],
            ['合计', '10,000', '9,660', '+3.5%']]
    table(sl, M, 2.14, 6.60, [2.00, 1.55, 1.55, 1.50], rows, row_h=0.50, header_h=0.40,
          fs=F14, hfs=F13, bold_first_col=True,
          aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 3,
          cell_color={(5, 0): NAVY, (5, 1): NAVY, (5, 2): NAVY, (5, 3): ORANGE})

    x = 7.30
    chip(sl, x, 1.72, '降本八项完成率', size=15)
    data = [('备件联合储备', 0.91), ('检修自主化', 0.87), ('以大代小置换', 0.72),
            ('功率预测优化', 0.83), ('无人值守减员', 0.60), ('备件国产化', 0.50)]
    for i, (lab, r) in enumerate(data):
        yy = 2.16 + i * 0.52
        txt(sl, x, yy, 1.95, 0.28, lab, size=F13, color=GRAY)
        rect(sl, x + 2.00, yy + 0.06, 2.40, 0.16, fill='E3EDF6')
        rect(sl, x + 2.00, yy + 0.06, 2.40 * r, 0.16, fill=BLUE if r >= 0.75 else SKY)
        txt(sl, x + 4.45, yy, 0.80, 0.28, f'{r*100:.0f}%', size=F13, color=NAVY,
            bold=True, align=PP_ALIGN.RIGHT)

    note_panel(sl, M, 5.58, CW, 0.90,
               '判断：成本总量可控，压力在材料费（备件）。降本目标1,500万元已完成80%，'
               '缺口主要来自备件国产化认证周期。', size=F14)


# ================= 07 设备 + 无人值守（合并页） =================
def p07(prs, num):
    sl = blank(prs)
    page_title(sl, '六  设备治理与无人值守（合并）', num)
    H(sl, '集电线路跳闸20次、同比减15次；无人值守完成90.0%。')

    chip(sl, M, 1.74, '四类设备风险', size=15)
    cards = [('01', '大部件损坏', '叶片、齿轮箱长停，单次14天。', RED),
             ('02', '集电线路跳闸', '20次，同比减15次。', ORANGE),
             ('03', '变桨系统', '故障率0.8次/台年。', ORANGE),
             ('04', '功率预测偏差', '准确率82%，低对标3pp。', BLUE)]
    for i, (no, title, desc, col) in enumerate(cards):
        cx = M + i * 3.14
        rect(sl, cx, 2.18, 2.94, 1.66, fill=LIGHT2, line=LINE, lw=0.75)
        rect(sl, cx, 2.18, 2.94, 0.07, fill=col)
        txt(sl, cx + 0.18, 2.40, 0.60, 0.36, no, size=18, color=col, bold=True)
        txt(sl, cx + 0.18, 2.82, 2.60, 0.32, title, size=F14, color=NAVY, bold=True)
        txt(sl, cx + 0.18, 3.18, 2.60, 0.56, desc, size=F12, color=GRAY, ls=1.24)

    chip(sl, M, 4.06, '无人值守进度（MW）', size=15)
    for i, (lab, val, r, col) in enumerate([('已投运', '1,800', 0.90, GREEN),
                                            ('在建', '160', 0.08, SKY),
                                            ('未启动', '40', 0.02, PALE)]):
        yy = 4.52 + i * 0.52
        txt(sl, M, yy, 1.40, 0.28, lab, size=F14, color=GRAY)
        rect(sl, M + 1.50, yy + 0.06, 4.20, 0.18, fill='E3EDF6')
        rect(sl, M + 1.50, yy + 0.06, 4.20 * r, 0.18, fill=col)
        txt(sl, M + 5.85, yy, 1.00, 0.28, val, size=F14, color=NAVY, bold=True,
            align=PP_ALIGN.RIGHT)

    x = 7.75
    chip(sl, x, 4.06, '两个待办', size=15)
    body(sl, x, 4.50, 5.16, 1.40, [
        ('通信改造：', '未投运200MW卡在通信通道。'),
        ('主机扩容：', '三期平台需扩容后方可接入。'),
    ], size=F14, gap=10)
    note_panel(sl, M, 6.02, CW, 0.62,
               '口径：无人值守完成率＝已投运容量÷可改造容量；在建含已签合同但未并网项目。',
               size=F13)


# ================= 08 全年目标 + 风险 =================
def p08(prs, num):
    sl = blank(prs)
    page_title(sl, '七  全年目标与风险难点', num)
    H(sl, '全年预计41.5亿千瓦时、完成98.8%，缺口0.5亿可控。')

    band(sl, M, 1.72, 7.30, 0.34, '全年电量：目标 / 预计（亿千瓦时）', fill=BLUE, size=13)
    schart(sl, M - 0.05, 2.14, 7.40, 3.40,
           ['一季度', '二季度', '三季度', '四季度E'],
           [('目标', [10.0, 11.5, 10.5, 10.0]), ('预计', [9.8, 11.0, 9.2, 11.5])],
           kind='column', colors=[PALE, BLUE], gap=90, fmt='0.0', label_size=12,
           legend=True, vsize=10, show_vaxis=False, vmax=14, label_color=NAVY, cat_size=F13)
    txt(sl, M, 5.66, 7.30, 0.30, '四季度预计11.5亿千瓦时，须与计划持平才能补平缺口。',
        size=F14, color=RED, bold=True)

    x = 8.00
    chip(sl, x, 1.72, '四个风险', size=15)
    body(sl, x, 2.16, 4.91, 3.40, [
        ('风资源：', '若再降5%，缺口扩大至1.5亿。'),
        ('限电：', '断面约束未解，四季度仍高压。'),
        ('新投：', '0.3GW并网时点存在滞后风险。'),
        ('考核：', '新政口径调整可能抬高考核。'),
    ], size=F14, gap=13)


# ================= 09 收尾 =================
def p09(prs, num):
    sl = blank(prs)
    page_title(sl, '', num, rule=False, label=False)
    txt(sl, 0, 2.62, SW, 0.50, '汇报完毕　请指正', size=30, color=NAVY, bold=True,
        align=PP_ALIGN.CENTER)
    rect(sl, SW / 2 - 0.45, 3.28, 0.9, 0.030, fill=BLUE)
    txt(sl, 0, 3.52, SW, 0.30, '生产管理部　|　20XX年10月', size=13, color=GRAY2,
        align=PP_ALIGN.CENTER)


def main():
    prs = new_deck()
    p02(prs, 1); p03(prs, 2); p04(prs, 3); p05(prs, 4)
    p06(prs, 5); p07(prs, 6); p08(prs, 7); p09(prs, 8)
    apply_chart_anim(prs)   # ANIM_CHARTS=False，等于空操作
    prs.save(OUT)
    print('saved:', OUT)


if __name__ == '__main__':
    main()
