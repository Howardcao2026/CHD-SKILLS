# -*- coding: utf-8 -*-
"""季度汇报 PPT 构建库：统一版式、色板、图表与表格封装。"""
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_MARK,
                                XL_TICK_LABEL_POSITION)
from pptx.oxml.ns import qn
from lxml import etree
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

# ---------- 色板 ----------
# 企业标准色卡：把公司 VI 手册里的主色/辅助色填进来即可，此处为示例值（蓝系 + 绿 + 橙）
NAVY   = '004A86'   # 主色（深蓝）—— 结论句、表头、强调条
BLUE   = '006AAF'   # 主数据色 —— 柱、条、折线
BLUE_D = '003D6E'   # 主色加深，仅用于深底反白
SKY    = '00A2E9'   # 次数据色 —— 第二序列、辅助条
PALE   = '9FD5F2'   # 次数据浅化 —— 计划值、背景柱
LIGHT  = 'EAF3FA'
LIGHT2 = 'F6FAFD'
LINE   = 'D5E3F0'
ORANGE = 'ED6C00'   # 负向色 —— 注脚/警示（色卡若无正红，用它替代红）
AMBER  = 'F6AD02'   # 中间态 —— 色条、高亮（太浅，不要用于文字）
GREEN  = '03994C'   # 正向色 —— 达标、正向差值
RED    = 'ED6C00'   # 未达标（示例色卡无正红，与 ORANGE 同值；有正红时改为此处）
GRAY   = '44515E'
GRAY2  = '76879A'
GRAY3  = '9AA9B7'
WHITE  = 'FFFFFF'

YH = '微软雅黑'

# 动画总开关：默认 False = 成品不带任何动画（改 True 可恢复图表入场，但不建议）
ANIM_CHARTS = False

SW, SH = 13.333, 7.5
M = 0.42
CW = SW - 2 * M

TITLE_Y = 0.95     # 章节小标（母版 logo 与横线占 0—0.85in，不得侵入）
RULE_Y = 1.30      # 保留引用，不再绘制（母版自带横线）
HEAD_Y = 1.20      # 一句话结论，20pt
BODY_T = 1.66
BODY_B = 6.72      # 底部让出母版页码与 slogan
FOOT_Y = 6.92      # 仅左侧栏目名，宽度受限，避开右下角

# 页脚左侧落款：换成自家公司名 + 部门
COMPANY_LINE = '示例公司 · 生产管理部'


def rgb(h):
    return RGBColor.from_string(h)


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    return prs


def blank(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl._anim_groups = []
    return sl


def _fix_font(run, name=YH):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', name)


def _spacing(p, ls=None, before=0, after=0):
    if ls is not None:
        p.line_spacing = ls
    p.space_before = Pt(before)
    p.space_after = Pt(after)


def _hang(p, marL=0.16, indent=-0.16):
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(int(Inches(marL))))
    pPr.set('indent', str(int(Inches(indent))))


def _strip_style(sh):
    el = sh._element
    st = el.find(qn('p:style'))
    if st is not None:
        el.remove(st)


def rect(sl, x, y, w, h, fill=None, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, adj=None):
    sh = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    _strip_style(sh)
    if adj is not None:
        try:
            sh.adjustments[0] = adj
        except Exception:
            pass
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def txt(sl, x, y, w, h, text, size=10, color=GRAY, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, ls=1.0, font=YH, alpha=None, italic=False):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    tf.auto_size = None
    lines = text.split('\n')
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        _spacing(p, ls)
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = rgb(color)
        _fix_font(r, font)
    return tb


def paras(sl, x, y, w, h, items, anchor=MSO_ANCHOR.TOP, font=YH):
    """items: list of dict(text|runs, size, color, bold, align, ls, before, after, hang, bullet)"""
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    tf.auto_size = None
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = it.get('align', PP_ALIGN.LEFT)
        _spacing(p, it.get('ls', 1.0), it.get('before', 0), it.get('after', 0))
        if it.get('hang'):
            _hang(p, it.get('marL', 0.16), it.get('indent', -0.16))
        runs = it.get('runs')
        if runs is None:
            runs = [(it.get('text', ''), it.get('size', 10), it.get('color', GRAY), it.get('bold', False))]
        for t, s, c, b in runs:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(s)
            r.font.bold = b
            r.font.color.rgb = rgb(c)
            _fix_font(r, font)
    return tb


def bullets(sl, x, y, w, h, items, size=9, color=GRAY, gap=6, bullet='•', bcolor=None,
            ls=1.15, lead_bold=True):
    """items: list of (lead, rest) or str"""
    out = []
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            lead, rest = it
            runs = [(bullet + ' ', size, bcolor or BLUE, True)]
            if lead:
                runs.append((lead, size, NAVY, True))
            runs.append((rest, size, color, False))
        else:
            runs = [(bullet + ' ', size, bcolor or BLUE, True), (it, size, color, False)]
        out.append({'runs': runs, 'hang': True, 'marL': 0.15, 'indent': -0.15,
                    'ls': ls, 'before': 0 if i == 0 else gap})
    return paras(sl, x, y, w, h, out)


# ---------- 页面骨架 ----------
def page_title(sl, text, num=None, rule=False, label=True):
    """顶部只放一行小灰标；横线与页码由公司母版提供，此处不画。"""
    if text:
        txt(sl, M, TITLE_Y, CW, 0.26, text, size=12, color=GRAY2, bold=True)
    if num is not None and label:
        txt(sl, M, FOOT_Y, 6.20, 0.22, COMPANY_LINE,
            size=8, color=GRAY3)


def headline(sl, text, y=HEAD_Y, size=17, w=CW, color=NAVY):
    txt(sl, M, y, w, 0.40, text, size=size, color=color, bold=True)


def chip(sl, x, y, text, size=10.5, color=NAVY, bar=BLUE, w=None, bar_h=0.22):
    """▎小标题"""
    rect(sl, x, y + 0.02, 0.055, bar_h, fill=bar)
    width = min(w or 6.0, SW - x - M - 0.13)
    txt(sl, x + 0.13, y - 0.02, width, 0.32, text, size=size, color=color, bold=True)


def band(sl, x, y, w, h, text, fill=BLUE, size=9.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER):
    """深色标题带（图表面板标题）"""
    rect(sl, x, y, w, h, fill=fill)
    txt(sl, x, y + (h - 0.20) / 2 - 0.02, w, 0.24, text, size=size, color=color, bold=bold, align=align)


def kpi(sl, x, y, w, h, label, value, note, accent=BLUE, vcolor=NAVY, vsize=24,
        unit='', notesize=8.5, ncolor=ORANGE, fill=LIGHT2):
    rect(sl, x, y, w, h, fill=fill, line=LINE, lw=0.75)
    rect(sl, x, y, w, 0.07, fill=accent)
    small = h < 1.20
    txt(sl, x + 0.14, y + (0.11 if small else 0.16), w - 0.22, 0.22, label,
        size=(8.3 if small else 8.8), color=GRAY2)
    vy = y + (0.30 if small else 0.42)
    tb = sl.shapes.add_textbox(Inches(x + 0.13), Inches(vy), Inches(w - 0.18), Inches(0.50))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.auto_size = None
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = value
    r.font.size = Pt(min(vsize, 18) if small else vsize); r.font.bold = True
    r.font.color.rgb = rgb(vcolor)
    _fix_font(r)
    if unit:
        r2 = p.add_run(); r2.text = unit
        r2.font.size = Pt(8.5); r2.font.bold = False; r2.font.color.rgb = rgb(GRAY2)
        _fix_font(r2)
    txt(sl, x + 0.14, y + h - (0.30 if small else 0.38), w - 0.22, 0.26, note,
        size=(7.6 if small else notesize), color=ncolor)


def stat(sl, x, y, w, h, value, label, vcolor=NAVY, vsize=19, lsize=8.2, align=PP_ALIGN.LEFT):
    txt(sl, x, y, w, 0.34, value, size=vsize, color=vcolor, bold=True, align=align)
    txt(sl, x, y + 0.36, w, 0.44, label, size=lsize, color=GRAY2, ls=1.05, align=align)


def panel(sl, x, y, w, h, fill=LIGHT2, line=LINE):
    rect(sl, x, y, w, h, fill=fill, line=line, lw=0.75)


# ---------- 表格 ----------
NO_STYLE = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'


def _cell_border(cell, color=LINE, w=0.75, sides=('L', 'R', 'T', 'B')):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
    idx = 0
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        if tag[2:] not in sides:
            idx += 1
            continue
        ln = etree.SubElement(tcPr, qn(tag))
        tcPr.remove(ln)
        tcPr.insert(idx, ln)
        idx += 1
        ln.set('w', str(int(w * 12700)))
        ln.set('cap', 'flat'); ln.set('cmpd', 'sng'); ln.set('algn', 'ctr')
        sf = etree.SubElement(ln, qn('a:solidFill'))
        c = etree.SubElement(sf, qn('a:srgbClr')); c.set('val', color)
        d = etree.SubElement(ln, qn('a:prstDash')); d.set('val', 'solid')


def table(sl, x, y, w, col_w, rows, row_h=0.28, header_h=0.32, fs=8.5, hfs=8.5,
          aligns=None, header_fill=BLUE, zebra=True, bold_first_col=False,
          cell_fill=None, cell_color=None, valign=MSO_ANCHOR.MIDDLE, pad=0.06,
          border=LINE, header_color=WHITE, row_hs=None):
    nr, nc = len(rows), len(rows[0])
    total_h = header_h + row_h * (nr - 1)
    gf = sl.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(total_h))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tid = tblPr.find(qn('a:tableStyleId'))
    if tid is None:
        tid = etree.SubElement(tblPr, qn('a:tableStyleId'))
    tid.text = NO_STYLE
    tbl.first_row = False
    tbl.horz_banding = False
    tbl.first_col = False
    tbl.last_row = False
    tbl.last_col = False

    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    tbl.rows[0].height = Inches(header_h)
    for r in range(1, nr):
        tbl.rows[r].height = Inches(row_hs[r - 1] if row_hs else row_h)

    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.margin_left = Inches(pad)
            cell.margin_right = Inches(pad)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = valign
            # fill
            f = None
            if r == 0:
                f = header_fill
            elif cell_fill and (r, c) in cell_fill:
                f = cell_fill[(r, c)]
            elif zebra and r % 2 == 0:
                f = LIGHT
            else:
                f = WHITE
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(f)
            _cell_border(cell, color=border, w=0.75)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = (aligns[c] if aligns else PP_ALIGN.CENTER)
            _spacing(p, 0.95)
            txtv = '' if val is None else str(val)
            parts = txtv.split('\n')
            for k, part in enumerate(parts):
                if k > 0:
                    p = tf.add_paragraph()
                    p.alignment = (aligns[c] if aligns else PP_ALIGN.CENTER)
                    _spacing(p, 0.95)
                r_ = p.add_run(); r_.text = part
                r_.font.size = Pt(hfs if r == 0 else fs)
                r_.font.bold = (r == 0) or (bold_first_col and c == 0)
                if r == 0:
                    r_.font.color.rgb = rgb(header_color)
                elif cell_color and (r, c) in cell_color:
                    r_.font.color.rgb = rgb(cell_color[(r, c)])
                else:
                    r_.font.color.rgb = rgb(GRAY if not (bold_first_col and c == 0) else NAVY)
                _fix_font(r_)
    return gf


# ---------- 图表 ----------
def _clean_frame(gf):
    cs = gf.chart._chartSpace
    ch = cs.find(qn('c:chart'))
    sp = etree.SubElement(cs, qn('c:spPr'))
    cs.remove(sp)
    ch.addnext(sp)
    sf = etree.SubElement(sp, qn('a:noFill'))
    ln = etree.SubElement(sp, qn('a:ln'))
    etree.SubElement(ln, qn('a:noFill'))


def _axis_style(ax, size=8, color=GRAY2, line=LINE, gridlines=False, visible=True, fmt=None):
    ax.has_major_gridlines = gridlines
    if gridlines:
        ax.major_gridlines.format.line.color.rgb = rgb('E8EFF6')
        ax.major_gridlines.format.line.width = Pt(0.6)
    ax.has_minor_gridlines = False
    ax.format.line.color.rgb = rgb(line)
    ax.format.line.width = Pt(0.75)
    ax.major_tick_mark = XL_TICK_MARK.NONE
    ax.minor_tick_mark = XL_TICK_MARK.NONE
    ax.tick_labels.font.size = Pt(size)
    ax.tick_labels.font.color.rgb = rgb(color)
    ax.tick_labels.font.name = YH
    if not visible:
        ax.visible = False


def chart(sl, x, y, w, h, cats, series, kind='column', colors=None, horizontal=False,
          stacked=False, gap=55, overlap=-5, labels=True, fmt='0.0', label_pos='outside',
          legend=False, leg_pos=XL_LEGEND_POSITION.BOTTOM, vsize=8, csize=8,
          vmax=None, vmin=None, show_vaxis=False, gridlines=False, point_colors=None,
          label_color=GRAY, label_size=8, cat_size=None, num_fmt_axis=None, marker=True):
    cd = CategoryChartData()
    cd.categories = cats
    for nm, vals in series:
        cd.add_series(nm, vals)

    if kind == 'column':
        ct = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED
    elif kind == 'bar':
        ct = XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
    elif kind == 'line':
        ct = XL_CHART_TYPE.LINE
    elif kind == 'line_marker':
        ct = XL_CHART_TYPE.LINE_MARKERS
    elif kind == 'doughnut':
        ct = XL_CHART_TYPE.DOUGHNUT
    elif kind == 'pie':
        ct = XL_CHART_TYPE.PIE
    else:
        ct = XL_CHART_TYPE.COLUMN_CLUSTERED

    gf = sl.shapes.add_chart(ct, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    chf = gf.chart
    chf.has_title = False
    _clean_frame(gf)

    plot = chf.plots[0]
    plot.gap_width = gap
    try:
        plot.overlap = overlap
    except Exception:
        pass
    if kind in ('line', 'line_marker'):
        plot.has_data_labels = labels
        for i, s in enumerate(plot.series):
            s.smooth = False
            s.format.line.width = Pt(2.0)
            s.format.line.color.rgb = rgb(colors[i % len(colors)])
            if kind == 'line_marker':
                s.marker.style = 8
                s.marker.size = 5
                s.marker.format.fill.solid()
                s.marker.format.fill.fore_color.rgb = rgb(colors[i % len(colors)])
                s.marker.format.line.color.rgb = rgb(colors[i % len(colors)])
            else:
                s.marker.style = 0
    else:
        plot.has_data_labels = labels
        for i, s in enumerate(plot.series):
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = rgb(colors[i % len(colors)])
            s.format.line.fill.background()
            try:
                s.invert_if_negative = False
            except Exception:
                pass

    if point_colors:
        for (si, pi, col) in point_colors:
            pt = plot.series[si].points[pi]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(col)
            pt.format.line.fill.background()
            dpt = pt._element
            idx_el = dpt.find(qn('c:idx'))
            old = dpt.find(qn('c:invertIfNegative'))
            if old is not None:
                dpt.remove(old)
            inv = etree.Element(qn('c:invertIfNegative'))
            inv.set('val', '0')
            if idx_el is not None:
                idx_el.addnext(inv)
            else:
                dpt.insert(0, inv)

    if labels:
        dl = plot.data_labels
        dl.font.size = Pt(label_size)
        dl.font.bold = False
        dl.font.color.rgb = rgb(label_color)
        dl.font.name = YH
        try:
            dl.number_format = fmt
            dl.number_format_is_linked = False
        except Exception:
            pass
        try:
            dl.position = {'outside': XL_LABEL_POSITION.OUTSIDE_END,
                           'inside': XL_LABEL_POSITION.INSIDE_END,
                           'center': XL_LABEL_POSITION.CENTER,
                           'above': XL_LABEL_POSITION.ABOVE,
                           'below': XL_LABEL_POSITION.BELOW,
                           'best': XL_LABEL_POSITION.BEST_FIT}[label_pos]
        except Exception:
            pass

    chf.has_legend = legend
    if legend:
        chf.legend.position = leg_pos
        chf.legend.include_in_layout = False
        chf.legend.font.size = Pt(8)
        chf.legend.font.color.rgb = rgb(GRAY2)
        chf.legend.font.name = YH

    if kind not in ('doughnut', 'pie'):
        ca = chf.category_axis
        _axis_style(ca, size=cat_size or csize, visible=True)
        if horizontal:
            ca.tick_labels.font.size = Pt(cat_size or csize)
            try:
                ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
            except Exception:
                pass
        va = chf.value_axis
        _axis_style(va, size=vsize, visible=show_vaxis, gridlines=gridlines)
        if vmax is not None:
            va.maximum_scale = vmax
        if vmin is not None:
            va.minimum_scale = vmin
        if num_fmt_axis:
            va.tick_labels.number_format = num_fmt_axis
            va.tick_labels.number_format_is_linked = False
    return gf


def progress(sl, x, y, w, h, ratio, fill=BLUE, bg='E3EDF6', label=None, lcolor=None):
    rect(sl, x, y, w, h, fill=bg)
    rect(sl, x, y, max(w * ratio, 0.02), h, fill=fill)
    if label:
        txt(sl, x + w + 0.10, y - 0.04, 0.9, 0.26, label, size=8.5, color=lcolor or GRAY, bold=True)


def turbine(sl, hx, hy, hub_h, blade, color=PALE, lw=1.6, rot=0):
    rect(sl, hx - 0.015, hy, 0.030, hub_h, fill=color)
    rect(sl, hx - 0.048, hy - 0.030, 0.096, 0.060, fill=color, shape=MSO_SHAPE.OVAL)
    for k in range(3):
        a = math.radians(rot + 90 + k * 120)
        ex, ey = hx + blade * math.cos(a), hy - 0.01 - blade * math.sin(a)
        ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(hx), Inches(hy - 0.01), Inches(ex), Inches(ey))
        _strip_style(ln)
        ln.line.color.rgb = rgb(color)
        ln.line.width = Pt(lw)
        ln.shadow.inherit = False


def arrow_right(sl, x, y, w, color=BLUE, lw=1.25):
    ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(lw)
    head = rect(sl, x + w - 0.10, y - 0.045, 0.10, 0.09, fill=color, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
    head.rotation = 90


# ================= 图表入场动画 =================
def _entrance_xml(specs, dur):
    blocks = []
    cid = 3
    for shape, flt, delay in specs:
        spid = shape.shape_id
        preset = 22 if flt.startswith('wipe') else 10
        blocks.append(
            f'<p:par><p:cTn id="{cid}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst><p:par><p:cTn id="{cid+1}" fill="hold">'
            f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{cid+2}" presetID="{preset}" presetClass="entr" presetSubtype="0" '
            f'fill="hold" grpId="0" nodeType="afterEffect">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{cid+3}" dur="1" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr>'
            f'<p:cTn id="{cid+4}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'</p:cBhvr></p:animEffect>'
            f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par>'
        )
        cid += 5
    return (
        '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        + ''.join(blocks) +
        '</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
    )


def apply_chart_anim(prs, dur=450):
    """每页图表形状加简单入场：柱子擦除、线条圆点淡入，全部同时开始。
    默认关闭（ANIM_CHARTS=False），成品不带任何动画；保留代码仅供实验。"""
    if not ANIM_CHARTS:
        return
    from pptx.oxml import parse_xml
    for sl in prs.slides:
        groups = getattr(sl, '_anim_groups', None)
        if not groups:
            continue
        specs, seen = [], set()
        for g in groups:
            for sh in g:
                if sh.shape_id in seen:
                    continue
                seen.add(sh.shape_id)
                if sh.has_text_frame and sh.text_frame.text.strip():
                    continue
                if (sh.height or 0) < 27432 or (sh.width or 0) < 27432:
                    continue
                flt = 'wipe(up)' if sh.shape_type == MSO_SHAPE.RECTANGLE else 'fade'
                specs.append((sh, flt, 0))
        if not specs:
            continue
        try:
            sl._element.append(parse_xml(_entrance_xml(specs, dur)))
        except Exception as e:
            print('anim skip:', e)


# ================= 矢量图表（形状绘制） =================
# 说明：不生成 PPT 原生图表对象，全部用形状绘制，任何阅读器/转换器都不会丢图。
def _wlen(s):
    return sum(1.0 if ord(ch) > 127 else 0.56 for ch in str(s))


def _fm(v, fmt='0.0'):
    if v is None:
        return ''
    if fmt.startswith('"'):
        dec = len(fmt.split('.')[-1]) if '.' in fmt else 0
        return ('-' if v >= 0 else '') + f'{abs(v):.{dec}f}'
    if fmt == '0':
        return f'{v:.0f}'
    if fmt == '#,##0':
        return f'{v:,.0f}'
    if fmt.startswith('0.'):
        dec = len(fmt.split('.')[1])
        return f'{v:,.{dec}f}'
    return f'{v:g}'


def _catw(cats, size, extra=0.14):
    return max(_wlen(c) for c in cats) * size / 72.0 * 0.97 + extra


def schart(sl, x, y, w, h, cats, series, kind='column', colors=None, horizontal=False,
           stacked=False, gap=55, overlap=-5, labels=True, fmt='0.0', label_pos='outside',
           legend=False, leg_pos=None, vsize=8, csize=8, vmax=None, vmin=None,
           show_vaxis=False, gridlines=False, point_colors=None, label_color=GRAY,
           label_size=8, cat_size=None, num_fmt_axis=None, marker=True, **kw):
    n0 = len(sl.shapes)
    colors = colors or [BLUE]
    cat_size = cat_size or csize
    vals_all = [v for nm, vs in series for v in vs if v is not None]
    vhi = vmax if vmax is not None else max(vals_all)
    vlo = vmin if vmin is not None else min(0, min(vals_all))
    pc = {}
    if point_colors:
        for (si, pi, col) in point_colors:
            pc[(si, pi)] = col
    if kind in ('bar',):
        _hbar(sl, x, y, w, h, cats, series, colors, gap, fmt, label_size, label_color,
              cat_size, vhi, vlo, pc, legend, vsize)
    elif kind in ('line', 'line_marker'):
        _vline(sl, x, y, w, h, cats, series, colors, loc_fmt(fmt), label_size, label_color,
               cat_size, vhi, vlo, show_vaxis, gridlines, vsize, legend, kind == 'line_marker')
    else:
        _vcol(sl, x, y, w, h, cats, series, colors, gap, fmt, label_size, label_color,
              cat_size, vhi, vlo, show_vaxis, gridlines, vsize, legend, pc)
    try:
        sl._anim_groups.append([sl.shapes[i] for i in range(n0, len(sl.shapes))])
    except Exception:
        pass


def loc_fmt(fmt):
    return fmt


def _legend_row(sl, x, y, w, names, colors, size=8):
    tw = [(_wlen(nm) * size / 72.0 * 0.97 + 0.30) for nm in names]
    total = sum(tw) + 0.12 * (len(names) - 1)
    cx = x + (w - total) / 2.0
    for nm, col, wd in zip(names, colors, tw):
        rect(sl, cx, y + 0.045, 0.11, 0.11, fill=col)
        txt(sl, cx + 0.15, y, wd - 0.15, 0.22, nm, size=size, color=GRAY2)
        cx += wd + 0.12


def _hbar(sl, x, y, w, h, cats, series, colors, gap, fmt, label_size, label_color,
          cat_size, vhi, vlo, pc, legend, vsize):
    n = len(cats)
    leg_h = 0.26 if legend else 0.0
    cw = _catw(cats, cat_size)
    vw = max(_wlen(_fm(v, fmt)) for nm, vs in series for v in vs) * label_size / 72.0 * 0.95 + 0.12
    px = x + cw
    pw = w - cw - vw
    ph = h - leg_h
    row = ph / n
    m = len(series)
    if m > 1:
        bh, gapv = row * 0.29, row * 0.045
    else:
        bh, gapv = min(row / (1 + gap / 100.0), row * 0.80), 0.0
    tot = bh * m + gapv * (m - 1)
    base = px + (0 - vlo) / (vhi - vlo) * pw
    for i, c in enumerate(cats):
        top0 = y + ph - (i + 1) * row
        for k, (nm, vs) in enumerate(series):
            v = vs[i]
            if v is None:
                continue
            L = (v - vlo) / (vhi - vlo) * pw
            ty = top0 + (row - tot) / 2.0 + k * (bh + gapv)
            col = pc.get((k, i), colors[k % len(colors)])
            rect(sl, base, ty, max(L, 0.008), bh, fill=col)
            lab = _fm(v, fmt)
            lx = base + L + 0.05
            if lx + _wlen(lab) * label_size / 72.0 > x + w:
                lx = base + L - _wlen(lab) * label_size / 72.0 - 0.05
            txt(sl, lx, ty - 0.02, vw + 0.18, bh + 0.04, lab, size=label_size,
                color=label_color, anchor=MSO_ANCHOR.MIDDLE)
        txt(sl, x, top0 + (row - tot) / 2.0, cw - 0.10, tot, c, size=cat_size,
            color=GRAY2, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    if legend:
        _legend_row(sl, x, y + ph + 0.04, w, [s[0] for s in series], colors, vsize)


def _vcol(sl, x, y, w, h, cats, series, colors, gap, fmt, label_size, label_color,
          cat_size, vhi, vlo, show_vaxis, gridlines, vsize, legend, pc):
    n = len(cats)
    m = len(series)
    leg_h = 0.26 if legend else 0.0
    axw = 0.46 if show_vaxis else 0.05
    top_pad = label_size / 72.0 * 1.9 + 0.10
    bot = cat_size / 72.0 * 1.7 + 0.10
    px = x + axw
    pw = w - axw - 0.06
    py = y + top_pad
    ph = h - top_pad - bot - leg_h
    span = vhi - vlo
    zero_y = py + ph * (vhi / span)
    if gridlines or show_vaxis:
        for t in range(6):
            v = vlo + span * t / 5.0
            gy = py + ph * (1 - t / 5.0)
            if gridlines and t > 0:
                rect(sl, px, gy, pw, 0.006, fill='E8EFF6')
            if show_vaxis:
                lab = num_axis(v, span)
                txt(sl, x, gy - 0.09, axw - 0.08, 0.18, lab, size=vsize, color=GRAY2,
                    align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    slot = pw / n
    bw = min(slot / (1 + gap / 100.0) / m, slot * 0.62 / m)
    for i, c in enumerate(cats):
        cx = px + slot * (i + 0.5)
        for k, (nm, vs) in enumerate(series):
            v = vs[i]
            if v is None:
                continue
            by = py + ph * (1 - (v - vlo) / span)
            left = cx - bw * m / 2.0 + k * bw
            col = pc.get((k, i), colors[k % len(colors)])
            rect(sl, left, min(by, zero_y), bw, max(abs(zero_y - by), 0.008), fill=col)
            lab = _fm(v, fmt)
            ly = by - label_size / 72.0 * 1.5 - 0.03 if v >= 0 else zero_y + (by - zero_y) + 0.03
            txt(sl, left - 0.14, ly, bw + 0.28, 0.20, lab, size=label_size, color=label_color,
                align=PP_ALIGN.CENTER)
        txt(sl, cx - slot / 2.0, py + ph + 0.05, slot, bot, c, size=cat_size, color=GRAY2,
            align=PP_ALIGN.CENTER)
    rect(sl, px, zero_y, pw, 0.010, fill='D5E3F0')
    if legend:
        _legend_row(sl, x, y + h - leg_h + 0.02, w, [s[0] for s in series], colors, vsize)


def num_axis(v, span):
    if span >= 100:
        return f'{v:,.0f}'
    if span >= 10:
        return f'{v:.0f}'
    if span >= 1:
        return f'{v:.1f}'
    return f'{v:.2f}'


def _vline(sl, x, y, w, h, cats, series, colors, fmt, label_size, label_color,
           cat_size, vhi, vlo, show_vaxis, gridlines, vsize, legend, marker):
    n = len(cats)
    leg_h = 0.28 if legend else 0.0
    axw = 0.46 if show_vaxis else 0.08
    top_pad = 0.14
    bot = cat_size / 72.0 * 1.7 + 0.10
    px = x + axw
    pw = w - axw - 0.10
    py = y + top_pad
    ph = h - top_pad - bot - leg_h
    span = vhi - vlo
    for t in range(6):
        v = vlo + span * t / 5.0
        gy = py + ph * (1 - t / 5.0)
        if gridlines:
            rect(sl, px, gy, pw, 0.006, fill='E8EFF6')
        if show_vaxis:
            txt(sl, x, gy - 0.09, axw - 0.08, 0.18, num_axis(v, span), size=vsize,
                color=GRAY2, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    xs = [px + pw * (i + 0.5) / n for i in range(n)]
    for k, (nm, vs) in enumerate(series):
        col = colors[k % len(colors)]
        pts = [(xs[i], py + ph * (1 - (vs[i] - vlo) / span)) for i in range(n)]
        for a, b in zip(pts, pts[1:]):
            ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(a[0]), Inches(a[1]),
                                         Inches(b[0]), Inches(b[1]))
            ln.line.color.rgb = rgb(col)
            ln.line.width = Pt(2.0)
        if marker:
            for (mx, my) in pts:
                d = 0.072
                rect(sl, mx - d / 2, my - d / 2, d, d, fill=col, shape=MSO_SHAPE.OVAL)
    for i, c in enumerate(cats):
        txt(sl, xs[i] - pw / n / 2.0, py + ph + 0.05, pw / n, bot, c, size=cat_size,
            color=GRAY2, align=PP_ALIGN.CENTER)
    if legend:
        _legend_row(sl, x, y + h - leg_h + 0.03, w, [s[0] for s in series], colors, vsize)
