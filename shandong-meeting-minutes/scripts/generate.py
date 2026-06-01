#!/usr/bin/env python3
"""山东公司会议纪要 .docx 生成器

严格按照《山东公司日常文稿规范2026》排版。
用法:
    from generate import create_meeting_minutes
    create_meeting_minutes(output_path, title, meta_para, s1_data, s2_data, attendees, recorder)
"""

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# 辅助函数
# ============================================================

def _set_east_asia(run, font_name):
    """设置中文字体"""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)


def _set_line_spacing(para, spacing_pt=28.9):
    """设置固定行间距"""
    pPr = para._element.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    sp.set(qn('w:line'), str(int(spacing_pt * 20)))
    sp.set(qn('w:lineRule'), 'exact')


def _set_char_spacing(run, spacing_pt=-0.2):
    """设置字符间距紧缩"""
    rPr = run._element.get_or_add_rPr()
    cs = OxmlElement('w:spacing')
    rPr.append(cs)
    cs.set(qn('w:val'), str(int(spacing_pt * 20)))


# ============================================================
# 文档构建器
# ============================================================

class MinutesBuilder:
    """会议纪要文档构建器"""

    def __init__(self):
        self.doc = Document()
        for s in self.doc.sections:
            s.top_margin = Cm(2.54)
            s.bottom_margin = Cm(2.54)
            s.left_margin = Cm(3.18)
            s.right_margin = Cm(3.18)

    def add_title(self, text):
        """大标题：方正小标宋简体 / 宋体 二号 居中"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.font.size = Pt(22)
        r.font.name = '方正小标宋简体'
        _set_east_asia(r, '方正小标宋简体')

    def add_body(self, text, bold=False, indent=True):
        """正文：仿宋_GB2312 三号 缩进2字符 紧缩0.2磅 行距28.9磅 两端对齐"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if indent:
            p.paragraph_format.first_line_indent = Cm(0.74)
        _set_line_spacing(p)
        r = p.add_run(text)
        r.font.size = Pt(16)
        r.font.name = '仿宋_GB2312'
        r.bold = bold
        _set_east_asia(r, '仿宋_GB2312')
        _set_char_spacing(r)
        return p

    def add_h1(self, text):
        """一级标题：黑体三号加粗"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(0.74)
        _set_line_spacing(p)
        r = p.add_run(text)
        r.font.size = Pt(16)
        r.bold = True
        r.font.name = '黑体'
        _set_east_asia(r, '黑体')
        _set_char_spacing(r)
        return p

    def add_h2(self, text):
        """二级标题：楷体_GB2312三号"""
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(0.74)
        _set_line_spacing(p)
        r = p.add_run(text)
        r.font.size = Pt(16)
        r.bold = False
        r.font.name = '楷体_GB2312'
        _set_east_asia(r, '楷体_GB2312')
        _set_char_spacing(r)
        return p

    def add_section_items(self, items):
        """逐段输出分类好的内容列表（含heading2和body）"""
        for item in items:
            if item['type'] == 'heading2':
                self.add_h2(item['text'])
            elif item['type'] == 'body':
                self.add_body(item['text'])

    def save(self, path):
        self.doc.save(path)
        print(f'会议纪要已生成: {path}')


# ============================================================
# 顶层 API
# ============================================================

def create_meeting_minutes(
    output_path, title, meta_paragraph,
    section1_title, section1_content,
    section2_title, section2_content,
    attendees, recorder
):
    """生成标准格式的会议纪要 .docx 文件"""
    b = MinutesBuilder()
    b.add_title(title)
    b.doc.add_paragraph()  # 空行
    b.add_body(meta_paragraph)
    b.add_body('纪要如下：')

    # 第一部分
    b.add_h1(section1_title)
    b.add_section_items(section1_content)

    # 第二部分
    b.add_h1(section2_title)
    b.add_section_items(section2_content)

    # 出席/记录
    b.add_body(attendees)
    b.add_body(recorder)

    b.save(output_path)


# ============================================================
# CLI 入口（用于测试）
# ============================================================

if __name__ == '__main__':
    import sys
    # 简单冒烟测试
    if len(sys.argv) < 2:
        output = '/tmp/test_minutes.docx'
    else:
        output = sys.argv[1]

    create_meeting_minutes(
        output_path=output,
        title='山东公司2026年测试生产营销工作会纪要',
        meta_paragraph='2026年6月1日，山东公司副总经理张锋通过远程会议系统主持召开了测试会议。',
        section1_title='一、生产营销工作完成情况',
        section1_content=[
            {'type': 'heading2', 'text': '（一）年度PBA核心指标完成情况'},
            {'type': 'body', 'text': '测试数据显示各项指标正常。'},
            {'type': 'heading2', 'text': '（二）重点工作进展'},
            {'type': 'body', 'text': '1. 测试工作进展顺利。'},
        ],
        section2_title='二、后续安全生产工作安排',
        section2_content=[
            {'type': 'body', 'text': '1. 测试安排事项一。'},
            {'type': 'body', 'text': '2. 测试安排事项二。'},
        ],
        attendees='出席：测试人员。',
        recorder='记录：测试记录人。',
    )
