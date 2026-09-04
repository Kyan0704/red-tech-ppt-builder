#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
红黑科技风 PPT 生成器
基于"无人机时代"模板提取的设计规范，支持快速生成风格统一的演示文稿。

用法:
    from generate_ppt import RedTechPPTBuilder
    builder = RedTechPPTBuilder("output.pptx")
    builder.add_cover(title="主标题", subtitle="副标题")
    builder.add_toc(items=["第一章", "第二章", "第三章"])
    builder.add_section_divider(part_no="01", title="章节标题")
    builder.add_content_page(title="页面标题", body="正文内容...")
    builder.save()
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os


# ============================================================
# 设计规范常量（从模板精确提取）
# ============================================================

class DesignSpec:
    """红黑科技风设计规范"""

    # 幻灯片尺寸: 33.87 x 19.05 cm (16:9)
    SLIDE_WIDTH_CM = 33.87
    SLIDE_HEIGHT_CM = 19.05

    # 主色调
    COLOR_ACCENT = RGBColor(0xCA, 0x3B, 0x41)       # 强调红 #CA3B41
    COLOR_DARK_BG = RGBColor(0x1A, 0x1A, 0x1A)       # 深色背景
    COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)           # 白色文字
    COLOR_LIGHT_GRAY = RGBColor(0xCC, 0xCC, 0xCC)      # 浅灰文字
    COLOR_MID_GRAY = RGBColor(0x88, 0x88, 0x88)        # 中灰

    # 字体
    FONT_MAIN = "字魂59号-创粗黑"
    FONT_FALLBACK = "微软雅黑"  # 当主字体不可用时的备选

    # 字号体系
    FONT_SIZE_COVER_TITLE = 60      # 封面主标题
    FONT_SIZE_SECTION_PART = 36     # 章节编号 PART 01.
    FONT_SIZE_SECTION_TITLE = 44     # 章节大标题
    FONT_SIZE_TOC_TITLE = 36         # 目录标题
    FONT_SIZE_PAGE_TITLE = 24        # 页面标题
    FONT_SIZE_BODY = 14              # 正文
    FONT_SIZE_CARD_TITLE = 16        # 卡片标题
    FONT_SIZE_THANK_YOU = 36         # 结尾页
    FONT_SIZE_DATA_NUMBER = 28       # 数据数字
    FONT_SIZE_DATA_LABEL = 12        # 数据标签

    # 标题装饰组合固定位置（模板中"组合5"）
    TITLE_DECOR_LEFT_CM = 3.17
    TITLE_DECOR_TOP_CM = 1.64
    TITLE_DECOR_WIDTH_CM = 7.11
    TITLE_DECOR_HEIGHT_CM = 2.56

    # 内容区域起始位置
    CONTENT_START_TOP_CM = 5.0
    CONTENT_MARGIN_LEFT_CM = 3.17
    CONTENT_MARGIN_RIGHT_CM = 3.17


# ============================================================
# 工具函数
# ============================================================

def _set_font(run, font_name=DesignSpec.FONT_MAIN, size=14,
              bold=False, color=DesignSpec.COLOR_WHITE, italic=False):
    """统一设置字体属性"""
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    # 设置中文字体
    rPr = run._r.get_or_add_rPr()
    eaFont = rPr.find(qn('a:ea'))
    if eaFont is None:
        eaFont = rPr.makeelement(qn('a:ea'), {'typeface': font_name})
        rPr.append(eaFont)
    else:
        eaFont.set('typeface', font_name)


def _add_textbox(slide, left_cm, top_cm, width_cm, height_cm,
                 text, font_size=14, bold=False, color=DesignSpec.COLOR_WHITE,
                 alignment=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                 font_name=DesignSpec.FONT_MAIN):
    """添加文本框并设置统一样式"""
    txBox = slide.shapes.add_textbox(
        Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor

    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    _set_font(run, font_name=font_name, size=font_size, bold=bold, color=color)
    return txBox


def _add_rect(slide, left_cm, top_cm, width_cm, height_cm,
              fill_color=DesignSpec.COLOR_ACCENT, line_color=None):
    """添加矩形色块"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
    return shape


def _add_oval(slide, left_cm, top_cm, size_cm,
              fill_color=DesignSpec.COLOR_ACCENT):
    """添加圆形（同心圆数据标记用）"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Cm(left_cm), Cm(top_cm), Cm(size_cm), Cm(size_cm)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def _set_slide_bg(slide, color=DesignSpec.COLOR_DARK_BG):
    """设置幻灯片纯色背景"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_bg_image(slide, image_path):
    """添加全屏背景图片"""
    if image_path and os.path.exists(image_path):
        slide.shapes.add_picture(
            image_path, 0, 0,
            width=Cm(DesignSpec.SLIDE_WIDTH_CM),
            height=Cm(DesignSpec.SLIDE_HEIGHT_CM)
        )


# ============================================================
# 主生成器类
# ============================================================

class RedTechPPTBuilder:
    """红黑科技风 PPT 生成器"""

    def __init__(self, output_path="output.pptx"):
        self.prs = Presentation()
        self.prs.slide_width = Cm(DesignSpec.SLIDE_WIDTH_CM)
        self.prs.slide_height = Cm(DesignSpec.SLIDE_HEIGHT_CM)
        self.output_path = output_path
        self.blank_layout = self.prs.slide_layouts[6]  # 空白布局

    def _new_slide(self, bg_color=DesignSpec.COLOR_DARK_BG, bg_image=None):
        """创建新幻灯片并设置背景"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        if bg_image:
            _add_bg_image(slide, bg_image)
        else:
            _set_slide_bg(slide, bg_color)
        return slide

    def _add_title_decor(self, slide, title_text):
        """添加页面标题装饰（左上角红色色块+标题，对应模板中的"组合5"）"""
        # 红色装饰条
        _add_rect(slide,
                  DesignSpec.TITLE_DECOR_LEFT_CM,
                  DesignSpec.TITLE_DECOR_TOP_CM + 0.3,
                  0.15, 1.8,
                  fill_color=DesignSpec.COLOR_ACCENT)
        # 标题文字
        _add_textbox(slide,
                     DesignSpec.TITLE_DECOR_LEFT_CM + 0.5,
                     DesignSpec.TITLE_DECOR_TOP_CM,
                     DesignSpec.TITLE_DECOR_WIDTH_CM,
                     DesignSpec.TITLE_DECOR_HEIGHT_CM,
                     title_text,
                     font_size=DesignSpec.FONT_SIZE_PAGE_TITLE,
                     bold=True,
                     color=DesignSpec.COLOR_WHITE,
                     anchor=MSO_ANCHOR.MIDDLE)

    # ----------------------------------------------------------
    # 页面类型 1: 封面页
    # ----------------------------------------------------------
    def add_cover(self, title, subtitle="", bg_image=None,
                  title_size=None, subtitle_size=18):
        """
        添加封面页
        对应模板第1页：全屏背景 + 大标题 + 装饰
        """
        slide = self._new_slide(bg_image=bg_image)

        ts = title_size or DesignSpec.FONT_SIZE_COVER_TITLE

        # 主标题（居中偏上）
        _add_textbox(slide,
                     3.0, 5.5,
                     DesignSpec.SLIDE_WIDTH_CM - 6.0, 4.0,
                     title,
                     font_size=ts,
                     bold=True,
                     color=DesignSpec.COLOR_WHITE,
                     alignment=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)

        # 副标题
        if subtitle:
            _add_textbox(slide,
                         3.0, 10.0,
                         DesignSpec.SLIDE_WIDTH_CM - 6.0, 2.0,
                         subtitle,
                         font_size=subtitle_size,
                         bold=False,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.CENTER,
                         anchor=MSO_ANCHOR.MIDDLE)

        # 左上角装饰文字（对应模板 "General Report Template"）
        _add_textbox(slide,
                     2.0, 1.5, 4.0, 2.5,
                     "General\nReport\nTemplate",
                     font_size=12,
                     bold=False,
                     color=DesignSpec.COLOR_LIGHT_GRAY,
                     alignment=PP_ALIGN.LEFT)

        return slide

    # ----------------------------------------------------------
    # 页面类型 2: 目录页
    # ----------------------------------------------------------
    def add_toc(self, items, title="Contents", bg_image=None):
        """
        添加目录页
        对应模板第2页：居中红色标题 + 目录列表
        """
        slide = self._new_slide(bg_image=bg_image)

        # 标题（红色，居中）
        _add_textbox(slide,
                     13.76, 4.45, 6.34, 1.79,
                     title,
                     font_size=DesignSpec.FONT_SIZE_TOC_TITLE,
                     bold=True,
                     color=DesignSpec.COLOR_ACCENT,
                     alignment=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)

        # 目录条目
        item_height = 1.5
        start_top = 7.5
        for i, item in enumerate(items):
            top = start_top + i * item_height
            # 序号
            _add_textbox(slide,
                         12.0, top, 1.5, item_height,
                         f"0{i+1}",
                         font_size=16,
                         bold=True,
                         color=DesignSpec.COLOR_ACCENT,
                         alignment=PP_ALIGN.RIGHT,
                         anchor=MSO_ANCHOR.MIDDLE)
            # 条目文字
            _add_textbox(slide,
                         14.0, top, 8.0, item_height,
                         item,
                         font_size=16,
                         bold=False,
                         color=DesignSpec.COLOR_WHITE,
                         alignment=PP_ALIGN.LEFT,
                         anchor=MSO_ANCHOR.MIDDLE)

        return slide

    # ----------------------------------------------------------
    # 页面类型 3: 章节过渡页
    # ----------------------------------------------------------
    def add_section_divider(self, part_no, title, subtitle="",
                             bg_image=None, left_image=None):
        """
        添加章节过渡页
        对应模板第3/8/13/18页：左侧大图 + 右侧 PART 编号 + 大标题
        """
        slide = self._new_slide(bg_image=bg_image)

        # 左侧装饰图（如果提供）
        if left_image and os.path.exists(left_image):
            slide.shapes.add_picture(
                left_image,
                Cm(-0.05), Cm(-0.02),
                width=Cm(24.06), height=Cm(15.78)
            )

        # PART 编号（36pt，白色）
        _add_textbox(slide,
                     14.79, 8.23, 10.85, 1.5,
                     f"PART {part_no}.",
                     font_size=DesignSpec.FONT_SIZE_SECTION_PART,
                     bold=False,
                     color=DesignSpec.COLOR_WHITE,
                     alignment=PP_ALIGN.LEFT)

        # 章节大标题（44pt，加粗，白色）
        _add_textbox(slide,
                     14.79, 9.8, 16.0, 2.5,
                     title,
                     font_size=DesignSpec.FONT_SIZE_SECTION_TITLE,
                     bold=True,
                     color=DesignSpec.COLOR_WHITE,
                     alignment=PP_ALIGN.LEFT,
                     anchor=MSO_ANCHOR.TOP)

        # 副标题（可选）
        if subtitle:
            _add_textbox(slide,
                         14.79, 12.8, 16.0, 1.5,
                         subtitle,
                         font_size=14,
                         bold=False,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.LEFT)

        return slide

    # ----------------------------------------------------------
    # 页面类型 4: 标准内容页（带标题栏）
    # ----------------------------------------------------------
    def add_content_page(self, title, body="", bullets=None,
                         bg_image=None):
        """
        添加标准内容页（左上角标题装饰 + 正文）
        对应模板中多数内容页的基础布局
        """
        slide = self._new_slide(bg_image=bg_image)
        self._add_title_decor(slide, title)

        # 正文区域
        content_top = DesignSpec.CONTENT_START_TOP_CM
        content_left = DesignSpec.CONTENT_MARGIN_LEFT_CM
        content_width = DesignSpec.SLIDE_WIDTH_CM - 2 * content_left
        content_height = DesignSpec.SLIDE_HEIGHT_CM - content_top - 2.0

        if bullets:
            # 项目符号列表
            txBox = slide.shapes.add_textbox(
                Cm(content_left), Cm(content_top),
                Cm(content_width), Cm(content_height)
            )
            tf = txBox.text_frame
            tf.word_wrap = True
            for i, bullet in enumerate(bullets):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(12)
                run = p.add_run()
                run.text = f"▸ {bullet}"
                _set_font(run, size=DesignSpec.FONT_SIZE_BODY,
                          color=DesignSpec.COLOR_LIGHT_GRAY)
        elif body:
            # 段落正文
            _add_textbox(slide,
                         content_left, content_top,
                         content_width, content_height,
                         body,
                         font_size=DesignSpec.FONT_SIZE_BODY,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.LEFT,
                         anchor=MSO_ANCHOR.TOP)

        return slide

    # ----------------------------------------------------------
    # 页面类型 5: 三栏卡片页
    # ----------------------------------------------------------
    def add_three_column(self, title, cards, bg_image=None):
        """
        添加三栏卡片页
        对应模板第5页：三个等宽卡片并排，每个卡片含图片+红色标签+文字
        cards: [{"title":..., "desc":..., "image":...}, ...]
        """
        slide = self._new_slide(bg_image=bg_image)
        self._add_title_decor(slide, title)

        card_width = 8.0
        card_height = 9.0
        gap = 1.5
        total_width = 3 * card_width + 2 * gap
        start_left = (DesignSpec.SLIDE_WIDTH_CM - total_width) / 2
        top = 6.0

        for i, card in enumerate(cards):
            left = start_left + i * (card_width + gap)

            # 卡片图片（如果提供）
            if card.get("image") and os.path.exists(card["image"]):
                slide.shapes.add_picture(
                    card["image"],
                    Cm(left), Cm(top),
                    width=Cm(card_width), height=Cm(card_height * 0.55)
                )

            # 红色标签条
            _add_rect(slide,
                      left + 0.5, top + card_height * 0.55 + 0.3,
                      card_width - 1.0, 0.8,
                      fill_color=DesignSpec.COLOR_ACCENT)

            # 卡片标题
            _add_textbox(slide,
                         left + 0.5, top + card_height * 0.55 + 0.3,
                         card_width - 1.0, 0.8,
                         card.get("title", ""),
                         font_size=DesignSpec.FONT_SIZE_CARD_TITLE,
                         bold=True,
                         color=DesignSpec.COLOR_WHITE,
                         alignment=PP_ALIGN.CENTER,
                         anchor=MSO_ANCHOR.MIDDLE)

            # 卡片描述
            _add_textbox(slide,
                         left + 0.3, top + card_height * 0.55 + 1.5,
                         card_width - 0.6, card_height * 0.4,
                         card.get("desc", ""),
                         font_size=12,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.LEFT)

        return slide

    # ----------------------------------------------------------
    # 页面类型 6: 图文混排页（左图右文+红色色块）
    # ----------------------------------------------------------
    def add_image_text(self, title, image_path, body,
                       accent_text="", bg_image=None):
        """
        添加图文混排页
        对应模板第6页：左侧图片 + 右侧红色色块 + 白色文字
        """
        slide = self._new_slide(bg_image=bg_image)
        self._add_title_decor(slide, title)

        # 左侧图片
        if image_path and os.path.exists(image_path):
            slide.shapes.add_picture(
                image_path,
                Cm(3.17), Cm(5.5),
                width=Cm(12.0), height=Cm(10.0)
            )

        # 右侧红色色块
        _add_rect(slide,
                  18.53, 7.1, 15.31, 7.93,
                  fill_color=DesignSpec.COLOR_ACCENT)

        # 强调文字（红色色块上的白色文字）
        if accent_text:
            _add_textbox(slide,
                         19.5, 7.5, 13.5, 2.0,
                         accent_text,
                         font_size=20,
                         bold=True,
                         color=DesignSpec.COLOR_WHITE,
                         alignment=PP_ALIGN.LEFT)

        # 正文（色块下方或右侧）
        _add_textbox(slide,
                     19.5, 10.0, 13.5, 4.5,
                     body,
                     font_size=DesignSpec.FONT_SIZE_BODY,
                     color=DesignSpec.COLOR_WHITE,
                     alignment=PP_ALIGN.LEFT)

        return slide

    # ----------------------------------------------------------
    # 页面类型 7: 全图过渡页
    # ----------------------------------------------------------
    def add_full_image_divider(self, title, subtitle="",
                                bg_image=None, top_image=None):
        """
        添加全图过渡页
        对应模板第7/12/17页：全屏背景图 + 居中大标题 + 顶部小图
        """
        slide = self._new_slide(bg_image=bg_image)

        # 顶部装饰小图
        if top_image and os.path.exists(top_image):
            slide.shapes.add_picture(
                top_image,
                Cm(12.89), Cm(2.77),
                width=Cm(8.11), height=Cm(4.79)
            )

        # 居中大标题
        _add_textbox(slide,
                     3.0, 7.5,
                     DesignSpec.SLIDE_WIDTH_CM - 6.0, 3.0,
                     title,
                     font_size=40,
                     bold=True,
                     color=DesignSpec.COLOR_WHITE,
                     alignment=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)

        # 副标题
        if subtitle:
            _add_textbox(slide,
                         3.0, 11.0,
                         DesignSpec.SLIDE_WIDTH_CM - 6.0, 1.5,
                         subtitle,
                         font_size=14,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.CENTER)

        return slide

    # ----------------------------------------------------------
    # 页面类型 8: 数据圆环页
    # ----------------------------------------------------------
    def add_data_circles(self, title, data_items, body="",
                          bg_image=None):
        """
        添加数据圆环页
        对应模板第11页：红色同心圆标记 + 数据数字 + 说明文字
        data_items: [{"number": "125", "label": "指标名称"}, ...]
        """
        slide = self._new_slide(bg_image=bg_image)
        self._add_title_decor(slide, title)

        circle_size = 3.16
        gap = 2.0
        n = len(data_items)
        total_width = n * circle_size + (n - 1) * gap
        start_left = (DesignSpec.SLIDE_WIDTH_CM - total_width) / 2
        top = 6.5

        for i, item in enumerate(data_items):
            left = start_left + i * (circle_size + gap)

            # 红色圆环
            _add_oval(slide, left, top, circle_size,
                      fill_color=DesignSpec.COLOR_ACCENT)

            # 数据数字
            _add_textbox(slide,
                         left - 0.5, top + 0.3,
                         circle_size + 1.0, circle_size - 0.6,
                         item.get("number", ""),
                         font_size=DesignSpec.FONT_SIZE_DATA_NUMBER,
                         bold=True,
                         color=DesignSpec.COLOR_WHITE,
                         alignment=PP_ALIGN.CENTER,
                         anchor=MSO_ANCHOR.MIDDLE)

            # 数据标签
            _add_textbox(slide,
                         left - 1.0, top + circle_size + 0.3,
                         circle_size + 2.0, 1.5,
                         item.get("label", ""),
                         font_size=DesignSpec.FONT_SIZE_DATA_LABEL,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.CENTER)

        # 底部说明文字
        if body:
            _add_textbox(slide,
                         DesignSpec.CONTENT_MARGIN_LEFT_CM, 14.0,
                         DesignSpec.SLIDE_WIDTH_CM - 2 * DesignSpec.CONTENT_MARGIN_LEFT_CM, 3.0,
                         body,
                         font_size=DesignSpec.FONT_SIZE_BODY,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.CENTER)

        return slide

    # ----------------------------------------------------------
    # 页面类型 9: 左右对比页
    # ----------------------------------------------------------
    def add_compare_page(self, title, left_title, left_items,
                         right_title, right_items, bg_image=None):
        """
        添加左右对比页
        对应模板第14/19页：左侧红色强调标题 + 右侧多条内容
        """
        slide = self._new_slide(bg_image=bg_image)
        self._add_title_decor(slide, title)

        # 左侧区域
        _add_textbox(slide,
                     3.17, 6.0, 10.0, 2.0,
                     left_title,
                     font_size=18,
                     bold=True,
                     color=DesignSpec.COLOR_ACCENT,
                     alignment=PP_ALIGN.LEFT)

        for i, item in enumerate(left_items):
            _add_textbox(slide,
                         3.17, 8.5 + i * 1.8, 10.0, 1.5,
                         item,
                         font_size=DesignSpec.FONT_SIZE_BODY,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.LEFT)

        # 右侧区域
        _add_textbox(slide,
                         18.0, 6.0, 12.0, 2.0,
                         right_title,
                         font_size=18,
                         bold=True,
                         color=DesignSpec.COLOR_WHITE,
                         alignment=PP_ALIGN.LEFT)

        for i, item in enumerate(right_items):
            _add_textbox(slide,
                         18.0, 8.5 + i * 1.8, 12.0, 1.5,
                         item,
                         font_size=DesignSpec.FONT_SIZE_BODY,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.LEFT)

        return slide

    # ----------------------------------------------------------
    # 页面类型 10: 结尾页
    # ----------------------------------------------------------
    def add_thank_you(self, text="THANK YOU", subtitle="",
                       bg_image=None):
        """
        添加结尾页
        对应模板第22页：全屏背景 + 居中 THANK YOU
        """
        slide = self._new_slide(bg_image=bg_image)

        _add_textbox(slide,
                     3.0, 8.0,
                     DesignSpec.SLIDE_WIDTH_CM - 6.0, 3.0,
                     text,
                     font_size=DesignSpec.FONT_SIZE_THANK_YOU,
                     bold=True,
                     color=DesignSpec.COLOR_WHITE,
                     alignment=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)

        if subtitle:
            _add_textbox(slide,
                         3.0, 11.5,
                         DesignSpec.SLIDE_WIDTH_CM - 6.0, 1.5,
                         subtitle,
                         font_size=14,
                         color=DesignSpec.COLOR_LIGHT_GRAY,
                         alignment=PP_ALIGN.CENTER)

        return slide

    # ----------------------------------------------------------
    # 批量生成：从配置字典生成完整PPT
    # ----------------------------------------------------------
    def build_from_config(self, config):
        """
        从配置字典批量生成PPT

        config 示例:
        {
            "output": "my_presentation.pptx",
            "slides": [
                {"type": "cover", "title": "主标题", "subtitle": "副标题"},
                {"type": "toc", "items": ["第一章", "第二章"]},
                {"type": "section", "part_no": "01", "title": "章节标题"},
                {"type": "content", "title": "页面标题", "body": "正文"},
                {"type": "three_column", "title": "三栏页", "cards": [...]},
                {"type": "image_text", "title": "图文页", "image_path": "...", "body": "..."},
                {"type": "full_image", "title": "全图过渡", "bg_image": "..."},
                {"type": "data_circles", "title": "数据页", "data_items": [...]},
                {"type": "compare", "title": "对比页", "left_title": "...", "left_items": [...], ...},
                {"type": "thank_you", "text": "THANK YOU"}
            ]
        }
        """
        slide_handlers = {
            "cover": lambda s: self.add_cover(
                title=s.get("title", ""),
                subtitle=s.get("subtitle", ""),
                bg_image=s.get("bg_image"),
                title_size=s.get("title_size")
            ),
            "toc": lambda s: self.add_toc(
                items=s.get("items", []),
                title=s.get("title", "Contents"),
                bg_image=s.get("bg_image")
            ),
            "section": lambda s: self.add_section_divider(
                part_no=s.get("part_no", "01"),
                title=s.get("title", ""),
                subtitle=s.get("subtitle", ""),
                bg_image=s.get("bg_image"),
                left_image=s.get("left_image")
            ),
            "content": lambda s: self.add_content_page(
                title=s.get("title", ""),
                body=s.get("body", ""),
                bullets=s.get("bullets"),
                bg_image=s.get("bg_image")
            ),
            "three_column": lambda s: self.add_three_column(
                title=s.get("title", ""),
                cards=s.get("cards", []),
                bg_image=s.get("bg_image")
            ),
            "image_text": lambda s: self.add_image_text(
                title=s.get("title", ""),
                image_path=s.get("image_path"),
                body=s.get("body", ""),
                accent_text=s.get("accent_text", ""),
                bg_image=s.get("bg_image")
            ),
            "full_image": lambda s: self.add_full_image_divider(
                title=s.get("title", ""),
                subtitle=s.get("subtitle", ""),
                bg_image=s.get("bg_image"),
                top_image=s.get("top_image")
            ),
            "data_circles": lambda s: self.add_data_circles(
                title=s.get("title", ""),
                data_items=s.get("data_items", []),
                body=s.get("body", ""),
                bg_image=s.get("bg_image")
            ),
            "compare": lambda s: self.add_compare_page(
                title=s.get("title", ""),
                left_title=s.get("left_title", ""),
                left_items=s.get("left_items", []),
                right_title=s.get("right_title", ""),
                right_items=s.get("right_items", []),
                bg_image=s.get("bg_image")
            ),
            "thank_you": lambda s: self.add_thank_you(
                text=s.get("text", "THANK YOU"),
                subtitle=s.get("subtitle", ""),
                bg_image=s.get("bg_image")
            ),
        }

        for slide_config in config.get("slides", []):
            slide_type = slide_config.get("type", "content")
            handler = slide_handlers.get(slide_type)
            if handler:
                handler(slide_config)
            else:
                print(f"Warning: unknown slide type '{slide_type}', skipped")

        if config.get("output"):
            self.output_path = config["output"]

        return self.save()

    # ----------------------------------------------------------
    # 保存
    # ----------------------------------------------------------
    def save(self):
        """保存PPT文件"""
        self.prs.save(self.output_path)
        print(f"PPT已保存: {self.output_path}")
        print(f"共 {len(self.prs.slides)} 页")
        return self.output_path


# ============================================================
# 命令行入口
# ============================================================

if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) < 2:
        print("用法: python generate_ppt.py <config.json>")
        print("示例配置:")
        print(json.dumps({
            "output": "demo.pptx",
            "slides": [
                {"type": "cover", "title": "演示标题", "subtitle": "副标题"},
                {"type": "toc", "items": ["第一部分", "第二部分", "第三部分"]},
                {"type": "section", "part_no": "01", "title": "章节一"},
                {"type": "content", "title": "内容页", "body": "这是正文内容。"},
                {"type": "thank_you"}
            ]
        }, ensure_ascii=False, indent=2))
        sys.exit(0)

    config_path = sys.argv[1]
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    builder = RedTechPPTBuilder(config.get("output", "output.pptx"))
    builder.build_from_config(config)
