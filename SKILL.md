---
name: red-tech-ppt-builder
description: "红黑科技风PPT生成器。基于专业模板提取的设计规范（字魂59号-创粗黑字体、#CA3B41强调红、深色背景、16:9宽屏），支持一键生成风格统一的演示文稿。当用户需要制作科技感、商务感、课程培训、产品发布、汇报演示类PPT，或要求套用红黑配色/科技风格/一致排版时使用。支持10种页面类型：封面、目录、章节过渡、标准内容、三栏卡片、图文混排、全图过渡、数据圆环、左右对比、结尾页。可通过Python API或JSON配置批量生成。"
---

# 红黑科技风 PPT 生成器

基于"无人机时代"专业模板的设计规范封装，支持快速生成风格统一、排版专业的演示文稿。

## 快速开始

### 方式一：Python API（推荐，灵活控制）

```python
import sys
sys.path.insert(0, "<skill_dir>/scripts")
from generate_ppt import RedTechPPTBuilder

builder = RedTechPPTBuilder("output.pptx")
builder.add_cover(title="课程主题", subtitle="副标题")
builder.add_toc(items=["第一章 概述", "第二章 原理", "第三章 实践"])
builder.add_section_divider(part_no="01", title="课程概述")
builder.add_content_page(title="页面标题", body="正文内容...")
builder.add_three_column(title="三大特点", cards=[
    {"title": "特点一", "desc": "描述文字"},
    {"title": "特点二", "desc": "描述文字"},
    {"title": "特点三", "desc": "描述文字"},
])
builder.add_thank_you()
builder.save()
```

### 方式二：JSON 配置（批量生成）

```bash
python <skill_dir>/scripts/generate_ppt.py config.json
```

config.json 格式见下方「配置格式」章节。

## 设计规范概要

完整规范见 [references/design_spec.md](references/design_spec.md)。

| 项目 | 规范 |
|------|------|
| 尺寸 | 33.87 × 19.05 cm（16:9） |
| 主字体 | 字魂59号-创粗黑（备选：微软雅黑） |
| 强调色 | #CA3B41（红） |
| 背景色 | #1A1A1A（深灰/黑） |
| 标题文字 | 白色 #FFFFFF |
| 正文文字 | 浅灰 #CCCCCC，14pt |

### 字号体系

- 封面主标题：60pt 加粗
- 章节大标题：44pt 加粗
- 章节编号/目录/结尾：36pt
- 页面标题：24pt 加粗
- 正文：14pt
- 数据数字：28pt 加粗

## 页面类型与方法

所有方法均返回 slide 对象，可继续自定义。

### 1. 封面页 `add_cover(title, subtitle="", bg_image=None, title_size=None)`

全屏背景 + 居中大标题 + 左上角装饰文字。

### 2. 目录页 `add_toc(items, title="Contents", bg_image=None)`

居中红色标题 + 垂直目录列表（自动编号 01/02/03）。

### 3. 章节过渡页 `add_section_divider(part_no, title, subtitle="", bg_image=None, left_image=None)`

左侧装饰图 + 右侧 "PART 01." 编号（36pt）+ 大标题（44pt加粗）。

### 4. 标准内容页 `add_content_page(title, body="", bullets=None, bg_image=None)`

左上角红色竖条标题装饰 + 正文区域。支持 `body`（段落）或 `bullets`（项目列表）。

### 5. 三栏卡片页 `add_three_column(title, cards, bg_image=None)`

三个等宽卡片，每个含图片（可选）+ 红色标签条 + 描述文字。
`cards` 格式：`[{"title":..., "desc":..., "image":...}, ...]`

### 6. 图文混排页 `add_image_text(title, image_path, body, accent_text="", bg_image=None)`

左侧图片 + 右侧红色实心色块 + 白色文字。对应模板第6页布局。

### 7. 全图过渡页 `add_full_image_divider(title, subtitle="", bg_image=None, top_image=None)`

全屏背景图 + 居中大标题。适合章节间的视觉过渡。

### 8. 数据圆环页 `add_data_circles(title, data_items, body="", bg_image=None)`

红色圆环 + 白色数据数字 + 下方标签。
`data_items` 格式：`[{"number":"125", "label":"指标名"}, ...]`

### 9. 左右对比页 `add_compare_page(title, left_title, left_items, right_title, right_items, bg_image=None)`

左侧红色强调标题 + 列表，右侧白色标题 + 列表。适合对比/并列内容。

### 10. 结尾页 `add_thank_you(text="THANK YOU", subtitle="", bg_image=None)`

全屏背景 + 居中结尾文字（36pt加粗）。

## 配置格式（JSON 批量生成）

```json
{
  "output": "my_presentation.pptx",
  "slides": [
    {"type": "cover", "title": "主标题", "subtitle": "副标题"},
    {"type": "toc", "items": ["第一章", "第二章", "第三章"]},
    {"type": "section", "part_no": "01", "title": "章节标题", "subtitle": "可选副标题"},
    {"type": "content", "title": "页面标题", "body": "正文段落", "bullets": ["要点1", "要点2"]},
    {"type": "three_column", "title": "三栏页", "cards": [
      {"title": "卡片1", "desc": "描述"},
      {"title": "卡片2", "desc": "描述"},
      {"title": "卡片3", "desc": "描述"}
    ]},
    {"type": "image_text", "title": "图文页", "image_path": "img.jpg", "body": "正文", "accent_text": "强调语"},
    {"type": "full_image", "title": "全图过渡", "bg_image": "bg.jpg"},
    {"type": "data_circles", "title": "数据页", "data_items": [
      {"number": "98%", "label": "满意度"},
      {"number": "500+", "label": "案例数"}
    ]},
    {"type": "compare", "title": "对比页", "left_title": "方案A", "left_items": ["优点1"], "right_title": "方案B", "right_items": ["优点1"]},
    {"type": "thank_you", "text": "THANK YOU", "subtitle": "感谢观看"}
  ]
}
```

所有页面类型均支持可选的 `bg_image` 字段设置全屏背景图。

## 使用建议

1. **字体依赖**: 主字体"字魂59号-创粗黑"需在运行环境中安装；未安装时自动降级为微软雅黑，风格会有差异。
2. **图片素材**: 脚本生成结构和文字排版；背景图、产品图等视觉素材需自行提供路径，或生成后在PPT中替换。
3. **内容量控制**: 每页正文建议不超过 150 字，三栏卡片每栏不超过 50 字，保持模板的留白感。
4. **配色一致性**: 不要随意引入除规范色板外的颜色；如需强调，使用 #CA3B41 红色。
5. **基于原模板修改**: 如需完全还原模板的装饰图形和图片效果，可直接以原PPT为底本，用本skill的规范指导文字内容的增删改。

## 文件结构

```
red-tech-ppt-builder/
├── SKILL.md                          # 本文件
├── scripts/
│   └── generate_ppt.py               # 核心生成脚本（RedTechPPTBuilder 类）
└── references/
    └── design_spec.md                # 完整设计规范（字体/配色/布局/页面类型详解）
```
