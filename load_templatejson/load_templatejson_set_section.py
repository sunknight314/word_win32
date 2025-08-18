import os
import win32com.client as win32
import json
from constants.constants import *


def configure_document_headers(doc):
    """配置文档页眉：奇偶页处理、域代码插入、节间连接"""



    # 1. 全局启用奇偶页不同页眉/页脚
    doc.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = config['section']['differentOddEven']

    # 如果启用奇偶页
    if config['section']['differentOddEven']:
        first_section = doc.Sections(1)
        # 设置奇数页页眉
        odd_header = first_section.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        # 设置具体页眉内容
        if config['header']['primaryText'] == "chapter":
            set_header_chapterMode(odd_header)
        else:
            set_header_fixedMode(odd_header, config['header']['primaryText'])

        odd_header_range = odd_header.Range
        odd_header_range.Font.NameFarEast = config['header']['font']  # 设置字体为宋体
        odd_header_range.Font.NameAscii = "Times New Roman"
        odd_header_range.Font.NameOther = "Times New Roman"
        odd_header_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter  # 居中对齐
        odd_header_range.Font.Size = config['header']['size']  # 设置字号为12
        odd_header_range.Borders(WdBorderType.wdBorderBottom).LineStyle = config['header']['borderStyle']  # 添加下边框
        odd_header_range.Borders(WdBorderType.wdBorderBottom).LineWidth = config['header']['borderWidth']
        odd_header_range.Borders(WdBorderType.wdBorderBottom).Color = WdColor.wdColorBlack  # 设置边框颜色为黑色

    # 设置偶数页页眉
        even_header = first_section.Headers(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
        # 设置具体页眉内容
        if config['header']['evenText'] == "chapter":
            set_header_chapterMode(even_header)
        else:
            set_header_fixedMode(even_header, config['header']['evenText'])

        even_header_range = even_header.Range
        even_header_range.Font.NameFarEast = config['header']['font']  # 设置字体为宋体
        even_header_range.Font.NameAscii = "Times New Roman"
        even_header_range.Font.NameOther = "Times New Roman"
        even_header_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter  # 居中对齐
        even_header_range.Font.Size = config['header']['size']  # 设置字号为12
        even_header_range.Borders(WdBorderType.wdBorderBottom).LineStyle = config['header']['borderStyle']  # 添加下边框
        even_header_range.Borders(WdBorderType.wdBorderBottom).LineWidth = config['header']['borderWidth']
        even_header_range.Borders(WdBorderType.wdBorderBottom).Color = WdColor.wdColorBlack  # 设置边框颜色为黑色


    # 如果不启用奇偶页
    else:
        # 设置统一页眉
        first_section = doc.Sections(1)
        # 设置奇数页页眉
        primary_header = first_section.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        # 设置具体页眉内容
        if config['header']['primaryText'] == "chapter":
            set_header_chapterMode(primary_header)
        else:
            set_header_fixedMode(primary_header, config['header']['primaryText'])

        primary_header_range = primary_header.Range
        primary_header_range.Font.NameFarEast = config['header']['font']  # 设置字体为宋体
        primary_header_range.Font.NameAscii = "Times New Roman"
        primary_header_range.Font.NameOther = "Times New Roman"
        primary_header_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter  # 居中对齐
        primary_header_range.Font.Size = config['header']['size']  # 设置字号为12
        primary_header_range.Borders(WdBorderType.wdBorderBottom).LineStyle = config['header']['borderStyle']  # 添加下边框
        primary_header_range.Borders(WdBorderType.wdBorderBottom).LineWidth = config['header']['borderWidth']

    # 3. 连接后续节的页眉到前一节
    sections = doc.Sections
    for section_idx in range(2, sections.Count + 1):
        current_section = sections(section_idx)
        # 连接主/偶数页眉到前一节
        current_section.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = True
        current_section.Headers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).LinkToPrevious = True

    # 4. 刷新所有域代码
    doc.Fields.Update()

def set_header_chapterMode(header):
    header_range = header.Range
    header_range.Text = ""  # 清空原有内容
    # 插入标题编号域
    header_range.Fields.Add(
        Range=header_range,
        Type=WdFieldType.wdFieldEmpty,
        Text=r'STYLEREF  "my10标题 1" \n ',  # \n开关显示标题编号
        PreserveFormatting=True
    )
    header_range.Collapse(Direction=WdCollapseDirection.wdCollapseEnd)  # 移动光标到域结尾
    header_range.Text = " "  # 添加分隔空格
    header_range.Collapse(Direction=WdCollapseDirection.wdCollapseEnd)
    # 插入标题文本域
    header_range.Fields.Add(
        Range=header_range,
        Type=WdFieldType.wdFieldEmpty,
        Text=r'STYLEREF  "my10标题 1" ',  # 无开关显示纯文本标题
        PreserveFormatting=True
    )

def set_header_fixedMode(header, text):
    header.Range.Text = text


def configure_document_footers(doc):
    """配置文档页脚（可选）"""
    # 奇数页页码
    footer = doc.Sections(1).Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
    footer_rng = footer.Range
    doc.Fields.Add(footer_rng,WdFieldType.wdFieldPage)

    footer_range = footer.Range
    footer_range.Font.NameFarEast = "宋体"  # 设置字体为宋体
    footer_range.Font.NameAscii = "Times New Roman"
    footer_range.Font.NameOther = "Times New Roman"
    footer_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter  # 居中对齐
    footer_range.Font.Size = 12  # 设置字号为12


    # 偶数页页码
    footer = doc.Sections(1).Footers(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
    footer_rng = footer.Range
    doc.Fields.Add(footer_rng, WdFieldType.wdFieldPage)

    footer_range = footer.Range
    footer_range.Font.NameFarEast = "宋体"  # 设置字体为宋体
    footer_range.Font.NameAscii = "Times New Roman"
    footer_range.Font.NameOther = "Times New Roman"
    footer_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter  # 居中对齐
    footer_range.Font.Size = 12  # 设置字号为12

    sections = doc.Sections
    for section_idx in range(2, sections.Count + 1):
        current_section = sections(section_idx)
        # 连接主/偶数页眉到前一节
        current_section.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = True
        current_section.Footers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).LinkToPrevious = True

    # 4. 刷新所有域代码
    doc.Fields.Update()


# 1. 启动 Word
word = win32.Dispatch("Word.Application")
word.Visible = True  # 调试时建议可见，方便观察

# 2. 打开文档（替换为你的路径）
doc_path = os.path.abspath("source_outlined2.docx")
doc = word.Documents.Open(doc_path)


json_path = os.path.abspath("E:\wordtest\word_win32\load_templatejson\word_template_config(1).json")
with open(json_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

# 遍历段落，按大纲等级插分节符
for i, para in enumerate(doc.Paragraphs):
    # OutlineLevel: 1=标题1, 2=标题2, ... 10=正文
    if para.OutlineLevel == 1:
        # 跳过第一个段落，避免文首插分节符
        if i == 0:
            continue
        rng = para.Range
        # 在段落前插入下一页分节符
        rng.InsertBreak(config['section']['breakType'])


configure_document_headers(doc)
# configure_document_footers(doc)
# 保存并关闭
# doc.Save()
