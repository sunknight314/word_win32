import os
import win32com.client as win32

from constants.constants import *


def configure_document_headers(doc):
    """配置文档页眉：奇偶页处理、域代码插入、节间连接"""
    # 1. 全局启用奇偶页不同页眉/页脚
    doc.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True

    # 2. 第一节页眉配置
    first_section = doc.Sections(1)

    # 2.1 主页眉（奇数页）插入域代码
    primary_header = first_section.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
    header_range = primary_header.Range
    header_range.Text = ""  # 清空原有内容

    # 插入标题编号域
    header_range.Fields.Add(
        Range=header_range,
        Type=WdFieldType.wdFieldEmpty,
        Text=r'STYLEREF  "my7标题 1" \n ',  # \n开关显示标题编号
        PreserveFormatting=True
    )
    header_range.Collapse(Direction=WdCollapseDirection.wdCollapseEnd)  # 移动光标到域结尾
    header_range.Text = " "  # 添加分隔空格
    header_range.Collapse(Direction=WdCollapseDirection.wdCollapseEnd)

    # 插入标题文本域
    header_range.Fields.Add(
        Range=header_range,
        Type=WdFieldType.wdFieldEmpty,
        Text=r'STYLEREF  "my7标题 1" ',  # 无开关显示纯文本标题
        PreserveFormatting=True
    )

    header_range = primary_header.Range
    header_range.Font.NameFarEast = "宋体"  # 设置字体为宋体
    header_range.Font.NameAscii = "Times New Roman"
    header_range.Font.NameOther = "Times New Roman"
    header_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter  # 居中对齐
    header_range.Font.Size = 12  # 设置字号为12
    header_range.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleDouble    # 添加下边框
    header_range.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
    header_range.Borders(WdBorderType.wdBorderBottom).Color = WdColor.wdColorBlack  # 设置边框颜色为黑色


    # 2.2 设置第一节偶数页页眉
    even_header = first_section.Headers(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
    even_header_range = even_header.Range
    even_header_range.Text = "偶数页页眉"  # 静态文本（可替换为域代码）

    header_range = even_header.Range
    header_range.Font.NameFarEast = "宋体"  # 设置字体为宋体
    header_range.Font.NameAscii = "Times New Roman"
    header_range.Font.NameOther = "Times New Roman"
    header_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter  # 居中对齐
    header_range.Font.Size = 12  # 设置字号为12
    header_range.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleDouble    # 添加下边框
    header_range.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
    header_range.Borders(WdBorderType.wdBorderBottom).Color = WdColor.wdColorBlack  # 设置边框颜色为黑色

    # 3. 连接后续节的页眉到前一节
    sections = doc.Sections
    for section_idx in range(2, sections.Count + 1):
        current_section = sections(section_idx)
        # 连接主/偶数页眉到前一节
        current_section.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = True
        current_section.Headers(WdHeaderFooterIndex.wdHeaderFooterEvenPages).LinkToPrevious = True

    # 4. 刷新所有域代码
    doc.Fields.Update()

def configure_document_footers(doc, front_sections_count=1):
    """
    配置文档页脚：
    - 前序部分使用罗马数字，从1开始连续
    - 正文部分使用阿拉伯数字，从1开始，后续节连续
    front_sections_count: 前序部分节数
    """

    sections = doc.Sections
    for section_idx in range(1, sections.Count + 1):
        section = sections(section_idx)

        # 解除与前一节的页脚链接
        for hf_idx in [WdHeaderFooterIndex.wdHeaderFooterPrimary,
                       WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                       WdHeaderFooterIndex.wdHeaderFooterFirstPage]:
            section.Footers(hf_idx).LinkToPrevious = False

        # 判断该节页码类型
        if section_idx <= front_sections_count:
            style = WdPageNumberStyle.wdPageNumberStyleUppercaseRoman
            # 前序节第一节从1开始，其余节连续编号
            restart = True if section_idx == 1 else False
            start_num = 1 if section_idx == 1 else None
        else:
            style = WdPageNumberStyle.wdPageNumberStyleArabic
            # 正文第一节从1开始，其余节连续编号
            restart = True if section_idx == front_sections_count + 1 else False
            start_num = 1 if section_idx == front_sections_count + 1 else None

        # 奇偶页页脚统一页码
        for hf_idx in [WdHeaderFooterIndex.wdHeaderFooterPrimary,
                       WdHeaderFooterIndex.wdHeaderFooterEvenPages]:
            footer = section.Footers(hf_idx)

            # 删除旧页码
            while footer.PageNumbers.Count > 0:
                footer.PageNumbers(1).Delete()

            # 添加新页码
            pnums = footer.PageNumbers
            pnums.Add(
                PageNumberAlignment=WdPageNumberAlignment.wdAlignPageNumberCenter,
                FirstPage=True,
            )
            pnums.NumberStyle = style
            pnums.RestartNumberingAtSection = restart
            if start_num is not None:
                pnums.StartingNumber = start_num

            # 设置字体和对齐
            footer_range = footer.Range
            footer_range.Font.NameFarEast = "宋体"
            footer_range.Font.NameAscii = "Times New Roman"
            footer_range.Font.NameOther = "Times New Roman"
            footer_range.Font.Size = 12
            footer_range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter

    # 刷新所有域
    doc.Fields.Update()


# 1. 启动 Word
word = win32.Dispatch("Word.Application")
word.Visible = True  # 调试时建议可见，方便观察

# 2. 打开文档（替换为你的路径）
doc_path = os.path.abspath("source_outlined.docx")
doc = word.Documents.Open(doc_path)

# 遍历段落，按大纲等级插分节符
for i, para in enumerate(doc.Paragraphs):
    # OutlineLevel: 1=标题1, 2=标题2, ... 10=正文
    if para.OutlineLevel == 1:
        # 跳过第一个段落，避免文首插分节符
        if i == 0:
            continue
        rng = para.Range
        # 在段落前插入下一页分节符
        rng.InsertBreak(WdBreakType.wdSectionBreakOddPage)


configure_document_headers(doc)
configure_document_footers(doc)
# 保存并关闭
# doc.Save()
