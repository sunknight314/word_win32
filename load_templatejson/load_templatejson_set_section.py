import os
import win32com.client as win32
import json
from constants.constants import *


def configure_document_headers(doc, config):
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

def set_header_fixedMode(header, text):
    header.Range.Text = text


def configure_document_footers(doc, config):
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
        if section_idx <= config['footer']['main']['sectionStart']:
            style = config['footer']['preface']['style']
            # 前序节第一节从1开始，其余节连续编号
            restart = True if section_idx == 1 else False
            start_num = 1 if section_idx == 1 else None
        else:
            style = config['footer']['main']['style']
            # 正文第一节从1开始，其余节连续编号
            restart = True if section_idx == config['footer']['main']['sectionStart'] + 1 else False
            start_num = 1 if section_idx == config['footer']['main']['sectionStart'] + 1 else None

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
                PageNumberAlignment=config['footer']['main']['alignment'],
                FirstPage=True,
            )
            pnums.NumberStyle = style
            pnums.RestartNumberingAtSection = restart
            if start_num is not None:
                pnums.StartingNumber = start_num

            # 设置字体和对齐
            footer_range = footer.Range
            footer_range.Font.NameFarEast = config['footer']['font']['nameFarEast']
            footer_range.Font.NameAscii = config['footer']['font']['nameAscii']
            footer_range.Font.NameOther = "Times New Roman"
            footer_range.Font.Size = config['footer']['font']['size']
            footer_range.ParagraphFormat.Alignment = config['footer']['paragraph']['alignment']
    # 刷新所有域
    doc.Fields.Update()





def insert_sectionBreak(doc, config):
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

if __name__ == "__main__":

    # 1. 启动 Word
    word = win32.Dispatch("Word.Application")
    word.Visible = True  # 调试时建议可见，方便观察

    # 2. 打开文档（替换为你的路径）
    doc_path = os.path.abspath("source_outlined.docx")
    doc = word.Documents.Open(doc_path)


    json_path = os.path.abspath("E:\wordtest\word_win32\load_templatejson\word_template_config(页码测试)).json")
    with open(json_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    insert_sectionBreak(doc, config)

    configure_document_headers(doc, config)
    configure_document_footers(doc, config)
    # 保存并关闭
    # doc.Save()
