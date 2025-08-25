import os
import win32com.client as win32
import json
from constants.constants import *


def insert_contents(doc, config):
    doc.Paragraphs(1).Range.InsertParagraphBefore()

    # 新插入的第一段范围
    title_range = doc.Paragraphs(1).Range
    # 写入标题文本（此操作会覆盖该段范围内内容），为保险起见在末尾插入段落符
    title_range.Text = "目录"
    title_range.InsertParagraphAfter()

    title_range.Style = "标题 1"

    title_range.ParagraphFormat.OutlineLevel = 5

    # 目录应该插入到现在的第二段（标题之后）
    toc_insert_range = doc.Paragraphs(2).Range
    # 插入目录，使用 Heading 样式，启用超链接
    toc = doc.TablesOfContents.Add(
        Range=toc_insert_range,
        UseHeadingStyles=True,
        LowerHeadingLevel=3,
        UseHyperlinks=True
    )


    toc.Range.Collapse(WdCollapseDirection.wdCollapseEnd)
    # 插入奇数页分节符（若当前已经在奇数页，Word 会跳到下一奇数页）
    toc.Range.InsertBreak(WdBreakType.wdSectionBreakOddPage)


    doc.Fields.Update()


if __name__ == "__main__":
    # 1. 启动 Word
    word = win32.Dispatch("Word.Application")
    word.Visible = True  # 调试时建议可见，方便观察

    # 2. 打开文档（替换为你的路径）
    doc_path = os.path.abspath("content_test.docx")
    doc = word.Documents.Open(doc_path)


    json_path = os.path.abspath("E:\wordtest\word_win32\load_templatejson\word_template_config(页码测试).json")
    with open(json_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    insert_contents(doc, config)