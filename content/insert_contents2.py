import json
import os

import win32com.client as win32
from win32com.client import constants as c

from constants.constants import WdCollapseDirection, WdBreakType


def insert_contents(doc):
    # 在文档开头插入一个段落作为“目录”标题
    doc.Paragraphs(1).Range.InsertParagraphBefore()
    title_range = doc.Paragraphs(1).Range

    # 写入标题并在末尾加段落符号，保证独立段落
    title_range.Text = "目录" + "\r"

    # 设置样式为 Heading1（或 "标题 1"），但通常不希望“目录”自身被收进 TOC，
    # 所以把 OutlineLevel 设为 BodyText。若希望把“目录”也显示在 TOC 中，
    # 可注释掉下一行或改成合适的级别。

    title_range.Style = "标题 1"
    title_range.ParagraphFormat.OutlineLevel = 5

    # 把插入点移动到标题段末并插入奇数页分节符（使正文从新节的奇数页开始）
    temp_range = title_range.Duplicate
    temp_range.Collapse(WdCollapseDirection.wdCollapseEnd)
    temp_range.InsertBreak(WdBreakType.wdSectionBreakOddPage)

    # 分节符插入之后，新的节通常会在文档中产生新的段落，获取第二段作为 TOC 插入点
    toc_insert_range = doc.Paragraphs(2).Range
    toc_insert_range.Text = ""  # 清空，确保不会和正文的第一个标题混到一起

    # 添加目录（TOC），这里抓取 1-3 级标题，启用超链接和页码
    toc = doc.TablesOfContents.Add(
        Range=toc_insert_range,
        UseHeadingStyles=True,
        UpperHeadingLevel=1,
        LowerHeadingLevel=3,
        UseHyperlinks=True,
        IncludePageNumbers=True
    )

    # **关键**：添加 TOC 后立即更新域，使 TOC 展开并取得正确的 Range
    doc.Fields.Update()

    # TOC 展开后，如果需要在 TOC 末尾再插入某种分隔或页码重置，可以在这里操作
    return doc  # 可选，返回 doc 以便其他调用使用



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

    insert_contents(doc)