import os
import win32com.client as win32
import json
from constants.constants import *
from content.insert_contents import insert_contents
from init.init import clear_footers_only

from load_templatejson.load_templatejson_set_list_style import load_template_from_json
from load_templatejson.load_templatejson_set_section import insert_sectionBreak, configure_document_headers, \
    configure_document_footers

# 1. 启动 Word
word = win32.Dispatch("Word.Application")
word.Visible = True  # 调试时建议可见，方便观察

# 2. 打开文档（替换为你的路径）
doc_path = os.path.abspath("source_outlined.docx")
doc = word.Documents.Open(doc_path)


json_path = os.path.abspath("E:\wordtest\word_win32\load_templatejson\word_template_config(页码测试).json")
with open(json_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

# 3. 设置央样式列表与应用样式
load_template_from_json(doc, config)

# 4. 插入分节符
insert_sectionBreak(doc, config)

# 5.清除页脚
clear_footers_only(doc)

# 5. 插入目录
insert_contents(doc, config)

# 6. 配置页眉页脚
configure_document_headers(doc, config)
configure_document_footers(doc, config)
