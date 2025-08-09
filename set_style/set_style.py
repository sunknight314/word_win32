import os
import win32com.client as win32
from win32com.client import constants as c

# 启动 Word
word = win32.Dispatch("Word.Application")
word.Visible = True  # 可视化调试
doc = word.Documents.Open(os.path.abspath("source_outlined.docx"))

# 获取样式集合
styles = doc.Styles

# 创建或修改“标题 1”
style1 = styles.Add("my2标题 1", 1)
style1.Font.NameFarEast = "宋体"
style1.Font.NameAscii = "Times New Roman"
style1.Font.Size = 16
style1.ParagraphFormat.LeftIndent = 28
style1.ParagraphFormat.SpaceAfter = 12

# 创建或修改“标题 2”
style2 = styles.Add("my2标题 2", 1)
style2.Font.NameFarEast = "宋体"
style2.Font.NameAscii = "Times New Roman"
style2.Font.Size = 14
style2.ParagraphFormat.LeftIndent = 50
style2.ParagraphFormat.SpaceAfter = 8

# 创建或修改“标题 3”
style3 = styles.Add("my2标题 3", 1)
style3.Font.NameFarEast = "宋体"
style3.Font.NameAscii = "Times New Roman"
style3.Font.Size = 12
style3.ParagraphFormat.LeftIndent = 70
style3.ParagraphFormat.SpaceAfter = 6