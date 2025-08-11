import os
import win32com.client as win32
from constants.constants import *

word = win32.Dispatch('Word.Application')

docx_path = 'source_outlined.docx'
abs_doc_path = os.path.abspath(docx_path)
doc = word.Documents.Open(abs_doc_path)
# 获取样式集合
styles = doc.Styles

newListTemplate = doc.ListTemplates.Add(Name="新列表模板2")
newListTemplate.Convert(1)


# ---------------------------------------------------------------------
# 设置列表级别1的格式
listLevel1 = newListTemplate.ListLevels(1)
listLevel1.NumberFormat = "第%1章"
listLevel1.TrailingCharacter = TrailingCharacter.wdTrailingSpace
listLevel1.NumberStyle = NumberStyle.wdListNumberStyleSimpChinNum1
listLevel1.NumberPosition = 32
listLevel1.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
listLevel1.TextPosition = 0
listLevel1.TabPosition = 9999999
listLevel1.ResetOnHigher = 0
listLevel1.StartAt = 1

style1 = styles.Add("my7标题 1", 1)
style1.Font.NameFarEast = "黑体"
style1.Font.NameAscii = "Times New Roman"
style1.Font.NameOther = "Times New Roman"
style1.Font.Bold = False
style1.Font.Size = 16
style1.Font.Color = 0x000000  # 黑色

style1.ParagraphFormat.RightIndent = 0
style1.ParagraphFormat.SpaceBefore = 24
style1.ParagraphFormat.SpaceAfter = 18
style1.ParagraphFormat.LineSpacing = 20
style1.ParagraphFormat.OutlineLevel = 1

listLevel1.LinkedStyle = style1

# ---------------------------------------------------------------------
# 设置列表级别2的格式
listLevel2 = newListTemplate.ListLevels(2)
listLevel2.NumberFormat = "%1.%2"
listLevel2.TrailingCharacter = TrailingCharacter.wdTrailingSpace
listLevel2.NumberStyle = NumberStyle.wdListNumberStyleLegal
listLevel2.NumberPosition = 0
listLevel2.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
listLevel2.TextPosition = 0
listLevel2.TabPosition = 9999999
listLevel2.ResetOnHigher = 1
listLevel2.StartAt = 1

style2 = styles.Add("my7标题 2", 1)
style2.Font.NameFarEast = "宋体"
style2.Font.NameAscii = "Times New Roman"
style2.Font.NameOther = "Times New Roman"
style2.Font.Bold = True
style2.Font.Size = 15
style1.Font.Color = 0x000000  # 黑色

style2.ParagraphFormat.RightIndent = 0
style2.ParagraphFormat.SpaceBefore = 18
style2.ParagraphFormat.SpaceAfter = 12
style2.ParagraphFormat.LineSpacing = 20
style2.ParagraphFormat.OutlineLevel = 2
listLevel2.LinkedStyle = style2



# ---------------------------------------------------------------------
# 设置列表级别3的格式
listLevel3 = newListTemplate.ListLevels(3)
listLevel3.NumberFormat = "%1.%2.%3"
listLevel3.TrailingCharacter = TrailingCharacter.wdTrailingSpace
listLevel3.NumberStyle = NumberStyle.wdListNumberStyleLegal
listLevel3.NumberPosition = 28
listLevel3.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
listLevel3.TextPosition = 0
listLevel3.TabPosition = 9999999
listLevel3.ResetOnHigher = 2
listLevel3.StartAt = 1

style3 = styles.Add("my7标题 3", 1)
style3.Font.NameFarEast = "宋体"
style3.Font.NameAscii = "Times New Roman"
style3.Font.NameOther = "Times New Roman"
style3.Font.Bold = True
style3.Font.Size = 14

style3.ParagraphFormat.RightIndent = 0
style3.ParagraphFormat.SpaceBefore = 12
style3.ParagraphFormat.SpaceAfter = 6
style3.ParagraphFormat.LineSpacing = 20
style3.ParagraphFormat.OutlineLevel = 3
listLevel3.LinkedStyle = style3


for para in doc.Paragraphs:
    ol_level = para.OutlineLevel  # 0-based for Heading 1 = 0, Heading 2 = 1, etc.
    if 1 <= ol_level <= 3:  # 只处理第1~3级标题
        para.Range.ListFormat.ApplyListTemplateWithLevel(
            newListTemplate
        )
