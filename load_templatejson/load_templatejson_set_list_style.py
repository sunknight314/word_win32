import os
import json
import win32com.client as win32
from constants.constants import *

def load_template_from_json(json_path, docx_path):
    # 读取JSON配置文件
    with open(json_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # 启动Word应用程序
    word = win32.Dispatch('Word.Application')
    word.Visible = False  # 可以设置为True来查看Word操作过程


    # 打开文档
    abs_doc_path = os.path.abspath(docx_path)
    doc = word.Documents.Open(abs_doc_path)

    # 获取样式集合
    styles = doc.Styles

    # 创建新的列表模板
    newListTemplate = doc.ListTemplates.Add(Name="从JSON配置创建的模板")
    newListTemplate.Convert(1)

    # 遍历配置中的标题级别并设置格式
    for i, heading in enumerate(config['headings']):
        level_index = heading['OutlineLevel']  # Word中的级别从1开始

        # 获取对应级别的列表级别对象
        listLevel = newListTemplate.ListLevels(level_index)
        # 设置列表级别属性
        listLevel.NumberFormat = heading['NumberFormat']
        listLevel.TrailingCharacter = heading['TrailingCharacter']
        listLevel.NumberStyle = heading['NumberStyle']
        listLevel.NumberPosition = heading['NumberPosition']
        listLevel.Alignment = heading['Alignment']
        listLevel.TextPosition = heading['TextPosition']
        listLevel.TabPosition = heading['TabPosition']
        listLevel.ResetOnHigher = heading['ResetOnHigher']
        listLevel.StartAt = heading['StartAt']

        # 创建对应样式
        style_name = f"my11标题 {level_index}"

        # 创建新样式
        style = styles.Add(style_name, WdStyleType.wdStyleTypeParagraph)

        # 设置字体属性
        font = heading['Font']
        style.Font.NameFarEast = font['NameFarEast']
        style.Font.NameAscii = font['NameAscii']
        style.Font.NameOther = font['NameOther']
        style.Font.Bold = font['Bold']
        style.Font.Size = font['Size']
        # 注意：颜色需要转换十六进制到Word颜色值
        if isinstance(font['Color'], str) and font['Color'].startswith('#'):
            # 将十六进制颜色转换为RGB值
            color_hex = font['Color'].lstrip('#')
            color_value = int(color_hex, 16)
            # Word使用BGR格式，需要转换
            bgr_color = ((color_value & 0xFF) << 16) | (color_value & 0xFF00) | ((color_value & 0xFF0000) >> 16)
            style.Font.Color = bgr_color
        else:
            style.Font.Color = 0x000000  # 默认黑色

        # 设置段落格式
        para_format = heading['ParagraphFormat']
        style.ParagraphFormat.RightIndent = para_format['RightIndent']
        style.ParagraphFormat.SpaceBefore = para_format['SpaceBefore']
        style.ParagraphFormat.SpaceAfter = para_format['SpaceAfter']
        style.ParagraphFormat.LineSpacing = para_format['LineSpacing']
        style.ParagraphFormat.OutlineLevel = heading['OutlineLevel']

        # 链接列表级别和样式
        listLevel.LinkedStyle = style

    # 应用列表模板到文档中的段落
    for para in doc.Paragraphs:
        ol_level = para.OutlineLevel
        # 注意：Word中的OutlineLevel从1开始，而JSON配置中也是从1开始
        if 1 <= ol_level <= len(config['headings']):
            para.Range.ListFormat.ApplyListTemplateWithLevel(
                newListTemplate
            )

    # 保存并关闭文档
    # doc.Save()
    # doc.Close()

    print(f"成功应用配置到文档: {docx_path}")



if __name__ == "__main__":
    # 使用示例
    json_config_path = os.path.join(os.path.dirname(__file__), 'word_template_config(3).json')
    docx_file_path = 'source_outlined2.docx'  # 需要确保这个文件存在

    if os.path.exists(json_config_path) and os.path.exists(docx_file_path):
        load_template_from_json(json_config_path, docx_file_path)
    else:
        print("错误：找不到配置文件或文档文件")
        print(f"配置文件路径: {json_config_path}")
        print(f"文档文件路径: {os.path.abspath(docx_file_path)}")