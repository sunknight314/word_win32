import os
import json
import win32com.client as win32


def set_outline_levels(doc_path, json_path):
    """
    根据JSON分析结果设置Word文档大纲级别（仅设置大纲级别，不应用样式）
    :param doc_path: Word文档路径
    :param json_path: 包含段落分析的JSON文件路径
    """
    # 1. 加载JSON分析数据
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            analysis_data = json.load(f)
    except FileNotFoundError:
        print(f"[错误] JSON文件不存在: {json_path}")
        return
    except json.JSONDecodeError:
        print(f"[错误] JSON格式无效: {json_path}")
        return

    # 2. 验证数据结构
    if "analysis_result" not in analysis_data:
        print("[错误] JSON缺少analysis_result字段")
        return

    # 3. 创建映射关系（确保级别在0-9有效范围内）[2,5](@ref)
    level_mapping = {
        "Heading1": 1,  # wdOutlineLevel1
        "Heading2": 2,  # wdOutlineLevel2
        "Heading3": 3,  # wdOutlineLevel3
        "Heading4": 4,  # wdOutlineLevel4
        "Normal": 5,  # 正文文本（wdOutlineLevelBodyText）
        "图片": 5  # 按正文处理
    }

    # 4. 启动Word应用
    word_app = win32.DispatchEx("Word.Application")  # 独立进程
    word_app.Visible = False  # 后台运行提高效率

    try:
        # 5. 打开文档
        abs_doc_path = os.path.abspath(doc_path)
        if not os.path.exists(abs_doc_path):
            print(f"[错误] Word文件不存在: {abs_doc_path}")
            return

        doc = word_app.Documents.Open(abs_doc_path)
        paragraphs = doc.Paragraphs
        total_paragraphs = paragraphs.Count
        print(f"文档总段落数: {total_paragraphs}")

        # 6. 遍历JSON设置大纲级别（核心逻辑）
        for item in analysis_data["analysis_result"]:
            para_num = item["paragraph_number"]

            # 跳过超出范围的段落
            if para_num > total_paragraphs or para_num < 1:
                print(f"警告：段落{para_num}超出范围(1-{total_paragraphs})")
                continue

            # 获取映射级别（确保在0-9有效范围）
            level = level_mapping.get(item["type"], 0)
            if level < 0 or level > 9:
                level = 5  # 强制重置为正文级别
                print(f"警告：段落{para_num}的级别超出范围，已重置为5")

            para = paragraphs(para_num)  # Word索引从1开始

            # 7. 仅设置大纲级别（移除了样式应用代码）
            para.OutlineLevel = level
            print(f"段落 {para_num}: 类型={item['type']}, 级别={level}")

        # 8. 保存并关闭
        new_path = os.path.splitext(abs_doc_path)[0] + "_outlined.docx"
        doc.SaveAs(new_path)
        print(f"文档已保存至: {new_path}")
        doc.Close()

    except Exception as e:
        print(f"处理失败: {str(e)}")
    finally:
        # 9. 确保退出Word进程
        word_app.Quit()
        print("Word进程已释放")


# 示例调用
if __name__ == "__main__":
    set_outline_levels(
        doc_path="source.docx",
        json_path="para_type.json"  # 替换为实际JSON路径
    )