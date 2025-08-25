import win32com.client as win32
from constants.constants import *


def clear_headers_footers(doc):
    """
    清除文档中所有节的页眉和页脚内容
    
    Args:
        doc: Word文档对象
    """
    # 遍历文档中的所有节
    sections = doc.Sections
    for section_idx in range(1, sections.Count + 1):
        section = sections(section_idx)
        
        # 获取所有类型的页眉（首页、奇数页、偶数页）
        for header_type in [WdHeaderFooterIndex.wdHeaderFooterPrimary, 
                           WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                           WdHeaderFooterIndex.wdHeaderFooterEvenPages]:
            try:
                header = section.Headers(header_type)
                # 断开与前一节的链接
                header.LinkToPrevious = False
                # 清除页眉内容
                header.Range.Text = ""
                # 清除页眉中的所有域
                while header.Range.Fields.Count > 0:
                    header.Range.Fields(1).Delete()
            except:
                # 某些页眉类型可能不存在，跳过异常
                pass
        
        # 获取所有类型的页脚（首页、奇数页、偶数页）
        for footer_type in [WdHeaderFooterIndex.wdHeaderFooterPrimary, 
                           WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                           WdHeaderFooterIndex.wdHeaderFooterEvenPages]:
            try:
                footer = section.Footers(footer_type)
                # 断开与前一节的链接
                footer.LinkToPrevious = False
                # 清除页脚内容
                footer.Range.Text = ""
                # 清除页脚中的所有域（包括页码域）
                while footer.Range.Fields.Count > 0:
                    footer.Range.Fields(1).Delete()
                # 清除页码
                while footer.PageNumbers.Count > 0:
                    footer.PageNumbers(1).Delete()
            except:
                # 某些页脚类型可能不存在，跳过异常
                pass
    
    # 更新所有域
    try:
        doc.Fields.Update()
    except:
        pass


def clear_headers_only(doc):
    """
    只清除文档中所有节的页眉内容
    
    Args:
        doc: Word文档对象
    """
    # 遍历文档中的所有节
    sections = doc.Sections
    for section_idx in range(1, sections.Count + 1):
        section = sections(section_idx)
        
        # 获取所有类型的页眉（首页、奇数页、偶数页）
        for header_type in [WdHeaderFooterIndex.wdHeaderFooterPrimary, 
                           WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                           WdHeaderFooterIndex.wdHeaderFooterEvenPages]:
            try:
                header = section.Headers(header_type)
                # 断开与前一节的链接
                header.LinkToPrevious = False
                # 清除页眉内容
                header.Range.Text = ""
                # 清除页眉中的所有域
                while header.Range.Fields.Count > 0:
                    header.Range.Fields(1).Delete()
            except:
                # 某些页眉类型可能不存在，跳过异常
                pass
    
    # 更新所有域
    try:
        doc.Fields.Update()
    except:
        pass


def clear_footers_only(doc):
    """
    只清除文档中所有节的页脚内容
    
    Args:
        doc: Word文档对象
    """
    # 遍历文档中的所有节
    sections = doc.Sections
    for section_idx in range(1, sections.Count + 1):
        section = sections(section_idx)
        
        # 获取所有类型的页脚（首页、奇数页、偶数页）
        for footer_type in [WdHeaderFooterIndex.wdHeaderFooterPrimary, 
                           WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                           WdHeaderFooterIndex.wdHeaderFooterEvenPages]:
            try:
                footer = section.Footers(footer_type)
                # 断开与前一节的链接
                footer.LinkToPrevious = False
                # 清除页脚内容
                footer.Range.Text = ""
                # 清除页脚中的所有域（包括页码域）
                while footer.Range.Fields.Count > 0:
                    footer.Range.Fields(1).Delete()
                # 清除页码
                while footer.PageNumbers.Count > 0:
                    footer.PageNumbers(1).Delete()
            except:
                # 某些页脚类型可能不存在，跳过异常
                pass
    
    # 更新所有域
    try:
        doc.Fields.Update()
    except:
        pass