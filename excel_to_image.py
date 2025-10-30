"""
使用 Spire.XLS 将 Excel 转换为图片
支持三种转换模式：
1. 将整个工作表转换为图片
2. 将工作表转换为无白边图片
3. 将指定单元格区域转换为图片
"""

from spire.xls import *
from spire.xls.common import *
import os


def convert_worksheet_to_image(excel_file, output_file, sheet_index=0):
    """
    将 Excel 工作表转换为图片
    
    参数:
        excel_file: 输入的 Excel 文件路径
        output_file: 输出的图片文件路径
        sheet_index: 工作表索引，默认为 0（第一个工作表）
    """
    try:
        # 创建 Workbook 对象并载入 Excel 文件
        workbook = Workbook()
        workbook.LoadFromFile(excel_file)
        
        # 获取指定的工作表
        sheet = workbook.Worksheets.get_Item(sheet_index)
        
        # 将工作表转换为图片
        image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)
        
        # 保存图片
        image.Save(output_file)
        
        # 释放资源
        workbook.Dispose()
        
        print(f"成功将工作表转换为图片：{output_file}")
        
    except Exception as e:
        print(f"转换失败：{str(e)}")


def convert_worksheet_to_image_no_margin(excel_file, output_file, sheet_index=0):
    """
    将 Excel 工作表转换为无白边图片
    
    参数:
        excel_file: 输入的 Excel 文件路径
        output_file: 输出的图片文件路径
        sheet_index: 工作表索引，默认为 0（第一个工作表）
    """
    try:
        # 创建 Workbook 对象并载入 Excel 文件
        workbook = Workbook()
        workbook.LoadFromFile(excel_file)
        
        # 获取指定的工作表
        sheet = workbook.Worksheets.get_Item(sheet_index)
        
        # 将工作表的所有边距设置为零
        sheet.PageSetup.TopMargin = 0
        sheet.PageSetup.BottomMargin = 0
        sheet.PageSetup.LeftMargin = 0
        sheet.PageSetup.RightMargin = 0
        
        # 将工作表转换为图片
        image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)
        
        # 保存图片
        image.Save(output_file)
        
        # 释放资源
        workbook.Dispose()
        
        print(f"成功将工作表转换为无白边图片：{output_file}")
        
    except Exception as e:
        print(f"转换失败：{str(e)}")


def convert_range_to_image(excel_file, output_file, start_row, start_col, end_row, end_col, sheet_index=0):
    """
    将 Excel 工作表的指定单元格区域转换为图片
    
    参数:
        excel_file: 输入的 Excel 文件路径
        output_file: 输出的图片文件路径
        start_row: 起始行索引（从1开始）
        start_col: 起始列索引（从1开始）
        end_row: 结束行索引
        end_col: 结束列索引
        sheet_index: 工作表索引，默认为 0（第一个工作表）
    """
    try:
        # 创建 Workbook 对象并载入 Excel 文件
        workbook = Workbook()
        workbook.LoadFromFile(excel_file)
        
        # 获取指定的工作表
        sheet = workbook.Worksheets.get_Item(sheet_index)
        
        # 将指定单元格区域转换为图片
        image = sheet.ToImage(start_row, start_col, end_row, end_col)
        
        # 保存图片
        image.Save(output_file)
        
        # 释放资源
        workbook.Dispose()
        
        print(f"成功将单元格区域 (行 {start_row}-{end_row}, 列 {start_col}-{end_col}) 转换为图片：{output_file}")
        
    except Exception as e:
        print(f"转换失败：{str(e)}")


def convert_all_worksheets_to_images(excel_file, output_dir, no_margin=False):
    """
    将 Excel 文件的所有工作表转换为图片
    
    参数:
        excel_file: 输入的 Excel 文件路径
        output_dir: 输出目录
        no_margin: 是否去除白边，默认为 False
    """
    try:
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 创建 Workbook 对象并载入 Excel 文件
        workbook = Workbook()
        workbook.LoadFromFile(excel_file)
        
        # 遍历所有工作表
        for i in range(workbook.Worksheets.Count):
            sheet = workbook.Worksheets.get_Item(i)
            sheet_name = sheet.Name
            
            # 如果需要去除白边
            if no_margin:
                sheet.PageSetup.TopMargin = 0
                sheet.PageSetup.BottomMargin = 0
                sheet.PageSetup.LeftMargin = 0
                sheet.PageSetup.RightMargin = 0
            
            # 将工作表转换为图片
            image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)
            
            # 保存图片
            output_file = os.path.join(output_dir, f"{sheet_name}.png")
            image.Save(output_file)
            
            print(f"成功将工作表 '{sheet_name}' 转换为图片：{output_file}")
        
        # 释放资源
        workbook.Dispose()
        
    except Exception as e:
        print(f"转换失败：{str(e)}")


if __name__ == "__main__":
    # 示例用法
    
    # 输入的 Excel 文件
    input_file = "示例.xlsx"
    
    # 创建输出目录
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    
    print("=== Excel 转图片示例 ===\n")
    
    # 示例 1: 将第一个工作表转换为图片
    print("1. 将整个工作表转换为图片")
    convert_worksheet_to_image(
        input_file, 
        os.path.join(output_dir, "工作表转图片.png"),
        sheet_index=0
    )
    print()
    
    # 示例 2: 将工作表转换为无白边图片
    print("2. 将工作表转换为无白边图片")
    convert_worksheet_to_image_no_margin(
        input_file,
        os.path.join(output_dir, "工作表转无白边图片.png"),
        sheet_index=0
    )
    print()
    
    # 示例 3: 将指定单元格区域转换为图片
    print("3. 将指定单元格区域转换为图片")
    convert_range_to_image(
        input_file,
        os.path.join(output_dir, "单元格范围转图片.png"),
        start_row=10,  # 从第10行开始
        start_col=1,   # 从第1列开始
        end_row=17,    # 到第17行
        end_col=6,     # 到第6列
        sheet_index=0
    )
    print()
    
    # 示例 4: 将所有工作表转换为图片
    print("4. 将所有工作表转换为图片（无白边）")
    convert_all_worksheets_to_images(
        input_file,
        os.path.join(output_dir, "all_sheets"),
        no_margin=True
    )
    print()
    
    print("=== 转换完成 ===")
