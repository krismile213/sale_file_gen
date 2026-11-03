import openpyxl
import pandas as pd
import logging
from pathlib import Path
from typing import Callable

def unmerge_and_fill(file_path, sheet_name, logger:logging.Logger):
    """
    拆分指定工作表中的所有合并单元格，并用原始值填充新单元格。
    
    参数:
    file_path (str): Excel文件的路径。
    sheet_name (str): 需要处理的工作表名称。
    """
    try:
        # 加载工作簿
        workbook = openpyxl.load_workbook(file_path)
        if sheet_name != None:
            sheet = workbook[sheet_name]
        else:
            sheet = workbook.worksheets[0]   
        
        # 存储需要拆分和填充的单元格信息
        merge_cells_to_unmerge = []
        for cell_range in sheet.merged_cells:
            merge_cells_to_unmerge.append(cell_range)
            
        # 逐个处理合并单元格
        for cell_range_string in merge_cells_to_unmerge:
            # 找到合并单元格左上角的单元格，它的值是原始值
            first_cell = sheet[cell_range_string.coord.split(':')[0]]
            merged_value = first_cell.value
            
            # 拆分单元格
            sheet.unmerge_cells(str(cell_range_string))
            
            # 用原始值填充拆分后的所有单元格
            for row in sheet[str(cell_range_string)]:
                for cell in row:
                    cell.value = merged_value
                    
        # 保存为新文件，避免覆盖原始数据
        new_file_path = file_path.replace(".xlsx", "_unmerged.xlsx")
        workbook.save(new_file_path)
        if logger:
            logger.info(f"合并单元格已成功处理，并保存到新文件: {new_file_path}")
        return new_file_path
    except FileNotFoundError:
        if logger:
            logger.error(f"错误: 文件 '{file_path}' 未找到。请检查文件路径。")
    except KeyError:
        if logger:
            logger.error(f"错误: 工作表 '{sheet_name}' 未找到。请检查工作表名称。")
    except Exception as e:
        if logger:
            logger.error(f"发生未知错误: {e}")

def reset_excel_view(file_path, target_sheet_name, output_path):
    # 1. 加载工作簿
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    for sheet_name in wb.sheetnames:
        if sheet_name != target_sheet_name:
            # 获取 sheet 对象并从工作簿中移除
            wb.remove(wb[sheet_name])

    for sheet in wb.worksheets:
        # --- 取消筛选 ---
        # 如果存在自动筛选区域，将其清空
        if sheet.auto_filter.ref:
            sheet.auto_filter.ref = None
            
        # --- 取消隐藏行 ---
        # row_dimensions 存储了所有被修改过属性的行
        for row_idx in sheet.row_dimensions:
            sheet.row_dimensions[row_idx].hidden = False
            
        # --- 取消隐藏列 ---
        for col_idx in sheet.column_dimensions:
            sheet.column_dimensions[col_idx].hidden = False
            
    # 2. 保存结果
    wb.save(output_path)
    print("已成功取消所有筛选和隐藏,且转化了公式值。")

def filter_excel_sheet_with_pandas(
    file_path: str | Path, 
    sheet_name: str | int, 
    condition_func: Callable[[pd.DataFrame], pd.Series]
) -> None:
    """
    使用自定义 Lambda 条件过滤 Excel 指定 Sheet 的数据，并完整保留其他 Sheet。
    
    参数:
    - file_path: Excel 文件路径
    - sheet_name: 要处理的工作表名称
    - condition_func: 一个 Lambda 表达式，接收 df 并返回布尔 Series
    """
    # 1. 读取目标 Sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # 2. 执行 Lambda 过滤条件
    # condition_func(df) 会将整个 df 传入，返回布尔索引进行筛选
    filtered_df = df[condition_func(df)]
    
    # 3. 写回原文件，安全覆盖对应 Sheet 并保留其他 Sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
    print(f"工作表 '{sheet_name}' 过滤完成，已安全覆盖！")