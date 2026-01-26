import io
import re
import os
import pandas as pd
from typing import List, Dict, Tuple, Any
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

class ExcelProcessor:
    @staticmethod
    def load_file(file_obj) -> Tuple[pd.DataFrame, Any, bool, str]:
        """
        加载文件，同时返回:
        1. pandas DataFrame (用于读取数据)
        2. openpyxl Workbook (用于保留格式写入, CSV则为None)
        3. is_xlsx 标记
        4. 原始文件名 (用于后续生成输出文件名)
        """
        content = file_obj.blob
        
        # === 1. 获取真实文件名 (处理 Dify 内部可能存在的 Hash 文件名问题) ===
        # 优先级: original_filename > filename > name
        original_filename = None
        
        # 尝试查找包含真实文件名的属性
        candidates = ['original_filename', 'upload_filename', 'filename', 'name']
        for attr in candidates:
            if hasattr(file_obj, attr):
                val = getattr(file_obj, attr)
                if val and isinstance(val, str):
                    # 简单的过滤逻辑：有些系统会把 hash 放在 filename 里，
                    # 如果只有 hash 没有扩展名，大概率不是我们要展示的文件名
                    if '.' in val: 
                        original_filename = val
                        break
        
        # 兜底
        if not original_filename:
            original_filename = getattr(file_obj, 'filename', 'unknown.xlsx')

        # === 2. 扩展名校验 ===
        # os.path.splitext 能正确处理中文文件名
        _, ext = os.path.splitext(original_filename)
        ext = ext.lower()
        
        supported_extensions = ['.csv', '.xlsx', '.xls']
        if ext not in supported_extensions:
            raise ValueError(
                f"Unsupported file type: {original_filename}. "
                f"Only CSV or Excel files are supported."
            )
        
        wb = None
        is_xlsx = False

        # === 3. 文件加载 ===
        if ext == '.csv':
            try:
                # 优先尝试 UTF-8，失败则尝试 GBK (常见于中文 CSV)
                df = pd.read_csv(io.BytesIO(content), encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(content), encoding='gbk')
            df = df.fillna("")
        else:
            # Excel 文件
            is_xlsx = True
            # Pandas 用于数据处理
            df = pd.read_excel(io.BytesIO(content))
            df = df.fillna("")
            
            # Openpyxl 用于格式保留 (仅支持 .xlsx, 不支持老旧 .xls)
            if ext == '.xlsx':
                try:
                    wb = load_workbook(io.BytesIO(content))
                except Exception:
                    wb = None # 加载失败降级处理
            
        return df, wb, is_xlsx, original_filename

    @staticmethod
    def save_file(df: pd.DataFrame, wb: Any, is_xlsx: bool, original_filename: str) -> Tuple[bytes, str]:
        output = io.BytesIO()
        
        # === 修正后的文件名生成逻辑 ===
        # 目标：data.xlsx -> smart_data.xlsx
        # 目标：data.csv  -> smart_data.xlsx
        # 目标：中文.xlsx -> smart_中文.xlsx
        
        # 1. 去除原有扩展名，只取文件名部分
        # os.path.splitext 会自动处理 "路径/文件名.扩展名"
        base_name, _ = os.path.splitext(original_filename)
        
        # 2. 如果文件名包含路径，只取最后的文件名部分 (防御性编程)
        base_name = os.path.basename(base_name)
        
        # 3. 拼接新文件名
        new_filename = f"smart_{base_name}.xlsx"

        # === 保存逻辑 ===
        if is_xlsx and wb is not None:
            # 方案 A: 有 Workbook 对象，直接保存 (保留原格式)
            wb.save(output)
        else:
            # 方案 B: 重写文件 (CSV转XLSX 或 无法加载Workbook的情况)
            # 清除 pandas 自动生成的 index 列名 (如 8, 9)
            new_columns = ["" if isinstance(c, int) else c for c in df.columns]
            df.columns = new_columns

            # 强制保存为 xlsx 格式 (即使原文件是 csv)
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
        
        output.seek(0)
        return output.read(), new_filename

    @staticmethod
    def validate_coord_format(coord: str, is_single_col_tool: bool) -> Tuple[bool, str]:
        if not coord or not coord.strip():
            return False, "Column expression cannot be empty."

        coord = coord.strip().upper().replace('：', ':').replace('，', ',')
        is_multi_expr = ',' in coord
        
        if is_single_col_tool and is_multi_expr:
            return False, "Format error: Single column analysis tool does not support multi-column syntax (e.g., 'A,B'). Please use 'D2' format."

        parts = coord.split(',')
        for part in parts:
            part = part.strip()
            if not re.search(r'[0-9]', part):
                return False, f"Format error: Expression '{part}' is missing starting row number (e.g., 'H2')."

            if not re.match(r'^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$', part):
                return False, f"Format error: Unable to parse '{part}'. Please check the format (examples: 'A2' or 'A2:A10')."

            if is_single_col_tool and ':' in part:
                sub_parts = part.split(':')
                col_a = re.match(r"([A-Z]+)", sub_parts[0]).group(1)
                col_b = re.match(r"([A-Z]+)", sub_parts[1]).group(1)
                if col_a != col_b: return False, "Logic error: Single column tool does not support cross-column ranges. Please use multi-column analysis tool."

        return True, ""

    @staticmethod
    def parse_range(range_str: str, max_rows: int) -> Dict[str, Any]:
        range_str = range_str.upper().strip().replace('：', ':')
        parts = range_str.split(':')
        
        def parse_single(s):
            match = re.match(r"([A-Z]+)([0-9]+)", s)
            if not match: return 0, 0
            col_str = match.group(1)
            row_num = int(match.group(2))
            
            col_idx = 0
            for char in col_str: col_idx = col_idx * 26 + (ord(char) - ord('A')) + 1
            col_idx -= 1
            
            row_idx = max(0, row_num - 2) 
            return col_idx, row_idx

        start_col, start_row = parse_single(parts[0])
        
        if len(parts) > 1:
            end_col, end_row_raw = parse_single(parts[1])
            end_row = min(end_row_raw, max_rows - 1)
        else:
            end_col = start_col
            end_row = max_rows - 1

        return {
            'col_idx': start_col,
            'start_row': start_row,
            'end_row': end_row,
            'col_name': re.match(r"([A-Z]+)", parts[0]).group(1)
        }

    @staticmethod
    def get_indices_list(coord_str: str, max_rows: int) -> List[Dict]:
        coord_str = coord_str.replace('，', ',').strip()
        return [ExcelProcessor.parse_range(c.strip(), max_rows) for c in coord_str.split(',')]