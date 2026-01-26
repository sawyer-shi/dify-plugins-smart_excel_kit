import io
import re
import pandas as pd
from typing import List, Dict, Tuple, Any
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

class ExcelProcessor:
    @staticmethod
    def load_file(file_obj) -> Tuple[pd.DataFrame, Any, bool, str]:
        """
        加载文件，同时返回 pandas DataFrame (用于读取数据) 和 openpyxl Workbook (用于保留格式写入)
        Returns: (df, wb, is_xlsx, filename)
        """
        content = file_obj.blob
        filename = file_obj.filename.lower()
        
        # File type validation
        supported_extensions = ['.csv', '.xlsx']
        if not any(filename.endswith(ext) for ext in supported_extensions):
            raise ValueError(
                f"Unsupported file type: {filename}. "
                f"Only the following formats are supported: {', '.join(supported_extensions)}. "
                f"Please upload a CSV or Excel file."
            )
        
        wb = None
        if filename.endswith('.csv'):
            try:
                df = pd.read_csv(io.BytesIO(content), encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(content), encoding='gbk')
            df = df.fillna("")
            is_xlsx = False
        else:
            # 1. 用 Pandas 读取数据用于分析 (速度快，处理方便)
            df = pd.read_excel(io.BytesIO(content))
            df = df.fillna("")
            
            # 2. 用 openpyxl 读取对象用于写入 (保留原始格式)
            try:
                wb = load_workbook(io.BytesIO(content))
            except Exception:
                wb = None # 如果加载失败，后续会回退到 Pandas 写入模式
            
            is_xlsx = True
            
        return df, wb, is_xlsx, file_obj.filename

    @staticmethod
    def save_file(df: pd.DataFrame, wb: Any, is_xlsx: bool, original_filename: str) -> Tuple[bytes, str]:
        output = io.BytesIO()
        
        base_name = original_filename.rsplit('.', 1)[0]
        new_filename = f"smart_{base_name}.xlsx"

        if is_xlsx and wb is not None:
            # === 方案 A: 使用 openpyxl 直接保存 Workbook 对象 ===
            # 这样可以保留原始文件的所有格式（颜色、字体、公式等）
            wb.save(output)
        else:
            # === 方案 B: CSV 或 Workbook 加载失败时的降级处理 ===
            # 清除自动生成的数字列名
            new_columns = ["" if isinstance(c, int) else c for c in df.columns]
            df.columns = new_columns

            if is_xlsx:
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
            else:
                df.to_csv(output, index=False, encoding='utf-8-sig')
        
        output.seek(0)
        return output.read(), new_filename

    @staticmethod
    def validate_coord_format(coord: str, is_single_col_tool: bool) -> Tuple[bool, str]:
        if not coord or not coord.strip():
            return False, "Column expression cannot be empty."

        coord = coord.strip().upper().replace('：', ':').replace('，', ',')
        is_multi_expr = ',' in coord
        
        if is_single_col_tool and is_multi_expr:
            return False, "格式错误: 单列分析工具不支持多列语法 (如 'A,B')，请使用 'D2' 格式。"

        parts = coord.split(',')
        for part in parts:
            part = part.strip()
            if not re.search(r'[0-9]', part):
                return False, f"格式错误: 表达式 '{part}' 缺少起始行号 (如 'H2')。"

            if not re.match(r'^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$', part):
                return False, f"格式错误: 无法解析 '{part}'。请检查格式 (示例: 'A2' 或 'A2:A10')。"

            if is_single_col_tool and ':' in part:
                sub_parts = part.split(':')
                col_a = re.match(r"([A-Z]+)", sub_parts[0]).group(1)
                col_b = re.match(r"([A-Z]+)", sub_parts[1]).group(1)
                if col_a != col_b: return False, "逻辑错误: 单列工具不支持跨列范围。请使用多列分析工具。"

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