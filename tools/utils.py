import io
import re
import pandas as pd
from typing import List, Dict, Tuple, Any

class ExcelProcessor:
    @staticmethod
    def load_file(file_obj) -> Tuple[pd.DataFrame, bool, str]:
        content = file_obj.blob
        filename = file_obj.filename.lower()
        
        # File type validation
        supported_extensions = ['.csv', '.xlsx', '.xls']
        if not any(filename.endswith(ext) for ext in supported_extensions):
            raise ValueError(
                f"Unsupported file type: {filename}. "
                f"Only the following formats are supported: {', '.join(supported_extensions)}. "
                f"Please upload a CSV or Excel file."
            )
        
        if filename.endswith('.csv'):
            try:
                df = pd.read_csv(io.BytesIO(content), encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(content), encoding='gbk')
            is_xlsx = False
        else:
            df = pd.read_excel(io.BytesIO(content))
            is_xlsx = True
            
        # 填充 nan 为空字符串，防止处理时报错
        df = df.fillna("")
        return df, is_xlsx, file_obj.filename

    @staticmethod
    def save_file(df: pd.DataFrame, is_xlsx: bool, original_filename: str) -> Tuple[bytes, str]:
        output = io.BytesIO()
        prefix = "analyzed_"
        new_filename = f"{prefix}{original_filename}"

        # === 修复：清除自动生成的数字列名 ===
        # Pandas 新增列时默认使用整数索引 (如 8) 作为列名
        # 我们在这里把所有整数类型的列名替换为空字符串，防止在 Excel 第一行显示 "8"
        new_columns = []
        for col in df.columns:
            if isinstance(col, int):
                new_columns.append("") # 将数字标题改为空白
            else:
                new_columns.append(col)
        df.columns = new_columns
        # =================================

        if is_xlsx:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # index=False 表示不写入行号(0,1,2...)，但默认会写入列名(Header)
                df.to_excel(writer, index=False)
        else:
            df.to_csv(output, index=False, encoding='utf-8-sig')
        
        output.seek(0)
        return output.read(), new_filename

    @staticmethod
    def validate_coord_format(coord: str, is_single_col_tool: bool) -> Tuple[bool, str]:
        """
        表达式最严校验：必须包含行号
        """
        if not coord or not coord.strip():
            return False, "Column expression cannot be empty."

        coord = coord.strip().upper().replace('：', ':').replace('，', ',')
        is_multi_expr = ',' in coord
        
        if is_single_col_tool and is_multi_expr:
            return False, (
                f"格式错误: 单列分析工具不支持多列语法 '{coord}'。\n"
                f"请使用 'D2' 或 'D2:D10' 格式。"
            )

        parts = coord.split(',')
        for part in parts:
            part = part.strip()
            # 规则1: 必须包含数字 (行号)。拒绝 "HHH", "I", "A:B"
            if not re.search(r'[0-9]', part):
                return False, (
                    f"格式错误: 表达式 '{part}' 缺少起始行号。\n"
                    f"请明确指定开始行，例如 'H2' (代表H列从第2行开始) 或 'I1'。"
                )

            # 规则2: 正则严格匹配 [字母][数字] optionally [: [字母][数字]]
            if not re.match(r'^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$', part):
                return False, f"格式错误: 无法解析 '{part}'。请检查格式 (示例: 'A2' 或 'A2:A10')。"

            # 规则3: 校验冒号左右是否同一列 (仅针对单列工具)
            if is_single_col_tool and ':' in part:
                sub_parts = part.split(':')
                col_a = re.match(r"([A-Z]+)", sub_parts[0]).group(1)
                col_b = re.match(r"([A-Z]+)", sub_parts[1]).group(1)
                if col_a != col_b:
                    return False, f"逻辑错误: 单列工具不支持跨列范围 '{part}'。请使用多列分析工具。"

        return True, ""

    @staticmethod
    def parse_range(range_str: str, max_rows: int) -> Dict[str, Any]:
        """
        解析 Excel 坐标范围
        """
        range_str = range_str.upper().strip().replace('：', ':')
        parts = range_str.split(':')
        
        def parse_single(s):
            match = re.match(r"([A-Z]+)([0-9]+)", s)
            if not match: return 0, 0
            col_str = match.group(1)
            row_num = int(match.group(2))
            
            col_idx = 0
            for char in col_str:
                col_idx = col_idx * 26 + (ord(char) - ord('A')) + 1
            col_idx -= 1
            
            row_idx = max(0, row_num - 2) 
            return col_idx, row_idx

        start_col, start_row = parse_single(parts[0])
        
        if len(parts) > 1:
            end_col, end_row_raw = parse_single(parts[1])
            # 修正: 用户输入的范围是 inclusive 的，但 range() 也是 inclusive 处理逻辑在外面
            # 这里只需确保不超过 max_rows
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