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
        return df, is_xlsx, file_obj.filename

    @staticmethod
    def save_file(df: pd.DataFrame, is_xlsx: bool, original_filename: str) -> Tuple[bytes, str]:
        output = io.BytesIO()
        prefix = "analyzed_"
        new_filename = f"{prefix}{original_filename}"

        if is_xlsx:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
        else:
            df.to_csv(output, index=False, encoding='utf-8-sig')
        
        output.seek(0)
        return output.read(), new_filename

    @staticmethod
    def parse_range(range_str: str, max_rows: int) -> Dict[str, Any]:
        """
        Parse Excel coordinate range expression
        
        Supported formats:
        - Single column range: "D2" means column D from row 2 to the last row
        - Column range: "D2:D10" means column D from row 2 to row 10
        - Multi-column range: "A2,B2" means column A from row 2, column B from row 2 (two full columns)
        - Multi-column range: "A2:A12,B2:B12" means column A from row 2 to row 12, column B from row 2 to row 12
        
        Returns: {
            'col_idx': column index (0-based),
            'start_row': start row number (0-based),
            'end_row': end row number (0-based, exclusive)
        }
        """
        range_str = range_str.upper().strip()
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
            return col_idx, max(0, row_num - 2)

        start_col, start_row = parse_single(parts[0])
        
        if len(parts) > 1:
            end_col, end_row = parse_single(parts[1])
        else:
            end_col, end_row = start_col, max_rows - 1

        return {
            'col_idx': start_col,
            'start_row': start_row,
            'end_row': min(end_row, max_rows - 1)
        }

    @staticmethod
    def get_indices_list(coord_str: str, max_rows: int) -> List[Dict]:
        """
        Parse multiple coordinate range expressions
        
        Supported formats:
        - "A2,B2" - multiple independent ranges, each can be a single column or column range
        - "A2:A12,B2:B12" - multiple ranges with row limits
        
        Returns: list of range dictionaries
        """
        return [ExcelProcessor.parse_range(c.strip(), max_rows) for c in coord_str.split(',')]
