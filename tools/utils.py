import io
import re
import os
import requests
import base64
import tempfile
import shutil
import mimetypes
import pandas as pd
from typing import List, Dict, Tuple, Any, Optional
from openpyxl import load_workbook

class ExcelProcessor:
    @staticmethod
    def load_file_with_copy(file_obj, sheet_number: int = 1) -> Tuple[pd.DataFrame, Any, bool, str, str, str]:
        """
        核心加载逻辑：
        Args:
            file_obj: 文件对象
            sheet_number: sheet页码（从1开始）
        Returns: (df, wb, is_xlsx, best_filename, input_path, output_path)
        """
        content = file_obj.blob
        
        # 1. 解析文件名
        candidates = []
        seen = set()
        common_attrs = ['original_filename', 'upload_filename', 'filename', 'name']
        for attr in common_attrs:
            if hasattr(file_obj, attr):
                val = getattr(file_obj, attr)
                if isinstance(val, str) and val not in seen:
                    candidates.append(val)
                    seen.add(val)
        try:
            if hasattr(file_obj, '__dict__'):
                val = file_obj.__dict__.get('filename')
                if isinstance(val, str) and val not in seen:
                    candidates.append(val)
                    seen.add(val)
        except: pass

        uuid_pattern = re.compile(r'^[a-fA-F0-9-]{32,36}\.(xlsx|csv|xls)$')
        best_filename = None
        for name in candidates:
            base_name = os.path.basename(name)
            if not any(base_name.lower().endswith(ext) for ext in ['.xlsx', '.csv', '.xls']): continue
            if uuid_pattern.match(base_name): continue
            best_filename = base_name; break
        
        if not best_filename: best_filename = "output.xlsx"
        _, ext = os.path.splitext(best_filename)
        ext = ext.lower()

        # 2. 创建物理文件
        fd_in, temp_input_path = tempfile.mkstemp(suffix=ext)
        with os.fdopen(fd_in, 'wb') as f:
            f.write(content)
            
        # Output: 副本
        fd_out, temp_output_path = tempfile.mkstemp(suffix=ext)
        os.close(fd_out)
        shutil.copy2(temp_input_path, temp_output_path)
        
        wb = None
        is_xlsx = False
        df = None

        try:
            if ext == '.csv':
                try: df = pd.read_csv(temp_input_path, encoding='utf-8')
                except UnicodeDecodeError: df = pd.read_csv(temp_input_path, encoding='gbk')
                df = df.fillna("")
            else:
                is_xlsx = True
                # header=0 意味着第一行是表头
                df = pd.read_excel(temp_input_path, header=0, sheet_name=sheet_number - 1)
                df = df.fillna("")
                
                if ext in ['.xlsx', '.xlsm']:
                    try: 
                        # data_only=False 保证读到公式本身而不是结果，便于写入保留
                        wb = load_workbook(temp_output_path)
                    except Exception as e: 
                        print(f"Workbook load error: {e}")
                        wb = None
                
        except Exception as e:
            ExcelProcessor.clean_paths([temp_input_path, temp_output_path])
            raise e

        return df, wb, is_xlsx, best_filename, temp_input_path, temp_output_path

    @staticmethod
    def load_file(file_obj):
        """兼容旧调用的适配器，防止 AttributeError"""
        df, wb, is_xlsx, best_filename, p_in, p_out = ExcelProcessor.load_file_with_copy(file_obj)
        # 旧接口返回5个参数，其中最后一个通常是 temp_path。
        # 我们返回 output_path 作为 temp_path
        # 注意：使用此旧接口将无法手动清理 input_path，可能会有极其微小的临时文件残留，建议更新调用方
        return df, wb, is_xlsx, best_filename, p_out

    @staticmethod
    def extract_image_map(temp_file_path: str) -> Dict[Tuple[int, int], List[str]]:
        """
        强力版图片提取：兼容多种 Anchor 格式
        """
        image_map = {}
        if not temp_file_path or not os.path.exists(temp_file_path): return image_map
        if not temp_file_path.lower().endswith('.xlsx'): return image_map # 仅支持 xlsx

        wb_temp = None
        try:
            wb_temp = load_workbook(temp_file_path, data_only=False)
            ws = wb_temp.active
            
            # 尝试获取所有可能的图片列表
            images = getattr(ws, '_images', []) or getattr(ws, 'images', [])
            
            for img in images:
                try:
                    # === 核心：坐标确立 ===
                    anchor_row = None
                    anchor_col = None

                    # 策略1: 尝试访问 .anchor._from (OneCellAnchor / TwoCellAnchor)
                    if hasattr(img, 'anchor'):
                        a = img.anchor
                        # 兼容不同版本的属性名
                        marker = getattr(a, '_from', None) or getattr(a, 'from', None)
                        
                        if marker:
                            anchor_col = marker.col
                            anchor_row = marker.row
                    
                    # 如果策略1失败，且没有任何坐标信息，则跳过
                    if anchor_row is None or anchor_col is None:
                        continue

                    # === 坐标系转换 ===
                    # OpenPyxl Row 0 = Excel Row 1 (Header)
                    # OpenPyxl Row 1 = Excel Row 2 (Data Row 0)
                    # Pandas Index = OpenPyxl Row - 1
                    
                    pd_row_idx = anchor_row - 1
                    
                    if pd_row_idx < 0: continue # 图片在表头或更上面

                    # === 图片数据提取 ===
                    img_data = None
                    
                    # 方式 A: ref (通常是 BytesIO)
                    if hasattr(img, 'ref') and hasattr(img.ref, 'read'):
                        img.ref.seek(0)
                        img_data = img.ref.read()
                    # 方式 B: fp (文件路劲或流)
                    elif hasattr(img, 'fp') and hasattr(img.fp, 'read'):
                        img.fp.seek(0)
                        img_data = img.fp.read()
                    # 方式 C: _data 闭包
                    elif hasattr(img, '_data'):
                        try: img_data = img._data()
                        except: pass
                    
                    if not img_data: continue

                    # === 格式与 Base64 ===
                    fmt = "png"
                    if hasattr(img, 'format') and img.format:
                        fmt = img.format.lower()
                    
                    # 过滤不支持的格式
                    if fmt in ['emf', 'wmf', 'vml']: continue
                    if fmt not in ['png', 'jpg', 'jpeg', 'bmp', 'gif', 'webp']: fmt = 'png'

                    b64 = base64.b64encode(img_data).decode('utf-8')
                    data_uri = f"data:image/{fmt};base64,{b64}"
                    
                    key = (pd_row_idx, anchor_col)
                    if key not in image_map: image_map[key] = []
                    image_map[key].append(data_uri)

                except Exception as e:
                    print(f"Warning: Single image extract failed: {e}")
                    continue

        except Exception as e:
            print(f"Extract image map failed: {e}")
        finally:
            if wb_temp: wb_temp.close()
            
        return image_map

    @staticmethod
    def download_url_to_base64(url: str) -> Optional[str]:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
        }
        try:
            # timeout防止卡死，verify=False防止内网证书报错
            response = requests.get(url, headers=headers, timeout=10, verify=False)
            if response.status_code == 200:
                content_type = response.headers.get('Content-Type', '').lower()
                
                # 宽松判断，防止服务器不返回标准 image 类型
                if 'image' in content_type or any(url.lower().endswith(x) for x in ['.jpg','.png','.jpeg','.webp']):
                    b64 = base64.b64encode(response.content).decode('utf-8')
                    
                    mime = content_type if 'image' in content_type else mimetypes.guess_type(url)[0]
                    if not mime: mime = 'image/png'
                    
                    return f"data:{mime};base64,{b64}"
        except Exception: pass
        return None

    @staticmethod
    def get_image_info(base64_str: str) -> Dict[str, str]:
        try:
            if base64_str.startswith('data:'):
                h = base64_str.split(';')[0]
                m = h.split(':')[1]
                f = m.split('/')[1]
                return {"mime_type": m, "format": f}
        except: pass
        return {"mime_type": "image/png", "format": "png"}

    @staticmethod
    def save_output_file(wb: Any, output_path: str, original_filename: str, custom_file_name: str = None) -> Tuple[bytes, str]:
        base_name, _ = os.path.splitext(original_filename)
        if custom_file_name and custom_file_name.strip():
            new_filename = f"{custom_file_name.strip()}.xlsx"
        else:
            new_filename = f"smart_{base_name}.xlsx"

        if wb:
            try: wb.save(output_path)
            except: pass
            
        with open(output_path, 'rb') as f:
            data = f.read()
        return data, new_filename

    @staticmethod
    def clean_paths(paths: List[str]):
        for p in paths:
            if p and os.path.exists(p):
                try: os.remove(p)
                except: pass
    
    @staticmethod
    def parse_range(range_str: str, max_rows: int) -> Dict[str, Any]:
        parts = range_str.upper().strip().replace('：', ':').split(':')
        def parse(s):
            m = re.match(r"([A-Z]+)([0-9]+)", s)
            if not m: return 0, 0
            c_str, r_num = m.groups()
            c_idx = 0
            for char in c_str: c_idx = c_idx * 26 + (ord(char) - ord('A')) + 1
            return c_idx - 1, max(0, int(r_num) - 2)
        sc, sr = parse(parts[0])
        ec, er = parse(parts[1]) if len(parts) > 1 else (sc, max_rows - 1)
        return {'col_idx': sc, 'start_row': sr, 'end_row': min(er, max_rows - 1), 'col_name': re.match(r"([A-Z]+)", parts[0]).group(1)}

    @staticmethod
    def get_indices_list(coord_str: str, max_rows: int) -> List[Dict]:
        return [ExcelProcessor.parse_range(c.strip(), max_rows) for c in coord_str.replace('，', ',').split(',')]

    @staticmethod
    def validate_coord_format(coord: str, is_single_col_tool: bool) -> Tuple[bool, str]:
        if not coord or not coord.strip(): return False, "Column cannot be empty."
        coord = coord.strip().upper().replace('：', ':').replace('，', ',')
        if is_single_col_tool and ',' in coord: return False, "Single column tool does not support multiple columns."
        parts = coord.split(',')
        for part in parts:
            if not re.match(r'^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$', part.strip()): return False, f"Invalid coords: {part}"
            if is_single_col_tool and ':' in part:
                 if part.split(':')[0][0] != part.split(':')[1][0]: return False, "Single col tool cannot span columns."
        return True, ""
