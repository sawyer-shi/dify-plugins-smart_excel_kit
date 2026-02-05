import json
import re
import os
import pandas as pd
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent
from tools.utils import ExcelProcessor

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import (
    BarChart, LineChart, PieChart, ScatterChart, 
    AreaChart, RadarChart, BubbleChart, DoughnutChart,
    Reference, Series
)
from openpyxl.chart.axis import TextAxis

class ChartGeneratorTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        user_prompt = tool_parameters.get('prompt')
        
        if not isinstance(llm_model, dict): yield self.create_text_message("Error: model_config invalid."); return
        if not file_obj: yield self.create_text_message("Error: No file uploaded."); return

        # 1. 加载数据
        df, wb, is_xlsx, origin_name, path_in, path_out = ExcelProcessor.load_file_with_copy(file_obj)
        
        try:
            # 数据预处理：剔除全空的行和列
            if df is not None:
                df.dropna(how='all', axis=0, inplace=True)
                df.dropna(how='all', axis=1, inplace=True)
            
            # 如果是 CSV 转 Excel 或 wb 未加载成功
            if not is_xlsx or wb is None:
                wb = Workbook()
                ws = wb.active
                # 写入 CSV 数据到 Excel
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                # 更新输出文件名路径为 .xlsx
                base_dir = os.path.dirname(path_out)
                new_path_out = os.path.join(base_dir, os.path.splitext(os.path.basename(path_out))[0] + ".xlsx")
                if os.path.exists(path_out): 
                    try: os.remove(path_out)
                    except: pass
                path_out = new_path_out
                origin_name = os.path.splitext(origin_name)[0] + ".xlsx"
            
            ws = wb.active
            valid_max_row = len(df) + 1 # 表头 + 数据
            max_col = ws.max_column

            # === Step 1: 智能数据画像 (Data Profiling) ===
            col_info = []
            for col in df.columns:
                dtype = df[col].dtype
                is_num = pd.api.types.is_numeric_dtype(dtype)
                unique_count = df[col].nunique()
                example_val = str(df[col].iloc[0]) if len(df) > 0 else ""
                
                col_desc = f"Column '{col}': Type={'Numeric' if is_num else 'Text/Date'}, UniqueValues={unique_count}, Example='{example_val}'"
                col_info.append(col_desc)
            
            col_info_str = "\n".join(col_info)
            
            # [Fix]: 使用 to_string() 替代 to_markdown()，移除对 tabulate 的依赖
            try:
                data_preview = df.head(3).to_string(index=False)
            except Exception:
                data_preview = str(df.head(3))
            
            # === Step 2: 构建“理解者” Prompt ===
            system_instruction = f"""
            You are an expert Data Analyst & Excel Chart Architect.
            Your goal is to fully understand the dataset structure and the user's intent to construct the perfect Excel chart.

            USER INTENT: "{user_prompt}"

            === DATASET PROFILE ===
            {col_info_str}

            === DATA PREVIEW ===
            {data_preview}

            === YOUR REASONING TASK (Internal) ===
            1. Identify the **Category Axis (X-Axis)**:
               - Usually a Text/Date column with valid labels (e.g., Product Name, Month).
               - Avoid ID columns if descriptive names exist.
            2. Identify the **Value Axis (Y-Axis)**:
               - MUST be Numeric columns.
               - If user says "Analyze Sales", pick the 'Sales' column.
               - If user says "Compare everything", pick all meaningful numeric columns.
            3. Choose the **Chart Type**:
               - Comparison -> 'column' (Bar Chart)
               - Trend -> 'line'
               - Part-to-Whole -> 'pie'
               - Correlation -> 'scatter'
            
            === OUTPUT FORMAT ===
            Return strictly valid JSON:
            {{
                "chart_type": "column" | "bar" | "line" | "pie" | "scatter" | "radar" | "area",
                "title": "A descriptive title based on columns selected (e.g. Sales by Region)",
                "x_axis_col": "Letter of the column for Categories (e.g. 'A')",
                "y_axis_cols": ["List of column letters for Values", "e.g. 'B'", "e.g. 'C'"],
                "reasoning": "Short explanation of why you picked these columns"
            }}
            """

            # === Step 3: 调用 LLM ===
            try:
                if hasattr(self, 'invoke_model'):
                    response = self.invoke_model(model=llm_model, messages=[UserPromptMessage(content=[TextPromptMessageContent(type='text', data=system_instruction)])])
                    llm_response = response.message.content
                elif hasattr(self, 'session'):
                    llm_service = self.session.model.llm
                    response = llm_service.invoke(model_config=llm_model, prompt_messages=[UserPromptMessage(content=[TextPromptMessageContent(type='text', data=system_instruction)])], stream=False)
                    llm_response = response.message.content
                else: 
                     raise AttributeError("No invoke interface")
            except Exception as e:
                yield self.create_text_message(f"AI Analysis Failed: {str(e)}")
                return

            # === Step 4: 执行绘图 ===
            try:
                json_str = re.sub(r'```json\s*', '', llm_response)
                json_str = re.sub(r'```', '', json_str)
                config = json.loads(json_str.strip())
            except:
                yield self.create_text_message(f"Failed to parse AI plan: {llm_response}")
                return

            ctype = config.get('chart_type', 'column').lower()
            title = config.get('title', 'Chart')
            x_letter = config.get('x_axis_col', 'A').upper()
            y_letters = config.get('y_axis_cols', [])
            reasoning = config.get('reasoning', '')

            # 插入位置：计算数据右侧空地
            insert_col_idx = max_col + 2
            
            def get_col_str(n):
                string = ""
                while n > 0:
                    n, remainder = divmod(n - 1, 26)
                    string = chr(65 + remainder) + string
                return string
            pos = f"{get_col_str(insert_col_idx)}2"

            def _c2i(letter):
                idx = 0
                for char in letter: idx = idx * 26 + (ord(char) - ord('A'))
                return idx + 1

            chart = None
            if ctype == 'column': chart = BarChart(); chart.type = "col"; chart.style = 10
            elif ctype == 'bar': chart = BarChart(); chart.type = "bar"; chart.style = 10
            elif ctype == 'line': chart = LineChart(); chart.style = 12
            elif ctype == 'pie': chart = PieChart()
            elif ctype == 'scatter': chart = ScatterChart(); chart.style = 13
            elif ctype == 'radar': chart = RadarChart()
            elif ctype == 'area': chart = AreaChart()
            elif ctype == 'doughnut': chart = DoughnutChart()
            else: chart = BarChart()

            chart.title = title
            
            is_xy = ctype in ['scatter', 'bubble']
            
            try:
                x_col_idx = _c2i(x_letter)
                # 校验列是否存在
                if x_col_idx > max_col + 5: raise ValueError(f"X-Axis Column {x_letter} is out of bounds")
                
                x_data = Reference(ws, min_col=x_col_idx, min_row=2, max_row=valid_max_row)

                if not is_xy and ctype in ['column', 'bar', 'line', 'area', 'radar']:
                    chart.x_axis = TextAxis()
                    chart.set_categories(x_data)
                elif ctype in ['pie', 'doughnut']:
                    chart.set_categories(x_data)

                for y_letter in y_letters:
                    y_col_idx = _c2i(y_letter)
                    # 校验
                    if y_col_idx > max_col + 5: continue
                    
                    y_data = Reference(ws, min_col=y_col_idx, min_row=1, max_row=valid_max_row)
                    
                    if is_xy:
                        series = Series(y_data, xvalues=x_data, title_from_data=True)
                        chart.series.append(series)
                    else:
                        chart.add_data(y_data, titles_from_data=True)
                
                ws.add_chart(chart, pos)
                
            except Exception as e:
                yield self.create_text_message(f"Error constructing chart: {str(e)}")
                # 不中断，继续返回文件（只包含数据转换）
            
            if wb: wb.save(path_out) 
            with open(path_out, 'rb') as f: data = f.read()

            yield self.create_text_message(f"**AI Reasoning:** {reasoning}\n\nChart '{title}' generated.")
            yield self.create_blob_message(blob=data, meta={'mime_type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'save_as': origin_name})

        finally:
            ExcelProcessor.clean_paths([path_in, path_out])