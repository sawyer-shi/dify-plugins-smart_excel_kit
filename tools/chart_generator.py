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
# 显式引入 TextAxis 来强制显示文本标签
from openpyxl.chart.axis import TextAxis

class ChartGeneratorTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        user_prompt = tool_parameters.get('prompt')
        
        if not isinstance(llm_model, dict): yield self.create_text_message("Error: model_config invalid."); return
        if not file_obj: yield self.create_text_message("Error: No file uploaded."); return

        # 1. 加载文件
        df, wb, is_xlsx, origin_name, path_in, path_out = ExcelProcessor.load_file_with_copy(file_obj)
        
        try:
            # === 数据清理与环境准备 ===
            if df is not None:
                df.dropna(how='all', axis=0, inplace=True)
                df.dropna(how='all', axis=1, inplace=True)
            
            # 处理 CSV 转 Excel 场景
            if not is_xlsx or wb is None:
                wb = Workbook()
                ws = wb.active
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                base_dir = os.path.dirname(path_out)
                new_path_out = os.path.join(base_dir, os.path.splitext(os.path.basename(path_out))[0] + ".xlsx")
                if os.path.exists(path_out): 
                    try: os.remove(path_out)
                    except: pass
                path_out = new_path_out
                origin_name = os.path.splitext(origin_name)[0] + ".xlsx"
            
            ws = wb.active
            # 使用 valid_max_row 确保不包含空行
            valid_max_row = len(df) + 1 
            max_col = ws.max_column

            # === Step 1: 数据画像生成 ===
            col_info = []
            for col in df.columns:
                dtype = df[col].dtype
                is_num = pd.api.types.is_numeric_dtype(dtype)
                col_desc = f"Column '{col}': Type={'Numeric' if is_num else 'Text'}"
                col_info.append(col_desc)
            col_info_str = "\n".join(col_info)
            
            try:
                data_preview = df.head(3).to_string(index=False)
            except:
                data_preview = str(df.head(3))

            # === Step 2: 构造 LLM Prompt ===
            system_instruction = f"""
            You are an expert Excel Chart Architect.
            Data Profile:
            {col_info_str}
            
            Preview:
            {data_preview}
            
            User Request: "{user_prompt}"

            Analyze columns and return JSON to draw the best chart.
            Format:
            {{
                "chart_type": "column" | "bar" | "line" | "pie" | "scatter" | "radar" | "area",
                "title": "Chart Title",
                "x_axis_col": "Column Letter for Categories (e.g. 'A')",
                "y_axis_cols": ["Column Letter(s) for Values", "e.g. 'B'"],
                "reasoning": "Why you chose this"
            }}
            """

            # === Step 3: LLM 决策 ===
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
                yield self.create_text_message(f"AI Failed: {str(e)}")
                return

            try:
                json_str = re.sub(r'```json\s*', '', llm_response)
                json_str = re.sub(r'```', '', json_str)
                config = json.loads(json_str.strip())
            except:
                yield self.create_text_message(f"Invalid AI Config: {llm_response}")
                return

            # === Step 4: 执行绘图 (关键修正部分) ===
            ctype = config.get('chart_type', 'column').lower()
            title = config.get('title', 'Chart')
            x_letter = config.get('x_axis_col', 'A').upper()
            y_letters = config.get('y_axis_cols', [])
            reasoning = config.get('reasoning', '')

            def _c2i(letter):
                idx = 0
                for char in letter: idx = idx * 26 + (ord(char) - ord('A'))
                return idx + 1

            # 工厂模式创建图表对象
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
            
            x_col_idx = _c2i(x_letter)
            # 定义 X 轴数据（从第2行开始，跳过表头）
            x_data = Reference(ws, min_col=x_col_idx, min_row=2, max_row=valid_max_row)

            # === 关键修正: 必须先添加 Y 轴数据 (add_data)，然后再绑定 X 轴分类 (set_categories) ===
            
            # 1. 添加 Y 轴数据
            for y_letter in y_letters:
                y_col_idx = _c2i(y_letter)
                # min_row=1 包含表头作为图例名称
                y_data = Reference(ws, min_col=y_col_idx, min_row=1, max_row=valid_max_row)
                
                if is_xy:
                    # 散点图特殊处理：同时绑定 X 和 Y
                    series = Series(y_data, xvalues=x_data, title_from_data=True)
                    chart.series.append(series)
                else:
                    # 类别图表：先加数据
                    chart.add_data(y_data, titles_from_data=True)

            # 2. 绑定 X 轴分类 (仅针对非 XY 图表)
            if not is_xy:
                chart.set_categories(x_data)
                
                # 3. (可选但推荐) 强制设置 X 轴为文本轴
                # 防止 Excel 把 "2023" 等年份数字识别成连续数值，导致显示位置偏移
                if ctype in ['column', 'bar', 'line', 'area', 'radar']:
                    # 这里不需要重新 new TextAxis，只是确保逻辑正确
                    # chart.x_axis = TextAxis() # 如果 set_categories 还是无效，可以取消这行注释强制覆盖
                    pass

            # 计算插入位置 (数据右侧 + 2列)
            insert_pos = f"{chr(min(max_col + 2 + 65, 90))}2"
            ws.add_chart(chart, insert_pos)
            
            if wb: wb.save(path_out) 
            with open(path_out, 'rb') as f: data = f.read()

            yield self.create_text_message(f"**AI Reasoning:** {reasoning}\n\nChart '{title}' created.")
            yield self.create_blob_message(blob=data, meta={'mime_type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'save_as': origin_name})

        finally:
            ExcelProcessor.clean_paths([path_in, path_out])