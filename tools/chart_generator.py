import json
import re
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent
from tools.utils import ExcelProcessor

from openpyxl.chart import (
    AreaChart, AreaChart3D,
    BarChart, BarChart3D,
    BubbleChart,
    DoughnutChart,
    LineChart, LineChart3D,
    PieChart, PieChart3D,
    RadarChart,
    ScatterChart,
    SurfaceChart, SurfaceChart3D,
    Reference, Series
)

class ChartGeneratorTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        user_prompt = tool_parameters.get('prompt')

        if not isinstance(llm_model, dict): yield self.create_text_message("Error: model_config invalid."); return
        if not file_obj: yield self.create_text_message("Error: No file uploaded."); return

        df, wb, is_xlsx, origin_name, path_in, path_out = ExcelProcessor.load_file_with_copy(file_obj)
        
        try:
            if not is_xlsx or not wb:
                yield self.create_text_message("Error: Chart generation is only supported for .xlsx files.")
                return

            ws = wb.active
            max_row = ws.max_row
            max_col = ws.max_column

            try:
                data_preview = df.head(3).to_markdown(index=False)
                columns = df.columns.tolist()
                col_index_map = {col: i+1 for i, col in enumerate(columns)}
            except Exception:
                data_preview = "Unable to read dataframe preview."
                columns = []

            chart_types_desc = """
            - 'column': Vertical Bar Chart (Clustered)
            - 'column_stacked': Vertical Bar Chart (Stacked)
            - 'bar': Horizontal Bar Chart
            - 'line': Line Chart
            - 'line_3d': 3D Line Chart
            - 'pie': Pie Chart
            - 'doughnut': Doughnut Chart
            - 'area': Area Chart
            - 'radar': Radar Chart
            - 'scatter': Scatter Chart (XY)
            - 'bubble': Bubble Chart
            - 'surface': Surface Chart
            """

            system_instruction = f"""
            You are an Excel Data Visualization Expert.
            
            User Request: "{user_prompt}"
            
            Data Preview (Top 3 rows):
            {data_preview}
            
            Columns detected: {columns}
            Total Data Rows: {max_row} (Header is row 1, Data starts row 2)

            TASK:
            Analyze the data and the user's intent. Generate a JSON configuration to create the most appropriate Excel chart.
            
            Supported 'chart_type' values:
            {chart_types_desc}

            RETURN JSON FORMAT ONLY (No markdown, no comments):
            {{
                "chart_type": "one of the supported types above",
                "title": "Chart Title String",
                "x_axis_col": "Column letter for Category/X-axis (e.g., 'A')",
                "y_axis_cols": ["List of Column letters for Values/Y-axis", "e.g., 'B'", "e.g., 'C'"],
                "cell_position": "Top-left cell to insert chart (e.g., 'E2')"
            }}
            
            LOGIC RULES:
            1. For 'scatter' or 'bubble' charts, 'x_axis_col' is numerical X-values. For others, it is Categories/Labels.
            2. 'y_axis_cols' must contain numerical columns.
            3. 'cell_position' should be to the right of the data to avoid overlapping.
            """

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
                yield self.create_text_message(f"LLM Config Generation Failed: {str(e)}")
                return

            try:
                json_str = re.sub(r'```json\s*', '', llm_response)
                json_str = re.sub(r'```', '', json_str)
                config = json.loads(json_str.strip())
            except Exception:
                yield self.create_text_message(f"Failed to parse LLM Response. Raw: {llm_response}")
                return

            ctype = config.get('chart_type', 'column').lower()
            title = config.get('title', 'Chart')
            x_letter = config.get('x_axis_col', 'A').upper()
            y_letters = config.get('y_axis_cols', [])
            pos = config.get('cell_position', f"{chr(min(max_col + 66, 90))}2")

            def _c2i(letter):
                idx = 0
                for char in letter: idx = idx * 26 + (ord(char) - ord('A'))
                return idx + 1

            try:
                chart = None
                
                if ctype == 'column':
                    chart = BarChart(); chart.type = "col"; chart.style = 10
                elif ctype == 'column_stacked':
                    chart = BarChart(); chart.type = "col"; chart.grouping = "stacked"; chart.overlap = 100
                elif ctype == 'bar':
                    chart = BarChart(); chart.type = "bar"; chart.style = 10
                elif ctype == 'line':
                    chart = LineChart(); chart.style = 12
                elif ctype == 'line_3d':
                    chart = LineChart3D()
                elif ctype == 'pie':
                    chart = PieChart()
                elif ctype == 'doughnut':
                    chart = DoughnutChart()
                elif ctype == 'area':
                    chart = AreaChart()
                elif ctype == 'radar':
                    chart = RadarChart()
                elif ctype == 'scatter':
                    chart = ScatterChart(); chart.style = 13
                elif ctype == 'bubble':
                    chart = BubbleChart()
                elif ctype == 'surface':
                    chart = SurfaceChart3D()
                else:
                    chart = BarChart()

                chart.title = title
                
                x_col_idx = _c2i(x_letter)
                x_data_ref = Reference(ws, min_col=x_col_idx, min_row=2, max_row=max_row)
                
                for y_letter in y_letters:
                    y_idx = _c2i(y_letter)
                    y_data_ref = Reference(ws, min_col=y_idx, min_row=1, max_row=max_row)
                    
                    if ctype in ['scatter', 'bubble']:
                        series = Series(y_data_ref, xvalues=x_data_ref, title_from_data=True)
                        chart.series.append(series)
                    else:
                        chart.add_data(y_data_ref, titles_from_data=True)

                if ctype not in ['scatter', 'bubble']:
                    chart.set_categories(x_data_ref)

                ws.add_chart(chart, pos)

            except Exception as e:
                yield self.create_text_message(f"Error drawing chart '{ctype}': {str(e)}")
                return

            data, fname = ExcelProcessor.save_output_file(wb, path_out, origin_name)
            mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            
            yield self.create_text_message(f"Generated {ctype} chart '{title}'.")
            yield self.create_blob_message(blob=data, meta={'mime_type': mime_type, 'save_as': fname})

        finally:
            ExcelProcessor.clean_paths([path_in, path_out])
