from collections.abc import Generator
from typing import Any
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent
from tools.utils import ExcelProcessor

class SingleColumnTextAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # --- 调试：如果遇到问题，可以先打印参数类型 ---
        # print(f"DEBUG Params: {tool_parameters.keys()}")
        
        # 1. 安全获取参数
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        input_coord = tool_parameters.get('input_column')
        output_coord = tool_parameters.get('output_column')
        user_prompt = tool_parameters.get('prompt')

        # 2. 校验参数是否有效
        if not isinstance(llm_model, dict):
            yield self.create_text_message(f"Error: Parameter 'model_config' is invalid. Expected dict, got {type(llm_model)}. Please delete and re-add this node in the workflow.")
            return
        
        if not file_obj:
            yield self.create_text_message("Error: No file uploaded.")
            return

        # 3. 读取文件
        try:
            df, is_xlsx, original_name = ExcelProcessor.load_file(file_obj)
        except Exception as e:
            yield self.create_text_message(f"Error loading file: {str(e)}")
            return
            
        max_rows = len(df)
        
        # 4. 解析坐标
        try:
            rows_range = ExcelProcessor.parse_range(output_coord, max_rows)
            in_info = ExcelProcessor.parse_range(input_coord, max_rows)
        except Exception as e:
            yield self.create_text_message(f"Excel coordinate error: {str(e)}")
            return

        target_rows = range(rows_range['start_row'], rows_range['end_row'] + 1)
        
        # 5. 循环处理
        for i in target_rows:
            try:
                content = str(df.iat[i, in_info['col_idx']])
            except IndexError:
                content = ""

            if not content or content.lower() == 'nan': 
                continue

            full_content = f"{user_prompt}\n\n[待分析内容]:\n{content}"
            messages = [UserPromptMessage(content=[
                TextPromptMessageContent(type='text', data=full_content)
            ])]
            
            # 调用模型
            try:
                # 这里的 llm_model 应当是字典，直接传给 invoke_model
                response = self.invoke_model(model=llm_model, messages=messages)
                result = response.message.content
            except Exception as e:
                # 捕获模型调用错误
                result = f"LLM Error: {str(e)}"

            while rows_range['col_idx'] >= len(df.columns): 
                df[len(df.columns)] = ""
            
            df.iat[i, rows_range['col_idx']] = result

        # 6. 保存并返回
        data, fname = ExcelProcessor.save_file(df, is_xlsx, original_name)
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if is_xlsx else 'text/csv'
        yield self.create_blob_message(
            blob=data,
            meta={
                'mime_type': mime_type,
                'save_as': fname
            }
        )