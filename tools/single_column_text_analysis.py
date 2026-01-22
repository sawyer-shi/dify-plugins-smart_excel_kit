from collections.abc import Generator
from typing import Any
import re

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent
from tools.utils import ExcelProcessor

class SingleColumnTextAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        input_coord = tool_parameters.get('input_column', '').strip()
        output_coord = tool_parameters.get('output_column', '').strip()
        user_prompt = tool_parameters.get('prompt')

        if not isinstance(llm_model, dict):
            yield self.create_text_message(f"Error: model_config invalid.")
            return

        if not file_obj:
            yield self.create_text_message("Error: No file uploaded.")
            return

        # 1. 严格语法校验 (禁止纯字母)
        is_valid, err_msg = ExcelProcessor.validate_coord_format(input_coord, is_single_col_tool=True)
        if not is_valid:
            yield self.create_text_message(f"[输入列错误] {err_msg}")
            return
        
        is_valid_out, err_msg_out = ExcelProcessor.validate_coord_format(output_coord, is_single_col_tool=True)
        if not is_valid_out:
            yield self.create_text_message(f"[输出列错误] {err_msg_out}")
            return

        df, is_xlsx, origin_name = ExcelProcessor.load_file(file_obj)
        max_rows = len(df)
        
        try:
            in_info = ExcelProcessor.parse_range(input_coord, max_rows)
            out_info = ExcelProcessor.parse_range(output_coord, max_rows)
        except Exception as e:
            yield self.create_text_message(f"Coordinate Parse Error: {str(e)}")
            return

        # 2. 严格数据边界校验 (检查 HHH 是否存在)
        if in_info['col_idx'] >= len(df.columns):
            yield self.create_text_message(
                f"错误: 您请求分析列 '{in_info['col_name']}'，但上传的文件只有 {len(df.columns)} 列。\n"
                f"Error: Column '{in_info['col_name']}' does not exist in the file."
            )
            return

        target_rows = range(in_info['start_row'], in_info['end_row'] + 1)
        
        for i in target_rows:
            try:
                content = str(df.iat[i, in_info['col_idx']])
            except IndexError:
                content = "" # 如果行超出了，当做空处理

            if not content or content.lower() == 'nan' or content.strip() == "": 
                continue

            full_content = f"{user_prompt}\n\n[待分析内容]:\n{content}"
            messages = [UserPromptMessage(content=[TextPromptMessageContent(type='text', data=full_content)])]

            try:
                if hasattr(self, 'invoke_model'):
                    response = self.invoke_model(model=llm_model, messages=messages)
                    result = response.message.content
                elif hasattr(self, 'session') and hasattr(self.session, 'model'):
                    llm_service = getattr(self.session.model, 'llm', None)
                    if not llm_service: raise AttributeError("No 'llm' service found.")
                    response = llm_service.invoke(model_config=llm_model, prompt_messages=messages, stream=False)
                    if hasattr(response, 'message'): result = response.message.content
                    else: result = getattr(response, 'content', str(response))
                else:
                    raise AttributeError("No invoke interface found.")
            except Exception as e:
                import traceback
                print(f"LLM Error: {e}")
                result = f"LLM Error: {str(e)}"

            if result and isinstance(result, str):
                result = re.sub(r'<think>.*?</think>', '', result, flags=re.DOTALL)
                result = re.sub(r'<thought>.*?</thought>', '', result, flags=re.DOTALL)
                result = result.strip()

            # 扩展列以保证输出位置存在
            while out_info['col_idx'] >= len(df.columns): 
                df[len(df.columns)] = ""
            
            df.iat[i, out_info['col_idx']] = result

        data, fname = ExcelProcessor.save_file(df, is_xlsx, origin_name)
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if is_xlsx else 'text/csv'
        yield self.create_blob_message(blob=data, meta={'mime_type': mime_type, 'save_as': fname})