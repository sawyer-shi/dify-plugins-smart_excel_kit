from collections.abc import Generator
from typing import Any
import re

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent, ImagePromptMessageContent
from tools.utils import ExcelProcessor

class SingleColumnImageAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        img_col = tool_parameters.get('image_column', '').strip()
        out_col = tool_parameters.get('output_column', '').strip()
        user_prompt = tool_parameters.get('prompt')

        if not isinstance(llm_model, dict): yield self.create_text_message(f"Error: model_config invalid."); return
        if not file_obj: yield self.create_text_message("Error: No file uploaded."); return

        is_valid, err_msg = ExcelProcessor.validate_coord_format(img_col, is_single_col_tool=True)
        if not is_valid: yield self.create_text_message(f"[Input Column Error] {err_msg}"); return
        
        is_valid_out, err_msg_out = ExcelProcessor.validate_coord_format(out_col, is_single_col_tool=True)
        if not is_valid_out: yield self.create_text_message(f"[Output Column Error] {err_msg_out}"); return

        # === 接收 wb ===
        df, wb, is_xlsx, origin_name = ExcelProcessor.load_file(file_obj)
        max_rows = len(df)
        
        try:
            in_info = ExcelProcessor.parse_range(img_col, max_rows)
            out_info = ExcelProcessor.parse_range(out_col, max_rows)
        except Exception as e: yield self.create_text_message(f"Excel coordinate error: {str(e)}"); return

        if in_info['col_idx'] >= len(df.columns):
             yield self.create_text_message(f"Error: Image column '{in_info['col_name']}' exceeds file range."); return

        target_rows = range(in_info['start_row'], in_info['end_row'] + 1)
        ws = wb.active if (is_xlsx and wb) else None

        for i in target_rows:
            try:
                url = str(df.iat[i, in_info['col_idx']]).strip()
            except IndexError: continue

            if not url or not url.startswith(('http', 'https')): continue

            content_list = [
                TextPromptMessageContent(type='text', data=user_prompt),
                ImagePromptMessageContent(type='image', url=url)
            ]
            
            try:
                if hasattr(self, 'invoke_model'):
                    response = self.invoke_model(model=llm_model, messages=[UserPromptMessage(content=content_list)])
                    result = response.message.content
                elif hasattr(self, 'session') and hasattr(self.session, 'model'):
                    llm_service = getattr(self.session.model, 'llm', None)
                    if not llm_service: raise AttributeError("No 'llm' service found.")
                    response = llm_service.invoke(model_config=llm_model, prompt_messages=[UserPromptMessage(content=content_list)], stream=False)
                    if hasattr(response, 'message'): result = response.message.content
                    else: result = getattr(response, 'content', str(response))
                else: raise AttributeError("No invoke interface found.")
            except Exception as e: result = f"LLM Error: {str(e)}"

            if result and isinstance(result, str):
                result = re.sub(r'<think>.*?</think>', '', result, flags=re.DOTALL)
                result = re.sub(r'<thought>.*?</thought>', '', result, flags=re.DOTALL)
                result = result.strip()
            
            # === 写入 WB ===
            if ws:
                try: ws.cell(row=i+2, column=out_info['col_idx']+1).value = result
                except: pass
            else:
                while out_info['col_idx'] >= len(df.columns): df[len(df.columns)] = ""
                df.iat[i, out_info['col_idx']] = result
        
        data, fname = ExcelProcessor.save_file(df, wb, is_xlsx, origin_name)
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        yield self.create_blob_message(blob=data, meta={'mime_type': mime_type, 'save_as': fname})