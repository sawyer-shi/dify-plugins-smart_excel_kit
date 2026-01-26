from collections.abc import Generator
from typing import Any
import re

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent
from tools.utils import ExcelProcessor

class MultiColumnTextAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        user_prompt = tool_parameters.get('prompt')

        input_coords = tool_parameters.get('input_columns') or tool_parameters.get('input_column') or ''
        input_coords = str(input_coords).strip()
        output_coord = tool_parameters.get('output_column', '').strip()

        if not isinstance(llm_model, dict): yield self.create_text_message(f"Error: model_config invalid."); return
        if not file_obj: yield self.create_text_message("Error: No file uploaded."); return

        is_valid, err_msg = ExcelProcessor.validate_coord_format(input_coords, is_single_col_tool=False)
        if not is_valid: yield self.create_text_message(f"[Input Column Error] {err_msg}"); return

        is_valid_out, err_msg_out = ExcelProcessor.validate_coord_format(output_coord, is_single_col_tool=True)
        if not is_valid_out: yield self.create_text_message(f"[Output Column Error] {err_msg_out}"); return

        # === 接收 wb ===
        df, wb, is_xlsx, origin_name = ExcelProcessor.load_file(file_obj)
        max_rows = len(df)
        
        try:
            in_infos = ExcelProcessor.get_indices_list(input_coords, max_rows)
            out_info = ExcelProcessor.parse_range(output_coord, max_rows)
        except Exception as e: yield self.create_text_message(f"Excel coordinate error: {str(e)}"); return

        for info in in_infos:
            if info['col_idx'] >= len(df.columns):
                yield self.create_text_message(f"Error: Input column '{info['col_name']}' does not exist (exceeds maximum column count)."); return

        if not in_infos: target_rows = range(0, 0)
        else:
            min_input_row = min(info['start_row'] for info in in_infos)
            max_input_row = max(info['end_row'] for info in in_infos)
            target_rows = range(min_input_row, max_input_row + 1)
        
        ws = wb.active if (is_xlsx and wb) else None

        for i in target_rows:
            row_data = []
            for info in in_infos:
                if info['start_row'] <= i <= info['end_row']:
                    try:
                        val = str(df.iat[i, info['col_idx']])
                        row_data.append(val)
                    except IndexError: pass
            
            content_str = " | ".join(row_data)
            if not content_str.strip(): continue

            full_content = f"{user_prompt}\n\n[Data to analyze]: {content_str}"
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