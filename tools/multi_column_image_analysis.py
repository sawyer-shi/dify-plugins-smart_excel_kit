from collections.abc import Generator
from typing import Any
import re

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent, ImagePromptMessageContent
from tools.utils import ExcelProcessor

class MultiColumnImageAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        user_prompt = tool_parameters.get('prompt')

        # === 核心修复：兼容参数名 ===
        # 兼容 'image_columns' 和 'image_column'
        img_cols = tool_parameters.get('image_columns') or tool_parameters.get('image_column') or ''
        img_cols = str(img_cols).strip()

        output_coord = tool_parameters.get('output_column', '').strip()
        # ==========================

        if not isinstance(llm_model, dict):
            yield self.create_text_message(f"Error: model_config invalid.")
            return

        if not file_obj:
            yield self.create_text_message("Error: No file uploaded.")
            return

        # 1. 校验
        is_valid, err_msg = ExcelProcessor.validate_coord_format(img_cols, is_single_col_tool=False)
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
            in_infos = ExcelProcessor.get_indices_list(img_cols, max_rows)
            out_info = ExcelProcessor.parse_range(output_coord, max_rows)
        except Exception as e:
            yield self.create_text_message(f"Excel coordinate error: {str(e)}")
            return

        # 2. 边界检查
        for info in in_infos:
            if info['col_idx'] >= len(df.columns):
                 yield self.create_text_message(f"错误: 图片输入列 '{info['col_name']}' 不存在.")
                 return

        if not in_infos:
            target_rows = range(0, 0)
        else:
            min_input_row = min(info['start_row'] for info in in_infos)
            max_input_row = max(info['end_row'] for info in in_infos)
            target_rows = range(min_input_row, max_input_row + 1)

        for i in target_rows:
            content_list = [TextPromptMessageContent(type='text', data=user_prompt)]
            has_img = False

            for info in in_infos:
                if info['start_row'] <= i <= info['end_row']:
                    try:
                        url = str(df.iat[i, info['col_idx']]).strip()
                        if url.startswith(('http', 'https')):
                            content_list.append(ImagePromptMessageContent(type='image', url=url))
                            has_img = True
                    except IndexError:
                        pass
            
            if not has_img: continue
            
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
                else:
                    raise AttributeError("No invoke interface found.")
            except Exception as e:
                result = f"LLM Error: {str(e)}"

            if result and isinstance(result, str):
                result = re.sub(r'<think>.*?</think>', '', result, flags=re.DOTALL)
                result = re.sub(r'<thought>.*?</thought>', '', result, flags=re.DOTALL)
                result = result.strip()

            while out_info['col_idx'] >= len(df.columns): 
                df[len(df.columns)] = ""
            df.iat[i, out_info['col_idx']] = result

        data, fname = ExcelProcessor.save_file(df, is_xlsx, origin_name)
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if is_xlsx else 'text/csv'
        yield self.create_blob_message(blob=data, meta={'mime_type': mime_type, 'save_as': fname})