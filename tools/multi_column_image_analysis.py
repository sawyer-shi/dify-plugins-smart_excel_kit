from collections.abc import Generator
from typing import Any
import os
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
        sheet_number = tool_parameters.get('sheet_number', 1)
        output_file_name = tool_parameters.get('output_file_name')

        img_cols = tool_parameters.get('image_columns') or tool_parameters.get('image_column') or ''
        img_cols = str(img_cols).strip()
        output_coord = tool_parameters.get('output_column', '').strip()

        if not isinstance(llm_model, dict): yield self.create_text_message(f"Error: model_config invalid."); return
        if not file_obj: yield self.create_text_message("Error: No file uploaded."); return
        if not isinstance(sheet_number, int) or sheet_number <= 0: yield self.create_text_message("Error: sheet_number must be greater than 0."); return

        is_valid, err_msg = ExcelProcessor.validate_coord_format(img_cols, is_single_col_tool=False)
        if not is_valid: yield self.create_text_message(f"[Input Error] {err_msg}"); return

        is_valid_out, err_msg_out = ExcelProcessor.validate_coord_format(output_coord, is_single_col_tool=True)
        if not is_valid_out: yield self.create_text_message(f"[Output Error] {err_msg_out}"); return

        # === 核心加载 ===
        df, wb, is_xlsx, origin_name, path_in, path_out = ExcelProcessor.load_file_with_copy(file_obj, sheet_number)
        if not output_file_name or not str(output_file_name).strip():
            output_file_name = os.path.splitext(origin_name)[0]
        max_rows = len(df)
        
        try:
            image_map = {}
            if is_xlsx and path_in:
                # 提取上浮图片
                image_map = ExcelProcessor.extract_image_map(path_in)

            try:
                in_infos = ExcelProcessor.get_indices_list(img_cols, max_rows)
                out_info = ExcelProcessor.parse_range(output_coord, max_rows)
            except Exception as e: yield self.create_text_message(f"Excel coordinate error: {str(e)}"); return

            if not in_infos: target_rows = range(0, 0)
            else:
                min_input_row = min(info['start_row'] for info in in_infos)
                max_input_row = max(info['end_row'] for info in in_infos)
                target_rows = range(min_input_row, max_input_row + 1)
            
            if wb and sheet_number <= len(wb.worksheets):
                ws = wb.worksheets[sheet_number - 1]
            else:
                ws = wb.active if (is_xlsx and wb) else None

            for i in target_rows:
                content_list = [TextPromptMessageContent(type='text', data=user_prompt)]
                has_content = False

                for info in in_infos:
                    if info['start_row'] <= i <= info['end_row']:
                        col_idx = info['col_idx']
                        cell_key = (i, col_idx)

                        # A. 上浮图片 (优先)
                        if cell_key in image_map:
                            for b64_img in image_map[cell_key]:
                                meta = ExcelProcessor.get_image_info(b64_img)
                                content_list.append(ImagePromptMessageContent(
                                    type='image', 
                                    url=b64_img,
                                    mime_type=meta['mime_type'],
                                    format=meta['format']
                                ))
                                has_content = True
                        
                        # B. 内容 (URL 或 文本)
                        try:
                            val = str(df.iat[i, col_idx]).strip()
                            if val and val.lower() != 'nan' and val != "":
                                # 如果看起来像 URL，尝试下载
                                if re.match(r'^https?://', val):
                                    b64_dl = ExcelProcessor.download_url_to_base64(val)
                                    if b64_dl:
                                        meta = ExcelProcessor.get_image_info(b64_dl)
                                        content_list.append(ImagePromptMessageContent(
                                            type='image', 
                                            url=b64_dl,
                                            mime_type=meta['mime_type'],
                                            format=meta['format']
                                        ))
                                        has_content = True
                                    else:
                                        # 重点：下载失败也保留文本，防止模型瞎编或拒绝
                                        content_list.append(TextPromptMessageContent(type='text', data=f"\n[Image URL (Process Failed)]: {val}"))
                                        has_content = True
                                else:
                                    content_list.append(TextPromptMessageContent(type='text', data=f"\n[Text Content]: {val}"))
                                    has_content = True
                        except IndexError: pass
                
                # 如果既没有图片也没有文本，跳过
                if not has_content: continue
                
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

                if ws:
                    try: ws.cell(row=i+2, column=out_info['col_idx']+1).value = result
                    except: pass
                else:
                    while out_info['col_idx'] >= len(df.columns): df[len(df.columns)] = ""
                    df.iat[i, out_info['col_idx']] = result

            data, fname = ExcelProcessor.save_output_file(wb, path_out, origin_name, output_file_name)
            mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            yield self.create_blob_message(blob=data, meta={'mime_type': mime_type, 'save_as': fname, 'filename': fname})

        finally:
            ExcelProcessor.clean_paths([path_in, path_out])
