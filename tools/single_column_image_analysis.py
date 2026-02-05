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
        sheet_number = tool_parameters.get('sheet_number', 1)

        if not isinstance(llm_model, dict): yield self.create_text_message(f"Error: model_config invalid."); return
        if not file_obj: yield self.create_text_message("Error: No file uploaded."); return
        if not isinstance(sheet_number, int) or sheet_number <= 0: yield self.create_text_message("Error: sheet_number must be greater than 0."); return

        # 校验参数
        is_valid, err_msg = ExcelProcessor.validate_coord_format(img_col, is_single_col_tool=True)
        if not is_valid: yield self.create_text_message(f"[Input Error] {err_msg}"); return
        
        is_valid_out, err_msg_out = ExcelProcessor.validate_coord_format(out_col, is_single_col_tool=True)
        if not is_valid_out: yield self.create_text_message(f"[Output Error] {err_msg_out}"); return

        # === 核心修改：使用副本模式加载 ===
        df, wb, is_xlsx, origin_name, path_in, path_out = ExcelProcessor.load_file_with_copy(file_obj, sheet_number)
        max_rows = len(df)
        
        try:
            image_map = {}
            if is_xlsx and path_in:
                # 从原文件(path_in)提取上浮图片
                image_map = ExcelProcessor.extract_image_map(path_in)

            try:
                in_info = ExcelProcessor.parse_range(img_col, max_rows)
                out_info = ExcelProcessor.parse_range(out_col, max_rows)
            except Exception as e: yield self.create_text_message(f"Excel coordinate error: {str(e)}"); return

            if in_info['col_idx'] >= len(df.columns):
                yield self.create_text_message(f"Error: Column '{in_info['col_name']}' exceeds file range."); return

            target_rows = range(in_info['start_row'], in_info['end_row'] + 1)
            if wb and sheet_number <= len(wb.worksheets):
                ws = wb.worksheets[sheet_number - 1]
            else:
                ws = wb.active if (is_xlsx and wb) else None

            for i in target_rows:
                # 构建消息列表：Prompt + (Images / URL-Images / Text)
                content_list = [TextPromptMessageContent(type='text', data=user_prompt)]
                has_content = False
                
                col_idx = in_info['col_idx']
                cell_key = (i, col_idx)

                # A. 检查这里是否有上浮图片
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
                
                # B. 检查单元格内的 URL 或文本
                try:
                    val = str(df.iat[i, col_idx]).strip()
                    if val and val.lower() != 'nan' and val != "":
                        # 尝试判断是否为 URL
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
                                # 下载失败，作为纯文本 URL 放入上下文，防止 info 丢失
                                content_list.append(TextPromptMessageContent(type='text', data=f"\n[Image URL (Process Failed)]: {val}"))
                                has_content = True
                        else:
                            # 纯文本内容
                            content_list.append(TextPromptMessageContent(type='text', data=f"\n[Content]: {val}"))
                            has_content = True
                except IndexError: pass

                # 只有当没有任何图片且没有任何文本内容时才跳过
                if not has_content: continue
                
                # 调用模型
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

                # 清理
                if result and isinstance(result, str):
                    result = re.sub(r'<think>.*?</think>', '', result, flags=re.DOTALL)
                    result = re.sub(r'<thought>.*?</thought>', '', result, flags=re.DOTALL)
                    result = result.strip()
                
                # 写入到副本
                if ws:
                    try: ws.cell(row=i+2, column=out_info['col_idx']+1).value = result
                    except: pass
                else:
                    while out_info['col_idx'] >= len(df.columns): df[len(df.columns)] = ""
                    df.iat[i, out_info['col_idx']] = result
            
            # 保存副本
            data, fname = ExcelProcessor.save_output_file(wb, path_out, origin_name)
            mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            yield self.create_blob_message(blob=data, meta={'mime_type': mime_type, 'save_as': fname})

        finally:
            ExcelProcessor.clean_paths([path_in, path_out])