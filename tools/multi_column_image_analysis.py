from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent, ImagePromptMessageContent
from tools.utils import ExcelProcessor

class MultiColumnImageAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 1. 获取参数
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        img_cols = tool_parameters.get('image_columns')
        out_col = tool_parameters.get('output_column')
        user_prompt = tool_parameters.get('prompt')

        # 2. 校验参数是否有效
        if not isinstance(llm_model, dict):
            yield self.create_text_message(f"Error: Parameter 'model_config' is invalid. Expected dict, got {type(llm_model)}. Please delete and re-add this node in the workflow.")
            return

        if not file_obj:
            yield self.create_text_message("Error: No file uploaded.")
            return

        # 3. 读取文件
        df, is_xlsx, origin_name = ExcelProcessor.load_file(file_obj)
        max_rows = len(df)
        
        # 3. 解析坐标
        try:
            out_info = ExcelProcessor.parse_range(out_col, max_rows)
            in_infos = ExcelProcessor.get_indices_list(img_cols, max_rows)
        except Exception as e:
            yield self.create_text_message(f"Excel coordinate error: {str(e)}")
            return

        target_rows = range(out_info['start_row'], out_info['end_row'] + 1)

        # 4. 循环处理
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
            
            if not has_img: 
                continue
            
            # 调用模型
            try:
                response = self.invoke_model(model=llm_model, messages=[UserPromptMessage(content=content_list)])
                result = response.message.content
            except Exception as e:
                result = f"LLM Error: {str(e)}"

            # 写入结果
            while out_info['col_idx'] >= len(df.columns): 
                df[len(df.columns)] = ""
            df.iat[i, out_info['col_idx']] = result

        # 5. 保存并返回文件
        data, fname = ExcelProcessor.save_file(df, is_xlsx, origin_name)
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if is_xlsx else 'text/csv'
        yield self.create_blob_message(
            blob=data,
            meta={
                'mime_type': mime_type,
                'save_as': fname
            }
        )
