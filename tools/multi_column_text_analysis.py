from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent
from tools.utils import ExcelProcessor

class MultiColumnTextAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 1. 获取参数
        llm_model = tool_parameters.get('model_config')
        file_obj = tool_parameters.get('upload_file')
        input_coords = tool_parameters.get('input_columns')
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
        df, is_xlsx, origin_name = ExcelProcessor.load_file(file_obj)
        max_rows = len(df)
        
        # 3. 解析坐标
        try:
            out_info = ExcelProcessor.parse_range(output_coord, max_rows)
            in_infos = ExcelProcessor.get_indices_list(input_coords, max_rows)
        except Exception as e:
            yield self.create_text_message(f"Excel coordinate error: {str(e)}")
            return

        target_rows = range(out_info['start_row'], out_info['end_row'] + 1)

        # 4. 循环处理
        for i in target_rows:
            row_data = []
            for info in in_infos:
                if info['start_row'] <= i <= info['end_row']:
                    try:
                        val = str(df.iat[i, info['col_idx']])
                        row_data.append(val)
                    except IndexError:
                        pass
            
            content_str = " | ".join(row_data)
            if not content_str.strip(): 
                continue

            # 构造 Prompt
            full_content = f"{user_prompt}\n\n[待分析数据]: {content_str}"
            messages = [UserPromptMessage(content=[
                TextPromptMessageContent(type='text', data=full_content)
            ])]

            # === 最终修复：调用模型 ===
            try:
                # 方案 A: 新版 SDK 标准调用 (未来兼容)
                if hasattr(self, 'invoke_model'):
                    response = self.invoke_model(model=llm_model, messages=messages)
                    result = response.message.content
                
                # 方案 B: 针对你当前环境的修复 (仿照 DataSummaryTool)
                # 路径: self.session -> .model -> .llm -> .invoke
                elif hasattr(self, 'session') and hasattr(self.session, 'model'):
                    # 1. 获取 LLM 能力对象
                    llm_service = getattr(self.session.model, 'llm', None)
                    
                    if not llm_service:
                        raise AttributeError(f"ModelInvocations 缺少 'llm' 属性。可用属性: {dir(self.session.model)}")
                    
                    # 2. 调用 invoke
                    # 注意：Excel处理通常设为 stream=False 以直接获取完整结果
                    response = llm_service.invoke(
                        model_config=llm_model,
                        prompt_messages=messages,  # 注意参数名变成了 prompt_messages
                        stream=False              # 关闭流式，直接拿结果
                    )
                    
                    # 3. 解析结果 (非流式返回的是 LLMResult)
                    if hasattr(response, 'message'):
                         result = response.message.content
                    else:
                         # 防御性代码：有些旧版本直接返回 message 对象
                         result = getattr(response, 'content', str(response))

                else:
                    raise AttributeError("无法找到可用的模型调用接口 (invoke_model 或 session.model.llm)")

            except Exception as e:
                # 错误处理
                import traceback
                print(f"Error invoking model: {e}\n{traceback.format_exc()}")
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
