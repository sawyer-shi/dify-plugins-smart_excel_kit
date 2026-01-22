from collections.abc import Generator
from typing import Any
import json
import re

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

        # 表达式校验
        # 单列分析只支持单列表达式（如 "D2" 或 "D2:D10"）
        if ',' in input_coord:
            # 不支持多列表达式
            yield self.create_text_message(
                f"Error: Single column analysis only supports single column expressions (e.g., 'D2' or 'D2:D10'). "
                f"For multiple columns, please use Multi-Column Analysis tool instead."
            )
            return

        # 使用输入列的行范围来确定要分析的行
        target_rows = range(in_info['start_row'], in_info['end_row'] + 1)
        
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

            # ==========================================
            # 清洗思考过程
            # ==========================================
            if result and isinstance(result, str):
                # 移除 DeepSeek 等模型的思考标签 <think>...</think>
                result = re.sub(r'<think>.*?</think>', '', result, flags=re.DOTALL)
                # 移除其他可能的思考标签
                result = re.sub(r'<thought>.*?</thought>', '', result, flags=re.DOTALL)
                # 去除首尾空白
                result = result.strip()
            # ==========================================

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