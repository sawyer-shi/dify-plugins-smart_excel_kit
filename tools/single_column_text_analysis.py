from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.entities.model.message import UserPromptMessage, TextPromptMessageContent
from tools.utils import ExcelProcessor

class SingleColumnTextAnalysisTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # 1. 获取参数
        llm_model = tool_parameters['llm_model']
        file_obj = tool_parameters['upload_file']
        input_coord = tool_parameters['input_column']
        output_coord = tool_parameters['output_column']
        user_prompt = tool_parameters['prompt']

        # 2. 读取文件 (使用 utils.py 中的逻辑)
        # 注意：ExcelProcessor 需要能处理 blob 数据
        df, is_xlsx, original_name = ExcelProcessor.load_file(file_obj)
        max_rows = len(df)
        
        # 3. 解析坐标
        try:
            rows_range = ExcelProcessor.parse_range(output_coord, max_rows)
            in_info = ExcelProcessor.parse_range(input_coord, max_rows)
        except Exception as e:
            # 返回错误文本提示
            yield self.create_text_message(f"Excel coordinate error: {str(e)}")
            return

        target_rows = range(rows_range['start_row'], rows_range['end_row'] + 1)
        
        # 4. 循环处理
        for i in target_rows:
            try:
                # 获取数据
                content = str(df.iat[i, in_info['col_idx']])
            except IndexError:
                content = ""

            # 跳过空值或 nan
            if not content or content.lower() == 'nan': 
                continue

            # 构造 Prompt
            full_content = f"{user_prompt}\n\n[待分析内容]:\n{content}"
            messages = [UserPromptMessage(content=[
                TextPromptMessageContent(type='text', data=full_content)
            ])]
            
            # 调用模型
            try:
                # invoke_model 返回的是 LLMResult，需要获取 message.content
                response = self.invoke_model(model=llm_model, messages=messages)
                result = response.message.content
            except Exception as e:
                result = f"LLM Error: {str(e)}"

            # 写入结果 (扩展列以防越界)
            while rows_range['col_idx'] >= len(df.columns): 
                df[len(df.columns)] = ""
            
            df.iat[i, rows_range['col_idx']] = result

        # 5. 保存并返回文件
        data, fname = ExcelProcessor.save_file(df, is_xlsx, original_name)
        
        # 使用 yield 返回 Blob 消息
        yield self.create_blob_message(blob=data, meta={'mime_type': 'application/octet-stream'}, save_as=fname)