# Privacy Policy / 隐私政策

**Effective Date / 生效日期**: February 11, 2026 / 2026年2月11日
**Version / 版本**: v0.0.1
**Plugin Name / 插件名称**: Smart Excel Kit / 智能多维表
**License / 许可证**: Apache License 2.0

---

## English Version

### Overview

The Smart Excel Kit Plugin ("Plugin") is designed to facilitate local Excel/CSV data analysis, image analysis, and chart generation operations while prioritizing user privacy and data security. This privacy policy explains how we handle your data when you use our plugin within the Dify platform.

### License Information

This plugin is licensed under Apache License 2.0, which is a permissive open-source license that allows for the use, modification, and distribution of software. The full license text is available in the LICENSE file included with this plugin.

### Data Collection

**What we DO NOT collect:**
- Personal identification information
- User account details
- Usage analytics or tracking data
- Device information
- Location data
- Cookies or similar tracking technologies

**What we process:**
- Excel/CSV files you upload for analysis
- Text content within your Excel/CSV files
- Image URLs referenced in your Excel/CSV files
- File metadata (e.g., filename, size, format) required for processing
- Temporary processing data required for file operations

### Data Processing

**Local Processing:**
- All file reading, writing, and processing operations are performed locally on your system
- Excel/CSV files are processed using openpyxl and pandas libraries locally
- No files are uploaded to external servers during the analysis process
- Temporary files are created and deleted securely during processing

**LLM Integration:**
- Text and image data is sent to your configured LLM model for analysis
- Only the specific data you select (text columns, image URLs) is transmitted to the LLM
- No data is sent to any LLM service other than the one you configure
- Chart generation configurations are determined by LLM analysis

### Data Storage

**No Data Retention:**
- We do not store or retain copies of your files after processing
- We do not access or view the contents of your files beyond what is necessary for analysis
- Temporary files are automatically deleted after the processing completes
- No data is transmitted to any third-party services except your configured LLM

### Third-Party Services

**LLM Services:**
- The plugin sends data to the LLM model you configure in Dify
- You have full control over which LLM provider to use
- Data transmission to LLM services follows the LLM provider's privacy policy
- We recommend reviewing your chosen LLM provider's privacy policy

### Security Measures

- **Local-First Architecture**: All file operations are performed locally
- **Secure Temporary Files**: Temporary files are created with restricted permissions
- **No Persistent Storage**: No data is persistently stored by the plugin
- **Memory Safety**: Data is cleared from memory after processing

### User Rights

You have the right to:
- Know what data is being processed
- Control which LLM model is used for analysis
- Delete your files at any time
- Request information about data processing

### Changes to This Policy

We may update this privacy policy from time to time. Any changes will be posted on our GitHub repository with an updated effective date.

### Contact Information

For privacy-related questions or concerns, please visit our GitHub repository:
https://github.com/sawyer-shi/dify-plugins-smart_excel_kit

---

## 中文版本

### 概述

智能多维表插件（"插件"）旨在促进本地 Excel/CSV 数据分析、图片分析和图表生成操作，同时优先考虑用户隐私和数据安全。本隐私政策解释当您在本插件中使用 Dify 平台时，我们如何处理您的数据。

### 许可证信息

本插件采用 Apache License 2.0 许可证，这是一种允许软件使用、修改和分发的宽松开源许可证。完整的许可证文本包含在本插件的 LICENSE 文件中。

### 数据收集

**我们**不会**收集的内容：**
- 个人身份信息
- 用户账户详情
- 使用分析或跟踪数据
- 设备信息
- 位置数据
- Cookie 或类似跟踪技术

**我们处理的内容：**
- 您上传用于分析的 Excel/CSV 文件
- Excel/CSV 文件中的文本内容
- Excel/CSV 文件中引用的图片 URL
- 处理所需的文件元数据（如文件名、大小、格式）
- 文件操作所需的临时处理数据

### 数据处理

**本地处理：**
- 所有文件读取、写入和处理操作都在您的系统上本地执行
- Excel/CSV 文件使用 openpyxl 和 pandas 库在本地处理
- 分析过程中不会将文件上传到外部服务器
- 临时文件在处理期间安全创建和删除

**大模型集成：**
- 文本和图片数据会发送至您配置的大模型进行分析
- 仅传输您选择的特定数据（文本列、图片 URL）至大模型
- 除了您配置的大模型外，不会向任何其他大模型服务发送数据
- 图表生成配置由大模型分析决定

### 数据存储

**无数据保留：**
- 处理后我们不会存储或保留您的文件副本
- 除了分析所需外，我们不会访问或查看您的文件内容
- 临时文件在处理完成后自动删除
- 除了您配置的大模型外，不会向任何第三方服务传输数据

### 第三方服务

**大模型服务：**
- 插件将数据发送至您在 Dify 中配置的大模型
- 您可以完全控制使用哪个大模型提供商
- 向大模型服务传输数据遵循该大模型提供商的隐私政策
- 我们建议您查看所选大模型提供商的隐私政策

### 安全措施

- **本地优先架构**: 所有文件操作都在本地执行
- **安全临时文件**: 临时文件以受限权限创建
- **无持久化存储**: 插件不会持久化存储任何数据
- **内存安全**: 处理后数据从内存中清除

### 用户权利

您有权：
- 了解正在处理的数据内容
- 控制使用哪个大模型进行分析
- 随时删除您的文件
- 请求有关数据处理的信息

### 政策变更

我们可能会不时更新本隐私政策。任何变更将发布在我们的 GitHub 仓库上，并更新生效日期。

### 联系信息

如有隐私相关的问题或疑虑，请访问我们的 GitHub 仓库：
https://github.com/sawyer-shi/dify-plugins-smart_excel_kit
