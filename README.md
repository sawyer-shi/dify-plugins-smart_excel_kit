# Smart Excel Kit

A powerful Dify plugin providing comprehensive **local** Excel/CSV data analysis and visualization capabilities. All operations are performed entirely on your local machine without requiring any external services beyond your chosen LLM, ensuring maximum data security and privacy. Supports text analysis, image analysis, and chart generation powered by your preferred LLM.

## Version Information

- **Current Version**: v0.0.1
- **Release Date**: 2026-01-24
- **Compatibility**: Dify Plugin Framework
- **Python Version**: 3.12

### Version History
- **v0.0.1** (2026-01-24): Initial release with text analysis, image analysis, and chart generation capabilities

## Quick Start

1. Download smart_excel_kit plugin from Dify marketplace
2. Install plugin in your Dify environment
3. Configure your preferred LLM model
4. Start analyzing your Excel/CSV data immediately

## Key Features

- **100% Local Processing**: All data processing operations are performed entirely on your local machine
- **Flexible LLM Integration**: Use your preferred LLM model for analysis and visualization
- **No External Data Transmission**: Your data never leaves your local environment except to your chosen LLM
- **Maximum Data Security**: Complete privacy and security for your sensitive data
- **Multiple Analysis Types**: Text analysis, image analysis, and chart generation in one tool
- **Excel Native Output**: All results are output as Excel files for seamless integration
- **Batch Processing**: Analyze multiple columns or rows in a single operation

## Core Features

### Text Analysis

#### Single Column Text Analysis (single_column_text_analysis)
Analyze text data in a single column using custom prompts and LLM.
- **Supported Formats**: Excel (.xlsx, .xls), CSV
- **Features**:
  - Custom prompt-based analysis
  - Flexible column selection (single cell or range)
  - Results written directly to specified output column
  - Batch processing support
  - Preserves original file structure

#### Multi Column Text Analysis (multi_column_text_analysis)
Analyze and correlate text data across multiple columns using custom prompts.
- **Supported Formats**: Excel (.xlsx, .xls), CSV
- **Features**:
  - Multi-column joint analysis
  - Custom prompt-based correlation analysis
  - Flexible column selection with range support
  - Results written to specified output column
  - Batch processing for multiple rows

### Image Analysis

#### Single Column Image Analysis (single_column_image_analysis)
Analyze images referenced by URLs in a single column using vision models.
- **Supported Formats**: Excel (.xlsx, .xls), CSV with image URLs
- **Features**:
  - Vision model-based image analysis
  - Support for multiple image formats (PNG, JPG, JPEG, WEBP, etc.)
  - Custom prompt-based image description
  - Flexible column selection
  - Results written to specified output column

#### Multi Column Image Analysis (multi_column_image_analysis)
Analyze and correlate multiple images across columns using vision models.
- **Supported Formats**: Excel (.xlsx, .xls), CSV with image URLs
- **Features**:
  - Multi-image joint analysis
  - Vision model-based correlation analysis
  - Custom prompt-based comparison
  - Flexible multi-column selection
  - Results written to specified output column

### Chart Generation

#### Chart Generator (chart_generator)
Automatically generate various types of Excel charts based on your data and LLM analysis.
- **Supported Chart Types**:
  - Column Chart (Clustered, Stacked)
  - Bar Chart (Horizontal)
  - Line Chart (2D, 3D)
  - Pie Chart
  - Doughnut Chart
  - Area Chart
  - Radar Chart
  - Scatter Chart (XY)
  - Bubble Chart
  - Surface Chart (3D)
- **Features**:
  - LLM-powered chart type selection
  - Automatic data range detection
  - Smart chart positioning
  - Multiple data series support
  - Custom chart titles and labels
  - Native Excel chart objects

### Data Manipulation

#### Excel Manipulator (excel_manipulator)
Use AI to intelligently modify, clean, or calculate Excel data using Python Pandas operations.
- **Supported Operations**:
  - Data cleaning and filtering
  - Row/column manipulation
  - Calculated columns creation
  - Data deduplication
  - Custom data transformations
- **Features**:
  - Natural language instructions
  - AI-generated Python Pandas code
  - Full Pandas library support
  - Preserves Excel file structure
  - Custom output filename support
  - Safe code execution environment

## Technical Advantages

- **Local Processing**: All file operations are performed locally, ensuring data privacy
- **LLM Flexibility**: Compatible with any LLM model supported by Dify
- **Native Excel Integration**: Uses openpyxl for seamless Excel file manipulation
- **Image URL Support**: Automatically downloads and processes images from URLs
- **Smart Data Detection**: Automatically detects data types and ranges
- **Error Handling**: Comprehensive error handling with user-friendly messages
- **Memory Efficient**: Optimized for large files with streaming processing

## Use Cases

### Content Analysis
- Analyze customer feedback or reviews in Excel columns
- Summarize long text content using LLM
- Extract key information from text data

### Image Processing
- Batch analyze product images from URLs
- Compare multiple images and generate descriptions
- Extract visual information from image datasets

### Data Visualization
- Generate sales charts from Excel data
- Create radar charts for performance comparison
- Build scatter plots for correlation analysis
- Design pie charts for data distribution

## Privacy & Security

- **Local-First Architecture**: All file processing happens locally
- **No Data Retention**: We don't store or retain copies of your files
- **LLM Only**: Data is only sent to your configured LLM model
- **No Third-Party Services**: No external APIs or services required
- **Secure Temporary Files**: Temporary files are securely deleted after processing

## License

This plugin is licensed under Apache License 2.0. See LICENSE file for details.

## Support

For issues, feature requests, or contributions, please visit our GitHub repository:
https://github.com/sawyer-shi/dify-plugins-smart_excel_kit
