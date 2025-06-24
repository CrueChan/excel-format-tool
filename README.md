# Excel Format Tool

A Python utility for batch processing and formatting Excel workbooks with automated styling, data validation, and worksheet protection.

![Python Version](https://img.shields.io/badge/python-3.12+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## Features

- **Batch Processing**: Process all Excel files in a specified directory
- **Conditional Formatting**: Apply different color schemes based on column position
  - Header rows: Dark gray background (#D9D9D9)
  - Content rows: Light gray background (#F2F2F2)
  - Editable columns: Yellow background for headers (#FFEB9C)
- **Data Validation**: Add dropdown lists with predefined options for status columns
- **Worksheet Protection**: Lock specific cells while keeping others editable
- **Auto-sizing**: Automatically adjust column widths based on content
- **Text Wrapping**: Enable text wrapping and vertical center alignment

## Requirements

- Python 3.12+
- openpyxl library

## Installation

### Option 1: Using uv (recommended)
```bash
git clone https://github.com/CrueChan/excel-format-tool.git
cd excel-format-tool
uv sync
```

### Option 2: Using pip
```bash
git clone https://github.com/CrueChan/excel-format-tool.git
cd excel-format-tool
pip install -e .
```

### Option 3: Direct installation
```bash
pip install openpyxl
```

## Usage

1. Create a folder named `按部门拆分` in the same directory as the script
2. Place your Excel files (.xlsx) in this folder
3. Run the script:

```bash
python main.py
```

The tool will automatically:
- Find the column containing "是否使用（必填）" (Usage Status Required)
- Apply conditional formatting based on column position
- Add data validation with dropdown options: "使用", "禁用（注销）", "已调离本单位"
- Set worksheet protection with password
- Adjust column widths and enable text wrapping

## File Structure

```
project/
├── main.py                 # Main script
├── pyproject.toml         # Project configuration
├── README.md              # This file
├── .python-version        # Python version specification
├── uv.lock               # Dependency lock file
└── 按部门拆分/            # Excel files directory
    ├── file1.xlsx
    ├── file2.xlsx
    └── ...
```

## Configuration

### Protection Password
The worksheet protection password is set to `E5T647kc`. You can modify this in the `format_workbook()` function.

### Data Validation Options
The dropdown list contains three options:
- 使用 (In Use)
- 禁用（注销） (Disabled/Deregistered)  
- 已调离本单位 (Transferred from Unit)

### Color Scheme
- **Header Gray**: #D9D9D9 (Dark gray for headers)
- **Content Gray**: #F2F2F2 (Light gray for content)
- **Yellow**: #FFEB9C (Yellow for editable column headers)

## API Reference

### `format_workbook(file_path)`
Processes a single Excel workbook with all formatting rules.

**Parameters:**
- `file_path` (str): Path to the Excel file

**Example:**
```python
from main import format_workbook
format_workbook('path/to/your/file.xlsx')
```

### `process_all_files(folder_path)`
Batch processes all Excel files in the specified folder.

**Parameters:**
- `folder_path` (str): Path to the folder containing Excel files

**Example:**
```python
from main import process_all_files
process_all_files('按部门拆分')
```

## Error Handling

The script includes error handling for:
- Missing target columns
- File processing errors
- Invalid file formats

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Notes

- Only processes `.xlsx` files
- Creates backup by overwriting original files
- Requires the presence of a column containing "是否使用（必填）" text
- Sets all worksheets to normal view mode
- Maintains data integrity while applying formatting

## Changelog

### v0.1.0
- Initial release
- Basic Excel formatting functionality
- Batch processing support
- Data validation and worksheet protection