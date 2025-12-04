# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**IneoXlsx v2.2.1** - Professional XML to Excel converter plugin developed by Ineo Solutions S.L.

This is a command-line tool that converts structured XML files into Excel (.xlsx) files with comprehensive styling, table formatting, column/row configuration, and advanced professional features.

## Core Commands

### Development
```bash
# Install dependencies
pip install -r requirements.txt

# Run the converter
python ineoXlsxCmdLine.py <input.xml> [output.xlsx]

# Example
python ineoXlsxCmdLine.py task/ineo_xlsx_small.xml output.xlsx
```

### Building Executable
```bash
# Build standalone .exe using PyInstaller
pyinstaller ineoXlsxCmdLine.spec

# The executable will be in dist/ineoXlsxCmdLine.exe
```

### Testing
```bash
# Run with test file
python ineoXlsxCmdLine.py task/ineo_xlsx_small.xml

# Test with debug logging enabled (XML must have log section)
# Check ineoXlsxCmdLine.log for detailed logs
```

## Architecture

### Main Entry Point
- **ineoXlsxCmdLine.py**: Command-line interface and main orchestration
  - Handles command-line arguments
  - Validates XML against XSD schema (schema.xsd)
  - Detects XML encoding automatically (UTF-8, ISO-8859-1, Windows-1252)
  - Sets up logging configuration
  - Delegates conversion to excel_funciones_exportacion.py

### Core Modules

**excel/excel_funciones_exportacion.py**
- Main conversion logic: `xml_to_excel()` function
- Processes XML structure and generates Excel workbooks
- Implements precedence system: Cell > Row > Column > Default
- Handles styles, alignment, formatting, tables
- Manages column width auto-adjustment with configurable limits
- Supports non-destructive updates to existing Excel files

**ineoXlsxGlobales.py**
- Global configuration constants
- Version information (INEOXLSXCMDLINE_VERSION, INEOXLSXCMDLINE_REVISION)
- Program metadata (PROGRAM_NAME, COMPANY_NAME)
- Execution timestamp management

**logging/ineoxlsx_logging.py**
- Logging utilities (if needed for extended logging)

### XML Schema Validation
- **schema.xsd**: Defines valid XML structure
- Validates before processing to catch errors early
- Supports two root elements: `<workbooks>` (legacy) and `<ineoDoc>` (new format)

### Key Design Patterns

1. **URI Prefixes**: Supports FILE://, BASE64://, URL:// prefixes for input/output
   - `FILE://path/to/file` - Local file
   - `BASE64://encoded_content` - Base64-encoded XML (input only)
   - `URL://https://...` - Remote URL (input only, not implemented for output)

2. **Precedence System**: Property application follows strict order
   - Cell properties override all
   - Row properties override Column and Default
   - Column properties override Default
   - Default values as fallback

3. **Non-Destructive Updates**: When Excel file exists
   - Reuses existing sheets when specified
   - Creates only new sheets as needed
   - Preserves all existing content
   - Applies only new/updated content from XML

4. **Auto-Adjustment Logic**
   - Global options (`<options>`) set defaults for all workbooks
   - Per-workbook attribute `autoAdjustColumnWidth` overrides global setting
   - Columns with explicit `width` are excluded from auto-adjustment
   - Configurable min/max width constraints

## XML Structure

### Complete Format (ineoDoc)
```xml
<?xml version="1.0" encoding="UTF-8"?>
<ineoDoc task="updateXlsx" task_id="unique_id">
    <data>
        <dataIn>FILE://input.xml</dataIn>
        <dataOut>FILE://output.xlsx</dataOut>
    </data>
    <log>
        <logLevel>DEBUG</logLevel>
        <logFile>FILE://./conversion.log</logFile>
        <logConsole>true</logConsole>
    </log>
    <options>
        <option name="autoAdjustColumnWidth" value="true"/>
        <option name="maxColumnWidth" value="50"/>
        <option name="minColumnWidth" value="8"/>
    </options>
    <workbooks>
        <styles>
            <!-- Style definitions -->
        </styles>
        <workbook name="SheetName" autoAdjustColumnWidth="true">
            <columnSettings>
                <!-- Column configurations -->
            </columnSettings>
            <rowSettings>
                <!-- Row configurations -->
            </rowSettings>
            <table name="TableName" ref="A1:E10" style="TableStyleMedium9"
                   showRowStripes="true" showColumnStripes="false" />
            <!-- Cell definitions -->
        </workbook>
    </workbooks>
</ineoDoc>
```

### Legacy Format (workbooks only)
```xml
<?xml version="1.0" encoding="UTF-8"?>
<workbooks>
    <styles>
        <style id="1">
            <font>Arial</font>
            <bold>true</bold>
        </style>
    </styles>
    <workbook name="Sheet1">
        <cell row="1" column="A" value="Data" style="1"/>
    </workbook>
</workbooks>
```

## Important Implementation Details

### Encoding Detection
The system automatically detects XML encoding from the `<?xml encoding="..." ?>` declaration. This is critical for processing files created in different locales (ANSI, UTF-8, etc.).

**Functions:**
- `detect_xml_encoding()` in both ineoXlsxCmdLine.py and excel_funciones_exportacion.py
- Uses binary mode to read first 200 bytes
- Falls back to UTF-8 if no encoding declaration found

### Excel Table Styles
60 predefined table styles available:
- TableStyleLight1-21 (subtle, formal)
- TableStyleMedium1-28 (balanced, most popular)
- TableStyleDark1-11 (high contrast)

Tables automatically include filter dropdowns in headers.

### Column Width Auto-Adjustment
1. Check global `<options>` for `autoAdjustColumnWidth` setting
2. Check workbook-level `autoAdjustColumnWidth` attribute (overrides global)
3. Skip any columns with explicit `width` attribute
4. Apply adjustment within min/max constraints
5. Log decisions at DEBUG level for transparency

### Logging Levels
- **DEBUG**: Style application, precedence decisions, auto-adjustment details
- **INFO**: General progress (sheets created, cells processed, tables added)
- **WARNING**: Non-critical issues
- **ERROR**: Critical failures that prevent conversion

### Building for Distribution
PyInstaller spec file (ineoXlsxCmdLine.spec) includes:
- schema.xsd bundled into executable
- Company icon (logo_ineosolutions.ico)
- Version information (version_info.rc)
- Console application (not windowed)

## Code Modification Guidelines

### When Adding New XML Features
1. Update schema.xsd with new elements/attributes
2. Add parsing logic in excel_funciones_exportacion.py
3. Document in README.md with examples
4. Consider precedence system if property can exist at multiple levels
5. Add appropriate logging at INFO or DEBUG level

### When Modifying Styles
- All style properties support the precedence system
- Test with cells, rows, and columns defining the same property
- Ensure openpyxl compatibility (Font, PatternFill, Alignment objects)

### When Changing Validation
- XSD validation happens before any processing
- Use lxml for XSD validation (not xml.etree)
- Provide clear error messages with line numbers

## Version Information

Current version tracked in `ineoXlsxGlobales.py`:
- INEOXLSXCMDLINE_VERSION = "2.2.1"
- INEOXLSXCMDLINE_REVISION = "20250926"

Update these constants when releasing new versions.

## Dependencies

Core libraries (requirements.txt):
- **openpyxl 3.1.5**: Excel file manipulation
- **lxml 5.3.0**: XML validation against XSD
- **et_xmlfile 2.0.0**: XML file handling
- **xlsxwriter 3.2.5**: Additional Excel features

## Files to Preserve

- **schema.xsd**: XSD schema validation - critical for XML structure validation
- **logo_ineosolutions.ico**: Company branding in executable
- **version_info.rc**: Windows executable version metadata
- **task/*.xml**: Test/example XML files
