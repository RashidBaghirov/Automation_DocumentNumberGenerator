# Payriff Documents Generator

An automated document generation system for creating sequential contract documents with automatic numbering, Word document processing, and printing capabilities.

## Overview

This project automates the process of generating contract documents for Payriff Bakubus Validator system. It creates sequential contract numbers, updates Word document templates, saves new documents, and automatically sends them to printers.

## Features

- **Sequential Contract Numbering**: Automatically generates contract numbers starting from 120
- **Word Document Automation**: Uses COM interface to manipulate Word documents
- **Template Processing**: Updates bookmarks in Word templates with contract numbers
- **Automatic Printing**: Detects and uses physical printers, avoiding virtual printers (PDF, FAX, etc.)
- **Smart Printer Detection**: Advanced printer discovery using WMIC, PowerShell, and Word APIs
- **Batch Processing**: Can generate multiple contracts in sequence
- **Data Persistence**: Saves contract numbers to JSON file for tracking

## Project Structure

```
Payriff_Documents_Generator/
├── document_generator.py          # Main application file
├── contract_numbers.json          # Contract numbers tracking (auto-generated)
├── Payriff_Bakubus_Validator.docx # Word template file
└── generated_contracts/           # Output folder for generated contracts
```

## Requirements

### Software Dependencies
- Python 3.7+
- Microsoft Word (installed and licensed)
- Windows OS (required for COM interface)

### Python Packages
```bash
pip install pywin32
```

### System Requirements
- Windows operating system
- Microsoft Word application
- At least one physical printer configured
- Administrative rights may be required for printer access

## Installation

1. **Clone or download the project files**
2. **Install Python dependencies:**
   ```bash
   pip install pywin32
   ```
3. **Ensure Microsoft Word is installed and working**
4. **Place your Word template file in the project directory**
5. **Update file paths in the code if necessary:**
   ```python
   self.template_path = r"C:\Path\To\Your\Template.docx"
   self.output_folder = r"C:\Path\To\Output\Folder"
   ```

##  Usage

### Basic Usage
```python
from document_generator import DocumentNumberGenerator

# Initialize the generator
generator = DocumentNumberGenerator()

# Generate a single contract
success = generator.auto_generate_and_save()

# Clean up
generator.close_word()
```

### Batch Generation
```python
# The run_auto_mode() method generates multiple contracts
generator = DocumentNumberGenerator()
generator.run_auto_mode()  # Currently set to generate 1 contract, modify target_count for more
```

### Manual Number Generation
```python
# Generate specific contract number
contract_number, number_only = generator.generate_sequential_contract_number(120)
print(f"Generated: {contract_number}")
```

##  Configuration

### Template Requirements
Your Word template must contain a bookmark named `document_number` where the contract number will be inserted.

**To add a bookmark in Word:**
1. Select the text/location where the number should appear
2. Go to Insert → Bookmark
3. Name it `document_number`
4. Click Add

### Contract Number Format
Generated numbers follow the pattern: `MQ-YYYYMMDD-XXX`
- `MQ`: Prefix
- `YYYYMMDD`: Current date
- `XXX`: Sequential number (starting from 120)

Example: `MQ-20250820-120`

##  Printer Configuration

The system automatically:
- Detects all available printers
- Filters out virtual printers (PDF, FAX, XPS, etc.)
- Prioritizes physical printers
- Falls back to any available printer if physical printers fail

### Supported Printer Detection Methods
1. **WMIC** - Windows Management Instrumentation
2. **PowerShell** - Get-Printer cmdlet
3. **Word COM** - Microsoft Word printer interface

##  Output

### Generated Files
- **Contract Documents**: `Contract_{number}_{timestamp}.docx`
- **Tracking Data**: `contract_numbers.json`

### Console Output
The application provides detailed logging:
- Printer detection results
- Document processing status
- Print job status
- Error messages and troubleshooting info

##  Troubleshooting

### Common Issues

**Word Application Not Starting**
- Ensure Microsoft Word is properly installed
- Check if Word is already running and close it
- Run the script with administrator privileges

**Template File Not Found**
- Verify the template path in the code
- Ensure the Word template file exists
- Check file permissions

**Printer Not Found**
- Ensure at least one printer is installed
- Check printer drivers are properly installed
- Verify printer is not in offline mode

**Bookmark Not Found**
- Ensure the Word template contains a bookmark named `document_number`
- Check bookmark spelling and case sensitivity

### Debug Mode
Enable detailed logging by checking console output for:
- Printer detection process
- Word document operations
- Print queue status

##  Data Storage

Contract numbers are stored in `contract_numbers.json`:
```json
{
  "sequential_20250820": {
    "current_number": 121
  }
}
```





---

**Note**: This project is specifically designed for Windows environments with Microsoft Word. The COM interface dependencies make it incompatible with other operating systems.
