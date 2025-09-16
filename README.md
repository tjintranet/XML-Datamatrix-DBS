# Excel to XML Batch Converter with DataMatrix Generator

A comprehensive web-based application that converts Excel files to XML format while simultaneously generating DataMatrix barcodes for each record. Designed specifically for book production workflows with intelligent ISBN handling and automated production route logic.

## Overview

This unified tool combines Excel-to-XML conversion with DataMatrix barcode generation, providing a complete solution for book manufacturing data processing. The application automatically processes Excel files and generates both XML job tickets and tracking barcodes in a single workflow.

## Features

### **Core Functionality**
- **Unified Processing**: Single upload generates both XML files and DataMatrix barcodes
- **Multi-Format Support**: Handles Excel (.xlsx, .xls) files seamlessly
- **Intelligent ISBN Detection**: Automatically handles both Limp_ISBN and Cased_ISBN columns
- **Dynamic Transfer Station Logic**: Sets barcode position 37 based on book type (1 for Limp, 2 for Cased)
- **Fixed Trim Off Head**: Uses 3mm (0030) setting for consistent production standards
- **Smart Column Mapping**: Automatically maps data fields with manual override option
- **Combined Download**: Single ZIP file containing both XML files and DataMatrix PDFs
- **Row Management**: Delete unwanted rows from preview before processing
- **Responsive Design**: Works on desktop and mobile devices

### **Production Route Intelligence**
- **Automatic Trim Width Adjustment**: Detects "Limp P/Bound 8pp Cover" production routes and increases Trim_Width by 10mm
- **Visual Indicators**: Adjusted rows highlighted in yellow with warning icons
- **Audit Trail**: Summary of all automatic adjustments made

### **DataMatrix Barcode Features**
- **37-Digit Format**: Precisely formatted according to book production specifications
- **Portrait Orientation**: All PDF barcodes generated in portrait format for consistent filing
- **Intelligent Filename**: `[Wi_Number]_[ISBN]_DBC.pdf` format for easy organization
- **Live Preview**: Visual barcode previews in data tables with hover zoom
- **Individual Downloads**: Click any barcode preview for instant PDF download

## Installation

1. Download the application files to your web server or local directory:
   - `index.html`
   - `script.js`
   - `datamatrix.js`

2. Open `index.html` in a web browser

3. No server-side technology or build process required - this is a completely client-side application

## Usage

### **Primary Workflow**

1. **Upload Excel File**: Click "Choose Excel File" or drag and drop
2. **Automatic Processing**: 
   - File is parsed and validated
   - Production route logic applied automatically
   - DataMatrix barcodes generated for each row
   - XML data prepared for download
3. **Review Data**: 
   - Preview table shows key information and barcode previews
   - Delete unwanted rows using red trash buttons (left column)
   - Switch to Edit Mode for detailed field editing
4. **Download**: Single button downloads ZIP containing both XML files and barcode PDFs

### **Row Management**
- **Delete Rows**: Use red trash buttons in leftmost column to remove unwanted records
- **Confirmation**: Each deletion requires user confirmation to prevent accidents
- **Safe Position**: Delete buttons positioned away from download actions to prevent misclicks

### **Edit Mode**
- **Field Editing**: Modify 12 key production fields with live barcode updates
- **Sticky Headers**: Column headers remain visible while scrolling through large datasets
- **Visual Feedback**: Changes immediately reflected in barcode previews
- **Save Changes**: Apply modifications and return to preview mode

## File Format Requirements

### **Excel Files (.xlsx, .xls)**

#### **Required Columns**:
- **Limp_ISBN** OR **Cased_ISBN**: 13-digit ISBN (determines Transfer Station setting)
- **Wi_Number**: Work item identifier (used for file naming)
- **Trim_Height**: Height dimension in mm
- **Trim_Width**: Width dimension in mm  
- **Spine_Size**: Spine thickness in mm
- **Cut_Off**: Book block height before trimming in mm

#### **Column Name Variations**:
The application recognizes common naming variations:
- **Height**: `Trim_Height`, `Product_Height`, `Height_MM`, `Book_Height`
- **Width**: `Trim_Width`, `Product_Width`, `Width_MM`, `Book_Width`
- **Spine**: `Spine_Size`, `Spine_Width`, `Thickness`, `Book_Thickness`
- **Cut-Off**: `Cut_Off`, `Cutoff`, `Bleed`, `Margin`, `Trim`

## DataMatrix Barcode Format

The application generates a 37-digit barcode string in the following format:

| Position | Length | Description            | Format/Rules                                           |
|----------|--------|------------------------|-------------------------------------------------------|
| 1-13     | 13     | ISBN                   | From Limp_ISBN or Cased_ISBN field                    |
| 14-17    | 4      | Endsheet Height        | Fixed as "0000"                                        |
| 18-20    | 3      | Spine Size             | 2 digits + trailing zero OR leading + trailing zero   |
| 21-24    | 4      | Book Block Height      | 3 digits from Cut-Off + trailing zero                 |
| 25-28    | 4      | Trim Off Head          | **Fixed at 0030 (3mm)**                               |
| 29-32    | 4      | Trim Height            | 3 digits from Height + trailing zero                  |
| 33-36    | 4      | Trim Width             | 3 digits from Width + trailing zero                   |
| 37       | 1      | Transfer Station       | **1 = Limp_ISBN used, 2 = Cased_ISBN used**           |

### **Key Formatting Rules**:
- **ISBN Selection Priority**: Limp_ISBN checked first, then Cased_ISBN if empty
- **Transfer Station Logic**: 
  - `1` if data comes from Limp_ISBN field
  - `2` if data comes from Cased_ISBN field
- **Spine Formatting**:
  - 2-digit values: Add trailing zero (e.g., 24 → "240")
  - 1-digit values: Add leading and trailing zero (e.g., 8 → "080")
- **Fixed Trim Off Head**: Always 0030 (3mm) for consistent production standards

## Generated File Formats

### **XML Structure**
Each row generates an XML file following this structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<csv>
  <Wi_Number>324623</Wi_Number>
  <Limp_ISBN>9781739665333</Limp_ISBN>
  <Cased_ISBN></Cased_ISBN>
  <Trim_Height>198</Trim_Height>
  <Trim_Width>129</Trim_Width>
  <Spine_Size>24</Spine_Size>
  <Cut_Off>209</Cut_Off>
  <!-- Additional fields as mapped from Excel -->
</csv>
```

### **PDF DataMatrix Files**
- **Format**: Portrait orientation, optimally sized for printing
- **Content**: Clean DataMatrix barcode without legends or labels
- **Filename**: `[Wi_Number]_[ISBN]_DBC.pdf`
- **Example**: `324623_9781739665333_DBC.pdf`

## Combined Download Features

### **Single ZIP Download**
- **Unified Package**: One download contains all XML files and all DataMatrix PDFs
- **Organized Structure**: Clear file naming for easy sorting and processing
- **Filename**: `xml_and_datamatrix_files.zip`

### **File Organization in ZIP**:
```
xml_and_datamatrix_files.zip
├── 324623.xml
├── 324623_9781739665333_DBC.pdf
├── 325366.xml
├── 325366_9781837732425_DBC.pdf
└── ... (additional files)
```

## Interface Elements

### **Main Controls**
- **File Upload**: Drag-and-drop or click to select Excel files
- **Mode Toggle**: Switch between Preview and Edit modes
- **Combined Download**: Single button for all files (XML + DataMatrix PDFs)
- **Clear All**: Reset interface for new file processing

### **Preview Table**
- **Delete Column**: Red trash buttons for row removal (leftmost, safe position)
- **Data Columns**: Wi Number, ISBN, Title, Production Route, Trim Width
- **DataMatrix Column**: Live barcode previews with click-to-download
- **Action Column**: Blue download buttons for individual XML files

### **Processing Indicators**
- **Progress Messages**: Real-time status updates during file processing
- **Adjustment Alerts**: Notifications when production route logic is applied
- **Error Handling**: Clear feedback for validation issues or processing errors

## Technical Implementation

### **Client-Side Processing**
- **No Server Required**: All processing happens in the browser
- **External Libraries**: 
  - DataMatrix.js for barcode generation
  - SheetJS (XLSX) for Excel parsing
  - jsPDF for PDF creation
  - JSZip for archive creation
  - Bootstrap 5 for responsive UI
- **Browser Compatibility**: Modern browsers with JavaScript enabled
- **Performance**: Efficiently handles large datasets with hundreds of records

### **Security & Privacy**
- **Local Processing**: No data transmitted to external servers
- **Session-Based**: No data persistence between browser sessions
- **Client-Only**: Complete functionality without server dependencies

## Error Handling & Troubleshooting

### **File Upload Issues**
- **Unsupported Format**: Ensure files are .xlsx or .xls Excel format
- **Missing Columns**: Use manual column mapping for non-standard headers
- **Empty Files**: Verify Excel file contains data rows beyond headers
- **Large Files**: Application handles hundreds of records efficiently

### **Barcode Generation Issues**
- **Missing ISBN**: Ensure either Limp_ISBN or Cased_ISBN contains valid data
- **Invalid Measurements**: Check that numeric fields contain valid mm measurements
- **Barcode Preview Shows N/A**: Missing required fields prevent barcode generation

### **Download Issues**
- **Browser Blocking**: Check popup/download blocker settings
- **Empty ZIP**: Verify data was processed successfully before download
- **PDF Generation Errors**: Check browser console for specific error messages

## Business Logic

### **Production Route Processing**
1. **Detection**: Scans for "Limp P/Bound 8pp Cover" in Production_Route column
2. **Adjustment**: Automatically increases Trim_Width by 10mm for matching rows
3. **Notification**: Displays summary of adjustments made
4. **Visual Marking**: Highlights adjusted rows in yellow throughout interface

### **ISBN Logic**
1. **Priority Check**: Limp_ISBN field checked first
2. **Fallback**: Cased_ISBN used if Limp_ISBN is empty
3. **Transfer Station**: Position 37 set to 1 (Limp) or 2 (Cased) accordingly
4. **Validation**: Ensures 13-digit format with leading zero padding if needed

## Browser Requirements

- **Modern Web Browser**: Chrome, Firefox, Safari, Edge (latest versions)
- **JavaScript**: Must be enabled
- **File API Support**: For Excel file processing
- **Canvas API**: For PDF generation
- **Download Support**: For ZIP file downloads
- **Local Storage**: Not required - all processing in memory

## Contributing

1. Fork the repository
2. Create a feature branch
3. Test thoroughly with various Excel formats
4. Ensure barcode validation accuracy
5. Submit pull request with clear description

## License

This project is available under the MIT License. Feel free to modify and distribute as needed for production workflows.

## Support

For issues or questions:
1. **Validation Errors**: Check Excel file format and required columns
2. **Processing Failures**: Verify browser console for JavaScript errors
3. **Download Problems**: Check browser download/popup settings
4. **Data Issues**: Ensure numeric fields contain valid measurements

## Version History

### **Current Version**
- **Combined Downloads**: Single ZIP with XML files and DataMatrix PDFs
- **Row Management**: Delete functionality with confirmation dialogs
- **Automatic Orientation**: PDF orientation selected based on barcode dimensions for optimal layout
- **Fixed Trim Off Head**: Standardized at 3mm (0030)
- **Enhanced Filename Format**: `[Wi_Number]_[ISBN]_DBC.pdf`
- **Improved Error Handling**: Comprehensive validation and user feedback
- **Responsive Layout**: Optimized interface for various screen sizes

### **Key Improvements**
- **Unified Workflow**: Eliminated need for separate XML and barcode downloads
- **User Safety**: Delete buttons positioned to prevent accidental clicks
- **Production Standards**: Fixed trim settings ensure consistency
- **Better Organization**: Intelligent file naming for easier management
- **Enhanced UX**: Streamlined interface with better visual feedback