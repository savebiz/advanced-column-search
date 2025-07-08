# 🔍 DataGuard - Advanced Excel Analysis Tools

A comprehensive, browser-based Excel data processing suite with advanced search, duplicate detection, sequence validation, and row insertion capabilities.

## ✨ Key Features

### 📁 File Management
- **Drag-and-drop** file upload for Excel files
- **File validation** with size limits and format checking
- **Visual feedback** for file selection status
- Support for `.xlsx` and `.xls` formats

### 📊 Smart Configuration
- **Dynamic sheet selection** dropdowns
- **Multi-column selection** capability with add/remove functionality
- **Real-time column population** based on sheet selection
- **Flexible matching logic** for different use cases

### ⚡ Performance Features
- **Progress tracking** with visual progress bars
- **Real-time status updates**
- **Memory-efficient processing** for large datasets
- **Responsive design** for mobile and desktop

### 📈 Results & Analytics
- **Comprehensive result summaries** with statistics
- **Color-coded result categories**
- **Detailed results export** to Excel format
- **Real-time progress updates**

### 🛡️ Robust Error Handling
- **Input validation** at every step
- **File corruption detection**
- **Memory management** for large files
- **User-friendly error messages**

## 🚀 Available Tools

### 1. 🔍 Advanced Column Search
**Purpose**: Search for multiple column values within a target column using exact or partial matching.

#### Features:
- **Multi-column search** with add/remove functionality
- **Exact Match**: Searches for exact string matches
- **Fuzzy Match**: Uses Levenshtein distance algorithm for similarity matching
- **Flexible Matching Logic**:
  - **All Must Match**: All selected columns must be found in target
  - **At Least One Must Match**: Any selected column can trigger a match
  - **Custom**: Specify how many out of total columns must match
- **Configurable similarity threshold** (10% - 100%)
- **Case-sensitive search** option
- **Complete source data export** with all columns

#### Usage:
1. Upload source and target Excel files
2. Select sheets from both files
3. Choose target column (where to search)
4. Add search columns (what to search for)
5. Configure search type and matching logic
6. Perform search and export results

### 2. 🔄 Duplicate Row Identifier
**Purpose**: Identify exact duplicate rows based on selected columns.

#### Features:
- **Multi-column duplicate checking** with dynamic column management
- **All duplicates marked**: Both original and duplicate rows get marked "DUPLICATE"
- **Comprehensive analysis**: Shows all source columns with status
- **Statistics**: Total rows, unique rows, duplicate rows, duplicate groups
- **Excel export** with detailed results and summary

#### Usage:
1. Upload Excel file
2. Select sheet
3. Add columns to check for duplicates
4. Run analysis
5. Review results and export

#### Example:
```
Row 1: John, Smith, john@email.com → UNIQUE
Row 5: John, Smith, john@email.com → DUPLICATE
Row 8: John, Smith, john@email.com → DUPLICATE
```

### 3. 📦 Box Number Sequence Validator
**Purpose**: Identify broken box number sequences in inventory files.

#### Features:
- **Box column selection** from uploaded file
- **Sequence validation** within data blocks
- **Issue detection** for broken sequences
- **Empty row handling** for block separators
- **Detailed reporting** with expected vs found values
- **Excel export** with issue details

#### Usage:
1. Upload Excel file
2. Select sheet
3. Choose box number column
4. Run validation
5. Review issues and export results

#### Example:
**Correct Format:**
```
0001  Customer A
0001  Customer B
0001  Customer C
[empty row]
0002  Customer D
0002  Customer E
```

**Issues Detected:**
```
0001  Customer A
0002  Customer B  ← ISSUE: Expected 0001, found 0002
0003  Customer C  ← ISSUE: Expected 0001, found 0003
[empty row]
0002  Customer D  ← OK
```

### 4. ➕ Insert Empty Rows by Box Number
**Purpose**: Automatically insert empty rows between groups of different box numbers to visually separate them in your Excel sheet.

#### Features:
- **Box column selection** from uploaded file
- **Automatic empty row insertion** between groups of different box numbers
- **Progress tracking** and summary of inserted rows
- **Excel export** with modified data and summary

#### Usage:
1. Upload Excel file
2. Select sheet
3. Choose box number column
4. Click "Insert Empty Rows"
5. Review summary and export the modified file

#### Example:
**Before:**
```
00201354  ...
00201354  ...
00201354  ...
00201356  ...
00201356  ...
```
**After:**
```
00201354  ...
00201354  ...
00201354  ...
[empty row]
00201356  ...
00201356  ...
```

## 🚀 Getting Started

### Prerequisites
- Modern web browser (Chrome, Firefox, Safari, Edge)
- Excel files (.xlsx or .xls format)
- No installation required - runs entirely in the browser

### Usage Instructions

1. **Open the Tool**
   - Open `index.html` in your web browser
   - The tool will load automatically

2. **Select Task**
   - Choose from the dropdown menu:
     - Advanced Column Search
     - Duplicate Row Identifier
     - Box Number Sequence Validator
     - Insert Empty Rows by Box Number

3. **Upload Files**
   - Use drag-and-drop or click to browse
   - Upload required files for the selected task

4. **Configure Settings**
   - Select appropriate sheets
   - Choose columns for analysis
   - Configure search parameters (for Advanced Search)

5. **Run Analysis**
   - Click the process button
   - Monitor progress in real-time
   - View results summary when complete

6. **Export Results**
   - Click "Export Results" to download Excel file
   - Results include both analysis data and summary

## 📋 File Requirements

### Supported Formats
- **Excel 2007+**: `.xlsx` files
- **Excel 97-2003**: `.xls` files

### File Size Limits
- **Maximum file size**: 50MB per file
- **Recommended**: Under 10MB for optimal performance

### Data Requirements
- Files must contain at least one sheet
- Sheets must have header row (first row)
- Data should be in tabular format

## 🔧 Search Algorithms

### Exact Match
- Searches for exact string matches
- All search values must be found in target value
- Case sensitivity can be toggled

### Fuzzy Match (Levenshtein Distance)
- Calculates string similarity using edit distance
- Configurable threshold (10% - 100%)
- Returns best match above threshold
- Handles typos, abbreviations, and variations

### Multi-Column Search
- Combines multiple source columns into single search
- All selected columns contribute to similarity calculation
- Flexible column selection and ordering

## 📊 Results Interpretation

### Advanced Column Search
- **Exact Match**: Perfect string match found
- **Partial Match**: Similarity above threshold
- **No Match**: No suitable match found

### Duplicate Row Identifier
- **UNIQUE**: Row has no duplicates
- **DUPLICATE**: Row has at least one duplicate elsewhere
- **No Data**: Row has no data in selected columns

### Box Sequence Validator
- **OK**: Box number follows expected sequence
- **ISSUE**: Box number doesn't match expected sequence
- **SEPARATOR**: Empty row separating blocks

### Insert Empty Rows by Box Number
- **Original Rows**: Number of data rows before insertion
- **Empty Rows Inserted**: Number of empty rows added between box number changes

### Similarity Scores
- **100%**: Perfect match
- **80-99%**: High similarity
- **60-79%**: Moderate similarity
- **Below threshold**: No match

### Export Format
- **Analysis Results Sheet**: Detailed results with all data
- **Summary Sheet**: Statistics and configuration details

## 🎨 UI Features

### Responsive Design
- Works on desktop, tablet, and mobile devices
- Adaptive layout for different screen sizes
- Touch-friendly interface

### Visual Feedback
- Progress bars for long operations
- Color-coded status indicators
- Real-time logging of operations

### Accessibility
- Keyboard navigation support
- Screen reader compatible
- High contrast mode support

## 🔍 Advanced Usage

### Large File Processing
- Tool processes files in chunks for memory efficiency
- Progress updates every 100 rows
- Automatic memory management

### Error Recovery
- Graceful handling of corrupted files
- Detailed error messages
- Recovery suggestions

### Performance Tips
- Use smaller files for faster processing
- Close other browser tabs during large analyses
- Ensure stable internet connection for file uploads

## 🛠️ Technical Details

### Browser Compatibility
- **Chrome**: 60+
- **Firefox**: 55+
- **Safari**: 12+
- **Edge**: 79+

### Dependencies
- **SheetJS**: Excel file processing
- **Vanilla JavaScript**: No additional frameworks
- **CSS3**: Modern styling and animations

### Security
- All processing happens locally in browser
- No data sent to external servers
- Files remain on your device

## 🐛 Troubleshooting

### Common Issues

**File won't upload**
- Check file format (.xlsx or .xls)
- Ensure file size is under 50MB
- Try refreshing the page

**No sheets found**
- Verify Excel file has data
- Check if file is corrupted
- Try opening in Excel first

**Analysis takes too long**
- Reduce file size
- Use fewer search columns
- Close other browser tabs

**No matches found**
- Check similarity threshold
- Verify column selections
- Try case-insensitive search

### Error Messages

**"File too large"**
- Use smaller Excel files
- Split large files into smaller ones

**"No sheets found"**
- File may be corrupted
- Try saving as .xlsx format

**"Processing error"**
- Refresh page and try again
- Check browser console for details

## 📝 License

This tool is provided as-is for educational and business use. No warranty is provided.

## 🤝 Contributing

To improve this tool:
1. Fork the repository
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## 📞 Support

For issues or questions:
1. Check the troubleshooting section
2. Review browser console for errors
3. Ensure files meet requirements
4. Try with smaller test files first

---

**Version**: 2.1.0  
**Last Updated**: 2024  
**Browser Support**: Modern browsers with ES6 support  
**Tools**: Advanced Column Search, Duplicate Row Identifier, Box Number Sequence Validator, Insert Empty Rows by Box Number 