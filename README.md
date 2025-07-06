# 🔍 Advanced Column Search Tool

A powerful, browser-based Excel data processing tool with advanced fuzzy matching and multi-column search capabilities.

## ✨ Key Features

### 📁 File Management
- **Drag-and-drop** file upload for both source and target Excel files
- **File validation** with size limits and format checking
- **Visual feedback** for file selection status
- Support for `.xlsx` and `.xls` formats

### 📊 Smart Configuration
- **Dynamic sheet selection** dropdowns
- **Multi-column search** capability with add/remove functionality
- **Flexible target column** selection
- **Real-time column population** based on sheet selection

### 🔍 Advanced Search Options
- **Exact Match**: Searches for exact string matches within target values
- **Fuzzy Match**: Uses Levenshtein distance algorithm for similarity matching
- **Configurable similarity threshold** (10% - 100%)
- **Case-sensitive search** option

### ⚡ Performance Features
- **Progress tracking** with visual progress bar
- **Real-time process logging**
- **Memory-efficient processing** for large datasets
- **Responsive design** for mobile and desktop

### 📈 Results & Analytics
- **Comprehensive search summary** with statistics
- **Color-coded result categories** (exact, partial, no match)
- **Detailed results export** to Excel format
- **Real-time progress updates**

### 🛡️ Robust Error Handling
- **Input validation** at every step
- **File corruption detection**
- **Memory management** for large files
- **User-friendly error messages**

## 🚀 Getting Started

### Prerequisites
- Modern web browser (Chrome, Firefox, Safari, Edge)
- Excel files (.xlsx or .xls format)
- No installation required - runs entirely in the browser

### Usage Instructions

1. **Open the Tool**
   - Open `index.html` in your web browser
   - The tool will load automatically

2. **Upload Files**
   - **Source File**: Contains the data you want to search for
   - **Target File**: Contains the data you want to search within
   - Use drag-and-drop or click to browse

3. **Configure Sheets and Columns**
   - Select the appropriate sheets from both files
   - Choose the target column (where to search)
   - Add one or more search columns (what to search for)

4. **Set Search Parameters**
   - Choose between Exact Match or Fuzzy Match
   - Adjust similarity threshold for fuzzy matching
   - Enable/disable case sensitivity

5. **Perform Search**
   - Click "Perform Search" to start processing
   - Monitor progress in real-time
   - View results summary when complete

6. **Export Results**
   - Click "Export Results" to download Excel file
   - Results include both search data and summary

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

### Match Types
- **Exact Match**: Perfect string match found
- **Partial Match**: Similarity above threshold
- **No Match**: No suitable match found

### Similarity Scores
- **100%**: Perfect match
- **80-99%**: High similarity
- **60-79%**: Moderate similarity
- **Below threshold**: No match

### Export Format
- **Search Results Sheet**: Detailed results with all data
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
- Close other browser tabs during large searches
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

**Search takes too long**
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

**Version**: 1.0.0  
**Last Updated**: 2024  
**Browser Support**: Modern browsers with ES6 support 