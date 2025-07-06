# 📊 Sample Data Guide

This guide provides examples of how to structure your Excel files for optimal use with the Advanced Column Search Tool.

## 📋 File Structure Requirements

### Source File (What you're searching for)
Your source file should contain the data you want to find in the target file.

**Example Source File Structure:**
```
| Name          | Company           | Email                    | Phone        |
|---------------|-------------------|--------------------------|---------------|
| John Smith    | TechCorp Inc      | john.smith@techcorp.com  | 555-0101     |
| Jane Doe      | DataSystems LLC   | jane.doe@datasys.com     | 555-0102     |
| Bob Johnson   | InnovateTech      | bob.j@innovate.com       | 555-0103     |
| Alice Brown   | FutureSoft        | alice@futuresoft.com     | 555-0104     |
```

### Target File (Where you're searching)
Your target file should contain the data you want to search within.

**Example Target File Structure:**
```
| ID | Full Name        | Organization          | Contact Email          | Phone Number | Department |
|----|------------------|----------------------|------------------------|--------------|------------|
| 1  | John A. Smith    | TechCorp Incorporated | john.smith@techcorp.com| 555-0101     | Engineering|
| 2  | Jane M. Doe      | DataSystems, LLC     | jane.doe@datasys.com   | 555-0102     | Marketing  |
| 3  | Robert Johnson   | InnovateTech Corp    | bob.j@innovate.com     | 555-0103     | Sales      |
| 4  | Alice R. Brown   | FutureSoft Solutions | alice@futuresoft.com   | 555-0104     | HR         |
| 5  | Mike Wilson      | TechCorp Inc         | mike.w@techcorp.com    | 555-0105     | Engineering|
```

## 🔍 Search Scenarios

### Scenario 1: Exact Name Matching
**Source Columns**: Name  
**Target Column**: Full Name  
**Search Type**: Exact Match  
**Expected Results**: High match rate for exact name matches

### Scenario 2: Fuzzy Company Matching
**Source Columns**: Company  
**Target Column**: Organization  
**Search Type**: Fuzzy Match (80% threshold)  
**Expected Results**: Matches variations like "TechCorp Inc" vs "TechCorp Incorporated"

### Scenario 3: Multi-Column Search
**Source Columns**: Name, Company  
**Target Column**: Full Name  
**Search Type**: Fuzzy Match (70% threshold)  
**Expected Results**: Combines name and company for better matching

### Scenario 4: Email Matching
**Source Columns**: Email  
**Target Column**: Contact Email  
**Search Type**: Exact Match  
**Expected Results**: Perfect matches for email addresses

## 📝 Creating Test Files

### Step 1: Create Source File
1. Open Excel or Google Sheets
2. Create headers in first row
3. Add sample data in subsequent rows
4. Save as `.xlsx` format

### Step 2: Create Target File
1. Create a new Excel file
2. Add more comprehensive data
3. Include variations and typos for testing
4. Save as `.xlsx` format

### Step 3: Test Different Scenarios
- Try exact matches first
- Test fuzzy matching with different thresholds
- Experiment with multi-column searches
- Test case sensitivity options

## 🎯 Best Practices

### Data Preparation
- **Clean your data** before uploading
- **Remove extra spaces** and formatting
- **Standardize formats** (dates, phone numbers, etc.)
- **Check for duplicates** in target file

### Column Selection
- **Choose relevant columns** for your search
- **Avoid empty columns** or columns with mostly null values
- **Consider data quality** when selecting search columns

### Search Configuration
- **Start with exact matching** to establish baseline
- **Use fuzzy matching** for data with variations
- **Adjust threshold** based on data quality
- **Test with small datasets** first

## 📊 Expected Results

### Exact Match Results
```
| Row | Source_Name | Target_Value    | Similarity_% | Match_Type |
|-----|-------------|-----------------|--------------|------------|
| 2   | Jane Doe    | Jane M. Doe     | 100%         | Exact      |
| 3   | Bob Johnson | Robert Johnson  | 100%         | Exact      |
```

### Fuzzy Match Results
```
| Row | Source_Company    | Target_Value           | Similarity_% | Match_Type |
|-----|-------------------|------------------------|--------------|------------|
| 1   | TechCorp Inc      | TechCorp Incorporated  | 85%          | Partial    |
| 2   | DataSystems LLC   | DataSystems, LLC      | 95%          | Partial    |
```

## 🚨 Common Issues

### No Matches Found
- **Check data formats** (spaces, punctuation, case)
- **Lower similarity threshold** for fuzzy matching
- **Verify column selections** are correct
- **Ensure data exists** in target file

### Too Many Matches
- **Increase similarity threshold**
- **Use more specific search columns**
- **Enable case sensitivity**
- **Refine your search criteria**

### Performance Issues
- **Reduce file sizes** (under 10MB recommended)
- **Use fewer search columns**
- **Close other browser tabs**
- **Process in smaller batches**

## 📈 Sample Data Files

You can create your own test files using these templates:

### Minimal Test File (5-10 rows)
Good for initial testing and understanding the tool.

### Medium Test File (100-500 rows)
Good for testing performance and accuracy.

### Large Test File (1000+ rows)
Good for stress testing and real-world scenarios.

## 🔧 Data Validation Tips

### Before Upload
- [ ] Check file format (.xlsx or .xls)
- [ ] Verify file size (under 50MB)
- [ ] Ensure headers are in first row
- [ ] Remove any empty rows/columns
- [ ] Check for special characters

### After Upload
- [ ] Verify sheets are detected
- [ ] Check column headers are correct
- [ ] Ensure data is visible in dropdowns
- [ ] Test with small search first

---

**Note**: These are example structures. Your actual data may vary, but following these patterns will help ensure optimal results with the Advanced Column Search Tool. 