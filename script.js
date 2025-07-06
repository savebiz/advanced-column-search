// Global variables
let searchSourceWorkbook = null;
let searchTargetWorkbook = null;
let advancedSearchInitialized = false;
window.searchResults = null;

// Initialize the application
function initializeAdvancedSearch() {
    if (advancedSearchInitialized) {
        console.log('[AdvancedSearch] Already initialized, skipping...');
        return;
    }

    // Element existence verification
    const sourceFileInput = document.getElementById('searchSourceFile');
    const targetFileInput = document.getElementById('searchTargetFile');
    
    if (!sourceFileInput || !targetFileInput) {
        console.error('[AdvancedSearch] Required elements not found');
        return;
    }

    // Task selector functionality
    const taskSelect = document.getElementById('taskSelect');
    if (taskSelect) {
        taskSelect.addEventListener('change', handleTaskChange);
    }

    // Event listener attachment
    sourceFileInput.addEventListener('change', handleSourceFileChange);
    targetFileInput.addEventListener('change', handleTargetFileChange);

    // Sheet change handlers
    document.getElementById('searchSourceSheet').addEventListener('change', handleSourceSheetChange);
    document.getElementById('searchTargetSheet').addEventListener('change', handleTargetSheetChange);

    // Search configuration handlers
    document.querySelectorAll('input[name="searchType"]').forEach(radio => {
        radio.addEventListener('change', handleSearchTypeChange);
    });

    // Similarity threshold handler
    document.getElementById('similarityThreshold').addEventListener('input', handleSimilarityChange);

    // Dynamic column management
    document.getElementById('addSearchColumn').addEventListener('click', addSearchColumn);
    document.getElementById('performSearchBtn').addEventListener('click', performAdvancedSearch);
    document.getElementById('exportSearchBtn').addEventListener('click', exportResults);

    // Initialize drag and drop
    initializeDragAndDrop();

    advancedSearchInitialized = true;
    console.log('[AdvancedSearch] Initialization complete');
}

// Task selector functionality
function handleTaskChange(event) {
    const selectedTask = event.target.value;
    const taskContents = document.querySelectorAll('.task-content');
    
    // Hide all task contents
    taskContents.forEach(content => {
        content.classList.remove('active');
    });
    
    // Show selected task content
    if (selectedTask) {
        const selectedContent = document.getElementById(selectedTask);
        if (selectedContent) {
            selectedContent.classList.add('active');
        }
    }
}

// Drag and Drop functionality
function initializeDragAndDrop() {
    const sourceSection = document.querySelector('.form-group:has(#searchSourceFile)');
    const targetSection = document.querySelector('.form-group:has(#searchTargetFile)');

    if (sourceSection) {
        sourceSection.addEventListener('dragover', handleDragOver);
        sourceSection.addEventListener('dragleave', handleDragLeave);
        sourceSection.addEventListener('drop', handleDrop);
    }

    if (targetSection) {
        targetSection.addEventListener('dragover', handleDragOver);
        targetSection.addEventListener('dragleave', handleDragLeave);
        targetSection.addEventListener('drop', handleDrop);
    }
}

function handleDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        const isSource = e.currentTarget.querySelector('#searchSourceFile');
        
        if (isSource) {
            document.getElementById('searchSourceFile').files = files;
            handleSourceFileChange({ target: { files: files } });
        } else {
            document.getElementById('searchTargetFile').files = files;
            handleTargetFileChange({ target: { files: files } });
        }
    }
}

// File handling functions
function handleSourceFileChange(event) {
    const file = event.target.files[0];
    if (!file) return;

    console.log('[AdvancedSearch] Source file selected:', file.name);
    
    // Update UI
    document.getElementById('sourceFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    updateFileInputLabel('searchSourceFile', true);

    // Validate file
    if (!validateFile(file)) return;

    // Process Excel file
    processExcelFile(file, 'source');
}

function handleTargetFileChange(event) {
    const file = event.target.files[0];
    if (!file) return;

    console.log('[AdvancedSearch] Target file selected:', file.name);
    
    // Update UI
    document.getElementById('targetFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    updateFileInputLabel('searchTargetFile', true);

    // Validate file
    if (!validateFile(file)) return;

    // Process Excel file
    processExcelFile(file, 'target');
}

function updateFileInputLabel(inputId, hasFile) {
    const label = document.querySelector(`label[for="${inputId}"]`);
    if (label) {
        if (hasFile) {
            label.classList.add('has-file');
            label.textContent = '📁 File Selected';
        } else {
            label.classList.remove('has-file');
            label.textContent = '📁 Choose Source File';
        }
    }
}

function validateFile(file) {
    // Check file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        alert('Please select a valid Excel file (.xlsx or .xls)');
        return false;
    }

    // Check file size (50MB limit)
    if (file.size > 50 * 1024 * 1024) {
        alert('The selected file is too large for browser processing. Please use a smaller file (recommended < 50MB).');
        return false;
    }

    return true;
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function processExcelFile(file, type) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error('No sheets found in the Excel file. The file may be corrupted or empty.');
            }

            if (type === 'source') {
                searchSourceWorkbook = workbook;
                populateSheetDropdown(workbook, 'searchSourceSheet');
                log('Source file loaded successfully with ' + workbook.SheetNames.length + ' sheets', 'success');
            } else {
                searchTargetWorkbook = workbook;
                populateSheetDropdown(workbook, 'searchTargetSheet');
                log('Target file loaded successfully with ' + workbook.SheetNames.length + ' sheets', 'success');
            }
        } catch (error) {
            console.error(`[AdvancedSearch] Error processing ${type} file:`, error);
            alert(`Error processing the Excel file: ${error.message}`);
            log(`Error processing ${type} file: ${error.message}`, 'error');
        }
    };
    
    reader.onerror = function() {
        alert('Error reading the file. Please try again.');
        log(`Error reading ${type} file`, 'error');
    };
    
    reader.readAsArrayBuffer(file);
}

// Sheet and column management
function populateSheetDropdown(workbook, dropdownId) {
    const dropdown = document.getElementById(dropdownId);
    dropdown.innerHTML = '<option value="">Select sheet...</option>';
    
    workbook.SheetNames.forEach(sheetName => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        dropdown.appendChild(option);
    });
    
    console.log('[populateSheetDropdown] Successfully populated:', dropdownId);
}

function populateColumnDropdown(workbook, sheetName, dropdownId) {
    const dropdown = document.getElementById(dropdownId);
    dropdown.innerHTML = '<option value="">Select column...</option>';
    
    if (!sheetName) return;
    
    try {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet || !sheet['!ref']) {
            log(`No data found in sheet: ${sheetName}`, 'warning');
            return;
        }
        
        const range = XLSX.utils.decode_range(sheet['!ref']);
        
        for (let col = range.s.c; col <= range.e.c; col++) {
            const colLetter = XLSX.utils.encode_col(col);
            const cellAddress = colLetter + '1';
            const headerCell = sheet[cellAddress];
            const headerValue = headerCell ? String(headerCell.v) : `Column ${colLetter}`;
            
            const option = document.createElement('option');
            option.value = colLetter;
            option.textContent = `${colLetter}: ${headerValue}`;
            dropdown.appendChild(option);
        }
        
        log(`Populated ${range.e.c - range.s.c + 1} columns for sheet: ${sheetName}`, 'success');
    } catch (error) {
        console.error('[populateColumnDropdown] Error:', error);
        log(`Error populating columns for sheet ${sheetName}: ${error.message}`, 'error');
    }
}

function populateSearchColumns(workbook, sheetName) {
    const selects = document.querySelectorAll('.search-column-select');
    selects.forEach(select => {
        const currentValue = select.value;
        select.innerHTML = '<option value="">Select column...</option>';
        
        if (!sheetName) return;
        
        try {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet || !sheet['!ref']) return;
            
            const range = XLSX.utils.decode_range(sheet['!ref']);
            
            for (let col = range.s.c; col <= range.e.c; col++) {
                const colLetter = XLSX.utils.encode_col(col);
                const cellAddress = colLetter + '1';
                const headerCell = sheet[cellAddress];
                const headerValue = headerCell ? String(headerCell.v) : `Column ${colLetter}`;
                
                const option = document.createElement('option');
                option.value = colLetter;
                option.textContent = `${colLetter}: ${headerValue}`;
                if (colLetter === currentValue) {
                    option.selected = true;
                }
                select.appendChild(option);
            }
        } catch (error) {
            console.error('[populateSearchColumns] Error:', error);
        }
    });
}

// Event handlers
function handleSourceSheetChange(event) {
    const sheetName = event.target.value;
    if (searchSourceWorkbook && sheetName) {
        populateSearchColumns(searchSourceWorkbook, sheetName);
        log('Source sheet changed to: ' + sheetName, 'success');
    }
}

function handleTargetSheetChange(event) {
    const sheetName = event.target.value;
    if (searchTargetWorkbook && sheetName) {
        populateColumnDropdown(searchTargetWorkbook, sheetName, 'searchTargetColumn');
        log('Target sheet changed to: ' + sheetName, 'success');
    }
}

function handleSearchTypeChange(event) {
    const similarityContainer = document.getElementById('similarityThreshold').parentElement.parentElement;
    if (event.target.value === 'partial') {
        similarityContainer.style.display = 'block';
        log('Switched to Fuzzy Match mode', 'success');
    } else {
        similarityContainer.style.display = 'none';
        log('Switched to Exact Match mode', 'success');
    }
}

function handleSimilarityChange(event) {
    const value = Math.round(event.target.value * 100);
    document.getElementById('similarityValue').textContent = value + '%';
}

// Dynamic column management
function addSearchColumn() {
    const container = document.getElementById('searchColumnsContainer');
    const newItem = document.createElement('div');
    newItem.className = 'search-column-item';
    
    const select = document.createElement('select');
    select.className = 'search-column-select form-control';
    select.innerHTML = '<option value="">Select column...</option>';
    
    // Populate with current sheet columns if available
    const sheetName = document.getElementById('searchSourceSheet').value;
    if (searchSourceWorkbook && sheetName) {
        try {
            const sheet = searchSourceWorkbook.Sheets[sheetName];
            if (sheet && sheet['!ref']) {
                const range = XLSX.utils.decode_range(sheet['!ref']);
                
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const colLetter = XLSX.utils.encode_col(col);
                    const cellAddress = colLetter + '1';
                    const headerCell = sheet[cellAddress];
                    const headerValue = headerCell ? String(headerCell.v) : `Column ${colLetter}`;
                    
                    const option = document.createElement('option');
                    option.value = colLetter;
                    option.textContent = `${colLetter}: ${headerValue}`;
                    select.appendChild(option);
                }
            }
        } catch (error) {
            console.error('[addSearchColumn] Error populating columns:', error);
        }
    }
    
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn btn-remove';
    removeBtn.textContent = 'Remove';
    removeBtn.onclick = function() { removeSearchColumn(this); };
    
    newItem.appendChild(select);
    newItem.appendChild(removeBtn);
    container.appendChild(newItem);
    
    log('Added new search column', 'success');
}

function removeSearchColumn(button) {
    const container = document.getElementById('searchColumnsContainer');
    const items = container.querySelectorAll('.search-column-item');
    if (items.length > 1) {
        button.parentElement.remove();
        log('Removed search column', 'success');
    } else {
        alert('At least one search column is required.');
        log('Cannot remove the last search column', 'warning');
    }
}

// Search functionality
function performAdvancedSearch() {
    // Validation
    if (!searchSourceWorkbook || !searchTargetWorkbook) {
        alert('Please select both source and target files.');
        return;
    }

    const sourceSheet = document.getElementById('searchSourceSheet').value;
    const targetSheet = document.getElementById('searchTargetSheet').value;
    const targetColumn = document.getElementById('searchTargetColumn').value;

    if (!sourceSheet || !targetSheet || !targetColumn) {
        alert('Please select all required sheets and columns.');
        return;
    }

    const searchColumns = Array.from(document.querySelectorAll('.search-column-select'))
        .map(select => select.value)
        .filter(value => value !== '');

    if (searchColumns.length === 0) {
        alert('Please select at least one search column.');
        return;
    }

    // Get search configuration
    const searchType = document.querySelector('input[name="searchType"]:checked').value;
    const similarityThreshold = parseFloat(document.getElementById('similarityThreshold').value);
    const caseSensitive = document.getElementById('searchCaseSensitive').checked;

    // Show progress
    showProgress();
    updateProgress(0, 'Initializing search...');

    // Perform search with progress updates
    setTimeout(() => {
        try {
            const results = executeSearch(
                searchSourceWorkbook, sourceSheet, searchColumns,
                searchTargetWorkbook, targetSheet, targetColumn,
                searchType, similarityThreshold, caseSensitive
            );
            
            displayResults(results);
            updateProgress(100, 'Search completed successfully!');
            
        } catch (error) {
            console.error('[AdvancedSearch] Search error:', error);
            alert('Error during search execution: ' + error.message);
            log('Search failed: ' + error.message, 'error');
            hideProgress();
        }
    }, 100);
}

function executeSearch(sourceWorkbook, sourceSheet, searchColumns, targetWorkbook, targetSheet, targetColumn, searchType, similarityThreshold, caseSensitive) {
    // Extract data
    const sourceData = XLSX.utils.sheet_to_json(sourceWorkbook.Sheets[sourceSheet], { header: 1 });
    const targetData = XLSX.utils.sheet_to_json(targetWorkbook.Sheets[targetSheet], { header: 1 });

    updateProgress(10, 'Data extracted, processing...');
    log(`Processing ${sourceData.length - 1} source records against ${targetData.length - 1} target records`, 'success');

    // Get column indices
    const sourceColumnIndices = searchColumns.map(col => XLSX.utils.decode_col(col));
    const targetColumnIndex = XLSX.utils.decode_col(targetColumn);

    // Prepare results
    const results = [];
    const headers = ['Row', ...searchColumns.map(col => `Source_${col}`), 'Target_Value', 'Similarity_%', 'Match_Type'];
    results.push(headers);

    let matchCount = 0;
    let partialMatchCount = 0;
    let noMatchCount = 0;

    // Process each source row
    const totalRows = sourceData.length - 1;
    for (let i = 1; i < sourceData.length; i++) {
        const sourceRow = sourceData[i];
        if (!sourceRow) continue;

        const progress = 10 + (i / sourceData.length) * 80;
        updateProgress(progress, `Processing row ${i} of ${sourceData.length}`);

        // Extract search values
        const searchValues = sourceColumnIndices.map(colIndex => 
            sourceRow[colIndex] ? String(sourceRow[colIndex]).trim() : ''
        ).filter(value => value !== '');

        if (searchValues.length === 0) {
            // No valid search values
            const resultRow = [
                i + 1,
                ...sourceColumnIndices.map(colIndex => sourceRow[colIndex] || ''),
                '',
                '0%',
                'No Match'
            ];
            results.push(resultRow);
            noMatchCount++;
            continue;
        }

        // Find best match in target data
        let bestMatch = null;
        let bestSimilarity = 0;
        let matchType = 'No Match';

        for (let j = 1; j < targetData.length; j++) {
            const targetRow = targetData[j];
            if (!targetRow || !targetRow[targetColumnIndex]) continue;

            const targetValue = String(targetRow[targetColumnIndex]).trim();
            if (!targetValue) continue;

            const similarity = calculateRowSimilarity(searchValues, targetValue, searchType, caseSensitive);

            if (searchType === 'exact' && similarity === 1.0) {
                bestMatch = targetValue;
                bestSimilarity = similarity;
                matchType = 'Exact';
                break; // Found exact match, no need to continue
            } else if (searchType === 'partial' && similarity >= similarityThreshold && similarity > bestSimilarity) {
                bestMatch = targetValue;
                bestSimilarity = similarity;
                matchType = bestSimilarity === 1.0 ? 'Exact' : 'Partial';
            }
        }

        // Add result
        const resultRow = [
            i + 1,
            ...sourceColumnIndices.map(colIndex => sourceRow[colIndex] || ''),
            bestMatch || '',
            Math.round(bestSimilarity * 100) + '%',
            matchType
        ];
        results.push(resultRow);

        // Update counters
        if (matchType === 'Exact') matchCount++;
        else if (matchType === 'Partial') partialMatchCount++;
        else noMatchCount++;

        // Log progress every 100 rows
        if (i % 100 === 0) {
            log(`Processed ${i} of ${totalRows} rows...`, 'success');
        }
    }

    updateProgress(90, 'Finalizing results...');

    // Store results and summary
    window.searchResults = {
        data: results,
        summary: {
            totalSourceRecords: sourceData.length - 1,
            matchCount: matchCount,
            partialMatchCount: partialMatchCount,
            noMatchCount: noMatchCount,
            searchType: searchType,
            similarityThreshold: similarityThreshold,
            searchColumns: searchColumns,
            targetColumn: targetColumn
        }
    };

    return window.searchResults;
}

function calculateRowSimilarity(searchValues, targetValue, searchType, caseSensitive) {
    if (searchType === 'exact') {
        // For exact match, check if all search values are found in target
        const targetStr = caseSensitive ? targetValue : targetValue.toLowerCase();
        for (const searchValue of searchValues) {
            if (!searchValue) continue;
            const searchStr = caseSensitive ? searchValue : searchValue.toLowerCase();
            if (!targetStr.includes(searchStr)) {
                return 0;
            }
        }
        return 1.0;
    } else {
        // For partial match, calculate average similarity
        let totalSimilarity = 0;
        let validValues = 0;
        
        for (const searchValue of searchValues) {
            if (!searchValue) continue;
            const similarity = calculateSimilarity(searchValue, targetValue, caseSensitive);
            totalSimilarity += similarity;
            validValues++;
        }
        
        return validValues > 0 ? totalSimilarity / validValues : 0;
    }
}

function calculateSimilarity(str1, str2, caseSensitive = false) {
    if (!str1 || !str2) return 0;
    
    const s1 = caseSensitive ? str1 : str1.toLowerCase();
    const s2 = caseSensitive ? str2 : str2.toLowerCase();
    
    // If strings are identical, return 1.0
    if (s1 === s2) return 1.0;
    
    // If one string contains the other, return high similarity
    if (s1.includes(s2) || s2.includes(s1)) {
        return 0.9;
    }
    
    const longer = s1.length > s2.length ? s1 : s2;
    const shorter = s1.length > s2.length ? s2 : s1;
    
    if (longer.length === 0) return 1.0;
    
    const editDistance = levenshteinDistance(longer, shorter);
    return (longer.length - editDistance) / longer.length;
}

function levenshteinDistance(str1, str2) {
    const matrix = [];
    
    for (let i = 0; i <= str2.length; i++) {
        matrix[i] = [i];
    }
    
    for (let j = 0; j <= str1.length; j++) {
        matrix[0][j] = j;
    }
    
    for (let i = 1; i <= str2.length; i++) {
        for (let j = 1; j <= str1.length; j++) {
            if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j] + 1
                );
            }
        }
    }
    
    return matrix[str2.length][str1.length];
}

// UI Management functions
function showProgress() {
    document.getElementById('progressContainer').style.display = 'block';
    document.getElementById('progressText').style.display = 'block';
    document.getElementById('logContainer').style.display = 'block';
    document.getElementById('searchSummary').style.display = 'none';
    document.getElementById('exportSearchBtn').style.display = 'none';
}

function hideProgress() {
    document.getElementById('progressContainer').style.display = 'none';
    document.getElementById('progressText').style.display = 'none';
}

function updateProgress(percentage, message) {
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    
    if (progressFill) {
        progressFill.style.width = `${percentage}%`;
    }
    if (progressText) {
        progressText.textContent = message || 'Processing...';
    }
    if (message) {
        log(message, 'success');
    }
}

function log(message, type = 'info') {
    const logContent = document.getElementById('logContent');
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry ${type}`;
    logEntry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
    logContent.appendChild(logEntry);
    logContent.scrollTop = logContent.scrollHeight;
}

function displayResults(results) {
    const summary = results.summary;
    
    // Create summary stats
    const statsHtml = `
        <div class="stat-card">
            <div class="stat-number">${summary.totalSourceRecords}</div>
            <div class="stat-label">Total Records</div>
        </div>
        <div class="stat-card exact">
            <div class="stat-number">${summary.matchCount}</div>
            <div class="stat-label">Exact Matches</div>
        </div>
        <div class="stat-card partial">
            <div class="stat-number">${summary.partialMatchCount}</div>
            <div class="stat-label">Partial Matches</div>
        </div>
        <div class="stat-card no-match">
            <div class="stat-number">${summary.noMatchCount}</div>
            <div class="stat-label">No Matches</div>
        </div>
    `;
    
    document.getElementById('summaryStats').innerHTML = statsHtml;
    document.getElementById('searchSummary').style.display = 'block';
    document.getElementById('exportSearchBtn').style.display = 'inline-flex';
    
    // Log final results
    log(`Search completed: ${summary.matchCount} exact matches, ${summary.partialMatchCount} partial matches, ${summary.noMatchCount} no matches`, 'success');
    
    setTimeout(() => {
        hideProgress();
    }, 1000);
}

function exportResults() {
    if (window.searchResults) {
        try {
            const newWorkbook = XLSX.utils.book_new();
            
            // Create main results sheet
            const searchSheet = XLSX.utils.aoa_to_sheet(window.searchResults.data);
            XLSX.utils.book_append_sheet(newWorkbook, searchSheet, 'Search Results');
            
            // Create summary sheet
            const summary = window.searchResults.summary;
            const summaryData = [
                ['Search Summary'],
                [''],
                ['Total Source Records', summary.totalSourceRecords],
                ['Exact Matches', summary.matchCount],
                ['Partial Matches', summary.partialMatchCount],
                ['No Matches', summary.noMatchCount],
                [''],
                ['Search Configuration'],
                ['Search Type', summary.searchType],
                ['Similarity Threshold', summary.similarityThreshold],
                ['Search Columns', summary.searchColumns.join(', ')],
                ['Target Column', summary.targetColumn],
                [''],
                ['Generated on', new Date().toLocaleString()]
            ];
            
            const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(newWorkbook, summarySheet, 'Summary');
            
            const filename = `advanced_search_results_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(newWorkbook, filename);
            
            log(`Results exported to ${filename}`, 'success');
        } catch (error) {
            console.error('[exportResults] Error:', error);
            alert('Error exporting results: ' + error.message);
            log('Export failed: ' + error.message, 'error');
        }
    } else {
        alert('No search results available for export');
        log('No search results to export', 'warning');
    }
}

// Initialize when page loads
document.addEventListener('DOMContentLoaded', function() {
    initializeAdvancedSearch();
    
    // Initialize search type handler
    handleSearchTypeChange({ target: { value: 'exact' } });
    
    // Initialize similarity display
    handleSimilarityChange({ target: { value: 0.8 } });
    
    log('DataGuard Advanced Column Search Tool initialized successfully', 'success');
});

// Test functions for debugging
window.testAdvancedSearchFileInput = function() {
    const input = document.getElementById('searchSourceFile');
    if (input) {
        input.click();
    }
};

window.debugAdvancedSearch = function() {
    console.log('Source workbook:', searchSourceWorkbook);
    console.log('Target workbook:', searchTargetWorkbook);
    console.log('Search results:', window.searchResults);
}; 