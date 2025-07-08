// Global variables
let searchSourceWorkbook = null;
let searchTargetWorkbook = null;
let advancedSearchInitialized = false;
window.searchResults = null;

// Duplicate identifier variables
let duplicateWorkbook = null;
let duplicateInitialized = false;
window.duplicateResults = null;

// Box validator variables
let boxWorkbook = null;
let boxInitialized = false;
window.boxResults = null;

// Insert Empty Rows by Box Number variables
let insertRowsWorkbook = null;
let insertRowsInitialized = false;
window.insertRowsResults = null;

// Batch processing state
let batchProcessing = {
    isActive: false,
    isPaused: false,
    cancelRequested: false,
    params: null,
    i: 1,
    results: [],
    matchCount: 0,
    partialMatchCount: 0,
    noMatchCount: 0,
    totalRows: 0,
    batchSize: 100,
    headers: [],
    onProgress: null,
    onComplete: null
};

function showBatchControls(showCancel, showNextBatch) {
    const controls = document.getElementById('batchControls');
    const cancelBtn = document.getElementById('cancelBtn');
    const nextBatchBtn = document.getElementById('nextBatchBtn');
    controls.style.display = (showCancel || showNextBatch) ? 'block' : 'none';
    cancelBtn.style.display = showCancel ? 'inline-block' : 'none';
    nextBatchBtn.style.display = showNextBatch ? 'inline-block' : 'none';
}

function resetBatchProcessingState() {
    batchProcessing.isActive = false;
    batchProcessing.isPaused = false;
    batchProcessing.cancelRequested = false;
    batchProcessing.params = null;
    batchProcessing.i = 1;
    batchProcessing.results = [];
    batchProcessing.matchCount = 0;
    batchProcessing.partialMatchCount = 0;
    batchProcessing.noMatchCount = 0;
    batchProcessing.totalRows = 0;
    batchProcessing.headers = [];
    batchProcessing.onProgress = null;
    batchProcessing.onComplete = null;
    showBatchControls(false, false);
}

function processBatchInternal() {
    if (!batchProcessing.isActive || batchProcessing.cancelRequested) {
        hideProgress();
        resetBatchProcessingState();
        return;
    }
    const {
        sourceData, targetData, sourceColumnIndices, targetColumnIndex,
        searchType, similarityThreshold, caseSensitive, onProgress, onComplete, incremental,
        matchingLogic, customMatchCount, allSourceColumns
    } = batchProcessing.params;
    let i = batchProcessing.i;
    const batchSize = batchProcessing.batchSize;
    const end = Math.min(i + batchSize, sourceData.length);

    for (; i < end; i++) {
        const sourceRow = sourceData[i];
        if (!sourceRow) continue;

        // Extract search values
        const searchValues = sourceColumnIndices.map(colIndex =>
            sourceRow[colIndex] ? String(sourceRow[colIndex]).trim() : ''
        ).filter(value => value !== '');

        if (searchValues.length === 0) {
            // No valid search values
            const resultRow = [
                i + 1,
                ...allSourceColumns.map(colIndex => sourceRow[colIndex] || ''),
                '',
                '0%',
                'No Match'
            ];
            batchProcessing.results.push(resultRow);
            batchProcessing.noMatchCount++;
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

            const similarity = calculateRowSimilarity(searchValues, targetValue, searchType, caseSensitive, matchingLogic, customMatchCount);

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

        // Add result with all source columns
        const resultRow = [
            i + 1,
            ...allSourceColumns.map(colIndex => sourceRow[colIndex] || ''),
            bestMatch || '',
            Math.round(bestSimilarity * 100) + '%',
            matchType
        ];
        batchProcessing.results.push(resultRow);

        // Update counters
        if (matchType === 'Exact') batchProcessing.matchCount++;
        else if (matchType === 'Partial') batchProcessing.partialMatchCount++;
        else batchProcessing.noMatchCount++;
    }

    batchProcessing.i = i;
    if (onProgress) {
        onProgress((i / sourceData.length) * 100);
    }

    if (i < sourceData.length) {
        if (batchProcessing.incremental) {
            // Pause and wait for user to continue
            batchProcessing.isPaused = true;
            showBatchControls(true, true);
            if (onComplete) {
                onComplete({
                    data: batchProcessing.results,
                    summary: {
                        totalSourceRecords: sourceData.length - 1,
                        matchCount: batchProcessing.matchCount,
                        partialMatchCount: batchProcessing.partialMatchCount,
                        noMatchCount: batchProcessing.noMatchCount,
                        searchType,
                        similarityThreshold,
                        matchingLogic,
                        customMatchCount,
                        searchColumns: sourceColumnIndices.map(idx => XLSX.utils.encode_col(idx)),
                        targetColumn: XLSX.utils.encode_col(targetColumnIndex)
                    }
                }, false);
            }
        } else {
            setTimeout(processBatchInternal, 0);
        }
    } else {
        // Done
        showBatchControls(false, false);
        if (onComplete) {
            onComplete({
                data: batchProcessing.results,
                summary: {
                    totalSourceRecords: sourceData.length - 1,
                    matchCount: batchProcessing.matchCount,
                    partialMatchCount: batchProcessing.partialMatchCount,
                    noMatchCount: batchProcessing.noMatchCount,
                    searchType,
                    similarityThreshold,
                    matchingLogic,
                    customMatchCount,
                    searchColumns: sourceColumnIndices.map(idx => XLSX.utils.encode_col(idx)),
                    targetColumn: XLSX.utils.encode_col(targetColumnIndex)
                }
            }, true);
        }
    }
}

function executeSearchInBatches(params) {
    // Cancel any previous batch
    resetBatchProcessingState();
    batchProcessing.isActive = true;
    batchProcessing.cancelRequested = false;
    batchProcessing.isPaused = false;
    batchProcessing.params = params;
    batchProcessing.i = 1;
    batchProcessing.results = [];
    batchProcessing.matchCount = 0;
    batchProcessing.partialMatchCount = 0;
    batchProcessing.noMatchCount = 0;
    batchProcessing.totalRows = params.sourceData.length;
    
    // Create headers with all source columns
    const sourceHeaders = params.allSourceColumns.map(col => `Source_${XLSX.utils.encode_col(col)}`);
    batchProcessing.headers = ['Row', ...sourceHeaders, 'Target_Value', 'Similarity_%', 'Match_Type'];
    batchProcessing.results.push(batchProcessing.headers);
    batchProcessing.onProgress = params.onProgress;
    batchProcessing.onComplete = params.onComplete;
    batchProcessing.incremental = params.incremental;

    showBatchControls(true, false);
    processBatchInternal();
}

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

    // Matching logic handlers
    document.querySelectorAll('input[name="matchingLogic"]').forEach(radio => {
        radio.addEventListener('change', handleMatchingLogicChange);
    });

    // Dynamic column management
    document.getElementById('addSearchColumn').addEventListener('click', addSearchColumn);
    document.getElementById('performSearchBtn').addEventListener('click', performAdvancedSearch);
    document.getElementById('exportSearchBtn').addEventListener('click', exportResults);

    // Initialize drag and drop
    initializeDragAndDrop();

    advancedSearchInitialized = true;
    console.log('[AdvancedSearch] Initialization complete');
}

// Initialize Duplicate Identifier
function initializeDuplicateIdentifier() {
    if (duplicateInitialized) {
        console.log('[DuplicateIdentifier] Already initialized, skipping...');
        return;
    }

    // Element existence verification
    const duplicateFileInput = document.getElementById('duplicateFile');
    
    if (!duplicateFileInput) {
        console.error('[DuplicateIdentifier] Required elements not found');
        return;
    }

    // Event listener attachment
    duplicateFileInput.addEventListener('change', handleDuplicateFileChange);

    // Sheet change handler
    document.getElementById('duplicateSheet').addEventListener('change', handleDuplicateSheetChange);

    // Dynamic column management
    document.getElementById('addDuplicateColumn').addEventListener('click', addDuplicateColumn);
    document.getElementById('performDuplicateBtn').addEventListener('click', performDuplicateAnalysis);
    document.getElementById('exportDuplicateBtn').addEventListener('click', exportDuplicateResults);

    duplicateInitialized = true;
    console.log('[DuplicateIdentifier] Initialization complete');
}

// Initialize Box Validator
function initializeBoxValidator() {
    if (boxInitialized) {
        console.log('[BoxValidator] Already initialized, skipping...');
        return;
    }

    // Element existence verification
    const boxFileInput = document.getElementById('boxFile');
    
    if (!boxFileInput) {
        console.error('[BoxValidator] Required elements not found');
        return;
    }

    // Event listener attachment
    boxFileInput.addEventListener('change', handleBoxFileChange);

    // Sheet change handler
    document.getElementById('boxSheet').addEventListener('change', handleBoxSheetChange);

    // Action button
    document.getElementById('performBoxBtn').addEventListener('click', performBoxValidation);
    document.getElementById('exportBoxBtn').addEventListener('click', exportBoxResults);

    boxInitialized = true;
    console.log('[BoxValidator] Initialization complete');
}

// Initialize Insert Empty Rows by Box Number
function initializeInsertRows() {
    if (insertRowsInitialized) {
        console.log('[InsertRows] Already initialized, skipping...');
        return;
    }

    // Element existence verification
    const fileInput = document.getElementById('insertRowsFile');
    if (!fileInput) {
        console.error('[InsertRows] Required elements not found');
        return;
    }

    // Event listener attachment
    fileInput.addEventListener('change', handleInsertRowsFileChange);
    document.getElementById('insertRowsSheet').addEventListener('change', handleInsertRowsSheetChange);
    document.getElementById('performInsertRowsBtn').addEventListener('click', performInsertRows);
    document.getElementById('exportInsertRowsBtn').addEventListener('click', exportInsertRowsResults);

    insertRowsInitialized = true;
    console.log('[InsertRows] Initialization complete');
}

function handleInsertRowsFileChange(event) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById('insertRowsFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    updateFileInputLabel('insertRowsFile', true);
    if (!validateFile(file)) return;
    processInsertRowsExcelFile(file);
}

function processInsertRowsExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error('No sheets found in the Excel file. The file may be corrupted or empty.');
            }
            insertRowsWorkbook = workbook;
            populateSheetDropdown(workbook, 'insertRowsSheet');
        } catch (error) {
            console.error('[InsertRows] Error processing file:', error);
            alert(`Error processing the Excel file: ${error.message}`);
        }
    };
    reader.onerror = function() {
        alert('Error reading the file. Please try again.');
    };
    reader.readAsArrayBuffer(file);
}

function handleInsertRowsSheetChange(event) {
    const sheetName = event.target.value;
    if (insertRowsWorkbook && sheetName) {
        populateColumnDropdown(insertRowsWorkbook, sheetName, 'insertRowsBoxColumn');
    }
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
            
            // Initialize the selected task
            if (selectedTask === 'advanced-search' && !advancedSearchInitialized) {
                initializeAdvancedSearch();
            } else if (selectedTask === 'duplicate-identifier' && !duplicateInitialized) {
                initializeDuplicateIdentifier();
            } else if (selectedTask === 'box-validator' && !boxInitialized) {
                initializeBoxValidator();
            } else if (selectedTask === 'insert-empty-rows' && !insertRowsInitialized) {
                initializeInsertRows();
            }
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
            } else {
                searchTargetWorkbook = workbook;
                populateSheetDropdown(workbook, 'searchTargetSheet');
            }
        } catch (error) {
            console.error(`[AdvancedSearch] Error processing ${type} file:`, error);
            alert(`Error processing the Excel file: ${error.message}`);
        }
    };
    
    reader.onerror = function() {
        alert('Error reading the file. Please try again.');
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
    } catch (error) {
        console.error('[populateColumnDropdown] Error:', error);
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
        updateTotalSearchColumns();
    }
}

function handleTargetSheetChange(event) {
    const sheetName = event.target.value;
    if (searchTargetWorkbook && sheetName) {
        populateColumnDropdown(searchTargetWorkbook, sheetName, 'searchTargetColumn');
    }
}

function handleSearchTypeChange(event) {
    const similarityContainer = document.getElementById('similarityThreshold').parentElement.parentElement;
    if (event.target.value === 'partial') {
        similarityContainer.style.display = 'block';
    } else {
        similarityContainer.style.display = 'none';
    }
}

function handleSimilarityChange(event) {
    const value = parseFloat(event.target.value);
    document.getElementById('similarityValue').textContent = Math.round(value * 100) + '%';
}

function handleMatchingLogicChange(event) {
    const customInput = document.getElementById('customMatchCount');
    const customLabel = document.querySelector('label[for="customMatch"]');
    
    if (event.target.value === 'custom') {
        customInput.style.display = 'inline-block';
        updateTotalSearchColumns();
    } else {
        customInput.style.display = 'none';
    }
}

function updateTotalSearchColumns() {
    const searchColumns = document.querySelectorAll('.search-column-select');
    const totalSpan = document.getElementById('totalSearchColumns');
    const customInput = document.getElementById('customMatchCount');
    
    if (totalSpan) {
        totalSpan.textContent = searchColumns.length;
    }
    
    if (customInput) {
        customInput.max = searchColumns.length;
        if (parseInt(customInput.value) > searchColumns.length) {
            customInput.value = Math.max(1, Math.ceil(searchColumns.length / 2));
        }
    }
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
    
    updateTotalSearchColumns();
}

function removeSearchColumn(button) {
    const container = document.getElementById('searchColumnsContainer');
    const items = container.querySelectorAll('.search-column-item');
    if (items.length > 1) {
        button.parentElement.remove();
        updateTotalSearchColumns();
    } else {
        alert('At least one search column is required.');
    }
}

// Duplicate Identifier file handling
function handleDuplicateFileChange(event) {
    const file = event.target.files[0];
    if (!file) return;

    console.log('[DuplicateIdentifier] File selected:', file.name);
    
    // Update UI
    document.getElementById('duplicateFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    updateFileInputLabel('duplicateFile', true);

    // Validate file
    if (!validateFile(file)) return;

    // Process Excel file
    processDuplicateExcelFile(file);
}

function processDuplicateExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error('No sheets found in the Excel file. The file may be corrupted or empty.');
            }

            duplicateWorkbook = workbook;
            populateSheetDropdown(workbook, 'duplicateSheet');
        } catch (error) {
            console.error('[DuplicateIdentifier] Error processing file:', error);
            alert(`Error processing the Excel file: ${error.message}`);
        }
    };
    
    reader.onerror = function() {
        alert('Error reading the file. Please try again.');
    };
    
    reader.readAsArrayBuffer(file);
}

function handleDuplicateSheetChange(event) {
    const sheetName = event.target.value;
    if (duplicateWorkbook && sheetName) {
        populateDuplicateColumns(duplicateWorkbook, sheetName);
    }
}

function populateDuplicateColumns(workbook, sheetName) {
    const selects = document.querySelectorAll('.duplicate-column-select');
    selects.forEach(select => {
        const currentValue = select.value;
        select.innerHTML = '<option value="">Select column...</option>';
        
        if (!sheetName) return;
        
        try {
            const sheet = workbook.Sheets[sheetName];
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
                    if (colLetter === currentValue) {
                        option.selected = true;
                    }
                    select.appendChild(option);
                }
            }
        } catch (error) {
            console.error('[populateDuplicateColumns] Error:', error);
        }
    });
}

// Duplicate Identifier column management
function addDuplicateColumn() {
    const container = document.getElementById('duplicateColumnsContainer');
    const newItem = document.createElement('div');
    newItem.className = 'search-column-item';
    
    const select = document.createElement('select');
    select.className = 'duplicate-column-select form-control';
    select.innerHTML = '<option value="">Select column...</option>';
    
    // Populate with current sheet columns if available
    const sheetName = document.getElementById('duplicateSheet').value;
    if (duplicateWorkbook && sheetName) {
        try {
            const sheet = duplicateWorkbook.Sheets[sheetName];
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
            console.error('[addDuplicateColumn] Error populating columns:', error);
        }
    }
    
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn btn-remove';
    removeBtn.textContent = 'Remove';
    removeBtn.onclick = function() { removeDuplicateColumn(this); };
    
    newItem.appendChild(select);
    newItem.appendChild(removeBtn);
    container.appendChild(newItem);
}

function removeDuplicateColumn(button) {
    const container = document.getElementById('duplicateColumnsContainer');
    const items = container.querySelectorAll('.search-column-item');
    if (items.length > 1) {
        button.parentElement.remove();
    } else {
        alert('At least one column is required for duplicate checking.');
    }
}

// Box Validator file handling
function handleBoxFileChange(event) {
    const file = event.target.files[0];
    if (!file) return;

    console.log('[BoxValidator] File selected:', file.name);
    
    // Update UI
    document.getElementById('boxFileInfo').textContent = `${file.name} (${formatFileSize(file.size)})`;
    updateFileInputLabel('boxFile', true);

    // Validate file
    if (!validateFile(file)) return;

    // Process Excel file
    processBoxExcelFile(file);
}

function processBoxExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error('No sheets found in the Excel file. The file may be corrupted or empty.');
            }

            boxWorkbook = workbook;
            populateSheetDropdown(workbook, 'boxSheet');
        } catch (error) {
            console.error('[BoxValidator] Error processing file:', error);
            alert(`Error processing the Excel file: ${error.message}`);
        }
    };
    
    reader.onerror = function() {
        alert('Error reading the file. Please try again.');
    };
    
    reader.readAsArrayBuffer(file);
}

function handleBoxSheetChange(event) {
    const sheetName = event.target.value;
    if (boxWorkbook && sheetName) {
        populateColumnDropdown(boxWorkbook, sheetName, 'boxNumberColumn');
    }
}

// UI Management functions
function showProgress() {
    document.getElementById('progressContainer').style.display = 'block';
    document.getElementById('progressText').style.display = 'block';
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
}

function displayResults(results) {
    window.searchResults = results; // Ensure export always has access to the latest results
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
                ['Matching Logic', summary.matchingLogic],
                ['Custom Match Count', summary.customMatchCount || 'N/A'],
                ['Search Columns', summary.searchColumns.join(', ')],
                ['Target Column', summary.targetColumn],
                [''],
                ['Generated on', new Date().toLocaleString()]
            ];
            
            const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(newWorkbook, summarySheet, 'Summary');
            
            const filename = `advanced_search_results_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(newWorkbook, filename);
            
        } catch (error) {
            console.error('[exportResults] Error:', error);
            alert('Error exporting results: ' + error.message);
        }
    } else {
        alert('No search results available for export');
    }
}

function calculateRowSimilarity(searchValues, targetValue, searchType, caseSensitive, matchingLogic = 'all', customMatchCount = null) {
    if (searchValues.length === 0) return 0;
    
    if (searchType === 'exact') {
        // For exact match, check based on matching logic
        const targetStr = caseSensitive ? targetValue : targetValue.toLowerCase();
        let matchCount = 0;
        
        for (const searchValue of searchValues) {
            if (!searchValue) continue;
            const searchStr = caseSensitive ? searchValue : searchValue.toLowerCase();
            if (targetStr.includes(searchStr)) {
                matchCount++;
            }
        }
        
        // Apply matching logic
        if (matchingLogic === 'all') {
            return matchCount === searchValues.filter(v => v).length ? 1.0 : 0;
        } else if (matchingLogic === 'any') {
            return matchCount > 0 ? 1.0 : 0;
        } else if (matchingLogic === 'custom') {
            const requiredMatches = customMatchCount || Math.ceil(searchValues.filter(v => v).length / 2);
            return matchCount >= requiredMatches ? 1.0 : 0;
        }
        
        return 0;
    } else {
        // For partial match, calculate individual similarities first
        const similarities = [];
        for (const searchValue of searchValues) {
            if (!searchValue) continue;
            const similarity = calculateSimilarity(searchValue, targetValue, caseSensitive);
            similarities.push(similarity);
        }
        
        if (similarities.length === 0) return 0;
        
        // Apply matching logic to determine overall similarity
        if (matchingLogic === 'all') {
            // All must be above threshold - return minimum similarity
            return Math.min(...similarities);
        } else if (matchingLogic === 'any') {
            // At least one must be above threshold - return maximum similarity
            return Math.max(...similarities);
        } else if (matchingLogic === 'custom') {
            // Custom logic - return average of top N matches
            const requiredMatches = customMatchCount || Math.ceil(similarities.length / 2);
            const sortedSimilarities = similarities.sort((a, b) => b - a);
            const topMatches = sortedSimilarities.slice(0, requiredMatches);
            return topMatches.reduce((sum, sim) => sum + sim, 0) / topMatches.length;
        }
        
        return 0;
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

// Cancel and next batch event handlers
function setupBatchControlHandlers() {
    const cancelBtn = document.getElementById('cancelBtn');
    const nextBatchBtn = document.getElementById('nextBatchBtn');
    if (cancelBtn) {
        cancelBtn.onclick = function() {
            batchProcessing.cancelRequested = true;
            hideProgress();
            resetBatchProcessingState();
        };
    }
    if (nextBatchBtn) {
        nextBatchBtn.onclick = function() {
            if (batchProcessing.isPaused && batchProcessing.isActive && !batchProcessing.cancelRequested) {
                batchProcessing.isPaused = false;
                showBatchControls(true, false);
                setTimeout(processBatchInternal, 0);
            }
        };
    }
}

// Add this function to trigger the search process from the UI
function performAdvancedSearch() {
    // Validate workbooks and selections
    if (!searchSourceWorkbook || !searchTargetWorkbook) {
        alert('Please upload both source and target Excel files.');
        return;
    }
    const sourceSheetName = document.getElementById('searchSourceSheet').value;
    const targetSheetName = document.getElementById('searchTargetSheet').value;
    if (!sourceSheetName || !targetSheetName) {
        alert('Please select both source and target sheets.');
        return;
    }
    const targetColumnLetter = document.getElementById('searchTargetColumn').value;
    if (!targetColumnLetter) {
        alert('Please select a target column to search in.');
        return;
    }
    // Collect search columns
    const searchColumnSelects = document.querySelectorAll('.search-column-select');
    const searchColumnLetters = Array.from(searchColumnSelects).map(sel => sel.value).filter(Boolean);
    if (searchColumnLetters.length === 0) {
        alert('Please select at least one search column from the source sheet.');
        return;
    }
    // Get search type and threshold
    const searchType = document.querySelector('input[name="searchType"]:checked').value;
    let similarityThreshold = parseFloat(document.getElementById('similarityThreshold').value);
    if (searchType === 'exact') similarityThreshold = 1.0;
    const caseSensitive = document.getElementById('searchCaseSensitive').checked;
    
    // Get matching logic
    const matchingLogic = document.querySelector('input[name="matchingLogic"]:checked').value;
    const customMatchCount = matchingLogic === 'custom' ? parseInt(document.getElementById('customMatchCount').value) : null;
    
    // Batch mode: incremental if file is large (or user wants it)
    const incremental = batchProcessing.incremental || false;
    // Prepare data
    const sourceSheet = searchSourceWorkbook.Sheets[sourceSheetName];
    const targetSheet = searchTargetWorkbook.Sheets[targetSheetName];
    const sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1 });
    const targetData = XLSX.utils.sheet_to_json(targetSheet, { header: 1 });
    // Map column letters to indices
    const sourceColumnIndices = searchColumnLetters.map(letter => XLSX.utils.decode_col(letter));
    const targetColumnIndex = XLSX.utils.decode_col(targetColumnLetter);
    
    // Get all source columns for export
    const sourceRange = XLSX.utils.decode_range(sourceSheet['!ref']);
    const allSourceColumns = [];
    for (let col = sourceRange.s.c; col <= sourceRange.e.c; col++) {
        allSourceColumns.push(col);
    }
    
    // Show progress
    showProgress();
    updateProgress(0, 'Starting search...');
    // Start batch search
    executeSearchInBatches({
        sourceData,
        targetData,
        sourceColumnIndices,
        targetColumnIndex,
        searchType,
        similarityThreshold,
        caseSensitive,
        onProgress: (percent) => {
            updateProgress(percent, `Processing... (${Math.round(percent)}%)`);
        },
        onComplete: (results, done) => {
            if (done) displayResults(results);
        },
        incremental: batchProcessing.incremental || false,
        matchingLogic,
        customMatchCount,
        allSourceColumns
    });
}

// Duplicate Analysis Functions
function performDuplicateAnalysis() {
    // Validate workbook and selections
    if (!duplicateWorkbook) {
        alert('Please upload an Excel file.');
        return;
    }
    
    const sheetName = document.getElementById('duplicateSheet').value;
    if (!sheetName) {
        alert('Please select a sheet.');
        return;
    }
    
    // Collect columns to check
    const columnSelects = document.querySelectorAll('.duplicate-column-select');
    const columnLetters = Array.from(columnSelects).map(sel => sel.value).filter(Boolean);
    if (columnLetters.length === 0) {
        alert('Please select at least one column to check for duplicates.');
        return;
    }
    
    // Prepare data
    const sheet = duplicateWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    if (data.length <= 1) {
        alert('The selected sheet has no data or only headers.');
        return;
    }
    
    // Map column letters to indices
    const columnIndices = columnLetters.map(letter => XLSX.utils.decode_col(letter));
    
    // Show progress
    showDuplicateProgress();
    updateDuplicateProgress(0, 'Starting duplicate analysis...');
    
    // Perform duplicate analysis
    setTimeout(() => {
        const results = analyzeDuplicates(data, columnIndices);
        displayDuplicateResults(results);
    }, 100);
}

function analyzeDuplicates(data, columnIndices) {
    const results = [];
    const duplicateMap = new Map(); // Maps row key to array of row indices
    const duplicateRows = new Set(); // Set of row indices that are duplicates
    
    // Create headers row
    const headers = ['Row', 'Status'];
    for (let col = 0; col < data[0].length; col++) {
        headers.push(`Column_${XLSX.utils.encode_col(col)}`);
    }
    results.push(headers);
    
    // Analyze each row
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        
        // Create key from selected columns
        const key = columnIndices.map(colIndex => 
            row[colIndex] ? String(row[colIndex]).trim() : ''
        ).join('|');
        
        // Skip empty keys
        if (!key || key.split('|').every(val => !val)) {
            const resultRow = [i + 1, 'No Data', ...row.map(cell => cell || '')];
            results.push(resultRow);
            continue;
        }
        
        // Check if this key already exists
        if (duplicateMap.has(key)) {
            // This is a duplicate
            duplicateRows.add(i);
            duplicateMap.get(key).push(i);
        } else {
            // First occurrence of this key
            duplicateMap.set(key, [i]);
        }
        
        // Add row to results
        const status = duplicateMap.get(key).length > 1 ? 'DUPLICATE' : 'UNIQUE';
        const resultRow = [i + 1, status, ...row.map(cell => cell || '')];
        results.push(resultRow);
        
        // Update progress
        if (i % 100 === 0) {
            updateDuplicateProgress((i / data.length) * 100, `Analyzing row ${i}...`);
        }
    }
    
    // Mark all rows in duplicate groups as DUPLICATE
    for (const [key, rowIndices] of duplicateMap) {
        if (rowIndices.length > 1) {
            for (const rowIndex of rowIndices) {
                results[rowIndex][1] = 'DUPLICATE'; // Update status column
            }
        }
    }
    
    // Calculate statistics
    const totalRows = data.length - 1;
    const duplicateCount = duplicateRows.size;
    const uniqueCount = totalRows - duplicateCount;
    const duplicateGroups = Array.from(duplicateMap.values()).filter(indices => indices.length > 1).length;
    
    return {
        data: results,
        summary: {
            totalRows,
            duplicateCount,
            uniqueCount,
            duplicateGroups,
            checkedColumns: columnIndices.map(idx => XLSX.utils.encode_col(idx))
        }
    };
}

function displayDuplicateResults(results) {
    window.duplicateResults = results;
    const summary = results.summary;
    
    // Create summary stats
    const statsHtml = `
        <div class="stat-card">
            <div class="stat-number">${summary.totalRows}</div>
            <div class="stat-label">Total Rows</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.uniqueCount}</div>
            <div class="stat-label">Unique Rows</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.duplicateCount}</div>
            <div class="stat-label">Duplicate Rows</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.duplicateGroups}</div>
            <div class="stat-label">Duplicate Groups</div>
        </div>
    `;
    
    document.getElementById('duplicateSummaryStats').innerHTML = statsHtml;
    document.getElementById('duplicateSummary').style.display = 'block';
    document.getElementById('exportDuplicateBtn').style.display = 'inline-flex';
    
    setTimeout(() => {
        hideDuplicateProgress();
    }, 1000);
}

function exportDuplicateResults() {
    if (window.duplicateResults) {
        try {
            const newWorkbook = XLSX.utils.book_new();
            
            // Create main results sheet
            const duplicateSheet = XLSX.utils.aoa_to_sheet(window.duplicateResults.data);
            XLSX.utils.book_append_sheet(newWorkbook, duplicateSheet, 'Duplicate Analysis');
            
            // Create summary sheet
            const summary = window.duplicateResults.summary;
            const summaryData = [
                ['Duplicate Analysis Summary'],
                [''],
                ['Total Rows', summary.totalRows],
                ['Unique Rows', summary.uniqueCount],
                ['Duplicate Rows', summary.duplicateCount],
                ['Duplicate Groups', summary.duplicateGroups],
                [''],
                ['Configuration'],
                ['Checked Columns', summary.checkedColumns.join(', ')],
                [''],
                ['Generated on', new Date().toLocaleString()]
            ];
            
            const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(newWorkbook, summarySheet, 'Summary');
            
            const filename = `duplicate_analysis_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(newWorkbook, filename);
            
        } catch (error) {
            console.error('[exportDuplicateResults] Error:', error);
            alert('Error exporting results: ' + error.message);
        }
    } else {
        alert('No duplicate analysis results available for export');
    }
}

// Box Sequence Validation Functions
function performBoxValidation() {
    // Validate workbook and selections
    if (!boxWorkbook) {
        alert('Please upload an Excel file.');
        return;
    }
    
    const sheetName = document.getElementById('boxSheet').value;
    if (!sheetName) {
        alert('Please select a sheet.');
        return;
    }
    
    const boxColumnLetter = document.getElementById('boxNumberColumn').value;
    if (!boxColumnLetter) {
        alert('Please select a box number column.');
        return;
    }
    
    // Prepare data
    const sheet = boxWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    if (data.length <= 1) {
        alert('The selected sheet has no data or only headers.');
        return;
    }
    
    // Map column letter to index
    const boxColumnIndex = XLSX.utils.decode_col(boxColumnLetter);
    
    // Show progress
    showBoxProgress();
    updateBoxProgress(0, 'Starting box sequence validation...');
    
    // Perform box validation
    setTimeout(() => {
        const results = validateBoxSequence(data, boxColumnIndex);
        displayBoxResults(results);
    }, 100);
}

function validateBoxSequence(data, boxColumnIndex) {
    const results = [];
    const issues = [];
    let currentBox = null;
    let expectedBox = null;
    let rowInBlock = 0;
    
    // Create headers row
    const headers = ['Row', 'Box Number', 'Status', 'Issue Description'];
    for (let col = 0; col < data[0].length; col++) {
        headers.push(`Column_${XLSX.utils.encode_col(col)}`);
    }
    results.push(headers);
    
    // Analyze each row
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        
        const boxNumber = row[boxColumnIndex] ? String(row[boxColumnIndex]).trim() : '';
        let status = 'OK';
        let issueDescription = '';
        
        // Check if row is empty (separator)
        const isEmptyRow = row.every(cell => !cell || String(cell).trim() === '');
        
        if (isEmptyRow) {
            // Empty row - reset block tracking
            currentBox = null;
            expectedBox = null;
            rowInBlock = 0;
            status = 'SEPARATOR';
            issueDescription = 'Empty row separator';
        } else if (boxNumber) {
            // Row has box number
            if (currentBox === null) {
                // Start of new block
                currentBox = boxNumber;
                expectedBox = boxNumber;
                rowInBlock = 1;
            } else {
                // Continuation of block
                rowInBlock++;
                
                if (boxNumber !== expectedBox) {
                    // Box number doesn't match expected
                    status = 'ISSUE';
                    issueDescription = `Expected box ${expectedBox}, found ${boxNumber}`;
                    issues.push({
                        row: i + 1,
                        expected: expectedBox,
                        found: boxNumber,
                        description: issueDescription
                    });
                }
            }
        } else {
            // Row has no box number but is not empty
            if (currentBox !== null) {
                // Should have box number in this block
                status = 'ISSUE';
                issueDescription = `Missing box number in block ${currentBox}`;
                issues.push({
                    row: i + 1,
                    expected: currentBox,
                    found: 'MISSING',
                    description: issueDescription
                });
            }
        }
        
        // Add row to results
        const resultRow = [i + 1, boxNumber || '', status, issueDescription, ...row.map(cell => cell || '')];
        results.push(resultRow);
        
        // Update progress
        if (i % 100 === 0) {
            updateBoxProgress((i / data.length) * 100, `Validating row ${i}...`);
        }
    }
    
    // Calculate statistics
    const totalRows = data.length - 1;
    const issueCount = issues.length;
    const okCount = totalRows - issueCount;
    
    return {
        data: results,
        summary: {
            totalRows,
            issueCount,
            okCount,
            boxColumn: XLSX.utils.encode_col(boxColumnIndex),
            issues: issues
        }
    };
}

function displayBoxResults(results) {
    window.boxResults = results;
    const summary = results.summary;
    
    // Create summary stats
    const statsHtml = `
        <div class="stat-card">
            <div class="stat-number">${summary.totalRows}</div>
            <div class="stat-label">Total Rows</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.okCount}</div>
            <div class="stat-label">Valid Rows</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.issueCount}</div>
            <div class="stat-label">Issues Found</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.boxColumn}</div>
            <div class="stat-label">Box Column</div>
        </div>
    `;
    
    document.getElementById('boxSummaryStats').innerHTML = statsHtml;
    document.getElementById('boxSummary').style.display = 'block';
    document.getElementById('exportBoxBtn').style.display = 'inline-flex';
    
    setTimeout(() => {
        hideBoxProgress();
    }, 1000);
}

function exportBoxResults() {
    if (window.boxResults) {
        try {
            const newWorkbook = XLSX.utils.book_new();
            
            // Create main results sheet
            const boxSheet = XLSX.utils.aoa_to_sheet(window.boxResults.data);
            XLSX.utils.book_append_sheet(newWorkbook, boxSheet, 'Box Validation');
            
            // Create summary sheet
            const summary = window.boxResults.summary;
            const summaryData = [
                ['Box Sequence Validation Summary'],
                [''],
                ['Total Rows', summary.totalRows],
                ['Valid Rows', summary.okCount],
                ['Issues Found', summary.issueCount],
                ['Box Column', summary.boxColumn],
                [''],
                ['Issues Details']
            ];
            
            // Add issue details
            if (summary.issues.length > 0) {
                summaryData.push(['Row', 'Expected', 'Found', 'Description']);
                summary.issues.forEach(issue => {
                    summaryData.push([issue.row, issue.expected, issue.found, issue.description]);
                });
            } else {
                summaryData.push(['No issues found']);
            }
            
            summaryData.push(['']);
            summaryData.push(['Generated on', new Date().toLocaleString()]);
            
            const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(newWorkbook, summarySheet, 'Summary');
            
            const filename = `box_validation_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(newWorkbook, filename);
            
        } catch (error) {
            console.error('[exportBoxResults] Error:', error);
            alert('Error exporting results: ' + error.message);
        }
    } else {
        alert('No box validation results available for export');
    }
}

// Progress Management for Duplicate Identifier
function showDuplicateProgress() {
    document.getElementById('duplicateProgressContainer').style.display = 'block';
    document.getElementById('duplicateProgressText').style.display = 'block';
    document.getElementById('duplicateSummary').style.display = 'none';
    document.getElementById('exportDuplicateBtn').style.display = 'none';
}

function hideDuplicateProgress() {
    document.getElementById('duplicateProgressContainer').style.display = 'none';
    document.getElementById('duplicateProgressText').style.display = 'none';
}

function updateDuplicateProgress(percentage, message) {
    const progressFill = document.getElementById('duplicateProgressFill');
    const progressText = document.getElementById('duplicateProgressText');
    
    if (progressFill) {
        progressFill.style.width = `${percentage}%`;
    }
    if (progressText) {
        progressText.textContent = message || 'Processing...';
    }
}

// Progress Management for Box Validator
function showBoxProgress() {
    document.getElementById('boxProgressContainer').style.display = 'block';
    document.getElementById('boxProgressText').style.display = 'block';
    document.getElementById('boxSummary').style.display = 'none';
    document.getElementById('exportBoxBtn').style.display = 'none';
}

function hideBoxProgress() {
    document.getElementById('boxProgressContainer').style.display = 'none';
    document.getElementById('boxProgressText').style.display = 'none';
}

function updateBoxProgress(percentage, message) {
    const progressFill = document.getElementById('boxProgressFill');
    const progressText = document.getElementById('boxProgressText');
    
    if (progressFill) {
        progressFill.style.width = `${percentage}%`;
    }
    if (progressText) {
        progressText.textContent = message || 'Processing...';
    }
}

// Insert Empty Rows by Box Number - Core Logic
function performInsertRows() {
    if (!insertRowsWorkbook) {
        alert('Please upload an Excel file.');
        return;
    }
    const sheetName = document.getElementById('insertRowsSheet').value;
    if (!sheetName) {
        alert('Please select a sheet.');
        return;
    }
    const boxColumnLetter = document.getElementById('insertRowsBoxColumn').value;
    if (!boxColumnLetter) {
        alert('Please select a box number column.');
        return;
    }
    const sheet = insertRowsWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (data.length <= 1) {
        alert('The selected sheet has no data or only headers.');
        return;
    }
    const boxColIdx = XLSX.utils.decode_col(boxColumnLetter);
    showInsertRowsProgress();
    updateInsertRowsProgress(0, 'Processing...');
    setTimeout(() => {
        const results = insertEmptyRowsByBoxNumber(data, boxColIdx);
        displayInsertRowsResults(results);
    }, 100);
}

function insertEmptyRowsByBoxNumber(data, boxColIdx) {
    const newData = [];
    let prevBox = null;
    let insertedCount = 0;
    // Always keep header
    newData.push(data[0]);
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const box = row[boxColIdx] ? String(row[boxColIdx]).trim() : '';
        // If box number changes and previous row is not empty, insert an empty row
        if (prevBox !== null && box && box !== prevBox) {
            // Check if previous row is already empty
            const prevRow = newData[newData.length - 1];
            const isPrevEmpty = prevRow.every(cell => !cell || String(cell).trim() === '');
            if (!isPrevEmpty) {
                // Insert empty row
                newData.push(Array(row.length).fill(''));
                insertedCount++;
            }
        }
        newData.push(row);
        prevBox = box;
        if (i % 100 === 0) {
            updateInsertRowsProgress((i / data.length) * 100, `Processing row ${i}...`);
        }
    }
    return {
        data: newData,
        summary: {
            totalRows: data.length - 1,
            insertedRows: insertedCount,
            boxColumn: XLSX.utils.encode_col(boxColIdx)
        }
    };
}

function displayInsertRowsResults(results) {
    window.insertRowsResults = results;
    const summary = results.summary;
    const statsHtml = `
        <div class="stat-card">
            <div class="stat-number">${summary.totalRows}</div>
            <div class="stat-label">Original Rows</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.insertedRows}</div>
            <div class="stat-label">Empty Rows Inserted</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${summary.boxColumn}</div>
            <div class="stat-label">Box Column</div>
        </div>
    `;
    document.getElementById('insertRowsSummaryStats').innerHTML = statsHtml;
    document.getElementById('insertRowsSummary').style.display = 'block';
    document.getElementById('exportInsertRowsBtn').style.display = 'inline-flex';
    setTimeout(() => {
        hideInsertRowsProgress();
    }, 1000);
}

function exportInsertRowsResults() {
    if (window.insertRowsResults) {
        try {
            const newWorkbook = XLSX.utils.book_new();
            const sheet = XLSX.utils.aoa_to_sheet(window.insertRowsResults.data);
            XLSX.utils.book_append_sheet(newWorkbook, sheet, 'With Empty Rows');
            // Add summary sheet
            const summary = window.insertRowsResults.summary;
            const summaryData = [
                ['Insert Empty Rows Summary'],
                [''],
                ['Original Rows', summary.totalRows],
                ['Empty Rows Inserted', summary.insertedRows],
                ['Box Column', summary.boxColumn],
                [''],
                ['Generated on', new Date().toLocaleString()]
            ];
            const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(newWorkbook, summarySheet, 'Summary');
            const filename = `with_empty_rows_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(newWorkbook, filename);
        } catch (error) {
            console.error('[exportInsertRowsResults] Error:', error);
            alert('Error exporting results: ' + error.message);
        }
    } else {
        alert('No results available for export');
    }
}

// Progress Management for Insert Empty Rows
function showInsertRowsProgress() {
    document.getElementById('insertRowsProgressContainer').style.display = 'block';
    document.getElementById('insertRowsProgressText').style.display = 'block';
    document.getElementById('insertRowsSummary').style.display = 'none';
    document.getElementById('exportInsertRowsBtn').style.display = 'none';
}
function hideInsertRowsProgress() {
    document.getElementById('insertRowsProgressContainer').style.display = 'none';
    document.getElementById('insertRowsProgressText').style.display = 'none';
}
function updateInsertRowsProgress(percentage, message) {
    const progressFill = document.getElementById('insertRowsProgressFill');
    const progressText = document.getElementById('insertRowsProgressText');
    if (progressFill) {
        progressFill.style.width = `${percentage}%`;
    }
    if (progressText) {
        progressText.textContent = message || 'Processing...';
    }
}

// Initialize when page loads
document.addEventListener('DOMContentLoaded', function() {
    initializeAdvancedSearch();
    
    // Initialize search type handler
    handleSearchTypeChange({ target: { value: 'exact' } });
    
    // Initialize similarity display
    handleSimilarityChange({ target: { value: 0.8 } });
    
    // Initialize matching logic
    handleMatchingLogicChange({ target: { value: 'all' } });
    updateTotalSearchColumns();
    
    setupBatchControlHandlers();
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