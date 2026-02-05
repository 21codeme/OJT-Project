// Set current date
document.addEventListener('DOMContentLoaded', async function() {
    const now = new Date();
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                   'July', 'August', 'September', 'October', 'November', 'December'];
    const dateStr = `${months[now.getMonth()]} ${now.getFullYear()}`;
    document.getElementById('currentDate').textContent = dateStr;
    
    // Enable add buttons even when no data
    document.getElementById('addItemBtn').disabled = false;
    document.getElementById('addPCBtn').disabled = false;
    
    // Check Supabase connection after a short delay to ensure config.js is loaded
    setTimeout(async () => {
        // Try to initialize Supabase if not already done
        if (typeof initSupabase === 'function') {
            initSupabase();
        }
        
        // Wait a bit more for Supabase to initialize
        setTimeout(async () => {
            if (checkSupabaseConnection()) {
                console.log('Supabase connected, loading data...');
                await loadFromSupabase();
            } else {
                console.log('Supabase not configured, using local storage only');
            }
        }, 300);
    }, 200);
});

// Sheet management
let sheets = {
    'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [],
        hasData: false
    }
};
let currentSheetId = 'sheet-1';
let sheetCounter = 1;
let originalWorkbook = null;

// Get current sheet data
function getCurrentSheet() {
    return sheets[currentSheetId];
}

// Set current sheet data
function setCurrentSheetData(data, hasDataFlag = true) {
    sheets[currentSheetId].data = data;
    sheets[currentSheetId].hasData = hasDataFlag;
}

// File input handler
document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                originalWorkbook = workbook;
                
                // Get the first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON with header row
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1, 
                    defval: '',
                    raw: false 
                });
                
                // Import all sheets from workbook
                workbook.SheetNames.forEach((sheetName, index) => {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                        header: 1, 
                        defval: '',
                        raw: false 
                    });
                    
                    if (index === 0) {
                        // Load first sheet into current sheet
                        setCurrentSheetData(jsonData, true);
                        displayData(jsonData);
                    } else {
                        // Create new sheets for other sheets
                        createNewSheet(sheetName, jsonData);
                    }
                });
                
                // Enable export and clear buttons
                document.getElementById('exportBtn').disabled = false;
                document.getElementById('clearBtn').disabled = false;
            } catch (error) {
                alert('Error reading Excel file: ' + error.message);
                console.error(error);
            }
        };
        reader.readAsArrayBuffer(file);
    }
});

// Apply condition-based row color
function applyConditionColor(row, condition) {
    // Remove existing condition classes
    row.classList.remove('condition-borrowed', 'condition-unserviceable');
    
    if (condition === 'Borrowed') {
        row.classList.add('condition-borrowed');
    } else if (condition === 'Unserviceable') {
        row.classList.add('condition-unserviceable');
    }
}

// Create editable cell
function createEditableCell(value, isPCHeader = false, cellIndex = -1, row = null) {
    const td = document.createElement('td');
    td.classList.add('editable');
    
    const input = document.createElement('input');
    input.type = 'text';
    input.value = value || '';
    
    // Auto uppercase for Article/It (index 0) and Description (index 1)
    const isAutoUppercase = cellIndex === 0 || cellIndex === 1;
    if (isAutoUppercase) {
        input.style.textTransform = 'uppercase';
        input.addEventListener('input', function() {
            // Convert to uppercase in real-time
            const cursorPos = this.selectionStart;
            this.value = this.value.toUpperCase();
            // Restore cursor position
            this.setSelectionRange(cursorPos, cursorPos);
        });
        input.addEventListener('paste', function(e) {
            e.preventDefault();
            const pastedText = (e.clipboardData || window.clipboardData).getData('text').toUpperCase();
            const start = this.selectionStart;
            const end = this.selectionEnd;
            this.value = this.value.substring(0, start) + pastedText + this.value.substring(end);
            // Restore cursor position
            const newPos = start + pastedText.length;
            this.setSelectionRange(newPos, newPos);
        });
    }
    
    input.addEventListener('blur', function() {
        // Ensure uppercase on blur for Article/It and Description
        if (isAutoUppercase) {
            this.value = this.value.toUpperCase();
        }
        updateDataFromTable();
        // If this is the condition cell (index 7), update row color
        if (cellIndex === 7 && row) {
            applyConditionColor(row, input.value);
        }
    });
    input.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            this.blur();
        }
    });
    
    td.appendChild(input);
    
    if (isPCHeader) {
        td.style.fontWeight = 'bold';
        input.style.fontWeight = 'bold';
    }
    
    return td;
}

// Create action cell with delete and highlight buttons
function createActionCell() {
    const td = document.createElement('td');
    td.classList.add('actions-cell');
    
    const buttonContainer = document.createElement('div');
    buttonContainer.style.display = 'flex';
    buttonContainer.style.gap = '5px';
    buttonContainer.style.flexWrap = 'wrap';
    buttonContainer.style.justifyContent = 'center';
    buttonContainer.style.alignItems = 'center';
    
    // Highlight button
    const highlightBtn = document.createElement('button');
    highlightBtn.className = 'highlight-btn';
    highlightBtn.textContent = '⭐ Highlight';
    highlightBtn.addEventListener('click', function() {
        const tr = this.closest('tr');
        if (tr && !tr.classList.contains('pc-header-row') && !tr.classList.contains('empty-row')) {
            if (tr.classList.contains('highlighted-row')) {
                tr.classList.remove('highlighted-row');
                highlightBtn.textContent = '⭐ Highlight';
                highlightBtn.style.background = '#ffc107';
            } else {
                tr.classList.add('highlighted-row');
                highlightBtn.textContent = '✅ Highlighted';
                highlightBtn.style.background = '#28a745';
            }
            updateDataFromTable();
        }
    });
    
    // Delete button
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'delete-btn';
    deleteBtn.textContent = '🗑️ Delete';
    deleteBtn.addEventListener('click', function() {
        if (confirm('Are you sure you want to delete this row?')) {
            const tr = this.closest('tr');
            if (tr) {
                tr.remove();
                updateDataFromTable();
                
                // Check if table is now empty
                const tbody = document.getElementById('tableBody');
                if (tbody.children.length === 0) {
                    tbody.innerHTML = '<tr class="empty-row"><td colspan="12" class="empty-message">No data loaded. Please import an Excel file or add items manually.</td></tr>';
                    setCurrentSheetData([], false);
                    document.getElementById('exportBtn').disabled = !hasAnyData();
                    document.getElementById('clearBtn').disabled = true;
                } else {
                    updateDataFromTable();
                }
            }
        }
    });
    
    buttonContainer.appendChild(highlightBtn);
    buttonContainer.appendChild(deleteBtn);
    td.appendChild(buttonContainer);
    return td;
}

// Add new item row
function addItemRow(isPCHeader = false) {
    const tbody = document.getElementById('tableBody');
    
    // Remove empty message if present
    const emptyRow = tbody.querySelector('.empty-row');
    if (emptyRow) {
        emptyRow.remove();
    }
    
    const tr = document.createElement('tr');
    if (isPCHeader) {
        tr.classList.add('pc-header-row');
    }
    
    // Create 10 editable cells
    for (let j = 0; j < 10; j++) {
        let defaultValue = '';
        if (isPCHeader && j === 0) {
            defaultValue = 'PC ' + (tbody.children.length + 1);
        }
        const td = createEditableCell(defaultValue, isPCHeader && j === 0, j, tr);
        tr.appendChild(td);
    }
    
    // Add action cell
    const actionCell = createActionCell();
    tr.appendChild(actionCell);
    
    tbody.appendChild(tr);
    setCurrentSheetData(getCurrentSheet().data, true);
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('clearBtn').disabled = false;
    
    // Focus on first input
    const firstInput = tr.querySelector('td input');
    if (firstInput) {
        firstInput.focus();
    }
}

// Modal functionality
const modal = document.getElementById('addItemModal');
const addItemForm = document.getElementById('addItemForm');
const closeModal = document.querySelector('.close-modal');
const cancelBtn = document.getElementById('cancelAddBtn');

// Auto uppercase for Article/It and Description in form
document.addEventListener('DOMContentLoaded', function() {
    const articleInput = document.getElementById('article');
    const descriptionInput = document.getElementById('description');
    
    if (articleInput) {
        articleInput.style.textTransform = 'uppercase';
        articleInput.addEventListener('input', function() {
            const cursorPos = this.selectionStart;
            this.value = this.value.toUpperCase();
            this.setSelectionRange(cursorPos, cursorPos);
        });
        articleInput.addEventListener('paste', function(e) {
            e.preventDefault();
            const pastedText = (e.clipboardData || window.clipboardData).getData('text').toUpperCase();
            const start = this.selectionStart;
            const end = this.selectionEnd;
            this.value = this.value.substring(0, start) + pastedText + this.value.substring(end);
            const newPos = start + pastedText.length;
            this.setSelectionRange(newPos, newPos);
        });
    }
    
    if (descriptionInput) {
        descriptionInput.style.textTransform = 'uppercase';
        descriptionInput.addEventListener('input', function() {
            const cursorPos = this.selectionStart;
            this.value = this.value.toUpperCase();
            this.setSelectionRange(cursorPos, cursorPos);
        });
        descriptionInput.addEventListener('paste', function(e) {
            e.preventDefault();
            const pastedText = (e.clipboardData || window.clipboardData).getData('text').toUpperCase();
            const start = this.selectionStart;
            const end = this.selectionEnd;
            this.value = this.value.substring(0, start) + pastedText + this.value.substring(end);
            const newPos = start + pastedText.length;
            this.setSelectionRange(newPos, newPos);
        });
    }
});

// Open modal when Add Item button is clicked
document.getElementById('addItemBtn').addEventListener('click', function() {
    modal.classList.add('show');
    // Focus on first input
    document.getElementById('article').focus();
});

// Picture preview functionality
const pictureInput = document.getElementById('picture');
const picturePreview = document.getElementById('picturePreview');
let selectedImageData = null;

pictureInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        if (file.type.startsWith('image/')) {
            const reader = new FileReader();
            reader.onload = function(e) {
                selectedImageData = e.target.result; // Base64 data
                picturePreview.innerHTML = `<img src="${selectedImageData}" alt="Preview">`;
                picturePreview.classList.remove('empty');
            };
            reader.readAsDataURL(file);
        } else {
            alert('Please select an image file.');
            pictureInput.value = '';
            picturePreview.innerHTML = '<span class="empty">No image selected</span>';
            picturePreview.classList.add('empty');
        }
    } else {
        selectedImageData = null;
        picturePreview.innerHTML = '<span class="empty">No image selected</span>';
        picturePreview.classList.add('empty');
    }
});

// Close modal
function closeModalFunc() {
    modal.classList.remove('show');
    addItemForm.reset();
    selectedImageData = null;
    picturePreview.innerHTML = '<span class="empty">No image selected</span>';
    picturePreview.classList.add('empty');
}

closeModal.addEventListener('click', closeModalFunc);
cancelBtn.addEventListener('click', closeModalFunc);

// Close modal when clicking outside
window.addEventListener('click', function(event) {
    if (event.target === modal) {
        closeModalFunc();
    }
});

// Handle form submission
addItemForm.addEventListener('submit', function(e) {
    e.preventDefault();
    
    // Get form values
    const formData = {
        article: document.getElementById('article').value.trim().toUpperCase(),
        description: document.getElementById('description').value.trim().toUpperCase(),
        oldProperty: document.getElementById('oldProperty').value.trim(),
        unitOfMeas: document.getElementById('unitOfMeas').value.trim(),
        unitValue: document.getElementById('unitValue').value.trim(),
        quantity: document.getElementById('quantity').value.trim(),
        location: document.getElementById('location').value.trim(),
        condition: document.getElementById('condition').value.trim(),
        remarks: document.getElementById('remarks').value.trim(),
        user: document.getElementById('user').value.trim(),
        picture: selectedImageData
    };
    
    // Add item to table
    addItemFromForm(formData);
    
    // Close modal and reset form
    closeModalFunc();
});

// Create picture cell
function createPictureCell(imageData = null) {
    const td = document.createElement('td');
    td.classList.add('picture-cell');
    
    if (imageData) {
        const img = document.createElement('img');
        img.src = imageData;
        img.alt = 'Item Picture';
        img.addEventListener('click', function() {
            showImageModal(imageData);
        });
        td.appendChild(img);
        
        // Add change picture button
        const changeBtn = document.createElement('button');
        changeBtn.className = 'upload-picture-btn';
        changeBtn.textContent = '📷 Change';
        changeBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'image/*';
            input.addEventListener('change', function(e) {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        img.src = e.target.result;
                        updateDataFromTable();
                    };
                    reader.readAsDataURL(file);
                }
            });
            input.click();
        });
        td.appendChild(changeBtn);
    } else {
        const noImage = document.createElement('span');
        noImage.className = 'no-image';
        noImage.textContent = 'No image';
        td.appendChild(noImage);
        
        const uploadBtn = document.createElement('button');
        uploadBtn.className = 'upload-picture-btn';
        uploadBtn.textContent = '📷 Upload';
        uploadBtn.addEventListener('click', function() {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'image/*';
            input.addEventListener('change', function(e) {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        td.innerHTML = '';
                        const img = document.createElement('img');
                        img.src = e.target.result;
                        img.alt = 'Item Picture';
                        img.addEventListener('click', function() {
                            showImageModal(e.target.result);
                        });
                        td.appendChild(img);
                        
                        const changeBtn = document.createElement('button');
                        changeBtn.className = 'upload-picture-btn';
                        changeBtn.textContent = '📷 Change';
                        changeBtn.addEventListener('click', function(e) {
                            e.stopPropagation();
                            const changeInput = document.createElement('input');
                            changeInput.type = 'file';
                            changeInput.accept = 'image/*';
                            changeInput.addEventListener('change', function(e) {
                                const file = e.target.files[0];
                                if (file) {
                                    const reader2 = new FileReader();
                                    reader2.onload = function(e) {
                                        img.src = e.target.result;
                                        updateDataFromTable();
                                    };
                                    reader2.readAsDataURL(file);
                                }
                            });
                            changeInput.click();
                        });
                        td.appendChild(changeBtn);
                        updateDataFromTable();
                    };
                    reader.readAsDataURL(file);
                }
            });
            input.click();
        });
        td.appendChild(uploadBtn);
    }
    
    return td;
}

// Show image in modal
function showImageModal(imageSrc) {
    let imageModal = document.getElementById('imageModal');
    if (!imageModal) {
        imageModal = document.createElement('div');
        imageModal.id = 'imageModal';
        imageModal.className = 'image-modal';
        const img = document.createElement('img');
        imageModal.appendChild(img);
        document.body.appendChild(imageModal);
        
        imageModal.addEventListener('click', function() {
            imageModal.classList.remove('show');
        });
    }
    imageModal.querySelector('img').src = imageSrc;
    imageModal.classList.add('show');
}

// Add item from form data
function addItemFromForm(formData) {
    const tbody = document.getElementById('tableBody');
    
    // Remove empty message if present
    const emptyRow = tbody.querySelector('.empty-row');
    if (emptyRow) {
        emptyRow.remove();
    }
    
    const tr = document.createElement('tr');
    
    // Create cells with form data
    const cells = [
        formData.article,
        formData.description,
        formData.oldProperty,
        formData.unitOfMeas,
        formData.unitValue,
        formData.quantity,
        formData.location,
        formData.condition,
        formData.remarks,
        formData.user
    ];
    
    cells.forEach((value, index) => {
        const td = createEditableCell(value, false, index, tr);
        tr.appendChild(td);
    });
    
    // Apply condition-based color
    applyConditionColor(tr, formData.condition);
    
    // Add picture cell
    const pictureCell = createPictureCell(formData.picture);
    tr.appendChild(pictureCell);
    
    // Add action cell
    const actionCell = createActionCell();
    tr.appendChild(actionCell);
    
    tbody.appendChild(tr);
    setCurrentSheetData(getCurrentSheet().data, true);
    document.getElementById('exportBtn').disabled = false;
    document.getElementById('clearBtn').disabled = false;
    
    // Update data
    updateDataFromTable();
}

// Add PC Section button
document.getElementById('addPCBtn').addEventListener('click', function() {
    const pcName = prompt('Enter PC Section Name (e.g., "PC USED BY: JOVEN T. CRUZ" or "SERVER"):');
    if (pcName !== null) {
        const tbody = document.getElementById('tableBody');
        const emptyRow = tbody.querySelector('.empty-row');
        if (emptyRow) {
            emptyRow.remove();
        }
        
        const tr = document.createElement('tr');
        tr.classList.add('pc-header-row');
        
        // Create first cell that spans all columns (12 columns total including Picture and Actions)
        const td = document.createElement('td');
        td.colSpan = 12;
        td.style.textAlign = 'center';
        td.style.fontWeight = 'bold';
        td.style.fontSize = '14px';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = pcName;
        input.style.width = '100%';
        input.style.textAlign = 'center';
        input.style.fontWeight = 'bold';
        input.style.fontSize = '14px';
        input.style.border = 'none';
        input.style.background = 'transparent';
        input.addEventListener('blur', function() {
            updateDataFromTable();
        });
        input.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                this.blur();
            }
        });
        
        td.appendChild(input);
        tr.appendChild(td);
        
        tbody.appendChild(tr);
        setCurrentSheetData(getCurrentSheet().data, true);
        document.getElementById('exportBtn').disabled = false;
        document.getElementById('clearBtn').disabled = false;
    }
});

// Display data in table
function displayData(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    if (!data || data.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="12" class="empty-message">No data found in the Excel file.</td></tr>';
        return;
    }
    
    // Find header row
    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(20, data.length); i++) {
        if (data[i] && data[i].length > 0) {
            const rowText = data[i].map(cell => cell ? cell.toString().toLowerCase() : '').join(' ');
            if (rowText.includes('article') && rowText.includes('description')) {
                headerRowIndex = i;
                break;
            }
        }
    }
    
    // Process all rows
    let dataRowIndex = 0; // Track index for highlight states (excluding PC headers)
    const sheet = getCurrentSheet();
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) {
            continue;
        }
        
        // Check if this is a header row
        const rowText = row.map(cell => cell ? cell.toString().toLowerCase() : '').join(' ');
        const isHeaderRow = rowText.includes('article') && rowText.includes('description');
        
        if (isHeaderRow) {
            continue; // Skip header row as it's already in thead
        }
        
        // Check if this is a PC header row
        const firstCell = row[0] ? row[0].toString().trim() : '';
        const isPCHeader = firstCell && (
            row.length === 1 || // Single cell = PC header
            firstCell.toUpperCase().includes('PC USED BY') ||
            firstCell.toUpperCase() === 'SERVER' ||
            firstCell.toUpperCase().startsWith('PC ') ||
            firstCell.toUpperCase() === 'PC'
        );
        
        const tr = document.createElement('tr');
        if (isPCHeader) {
            tr.classList.add('pc-header-row');
            
            // For PC headers, create a single cell that spans all columns
            const td = document.createElement('td');
            td.colSpan = 12;
            td.style.textAlign = 'center';
            td.style.fontWeight = 'bold';
            td.style.fontSize = '14px';
            
            const input = document.createElement('input');
            input.type = 'text';
            input.value = firstCell;
            input.style.width = '100%';
            input.style.textAlign = 'center';
            input.style.fontWeight = 'bold';
            input.style.fontSize = '14px';
            input.style.border = 'none';
            input.style.background = 'transparent';
            input.addEventListener('blur', function() {
                updateDataFromTable();
            });
            input.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    this.blur();
                }
            });
            
            td.appendChild(input);
            tr.appendChild(td);
        } else {
            // Create 10 editable cells for regular rows
            for (let j = 0; j < 10; j++) {
                const cellValue = row[j] !== undefined && row[j] !== null ? row[j].toString() : '';
                const td = createEditableCell(cellValue, false, j, tr);
                tr.appendChild(td);
            }
            
            // Apply condition-based color (condition is at index 7)
            const conditionValue = row[7] ? row[7].toString().trim() : '';
            applyConditionColor(tr, conditionValue);
            
            // Add picture cell (empty for imported data)
            const pictureCell = createPictureCell(null);
            tr.appendChild(pictureCell);
            
            // Add action cell
            const actionCell = createActionCell();
            tr.appendChild(actionCell);
            
            // Restore highlight state if exists
            if (sheet.highlightStates && sheet.highlightStates[dataRowIndex] === true) {
                tr.classList.add('highlighted-row');
                const highlightBtn = tr.querySelector('.highlight-btn');
                if (highlightBtn) {
                    highlightBtn.textContent = '✅ Highlighted';
                    highlightBtn.style.background = '#28a745';
                }
            }
            
            dataRowIndex++;
        }
        
        tbody.appendChild(tr);
    }
}

// Sheet Management Functions
function createNewSheet(name = null, data = null) {
    sheetCounter++;
    const sheetId = `sheet-${sheetCounter}`;
    const sheetName = name || `Sheet ${sheetCounter}`;
    
    sheets[sheetId] = {
        id: sheetId,
        name: sheetName,
        data: data || [],
        hasData: data ? true : false
    };
    
    // Create tab
    const sheetTabs = document.getElementById('sheetTabs');
    const tab = document.createElement('div');
    tab.className = 'sheet-tab';
    tab.setAttribute('data-sheet-id', sheetId);
    tab.innerHTML = `
        <span class="sheet-name">${sheetName}</span>
        <span class="close-sheet" data-sheet-id="${sheetId}">×</span>
    `;
    
    tab.addEventListener('click', function(e) {
        if (!e.target.classList.contains('close-sheet')) {
            switchToSheet(sheetId);
        }
    });
    
    const closeBtn = tab.querySelector('.close-sheet');
    closeBtn.addEventListener('click', function(e) {
        e.stopPropagation();
        deleteSheet(sheetId);
    });
    
    sheetTabs.appendChild(tab);
    switchToSheet(sheetId);
    
    return sheetId;
}

function switchToSheet(sheetId) {
    if (!sheets[sheetId]) return;
    
    // Save current sheet data
    saveCurrentSheetData();
    
    // Switch to new sheet
    currentSheetId = sheetId;
    
    // Update tabs
    document.querySelectorAll('.sheet-tab').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelector(`[data-sheet-id="${sheetId}"]`).classList.add('active');
    
    // Load sheet data
    const sheet = sheets[sheetId];
    displayData(sheet.data);
    
    // Update buttons
    document.getElementById('exportBtn').disabled = !hasAnyData();
    document.getElementById('clearBtn').disabled = !sheet.hasData;
}

function saveCurrentSheetData() {
    const tbody = document.getElementById('tableBody');
    const rows = tbody.querySelectorAll('tr:not(.empty-row)');
    
    const sheetData = [];
    const highlightStates = [];
    
    rows.forEach((row, index) => {
        // Check if this is a PC header row
        if (row.classList.contains('pc-header-row')) {
            // PC header row - get the text from the first cell (which spans all columns)
            const firstCell = row.querySelector('td');
            if (firstCell) {
                const input = firstCell.querySelector('input');
                const pcName = input ? input.value : firstCell.textContent.trim();
                // Store PC header as a single cell value (will be detected in export)
                sheetData.push([pcName]);
                highlightStates.push(false); // PC headers can't be highlighted
            }
        } else {
            // Regular row - get all editable cells
            const cells = row.querySelectorAll('td.editable');
            const rowData = [];
            cells.forEach(cell => {
                const input = cell.querySelector('input');
                rowData.push(input ? input.value : '');
            });
            if (rowData.length > 0) {
                sheetData.push(rowData);
                // Save highlight state
                highlightStates.push(row.classList.contains('highlighted-row'));
            }
        }
    });
    
    setCurrentSheetData(sheetData, sheetData.length > 0);
    // Store highlight states in sheet metadata
    sheets[currentSheetId].highlightStates = highlightStates;
    
    // Sync to Supabase if connected
    if (checkSupabaseConnection()) {
        syncToSupabase();
    }
}

// Supabase Sync Functions
async function syncToSupabase() {
    if (!checkSupabaseConnection()) return;
    
    try {
        const sheet = getCurrentSheet();
        
        // Save/update sheet in Supabase
        const { error: sheetError } = await supabase
            .from('sheets')
            .upsert({
                id: currentSheetId,
                name: sheet.name,
                updated_at: new Date().toISOString()
            }, { onConflict: 'id' });
        
        if (sheetError) {
            console.error('Error saving sheet to Supabase:', sheetError);
            return;
        }
        
        // Delete existing items for this sheet
        const { error: deleteError } = await supabase
            .from('inventory_items')
            .delete()
            .eq('sheet_id', currentSheetId);
        
        if (deleteError) {
            console.error('Error deleting old items:', deleteError);
            return;
        }
        
        // Prepare items to insert
        const itemsToInsert = [];
        const tbody = document.getElementById('tableBody');
        const rows = tbody.querySelectorAll('tr:not(.empty-row)');
        
        rows.forEach((row, rowIndex) => {
            if (row.classList.contains('pc-header-row')) {
                // PC header row
                const firstCell = row.querySelector('td');
                if (firstCell) {
                    const input = firstCell.querySelector('input');
                    const pcName = input ? input.value : firstCell.textContent.trim();
                    itemsToInsert.push({
                        sheet_id: currentSheetId,
                        sheet_name: sheet.name,
                        row_index: rowIndex,
                        article: pcName,
                        is_pc_header: true,
                        is_highlighted: false
                    });
                }
            } else {
                // Regular row
                const cells = row.querySelectorAll('td.editable');
                if (cells.length >= 10) {
                    const pictureCell = row.querySelector('.picture-cell');
                    let pictureUrl = null;
                    if (pictureCell) {
                        const img = pictureCell.querySelector('img');
                        if (img && img.src) {
                            pictureUrl = img.src; // Base64 or URL
                        }
                    }
                    
                    itemsToInsert.push({
                        sheet_id: currentSheetId,
                        sheet_name: sheet.name,
                        row_index: rowIndex,
                        article: cells[0]?.querySelector('input')?.value || '',
                        description: cells[1]?.querySelector('input')?.value || '',
                        old_property_n_assigned: cells[2]?.querySelector('input')?.value || '',
                        unit_of_meas: cells[3]?.querySelector('input')?.value || '',
                        unit_value: cells[4]?.querySelector('input')?.value || '',
                        quantity: cells[5]?.querySelector('input')?.value || '',
                        location: cells[6]?.querySelector('input')?.value || '',
                        condition: cells[7]?.querySelector('input')?.value || '',
                        remarks: cells[8]?.querySelector('input')?.value || '',
                        user: cells[9]?.querySelector('input')?.value || '',
                        picture_url: pictureUrl,
                        is_pc_header: false,
                        is_highlighted: row.classList.contains('highlighted-row')
                    });
                }
            }
        });
        
        // Insert all items
        if (itemsToInsert.length > 0) {
            const { error: insertError } = await supabase
                .from('inventory_items')
                .insert(itemsToInsert);
            
            if (insertError) {
                console.error('Error inserting items to Supabase:', insertError);
            } else {
                console.log('Data synced to Supabase successfully');
            }
        }
    } catch (error) {
        console.error('Error syncing to Supabase:', error);
    }
}

// Load data from Supabase
async function loadFromSupabase() {
    if (!checkSupabaseConnection()) return;
    
    try {
        // Load all sheets
        const { data: sheetsData, error: sheetsError } = await supabase
            .from('sheets')
            .select('*')
            .order('created_at', { ascending: true });
        
        if (sheetsError) {
            console.error('Error loading sheets from Supabase:', sheetsError);
            return;
        }
        
        if (!sheetsData || sheetsData.length === 0) {
            console.log('No sheets found in Supabase');
            return;
        }
        
        // Clear existing sheets
        sheets = {};
        sheetCounter = 0;
        
        // Load each sheet and its items
        for (const sheetData of sheetsData) {
            const { data: itemsData, error: itemsError } = await supabase
                .from('inventory_items')
                .select('*')
                .eq('sheet_id', sheetData.id)
                .order('row_index', { ascending: true });
            
            if (itemsError) {
                console.error('Error loading items for sheet:', itemsError);
                continue;
            }
            
            // Convert items back to row data format
            const rowData = [];
            const highlightStates = [];
            
            itemsData.forEach(item => {
                if (item.is_pc_header) {
                    rowData.push([item.article]);
                    highlightStates.push(false);
                } else {
                    rowData.push([
                        item.article || '',
                        item.description || '',
                        item.old_property_n_assigned || '',
                        item.unit_of_meas || '',
                        item.unit_value || '',
                        item.quantity || '',
                        item.location || '',
                        item.condition || '',
                        item.remarks || '',
                        item.user || ''
                    ]);
                    highlightStates.push(item.is_highlighted || false);
                }
            });
            
            // Create sheet
            sheets[sheetData.id] = {
                id: sheetData.id,
                name: sheetData.name,
                data: rowData,
                hasData: rowData.length > 0,
                highlightStates: highlightStates
            };
            
            // Update sheet counter
            const sheetNum = parseInt(sheetData.id.replace('sheet-', ''));
            if (sheetNum > sheetCounter) {
                sheetCounter = sheetNum;
            }
        }
        
        // Switch to first sheet if available
        const firstSheetId = Object.keys(sheets)[0];
        if (firstSheetId) {
            currentSheetId = firstSheetId;
            switchToSheet(firstSheetId);
        }
        
        // Update sheet tabs
        updateSheetTabs();
        
        console.log('Data loaded from Supabase successfully');
    } catch (error) {
        console.error('Error loading from Supabase:', error);
    }
}

// Update sheet tabs UI
function updateSheetTabs() {
    const sheetTabs = document.getElementById('sheetTabs');
    sheetTabs.innerHTML = '';
    
    Object.values(sheets).forEach(sheet => {
        const tab = document.createElement('div');
        tab.className = 'sheet-tab' + (sheet.id === currentSheetId ? ' active' : '');
        tab.setAttribute('data-sheet-id', sheet.id);
        tab.innerHTML = `
            <span class="sheet-name">${sheet.name}</span>
            <span class="close-sheet" data-sheet-id="${sheet.id}">×</span>
        `;
        
        tab.addEventListener('click', function(e) {
            if (!e.target.classList.contains('close-sheet')) {
                switchToSheet(sheet.id);
            }
        });
        
        const closeBtn = tab.querySelector('.close-sheet');
        closeBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            deleteSheet(sheet.id);
        });
        
        sheetTabs.appendChild(tab);
    });
}

async function deleteSheet(sheetId) {
    if (Object.keys(sheets).length <= 1) {
        alert('Cannot delete the last sheet. At least one sheet must exist.');
        return;
    }
    
    if (confirm(`Are you sure you want to delete "${sheets[sheetId].name}"?`)) {
        // Delete from Supabase if connected
        if (checkSupabaseConnection()) {
            try {
                // Delete items first
                const { error: itemsError } = await supabase
                    .from('inventory_items')
                    .delete()
                    .eq('sheet_id', sheetId);
                
                if (itemsError) {
                    console.error('Error deleting items from Supabase:', itemsError);
                }
                
                // Delete sheet
                const { error: sheetError } = await supabase
                    .from('sheets')
                    .delete()
                    .eq('id', sheetId);
                
                if (sheetError) {
                    console.error('Error deleting sheet from Supabase:', sheetError);
                }
            } catch (error) {
                console.error('Error deleting from Supabase:', error);
            }
        }
        
        delete sheets[sheetId];
        
        // Remove tab
        const tab = document.querySelector(`[data-sheet-id="${sheetId}"]`);
        if (tab) tab.remove();
        
        // Switch to another sheet if current was deleted
        if (currentSheetId === sheetId) {
            const remainingSheets = Object.keys(sheets);
            if (remainingSheets.length > 0) {
                switchToSheet(remainingSheets[0]);
            }
        }
    }
}

function hasAnyData() {
    return Object.values(sheets).some(sheet => sheet.hasData);
}

// Update inventoryData from table
function updateDataFromTable() {
    saveCurrentSheetData();
}

// Export to Excel with ExcelJS (supports images and styling)
document.getElementById('exportBtn').addEventListener('click', async function() {
    updateDataFromTable();
    
    if (!hasAnyData()) {
        alert('No data to export. Please add items or import an Excel file first.');
        return;
    }
    
    // Check if ExcelJS is available
    if (typeof ExcelJS === 'undefined') {
        alert('ExcelJS library is loading. Please wait a moment and try again.');
        return;
    }
    
    try {
        const workbook = new ExcelJS.Workbook();
        
        // Export all sheets
        for (const sheet of Object.values(sheets)) {
            if (!sheet.hasData || sheet.data.length === 0) continue;
            
            const worksheet = workbook.addWorksheet(sheet.name);
            
            // Get current date
            const now = new Date();
            const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December'];
            const currentDate = `${months[now.getMonth()]} ${now.getFullYear()}`;
            
            // Add logo in column A, rows 1-2 (0-indexed: rows 0-1)
            try {
                const logoResponse = await fetch('images/omsc.png');
                const logoBuffer = await logoResponse.arrayBuffer();
                const imageId = workbook.addImage({
                    buffer: logoBuffer,
                    extension: 'png',
                });
                
                // Insert logo in cell A1, spanning 2 rows
                worksheet.addImage(imageId, {
                    tl: { col: 0, row: 0 },
                    ext: { width: 120, height: 120 }
                });
            } catch (logoError) {
                console.warn('Could not load logo image:', logoError);
            }
            
            // Row 1: OCCIDENTAL MINDORO STATE COLLEGE (start from column B to avoid logo)
            const row1 = worksheet.addRow(['', 'OCCIDENTAL MINDORO STATE COLLEGE']);
            row1.getCell(2).font = { bold: true, size: 18 };
            row1.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.mergeCells(1, 2, 1, 11); // Merge columns B to K (2-11)
            
            // Row 2: Multimedia and Speech Laboratory
            const row2 = worksheet.addRow(['', 'Multimedia and Speech Laboratory']);
            row2.getCell(2).font = { bold: true, size: 14 };
            row2.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.mergeCells(2, 2, 2, 11);
            
            // Row 3: ICT Equipment, Devices & Accessories
            const row3 = worksheet.addRow(['', 'ICT Equipment, Devices & Accessories']);
            row3.getCell(2).font = { bold: true, size: 12 };
            row3.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.mergeCells(3, 2, 3, 11);
            
            // Row 4: AS OF date
            const row4 = worksheet.addRow(['', `AS OF ${currentDate.toUpperCase()}`]);
            row4.getCell(2).font = { bold: true, size: 11 };
            row4.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheet.mergeCells(4, 2, 4, 11);
            
            // Empty row for spacing
            worksheet.addRow([]);
            
            // Column headers row (row 6, 0-indexed: row 5)
            const headerRow = worksheet.addRow([
                'Article/It',
                'Description',
                'Old Property N Assigned',
                'Unit of meas',
                'Unit Value',
                'Quantity per Physical count',
                'Location/Whereabout',
                'Condition',
                'Remarks',
                'User',
                'Picture'
            ]);
            
            // Style header row
            headerRow.eachCell((cell) => {
                cell.font = { bold: true, size: 11 };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF495057' }
                };
                cell.font.color = { argb: 'FFFFFFFF' };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
            
            // Add data rows
            let currentRow = 7; // Start after headers (0-indexed: row 6 is column headers, so data starts at row 7)
            let dataRowIndex = 0; // Index for tracking highlight states
            
            sheet.data.forEach(rowData => {
                // Skip empty rows
                if (!rowData || rowData.length === 0) {
                    return;
                }
                
                const firstCell = rowData[0] ? rowData[0].toString().trim() : '';
                
                // Check if this is a PC header row
                // PC headers are stored as single-element arrays [pcName]
                // or have PC-related text in first cell
                const isPCHeader = firstCell && (
                    rowData.length === 1 || // Single cell = PC header
                    firstCell.toUpperCase().includes('PC USED BY') ||
                    firstCell.toUpperCase() === 'SERVER' ||
                    firstCell.toUpperCase().startsWith('PC ') ||
                    firstCell.toUpperCase() === 'PC'
                );
                
                if (isPCHeader) {
                    // PC Header row - merge across all columns and style
                    const pcRow = worksheet.addRow([firstCell, '', '', '', '', '', '', '', '', '', '']);
                    pcRow.getCell(1).font = { bold: true, size: 12 };
                    pcRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
                    pcRow.getCell(1).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFD3D3D3' } // Light gray
                    };
                    // Merge cells across all 11 columns (B to L, which is indices 1-11)
                    worksheet.mergeCells(currentRow, 1, currentRow, 11);
                    
                    // Add borders to merged cell
                    pcRow.getCell(1).border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                } else if (rowData.length >= 10) {
                    // Regular data row with full data
                    const exportRow = [...rowData];
                    // Add picture note if needed
                    if (exportRow.length === 10) {
                        exportRow.push(''); // Picture column
                    }
                    const dataRow = worksheet.addRow(exportRow);
                    
                    // Get condition value (index 7)
                    const conditionValue = rowData[7] ? rowData[7].toString().trim() : '';
                    
                    // Check if this row should be highlighted
                    const isHighlighted = sheet.highlightStates && sheet.highlightStates[dataRowIndex] === true;
                    
                    // Determine row color based on condition
                    let conditionColor = null;
                    if (conditionValue === 'Borrowed') {
                        conditionColor = { argb: 'FFFFFF3D' }; // Yellow
                    } else if (conditionValue === 'Unserviceable') {
                        conditionColor = { argb: 'FFF8D7DA' }; // Light red
                    }
                    
                    // Add borders, colors, and center alignment to data cells
                    dataRow.eachCell((cell) => {
                        cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        };
                        
                        // Center align all cells
                        cell.alignment = {
                            horizontal: 'center',
                            vertical: 'middle',
                            wrapText: true
                        };
                        
                        // Apply condition-based color (priority over highlight)
                        if (conditionColor) {
                            cell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: conditionColor
                            };
                        } else if (isHighlighted) {
                            // Apply yellow highlight if row is highlighted and no condition color
                            cell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFFFFF3D' } // Yellow background
                            };
                        }
                    });
                    
                    dataRowIndex++;
                }
                currentRow++;
            });
            
            // Auto-fit column widths based on content
            const maxColumnWidths = [12, 30, 20, 12, 15, 25, 25, 12, 20, 25, 20]; // Minimum widths
            const columnHeaders = ['Article/It', 'Description', 'Old Property N Assigned', 'Unit of meas', 
                                   'Unit Value', 'Quantity per Physical count', 'Location/Whereabout', 
                                   'Condition', 'Remarks', 'User', 'Picture'];
            
            // Calculate max width for each column
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                    if (colNumber <= 11) { // Only for data columns
                        const cellValue = cell.value ? cell.value.toString() : '';
                        const cellLength = cellValue.length;
                        // Add some padding (multiply by 1.2 for better fit)
                        const estimatedWidth = Math.max(cellLength * 1.2, columnHeaders[colNumber - 1].length * 1.2);
                        if (estimatedWidth > maxColumnWidths[colNumber - 1]) {
                            maxColumnWidths[colNumber - 1] = Math.min(estimatedWidth, 50); // Cap at 50 to prevent too wide
                        }
                    }
                });
            });
            
            // Set column widths with auto-fit
            worksheet.columns = maxColumnWidths.map((width, index) => ({
                width: Math.max(width, columnHeaders[index].length * 1.2)
            }));
            
            // Set row heights for header rows
            worksheet.getRow(1).height = 30; // OCCIDENTAL MINDORO STATE COLLEGE
            worksheet.getRow(2).height = 25; // Multimedia and Speech Laboratory
            worksheet.getRow(3).height = 22; // ICT Equipment
            worksheet.getRow(4).height = 20; // AS OF date
            worksheet.getRow(6).height = 25; // Column headers
            
            // Configure page setup for long bond paper and fit to one page
            worksheet.pageSetup = {
                paperSize: 9, // A4 (closest standard size, user can change to long bond in Excel)
                orientation: 'landscape', // Landscape for better fit of wide table
                fitToPage: true,
                fitToWidth: 1, // Fit all columns to 1 page width
                fitToHeight: 1, // Fit all rows to 1 page height
                margins: {
                    left: 0.3,
                    right: 0.3,
                    top: 0.5,
                    bottom: 0.5,
                    header: 0.3,
                    footer: 0.3
                },
                scale: 100, // 100% scale
                horizontalCentered: true, // Center horizontally
                verticalCentered: false
            };
            
            // Note: For long bond paper (8.5" x 13"), user can:
            // 1. Open Excel file
            // 2. Go to Page Layout > Size > More Paper Sizes
            // 3. Select "Custom" and set width: 8.5", height: 13"
            // The fitToPage settings above will ensure everything fits on one page
        }
        
        // Generate filename with current date
        const now = new Date();
        const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December'];
        const filename = `INVENTORY-AS-OF-${months[now.getMonth()].toUpperCase()}-${now.getFullYear()}.xlsx`;
        
        // Write file
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        window.URL.revokeObjectURL(url);
        
        alert('Excel file exported successfully with ' + Object.values(sheets).filter(s => s.hasData).length + ' sheet(s)!');
    } catch (error) {
        alert('Error exporting Excel file: ' + error.message);
        console.error(error);
    }
});

// New Sheet button
document.getElementById('newSheetBtn').addEventListener('click', function() {
    const sheetName = prompt('Enter sheet name:', `Sheet ${sheetCounter + 1}`);
    if (sheetName !== null && sheetName.trim() !== '') {
        createNewSheet(sheetName.trim());
    }
});

// Initialize sheet tab click handlers (run after DOM is ready)
setTimeout(function() {
    const firstTab = document.querySelector('.sheet-tab');
    if (firstTab) {
        firstTab.addEventListener('click', function(e) {
            if (!e.target.classList.contains('close-sheet')) {
                switchToSheet('sheet-1');
            }
        });
        
        const closeBtn = firstTab.querySelector('.close-sheet');
        if (closeBtn) {
            closeBtn.addEventListener('click', function(e) {
                e.stopPropagation();
                deleteSheet('sheet-1');
            });
        }
    }
}, 100);

// Clear data
document.getElementById('clearBtn').addEventListener('click', function() {
    const currentSheet = getCurrentSheet();
    if (confirm(`Are you sure you want to clear all data in "${currentSheet.name}"?`)) {
        setCurrentSheetData([], false);
        document.getElementById('tableBody').innerHTML = 
            '<tr class="empty-row"><td colspan="12" class="empty-message">No data loaded. Please import an Excel file or add items manually.</td></tr>';
        document.getElementById('exportBtn').disabled = !hasAnyData();
        document.getElementById('clearBtn').disabled = true;
    }
});
