// sidepanel.js

// Global variables
let workbook = null;
let columnInfo = [];
let originalFilename = 'sheet';
let addedRowCount = 0;
let toastTimeout;

// --- Utility Functions ---

function logToPanel(message) {
  console.log(message);
  const logEntry = document.createElement('p');
  const time = new Date().toLocaleTimeString();
  logEntry.textContent = `[${time}] ${message}`;
  const activityLog = document.getElementById('activityLog');
  if (activityLog) {
      activityLog.appendChild(logEntry);
      activityLog.scrollTop = activityLog.scrollHeight;
  }
}

function sanitizeText(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;
    return temp.textContent || "";
}

// Increased default duration slightly for better readability
function showToast(message, type = 'success', duration = 4000) {
    const toast = document.getElementById('toastNotification');
    if (!toast) return;

    clearTimeout(toastTimeout);

    toast.textContent = message;
    toast.className = `toast-${type}`;
    
    if (duration > 0) {
        toastTimeout = setTimeout(() => {
            toast.classList.add('hidden');
        }, duration);
    }
}

function updateLoadingOverlay(show, message = "Processing...") {
    const overlay = document.getElementById('loadingOverlay');
    const msgElement = document.getElementById('loadingMessage');
    if (overlay && msgElement) {
        msgElement.textContent = message;
        if (show) {
            overlay.classList.remove('hidden');
        } else {
            overlay.classList.add('hidden');
        }
    }
}

// --- UI State Management ---

// updateFileIndicator function removed.

function showEmptyState() {
    const fieldsContainer = document.getElementById('fieldsContainer');
    const guidingLabels = document.querySelector('.guiding-labels');
    const fileTypesLabel = document.getElementById('fileTypesLabel');
    const loadRecipeBtn = document.getElementById('loadRecipeBtn');
    const fileUploadBtn = document.querySelector('.file-upload-btn');
    const downloadBtn = document.getElementById('downloadBtn');

    if (fieldsContainer) {
        fieldsContainer.innerHTML = `
            <div class="empty-state">
                <h3>Get Started</h3>
                <p>Upload an .xlsx, .xls, or .csv file to define your data fields.</p>
                <p>Or, load a <strong>.json template</strong> to start a new sheet.</p>
            </div>
        `;
    }
    if (guidingLabels) guidingLabels.classList.add('hidden');
    if (fileTypesLabel) fileTypesLabel.classList.remove('hidden');
    
    // Reset custom file button text
    if (fileUploadBtn) {
        fileUploadBtn.textContent = 'Choose Sheet File';
    }

    // Allow loading a template even if no workbook is loaded yet
    if (loadRecipeBtn) loadRecipeBtn.disabled = false;
    
    // Ensure other buttons are reset
    document.getElementById('saveRecipeBtn').disabled = true;
    
    // Ensure download button and counter are hidden on reset
    if (downloadBtn) {
        downloadBtn.classList.add('hidden');
    }
    document.getElementById('updateCounter').classList.add('hidden');
    
    updateAddButtonState();
}

function updateAddButtonState() {
  const fieldsContainer = document.getElementById('fieldsContainer');
  const addToSheetBtn = document.getElementById('addToSheetBtn');
  
  if (!addToSheetBtn) return;

  if (!fieldsContainer || columnInfo.length === 0) {
    addToSheetBtn.disabled = true;
    return;
  }

  const inputs = fieldsContainer.querySelectorAll('input[type="text"]');
  const isAnyFieldFilled = Array.from(inputs).some(input => input.value.trim() !== '');
  
  const isConfirming = addToSheetBtn.textContent !== 'Add Row to Sheet';
  
  addToSheetBtn.disabled = !isAnyFieldFilled || isConfirming;
}

function updateClearRecipeButtonState() {
    const clearRecipeBtn = document.getElementById('clearRecipeBtn');
    if (!clearRecipeBtn) return;

    if (columnInfo.length === 0) {
        clearRecipeBtn.classList.add('hidden');
        return;
    }

    const isAnyInfoFilled = columnInfo.some(info =>
        (info.type && info.type.trim() !== '') ||
        (info.format && info.format.trim() !== '') ||
        (info.keywords && info.keywords.trim() !== '')
    );

    if (isAnyInfoFilled) {
        clearRecipeBtn.classList.remove('hidden');
    } else {
        clearRecipeBtn.classList.add('hidden');
    }
}

// --- Core Logic Functions ---

function addRowToSheetLogic() {
    if (!workbook && columnInfo.length > 0) {
         initializeWorkbookFromTemplate(columnInfo.map(c => c.name));
    }

    if (!workbook) {
        logToPanel("Cannot add row: No sheet structure available.");
        showToast("Please load a sheet or template first.", 'error');
        return false;
    }
    
    const newRowData = columnInfo.map(info => {
        const inputElement = document.getElementById(`field-${info.name}`);
        const value = inputElement ? inputElement.value.trim() : '';
        return (value === 'AI_PARSE_ERROR') ? '' : value;
    });
    
    if (newRowData.every(value => value === '')) {
        logToPanel("Skipped adding row: All fields are empty.");
        return false;
    }

    try {
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        XLSX.utils.sheet_add_aoa(worksheet, [newRowData], { origin: -1 });
        logToPanel("Data appended to in-memory sheet successfully.");

        columnInfo.forEach(info => {
            const lockBtn = document.getElementById(`lock-${info.name}`);
            if (lockBtn && lockBtn.dataset.locked === 'true') return;
            
            const inputElement = document.getElementById(`field-${info.name}`);
            if (inputElement) inputElement.value = '';
        });

        addedRowCount++;
        const updateCounter = document.getElementById('updateCounter');
        updateCounter.textContent = `${addedRowCount} Row(s) Added`;
        updateCounter.classList.remove('hidden');
        
        // FIX: Ensure the download button becomes visible using classList
        const downloadBtn = document.getElementById('downloadBtn');
        if (downloadBtn) {
            downloadBtn.classList.remove('hidden');
        }
        
        updateAddButtonState();
        return true;
        
    } catch (error) {
        logToPanel(`Error adding row: ${error.message}`);
        showToast("Error while updating the sheet.", 'error');
        return false;
    }
}

function initializeWorkbookFromTemplate(headers) {
    logToPanel("Initializing new workbook structure from template.");
    const worksheet = XLSX.utils.aoa_to_sheet([headers]);
    workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    originalFilename = 'Sheet_from_Template.xlsx';
    
    // Ensure download button and counter are reset when initializing a new workbook
    document.getElementById('downloadBtn').classList.add('hidden');
    addedRowCount = 0;
    document.getElementById('updateCounter').classList.add('hidden');
    // Feedback is handled in the calling function (loadRecipe)
}


// --- File Handling and AI Analysis ---

// Handles the selection of a spreadsheet file
async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    originalFilename = file.name;
    logToPanel(`File selected: ${originalFilename}`);
    updateLoadingOverlay(true, "Loading sheet...");
    
    // Reference to the custom upload button (for UI feedback)
    const fileUploadBtn = document.querySelector('.file-upload-btn');

    try {
        const data = await file.arrayBuffer();
        workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (!json || json.length === 0 || json[0].length === 0) {
            throw new Error("Excel sheet is empty or has no headers.");
        }

        const headers = json[0];
        const firstDataRow = json[1] || [];

        // AI Analysis Logic
        const settings = await chrome.storage.local.get(['analyzeHeadersWithAI']);
        const analyzeWithAI = settings.analyzeHeadersWithAI === true;

        if (analyzeWithAI) {
            if (typeof LanguageModel === 'undefined' || typeof LanguageModel.create !== 'function') {
                logToPanel("AI features unavailable in this browser. Skipping AI header analysis.");
                showToast("AI Header Analysis requires browser AI support.", 'info');
                initializeDefaultColumnInfo(headers);
            } else {
                await analyzeHeadersWithAI(headers, firstDataRow);
            }
        } else {
            logToPanel("Skipping AI header analysis (setting is disabled).");
            initializeDefaultColumnInfo(headers);
        }

        // Update the UI
        displayFields(columnInfo);
        
        // Reset UI state for the new file
        // Use classList for managing visibility
        const downloadBtn = document.getElementById('downloadBtn');
        if (downloadBtn) {
            downloadBtn.classList.add('hidden');
        }
        addedRowCount = 0;
        document.getElementById('updateCounter').classList.add('hidden');
        document.getElementById('fileTypesLabel').classList.add('hidden');
        
        // Use toast for feedback
        showToast(`Sheet loaded: ${originalFilename}`, 'success');

    } catch (error) {
        logToPanel(`Error processing file: ${error.message}`);
        showToast(`Error: ${error.message}`, 'error');
        // Reset state on failure
        workbook = null;
        columnInfo = [];
        showEmptyState();
    } finally {
        updateLoadingOverlay(false);
        
        // Reset the custom file button text
        if (fileUploadBtn) {
            fileUploadBtn.textContent = 'Choose Sheet File';
        }
        
        // Clear the file input value
        event.target.value = '';
    }
}

// (initializeDefaultColumnInfo, analyzeHeadersWithAI, displayFields, saveRecipe, loadRecipe, fillFieldsFromData implementations remain the same as the previous optimized version)

function initializeDefaultColumnInfo(headers) {
    columnInfo = headers.map(header => ({
        name: header, type: '', format: '', keywords: ''
    }));
}

async function analyzeHeadersWithAI(headers, firstDataRow) {
    logToPanel("Analyzing sheet structure with AI...");
    updateLoadingOverlay(true, "Analyzing Headers with AI...");
    let session;
    try {
        session = await LanguageModel.create();
        
        const examples = headers.map((header, index) => {
            const exampleValue = firstDataRow[index] ? String(firstDataRow[index]).substring(0, 50) : 'N/A';
            return `Header: "${header}", Example Value: "${exampleValue}"`;
        }).join('\n');

        const prompt = `Analyze the following spreadsheet headers and example values. For each header, determine:
1. Likely Data Type (Choose from: Text, Number, Date, Email, URL, Currency, Time, Other).
2. Example Format (e.g., yyyy-mm-dd, $0.00, or N/A).
3. 2-3 common Keywords or Aliases (comma separated).

Headers and Examples:
${examples}

CRITICAL: Respond ONLY with a valid JSON array of objects, where each object has the keys: "name", "type", "format", "keywords". The order must match the input headers. If format or keywords are N/A, use an empty string "". Do not include any other text or markdown formatting (like \`\`\`json).
JSON Output:`;

        const rawResult = await session.prompt(prompt);
        
        const jsonString = rawResult.replace(/```json/i, '').replace(/```/g, '').trim();
        const parsedResult = JSON.parse(jsonString);

        if (Array.isArray(parsedResult) && parsedResult.length === headers.length) {
            columnInfo = parsedResult.map((item, index) => ({
                name: item.name || headers[index],
                type: item.type || '',
                format: (item.format && item.format.toUpperCase() !== 'N/A') ? item.format : '',
                keywords: (item.keywords && item.keywords.toUpperCase() !== 'N/A') ? item.keywords : ''
            }));
            logToPanel("✅ AI analysis complete.");
        } else {
            throw new Error("AI returned invalid structure or incorrect number of items.");
        }

    } catch (aiError) {
        logToPanel(`❌ Error during AI analysis: ${aiError.message}. Falling back to manual setup.`);
        showToast("AI header analysis failed. Please define fields manually.", 'error');
        initializeDefaultColumnInfo(headers);
    } finally {
        if (session) session.destroy();
    }
}

function displayFields(infoArray) {
    const fieldsContainer = document.getElementById('fieldsContainer');
    const saveRecipeBtn = document.getElementById('saveRecipeBtn');
    const guidingLabels = document.querySelector('.guiding-labels');

    if (guidingLabels) {
        guidingLabels.classList.remove('hidden');
    }

    fieldsContainer.innerHTML = '';
    const hasInfo = infoArray.length > 0;
    
    if (!hasInfo) {
        showEmptyState();
        return;
    }

    saveRecipeBtn.disabled = false;

    infoArray.forEach((info, index) => {
        const fieldDiv = document.createElement('div');
        fieldDiv.className = 'field-wrapper';

        const headerDiv = document.createElement('div');
        headerDiv.className = 'field-header';
        
        const label = document.createElement('label');
        label.textContent = info.name;
        label.setAttribute('for', `field-${info.name}`);
        headerDiv.appendChild(label);

        const createTag = (prop, placeholderText, className) => {
            const tag = document.createElement('span');
            tag.className = `data-tag ${className}`;
            tag.contentEditable = true;
            tag.setAttribute('data-placeholder', placeholderText);

            if (info[prop]) {
                tag.textContent = info[prop];
            } else {
                tag.textContent = placeholderText;
                tag.classList.add('placeholder');
            }

            // Event listeners
            tag.addEventListener('input', (e) => {
                if (!e.target.classList.contains('placeholder')) {
                    columnInfo[index][prop] = e.target.textContent;
                }
                updateClearRecipeButtonState();
            });

            tag.addEventListener('focus', (e) => {
                if (e.target.classList.contains('placeholder')) {
                    e.target.textContent = '';
                    e.target.classList.remove('placeholder');
                }
            });

            tag.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    e.target.blur();
                }
            });

            tag.addEventListener('blur', (e) => {
                const sanitizedValue = sanitizeText(e.target.innerHTML).trim();
                if (sanitizedValue === '') {
                    e.target.textContent = placeholderText;
                    e.target.classList.add('placeholder');
                    columnInfo[index][prop] = '';
                } else {
                    e.target.textContent = sanitizedValue;
                    columnInfo[index][prop] = sanitizedValue;
                }
                updateClearRecipeButtonState();
            });
            
            headerDiv.appendChild(tag);
        };

        createTag('type', 'e.g., Text', 'data-type');
        createTag('format', 'e.g., YYYY-MM-DD', 'data-format');
        createTag('keywords', 'e.g., Alias', 'data-keywords');

        const inputContainer = document.createElement('div');
        inputContainer.className = 'input-container';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.id = `field-${info.name}`;
        input.name = info.name;
        input.addEventListener('input', updateAddButtonState);

        const lockBtn = document.createElement('button');
        lockBtn.className = 'lock-btn';
        lockBtn.id = `lock-${info.name}`;
        lockBtn.textContent = '🔓';
        lockBtn.dataset.locked = 'false';
        lockBtn.title = 'Lock field value (prevents clearing/overwriting)';
        
        lockBtn.addEventListener('click', () => {
            const isLocked = lockBtn.dataset.locked === 'true';
            const inputToToggle = document.getElementById(`field-${info.name}`);
            
            if (isLocked) {
                // Unlock
                lockBtn.dataset.locked = 'false';
                lockBtn.textContent = '🔓';
                inputToToggle.readOnly = false;
                logToPanel(`Field "${info.name}" unlocked.`);
            } else {
                // Lock
                lockBtn.dataset.locked = 'true';
                lockBtn.textContent = '🔒';
                inputToToggle.readOnly = true;
                logToPanel(`Field "${info.name}" locked.`);
            }
        });

        inputContainer.appendChild(input);
        inputContainer.appendChild(lockBtn);
        
        fieldDiv.appendChild(headerDiv);
        fieldDiv.appendChild(inputContainer);
        fieldsContainer.appendChild(fieldDiv);
    });

    logToPanel("Displayed fields in the panel.");
    updateAddButtonState();
    updateClearRecipeButtonState();
}

function saveRecipe() {
  if (columnInfo.length === 0) {
    logToPanel("No template info to save.");
    showToast("No fields loaded to save as a template.", 'error');
    return;
  }
  
  const baseName = originalFilename.replace(/\.(xlsx|xls|csv)$/i, '');
  const templateFilename = `${baseName}_template.json`;

  const recipeData = JSON.stringify(columnInfo, null, 2);
  const blob = new Blob([recipeData], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = templateFilename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  showToast('Template saved successfully!', 'success');
  logToPanel(`Template saved: ${templateFilename}`);
}

function loadRecipe(event) {
  const file = event.target.files[0];
  const recipeFileInput = event.target;
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const loadedInfo = JSON.parse(e.target.result);
      if (!Array.isArray(loadedInfo) || loadedInfo.length === 0) {
        throw new Error("Invalid or empty template file.");
      }

      const recipeHeaders = loadedInfo.map(c => c.name);

      if (!workbook) {
        // Template-First Workflow
        initializeWorkbookFromTemplate(recipeHeaders);
        columnInfo = loadedInfo;
        document.getElementById('fileTypesLabel').classList.add('hidden');
        // Use Toast for feedback
        showToast(`Template loaded. New sheet initialized.`, 'success');
        logToPanel(`Template "${file.name}" loaded. Initialized new workbook.`);

      } else {
        // Workbook-Loaded Workflow
        const currentHeaders = columnInfo.map(c => c.name);

        if (currentHeaders.length !== recipeHeaders.length || !currentHeaders.every((h, i) => h === recipeHeaders[i])) {
            if (confirm("Template headers do not match the currently loaded sheet. Do you want to overwrite the current setup and start a new sheet based on this template?")) {
                initializeWorkbookFromTemplate(recipeHeaders);
                columnInfo = loadedInfo;
                // Use Toast for feedback
                showToast(`Template loaded. Current sheet replaced.`, 'success');
            } else {
                throw new Error("Template headers do not match the sheet. Load cancelled.");
            }
        } else {
            columnInfo = loadedInfo;
            // Use Toast for feedback
            showToast(`Template "${file.name}" applied to sheet.`, 'success');
            logToPanel(`Template "${file.name}" applied.`);
        }
      }

      displayFields(columnInfo);

    } catch (error) {
      showToast(`Error loading template: ${error.message}`, 'error');
      logToPanel(`Error loading template: ${error.message}`);
    } finally {
      recipeFileInput.value = '';
    }
  };
  reader.readAsText(file);
}

// Helper function for filling fields from AI data
function fillFieldsFromData(data) {
    for (const field in data) {
        const lockBtn = document.getElementById(`lock-${field}`);
        if (lockBtn && lockBtn.dataset.locked === 'true') {
            logToPanel(` > Skipped locked field '${field}'`);
            continue;
        }
        
        const inputElement = document.getElementById(`field-${field}`);
        if (inputElement) {
            inputElement.value = data[field];
            inputElement.classList.add('flash-on-fill');
            inputElement.addEventListener('animationend', () => {
                inputElement.classList.remove('flash-on-fill');
            }, { once: true });
        }
    }
}

// --- Initialization and Event Listeners ---

document.addEventListener('DOMContentLoaded', () => {
    
    if (typeof XLSX === 'undefined') {
        logToPanel("Fatal Error: XLSX library not loaded.");
        showToast("Error: Spreadsheet library failed to load.", 'error', 10000);
        return;
    }

    // Get element references
    const fileInput = document.getElementById('fileInput');
    const addToSheetBtn = document.getElementById('addToSheetBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const saveRecipeBtn = document.getElementById('saveRecipeBtn');
    const loadRecipeBtn = document.getElementById('loadRecipeBtn');
    const settingsBtn = document.getElementById('settingsBtn');
    const settingsMenu = document.getElementById('settingsMenu');
    const guideBtn = document.getElementById('guideBtn');
    const aiAnalysisToggle = document.getElementById('aiAnalysisToggle');
    const clearRecipeBtn = document.getElementById('clearRecipeBtn');
    const stopAIBtn = document.getElementById('stopAIBtn');
    const toggleLogBtn = document.getElementById('toggleLogBtn');
    const activityLog = document.getElementById('activityLog');
    const batchModeToggle = document.getElementById('batchModeToggle');
    const fileUploadBtn = document.querySelector('.file-upload-btn'); // Reference for custom button UI updates

    // Create hidden file input for loading templates
    const recipeFileInput = document.createElement('input');
    recipeFileInput.type = 'file';
    recipeFileInput.accept = '.json';
    recipeFileInput.style.display = 'none';
    document.body.appendChild(recipeFileInput);

    // --- Event Listener Setup ---

    // File Upload Listener
    fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        
        // Update UI for the custom button
        if (file && fileUploadBtn) {
            // Update button text temporarily to show file is selected/processing
            const truncatedName = file.name.length > 20 ? file.name.substring(0, 20) + '...' : file.name;
            fileUploadBtn.textContent = `Processing: ${truncatedName}`;
        } else if (fileUploadBtn) {
            fileUploadBtn.textContent = 'Choose Sheet File';
        }
        
        // Call the main handler
        handleFileSelect(event);
    });

    // Template Loading
    saveRecipeBtn.addEventListener('click', saveRecipe);
    loadRecipeBtn.addEventListener('click', () => recipeFileInput.click());
    recipeFileInput.addEventListener('change', loadRecipe);

    // Clear Recipe Button
    clearRecipeBtn.addEventListener('click', () => {
        if (columnInfo.length > 0) {
            columnInfo.forEach(info => {
                info.type = '';
                info.format = '';
                info.keywords = '';
            });
            displayFields(columnInfo);
            showToast('Field info cleared.', 'success');
            logToPanel('Field metadata cleared by user.');
        }
    });

    // AI Control
    stopAIBtn.addEventListener('click', () => {
        logToPanel("Stop request sent. Halting AI process...");
        updateLoadingOverlay(true, "Stopping AI Task...");
        chrome.runtime.sendMessage({ type: 'CANCEL_AI_TASK' });
    });
    
    // Activity Log Toggle
    toggleLogBtn.addEventListener('click', () => {
        activityLog.classList.toggle('hidden');
        if (activityLog.classList.contains('hidden')) {
            toggleLogBtn.textContent = '˄';
            toggleLogBtn.title = 'Show Activity Log';
        } else {
            toggleLogBtn.textContent = '˅';
            toggleLogBtn.title = 'Hide Activity Log';
            activityLog.scrollTop = activityLog.scrollHeight;
        }
    });

    // Settings Menu Toggle
    settingsBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      settingsMenu.classList.toggle('hidden');
    });



    // Close settings menu when clicking outside
    window.addEventListener('click', (event) => {
      if (!settingsMenu.classList.contains('hidden')) {
        if (!event.target.closest('#settingsMenu') && !event.target.closest('.info-tooltip')) {
           settingsMenu.classList.add('hidden');
        }
      }
    });

    // Settings Toggles
    aiAnalysisToggle.addEventListener('change', () => {
        const isEnabled = aiAnalysisToggle.checked;
        chrome.storage.local.set({ analyzeHeadersWithAI: isEnabled }, () => {
            logToPanel(`AI Header Analysis setting saved: ${isEnabled ? 'ON' : 'OFF'}.`);
        });
    });

    batchModeToggle.addEventListener('change', () => {
        const isEnabled = batchModeToggle.checked;
        chrome.storage.local.set({ batchModeEnabled: isEnabled }, () => {
            logToPanel(`Batch Mode setting saved: ${isEnabled ? 'ON' : 'OFF'}.`);
        });
    });

    // Add Row Button
    addToSheetBtn.addEventListener('click', () => {
      const success = addRowToSheetLogic();
      if (success) {
          const originalBtnText = addToSheetBtn.textContent;
          addToSheetBtn.textContent = '✔ Row Added!';
          setTimeout(() => {
            addToSheetBtn.textContent = originalBtnText;
            updateAddButtonState();
          }, 1500);
      }
    });

    // Download Button
    downloadBtn.addEventListener('click', () => {
      if (!workbook) return;
      
      const bookType = originalFilename.toLowerCase().endsWith('.csv') ? 'csv' : 'xlsx';
      const mimeType = bookType === 'csv' ? 'text/csv' : 'application/octet-stream';

      const wbout = XLSX.write(workbook, { bookType: bookType, type: 'array' });
      const blob = new Blob([wbout], { type: mimeType });
      
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      
      const baseName = originalFilename.replace(/\.(xlsx|xls|csv)$/i, '');
      const newFilename = `updated_${baseName}.${bookType}`;
      link.download = newFilename; 

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      logToPanel(`Downloaded updated file: ${newFilename}`);
      showToast(`Downloaded: ${newFilename}`, 'success');
    });

    // --- Message Listener (Communication with background.js) ---
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        switch (request.type) {
          case 'GET_FIELDS':
            if (columnInfo.length === 0) {
                showToast("Please load a sheet or template before extracting data.", 'error');
                logToPanel("Extraction attempt failed: No fields loaded in panel.");
            }
            logToPanel("Sending fields and settings to background script.");
            sendResponse({ 
                fields: columnInfo,
                isBatchMode: batchModeToggle.checked
            });
            return true;

          case 'FILL_AND_ADD':
            logToPanel("Received batch item data, auto-adding row...");
            const batchData = request.data;
            for (const field in batchData) {
                const inputElement = document.getElementById(`field-${field}`);
                if (inputElement && !inputElement.readOnly) {
                    inputElement.value = batchData[field];
                    inputElement.classList.add('flash-on-fill');
                    inputElement.addEventListener('animationend', () => {
                        inputElement.classList.remove('flash-on-fill');
                    }, { once: true });
                }
            }
            setTimeout(() => {
                addRowToSheetLogic();
            }, 500);
            break;

          case 'FILL_FIELDS':
            logToPanel("Received AI data to fill fields.");
            fillFieldsFromData(request.data);
            updateAddButtonState();
            break;

          case 'SHOW_SPINNER':
            updateLoadingOverlay(true, "AI Extracting Data...");
            break;
            
          case 'HIDE_SPINNER':
            updateLoadingOverlay(false);
            break;
            
          case 'LOG_TO_PANEL':
            logToPanel(`[BG] ${request.message}`);
            break;

          case 'SHOW_TOAST':
            showToast(request.message, request.toastType || 'info', request.duration || 5000);
            break;
        }
    });

    // --- Initialization ---

    // Load settings from storage
    chrome.storage.local.get(['analyzeHeadersWithAI', 'batchModeEnabled'], (result) => {
        const aiEnabled = result.analyzeHeadersWithAI !== false; 
        aiAnalysisToggle.checked = aiEnabled;

        const batchEnabled = result.batchModeEnabled === true;
        batchModeToggle.checked = batchEnabled;
    });

    // Initialize UI state
    showEmptyState();
    logToPanel("Pull2Sheet side panel loaded and ready.");
});