// background.js

let isAITaskCancelled = false;

// Configuration Constants for Stability
const AI_SESSION_TIMEOUT_MS = 30000; // 30 seconds timeout for session creation
const AI_TIMEOUT_MS_SEQUENTIAL = 30000; // 20 seconds timeout PER FIELD in sequential mode
const AI_TIMEOUT_MS_BATCH_SPLIT = 55000; // 45 seconds for the initial batch split
const MAX_INPUT_LENGTH = 8000; // Max characters for single extraction
const BATCH_MAX_LENGTH = 12000; // Max characters for batch extraction

// --- Utility Functions ---

// Send logs to the side panel
function logToPanel(message) {
  console.log(message);
  try {
    // Use runtime.sendMessage, catching errors if the side panel is closed
    chrome.runtime.sendMessage({ type: 'LOG_TO_PANEL', message: message }).catch(err => console.error("Log send error:", err));
  } catch (error) {
    console.error("Could not send log message to side panel:", error);
  }
}

// Send toast notifications to the side panel
function showToastInPanel(message, type = 'success', duration = 5000) {
    try {
        chrome.runtime.sendMessage({ type: 'SHOW_TOAST', message: message, toastType: type, duration: duration }).catch(err => console.error("Toast send error:", err));
    } catch (error) {
        console.error("Failed to send toast message to panel:", error);
    }
}

/**
 * Utility function to add a timeout to a promise.
 * Prevents hanging indefinitely if the AI model fails to respond.
 */
function promiseWithTimeout(promise, timeoutMs, errorMessage = "Operation timed out") {
    let timeoutHandle;
    const timeoutPromise = new Promise((_, reject) => {
        timeoutHandle = setTimeout(() => {
            // Crucial: Set the cancellation flag when a timeout occurs to stop subsequent operations
            isAITaskCancelled = true;
            logToPanel(`[TIMEOUT] ${errorMessage}`);
            reject(new Error(errorMessage));
        }, timeoutMs);
    });

    // Race the main promise against the timeout
    return Promise.race([
        promise,
        timeoutPromise
    ]).finally(() => {
        // Crucial: Clear the timer if the main promise finished first
        clearTimeout(timeoutHandle);
    });
}

// --- Setup and Listeners ---

chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch((error) => console.error(error));

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: "extractToSheet",
    title: "Extract to Pull2Sheet",
    contexts: ["selection", "image"]
  });
  // Initialize default settings if not set
  chrome.storage.local.get(['analyzeHeadersWithAI', 'batchModeEnabled'], (result) => {
    if (result.analyzeHeadersWithAI === undefined) {
        chrome.storage.local.set({ analyzeHeadersWithAI: true }); // Default ON
    }
    if (result.batchModeEnabled === undefined) {
        chrome.storage.local.set({ batchModeEnabled: false }); // Default OFF
    }
  });
});

// Listener for cancellation requests
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.type === 'CANCEL_AI_TASK') {
        logToPanel("🛑 Cancellation request received. Process will stop soon.");
        isAITaskCancelled = true;
    }
    return true;
});

// --- AI Functions ---

/**
 * Checks AI Model Availability and provides guidance if unavailable.
 */
async function checkAIAvailability() {
    if (typeof LanguageModel === 'undefined' || typeof LanguageModel.availability !== 'function') {
        const errorMessage = "AI features are not supported in this browser version. Please update your browser.";
        showToastInPanel(errorMessage, 'error', 10000);
        throw new Error(errorMessage);
    }

    const availability = await LanguageModel.availability();
    // Check for common unavailability statuses
    if (availability === 'unavailable' || availability === 'no') {
        const errorMessage = `AI model is unavailable (Status: ${availability}). Please ensure the built-in AI model (e.g., Gemini Nano) is enabled in your browser settings (e.g., chrome://settings/ai).`;
        showToastInPanel(errorMessage, 'error', 15000);
        throw new Error(errorMessage);
    }
    return true;
}

/**
 * Helper function to create an AI session safely with a timeout.
 * This prevents hanging if LanguageModel.create() never resolves.
 */
async function createAISession(options = {}) {
    logToPanel("Attempting to create AI session...");
    try {
        const session = await promiseWithTimeout(
            LanguageModel.create(options),
            AI_SESSION_TIMEOUT_MS,
            `AI session creation timed out after ${AI_SESSION_TIMEOUT_MS / 1000} seconds.`
        );
        logToPanel("AI session created successfully.");
        return session;
    } catch (error) {
        logToPanel(`❌ Failed to create AI session: ${error.message}`);
        if (error.message.includes("timed out")) {
            showToastInPanel("AI initialization timed out. The model might be busy or not ready. Please try again later.", 'error', 10000);
        }
        throw error; // Propagate the error to the caller
    }
}

/**
 * Extracts data sequentially by asking the AI one question per field.
 * This approach is slower than single-shot JSON extraction but is more stable for local models.
 */
async function extractDataSequential(inputData, fields, session, inputType = 'text') {
    const results = {};
    logToPanel(`Starting sequential AI extraction for ${fields.length} fields...`);

    for (const fieldInfo of fields) {
        // Check the flag before processing each field
        if (isAITaskCancelled) {
            logToPanel("🛑 AI task stopped (cancelled or timed out) during sequential extraction.");
            break; // Exit the loop
        }

        logToPanel(`- Querying AI for: ${fieldInfo.name}`);

        // Construct the simple, direct prompt
        const keywordInstruction = fieldInfo.keywords ? `Look for labels similar to the field name or these keywords: "${fieldInfo.keywords}".` : '';
        const typeInstruction = fieldInfo.type ? `The data type should be a "${fieldInfo.type}".` : '';
        const formatInstruction = fieldInfo.format ? `The format should be similar to "${fieldInfo.format}".` : '';

        let promptInstructions = `Extract the value for "${fieldInfo.name}".
- ${keywordInstruction}
- ${typeInstruction}
- ${formatInstruction}
- CRITICAL: Respond with ONLY the extracted value. Do not add explanations or quotes.
- If the information is not present, respond with "N/A".`;

        let promptContent;

        if (inputType === 'text') {
            promptContent = `${promptInstructions}
---START OF TEXT---
${inputData}
---END OF TEXT---
Value:`;
        } else {
            // Multimodal prompt for images
            promptContent = [{
                role: 'user',
                content: [
                    { type: 'text', value: `From the image provided: ${promptInstructions}\nValue:` },
                    { type: 'image', value: inputData }
                ]
            }];
        }

        try {
            // Use promiseWithTimeout to prevent a single step from hanging
            const rawResult = await promiseWithTimeout(
                session.prompt(promptContent),
                AI_TIMEOUT_MS_SEQUENTIAL,
                `AI timed out while extracting field: "${fieldInfo.name}".`
            );

            // Clean the result
            const cleanedResult = rawResult.trim().replace(/^"|"$/g, '');

            // Standardize "N/A" responses to empty strings for the sheet
            if (cleanedResult.toUpperCase() === 'N/A' || cleanedResult === '') {
                results[fieldInfo.name] = "";
                logToPanel(`  - Result: Not Found (N/A)`);
            } else {
                results[fieldInfo.name] = cleanedResult;
                logToPanel(`  - Result: "${cleanedResult}"`);
            }

        } catch (error) {
            logToPanel(`  - ❌ Error on field "${fieldInfo.name}": ${error.message}`);
            results[fieldInfo.name] = 'AI_ERROR';

            // If one iteration fails (especially due to timeout), stop the process.
            // isAITaskCancelled is already set by promiseWithTimeout if it timed out.
            if (!error.message.includes("timed out")) {
                // If it's a non-timeout error, ensure we still stop the loop
                isAITaskCancelled = true;
            }
            break;
        }
    }

    return results;
}


// --- Extraction Handlers ---

// Handles extraction for a single item (text or image)
async function handleSingleExtraction(inputData, fields) {
  isAITaskCancelled = false;
  let dataToFill = {};
  let session;

  try {
    await checkAIAvailability();
    
    logToPanel("Starting single item extraction...");

    let processedInputData = inputData;
    const inputType = typeof inputData === 'string' ? 'text' : 'image';

    // Input Truncation for stability
    if (inputType === 'text' && inputData.length > MAX_INPUT_LENGTH) {
        logToPanel(`Input text too long (${inputData.length} chars). Truncating to ${MAX_INPUT_LENGTH} chars.`);
        showToastInPanel(`Warning: Input truncated to ${MAX_INPUT_LENGTH} characters for stability.`, 'info');
        processedInputData = inputData.substring(0, MAX_INPUT_LENGTH) + "... [TRUNCATED]";
    }

    // Create the correct session type using the safe creator function
    if (inputType === 'text') {
        // This will throw if it times out
        session = await createAISession();
    } else {
        // Create a session that explicitly supports images (multimodal)
        try {
            // This will throw if it times out
            session = await createAISession({ expectedInputs: [{ type: 'image' }, { type: 'text' }] });
        } catch (error) {
            // Handle specific case where multimodal might not be supported, distinct from timeout
            if (!error.message.includes("timed out")) {
                logToPanel(`Failed to create multi-modal session: ${error.message}`);
                showToastInPanel("This AI model configuration does not support image analysis.", 'error');
            }
            throw new Error("Image analysis failed or timed out.");
        }
    }

    // Use the sequential extraction function
    dataToFill = await extractDataSequential(processedInputData, fields, session, inputType);
    
    if (!isAITaskCancelled) {
        logToPanel("✅ AI extraction complete. Sending data to panel.");
        chrome.runtime.sendMessage({ type: 'FILL_FIELDS', data: dataToFill });
    } else {
        // Provide feedback if the task was cancelled (by user or timeout)
        if (Object.values(dataToFill).includes('AI_ERROR')) {
            showToastInPanel("AI operation stopped due to timeout or error. Partial results may be shown.", 'error', 10000);
        } else {
           logToPanel("🛑 AI task cancelled by user.");
           showToastInPanel("AI operation cancelled by user.", 'info');
        }
    }
  } catch (error) {
    // Catch errors from checkAIAvailability, createAISession, or extractDataSequential
    logToPanel(`An error occurred during extraction: ${error.message}`);
  } finally {
    // Ensure session is always destroyed
    if (session) {
        session.destroy();
        logToPanel("AI session destroyed.");
    }
    chrome.runtime.sendMessage({ type: 'HIDE_SPINNER' });
  }
}

// Handles extraction for multiple items from a list (text only)
async function handleBatchExtraction(inputData, fields) {
    isAITaskCancelled = false;
    let session;

    try {
        await checkAIAvailability();
        
        logToPanel("Starting Batch Mode extraction...");

        // Input Truncation for stability
        let processedInputData = inputData;
        if (inputData.length > BATCH_MAX_LENGTH) {
            logToPanel(`Batch input text too long (${inputData.length} chars). Truncating to ${BATCH_MAX_LENGTH} chars.`);
            showToastInPanel(`Warning: Batch input truncated to ${BATCH_MAX_LENGTH} characters for stability.`, 'info');
            processedInputData = inputData.substring(0, BATCH_MAX_LENGTH) + "... [TRUNCATED]";
        }

        // Create session safely (this will throw if it times out)
        session = await createAISession();

        // Step 1: Ask AI to identify and separate distinct items in the input text
        logToPanel("Asking AI to split text into individual items...");
        const DELIMITER = "|||ITEM_SEPARATOR|||";
        const metaPrompt = `The following text contains a list or collection of similar items (e.g., search results, contacts, products). Identify each distinct item and separate them using the unique delimiter "${DELIMITER}". 
CRITICAL: Do not change, summarize, or omit any content within each item. Preserve the original text of each item exactly.
---START OF TEXT---
${processedInputData}
---END OF TEXT---`;

        // Use promiseWithTimeout for the meta-prompt
        const delimitedText = await promiseWithTimeout(
            session.prompt(metaPrompt),
            AI_TIMEOUT_MS_BATCH_SPLIT,
            `AI batch splitting timed out.`
        );

        const chunks = delimitedText.split(DELIMITER).filter(c => c.trim() !== '');
        
        if (chunks.length === 0) {
            logToPanel("AI could not identify any items in the selection.");
            showToastInPanel("Batch Mode: AI could not find distinct items in the selected text.", 'error');
            return;
        }

        // If only one chunk is found, process as a single extraction
        if (chunks.length === 1) {
            logToPanel("AI identified only 1 item. Processing as single extraction.");
            // Efficiently reuse the session for the single extraction using the sequential approach
            const dataToFill = await extractDataSequential(processedInputData, fields, session, 'text');
            if (!isAITaskCancelled) {
                // Send FILL_FIELDS, not FILL_AND_ADD, as it's treated as a single item
                chrome.runtime.sendMessage({ type: 'FILL_FIELDS', data: dataToFill });
            }
            return;
        }

        logToPanel(`AI identified ${chunks.length} items to process.`);

        // Step 2: Process each chunk individually using sequential extraction
        for (let i = 0; i < chunks.length; i++) {
            // Check if cancelled by user or by a previous timeout/error
            if (isAITaskCancelled) {
                logToPanel("🛑 Batch task stopped.");
                break;
            }
            logToPanel(`--- Processing item ${i + 1} of ${chunks.length} ---`);
            const chunkData = chunks[i];
            
            // Use the sequential extraction function (timeout is handled within)
            const dataToFill = await extractDataSequential(chunkData, fields, session, 'text');
            
            // If timeout/error occurred during this item (isAITaskCancelled will be true), stop the loop
            if (isAITaskCancelled) {
                showToastInPanel("Batch operation stopped due to timeout or error.", 'error', 10000);
                break;
            }

            // Send data to the side panel to fill AND add the row automatically
            chrome.runtime.sendMessage({ type: 'FILL_AND_ADD', data: dataToFill });
            // Add a small delay (800ms) between items to allow the UI to update visually
            await new Promise(resolve => setTimeout(resolve, 800)); 
        }
        
        if (!isAITaskCancelled) {
            logToPanel("✅ Batch extraction complete.");
            showToastInPanel(`Batch complete: Successfully processed ${chunks.length} items.`, 'success');
        }
    } catch (error) {
        logToPanel(`An error occurred during batch extraction: ${error.message}`);
        // Handle timeout during the initial splitting phase (if not already caught by session creation timeout)
        if (error.message.includes("timed out") && !error.message.includes("session creation")) {
            showToastInPanel("AI batch processing timed out during analysis. Try selecting fewer items.", 'error');
        }
    } finally {
        // Ensure session is always destroyed
        if (session) {
            session.destroy();
            logToPanel("AI session destroyed.");
        }
        chrome.runtime.sendMessage({ type: 'HIDE_SPINNER' });
    }
}

// --- Context Menu Handler ---

// Listener for the context menu click
chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  logToPanel("Context menu clicked.");

  // Ensure the side panel is open before proceeding (required in MV3)
  if (chrome.sidePanel && chrome.sidePanel.open) {
    await chrome.sidePanel.open({ tabId: tab.id });
  }
  
  // Wait briefly for the panel to potentially initialize if it was just opened
  await new Promise(resolve => setTimeout(resolve, 200));

  // Request the current fields and settings from the side panel
  let response;
  try {
    response = await chrome.runtime.sendMessage({ type: 'GET_FIELDS' });
  } catch (error) {
    logToPanel("Error communicating with side panel. Please ensure Pull2Sheet is open and ready.");
    console.error("GET_FIELDS failed:", error);
    return;
  }

  const fields = response?.fields;
  const isBatchMode = response?.isBatchMode;

  if (!fields || fields.length === 0) {
    logToPanel("Action stopped: No fields defined in side panel.");
    // The side panel will also show a toast if it receives the GET_FIELDS when empty.
    return;
  }

  // Show the spinner in the side panel
  chrome.runtime.sendMessage({ type: 'SHOW_SPINNER' });

  // Handle Text Selection
  if (info.selectionText) {
    const textSnippet = info.selectionText.length > 150 ? info.selectionText.substring(0, 150) + '...' : info.selectionText;
    logToPanel(`Processing selected text: "${textSnippet}"`);
    
    if (isBatchMode) {
        await handleBatchExtraction(info.selectionText, fields);
    } else {
        await handleSingleExtraction(info.selectionText, fields);
    }

  // Handle Image Selection
  } else if (info.mediaType === 'image' && info.srcUrl) {
    if (isBatchMode) {
        logToPanel("Batch mode is not supported for images. Processing as a single item.");
        showToastInPanel("Batch Mode is only available for text extraction.", 'info');
    }
    
    logToPanel(`Processing image from URL: ${info.srcUrl.substring(0, 100)}...`);
    try {
        // Fetch the image data. This requires host_permissions due to CORS.
        const response = await fetch(info.srcUrl);
        if (!response.ok) throw new Error(`Failed to fetch image: ${response.statusText} (Status ${response.status}). Potential CORS issue.`);
        const imageBlob = await response.blob();
        
        // Check if the fetched content is actually an image
        if (!imageBlob.type.startsWith('image/')) {
            throw new Error(`Fetched content is not a recognized image type (MIME type: ${imageBlob.type}).`);
        }

        logToPanel("Image fetched successfully as a blob.");
        await handleSingleExtraction(imageBlob, fields);
    } catch (error) {
        logToPanel(`Error handling image: ${error.message}`);
        showToastInPanel(`Error processing image. It might be protected, inaccessible (CORS), or invalid. See log for details.`, 'error', 10000);
        chrome.runtime.sendMessage({ type: 'HIDE_SPINNER' });
    }
  } else {
    // No recognized selection
    logToPanel("Context menu clicked, but no text or image selected.");
    chrome.runtime.sendMessage({ type: 'HIDE_SPINNER' });
  }
});