================================================================
Pull2Sheet - Local Privacy First AI-Powered Web Data Extraction Tool
================================================================

1. Introduction
2. The Problem It Solves
3. Key Features
4. Requirements & Setup (IMPORTANT)
5. Installation (Sideloading)
6. How to Use
7. APIs and Libraries Used

----------------------------------------------------------------

1. Introduction

Pull2Sheet is a powerful browser extension designed to streamline the process of collecting data from the web and organizing it into spreadsheets. By leveraging advanced, on-device Artificial Intelligence (AI), Pull2Sheet allows users to extract text and analyze images directly from any webpage and map the information directly into the columns of an existing spreadsheet file (.xlsx, .xls, .csv).

----------------------------------------------------------------

2. The Problem It Solves

Manually copying and pasting information from websites into spreadsheets is tedious, time-consuming, and prone to errors. Researchers, recruiters, marketers, and data analysts often spend hours collecting data. Pull2Sheet automates this process. It provides an intuitive, flexible way to capture structured data from unstructured text and images, using AI to understand the context and categorize the information instantly, boosting productivity and accuracy.

----------------------------------------------------------------

3. Key Features

*   AI-Powered Extraction: Uses the browser's built-in Language Model (e.g., Gemini Nano) to intelligently extract structured data.
*   Privacy Focused: All AI processing and file handling occur locally on your device. Your data does not leave your browser.
*   Optimized Performance: Utilizes advanced prompting techniques (JSON output mode) to extract all fields simultaneously, ensuring fast performance.
*   Image Analysis (OCR): Supports extracting information directly from images via the right-click context menu (requires a multi-modal capable AI model).
*   Spreadsheet Integration: Supports .xlsx, .xls, and .csv files, or starts new sheets from templates.
*   Automatic Field Detection (Optional/Longer processing on file upload): AI analyzes your spreadsheet headers upon loading to automatically suggest data types, formats, and keywords.
*   Customizable Mapping: Manually define expected data types, formats, and keywords/aliases to fine-tune AI accuracy.
*   Template System: Save field configurations as JSON templates for reuse.
*   Data Locking: Lock specific fields to prevent them from being cleared or overwritten during subsequent extractions.


----------------------------------------------------------------

4. Requirements & Setup (IMPORTANT)

This extension relies entirely on the browser's built-in, on-device AI model (Gemini Nano). This feature is currently emerging and requires specific setup.

ACTION REQUIRED:
You must ensure the AI model is enabled and downloaded before use.

1. Browser Version: Ensure you are running Google Chrome version 127 or later.
2. Enable Experimental AI: Navigate to `chrome://flags/#prompt-api-for-gemini-nano-multimodal-input` in your browser and ensure the built-in AI features are enabled.
3. Verify Model Download: The model downloads automatically when enabled, but this can take time. To check the status, navigate to `chrome://on-device-internals/` and Ensure Model status is Ready, And navigate to `chrome://components/` and locate "Optimization Guide On Device Model". Ensure the status is "Up-to-date". 

If the download is missing, in progress, slow, or shows an error, the extension's AI features cannot function.

For detailed instructions, requirements, and troubleshooting regarding the AI model setup, please refer to the official documentation:
https://developer.chrome.com/docs/ai/get-started#model_download

----------------------------------------------------------------

5. Installation (Sideloading)

If you are installing this extension from source files, follow these steps to "sideload" it:

1. Download and unzip the extension files to a folder on your computer.
2. Open your Chrome browser and navigate to: `chrome://extensions/`
3. In the top right corner, enable "Developer mode".
4. Click the "Load unpacked" button that appears.
5. Select the folder where you unzipped the extension files and click "Select Folder".
6. The Pull2Sheet extension should now appear in your list of extensions.
7. Access the Extension: Click the Extensions icon (Pull2Sheet Logo) in the Chrome toolbar, and then click the "Pin" icon next to Pull2Sheet for easy access.

----------------------------------------------------------------

6. How to Use

1. Open the Side Panel: Click the Pull2Sheet extension icon in your toolbar to open the side panel interface.
2. Load Data Structure: Click "Choose Sheet File" to upload your spreadsheet OR click "Load Template (.json)" to start a new sheet from a saved template.
3. Configure Fields (Optional): Review the 'TYPE', 'FORMAT', and 'KEYWORDS' tags. Providing 'KEYWORDS' (aliases for your headers) significantly improves AI accuracy to the structured output you want.
4. Extract Data: Navigate to a webpage. Highlight the text you want to extract OR right-click an image. Select "Extract to Pull2Sheet" from the context menu.
5. Review and Add: The AI fills the fields in the side panel. Review the data for accuracy, make any necessary edits, and click "Add Row to Sheet".
6. Batch Mode (Optional): For lists or search results, open the settings (⚙️) and enable "Batch Extraction Mode" before extracting. The AI will automatically extract and add rows for each item found in the selection (Note: Limit to few items at a time since model context window is limited & processing time would be high).
7. Download: Once you have added one or more rows, the "Download Updated Sheet" button will appear. Click it to save your updated file.

----------------------------------------------------------------

7. APIs and Libraries Used

Pull2Sheet utilizes several modern web and browser APIs:

*   Chrome Extension APIs:
    - chrome.sidePanel: Provides the user interface docked to the side of the browser.
    - chrome.contextMenus: Enables the "Extract to Pull2Sheet" option in the right-click menu.
    - chrome.storage: Used to save user preferences (e.g., AI settings, Batch Mode toggle) locally.
    - chrome.runtime: Facilitates communication between the background script (AI tasks) and the side panel UI.

*   LanguageModel API:
    - The core of the AI functionality. This API provides access to the browser's built-in, on-device AI models (See Requirements).
    - It is used for analyzing spreadsheet headers, extracting structured data from text (using optimized JSON output mode), analyzing images (OCR), and splitting lists in Batch Mode.

*   Web APIs:
    - Fetch API and Blob API: Used in the background script to retrieve image data when the user selects an image for extraction (requires host permissions due to CORS).
    - File API and FileReader API: Used in the side panel to handle the uploading and reading of spreadsheet files and JSON templates.

*   Third-Party Libraries:
    - SheetJS (XLSX.js): A JavaScript library (xlsx.full.min.js) used for reading, processing, and writing various spreadsheet formats entirely within the browser.
