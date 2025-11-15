document.addEventListener('DOMContentLoaded', () => {    
    // --- URL Security Check (Dormant) --- 
    // To activate this check for deployment, change ACTIVATE_URL_CHECK to true.
    // This will ensure the application only runs from the specified domain.
    const ACTIVATE_URL_CHECK = false; 
    const encodedExpectedOrigin = 'aHR0cHM6Ly9maWxlc21hdGNoLmNvbQ==';

    // --- Feature Flags ---
    // Set to true to enable experimental support for XML and JSON file types.
    const ENABLE_XML_JSON_SUPPORT = true;

    if (ACTIVATE_URL_CHECK) {
        try {
            const expectedOrigin = atob(encodedExpectedOrigin);
            const currentOrigin = window.location.origin;

            if (currentOrigin !== expectedOrigin) {
                // If origins do not match, replace the page content with an error
                // and stop the rest of the script from executing.
                document.body.innerHTML = `
                    <div style="text-align: center; padding: 50px; font-family: sans-serif; color: #212529;">
                        <h1>Access Denied</h1>
                        <p>This application can only be run from the authorized domain.</p>
                    </div>`;
                throw new Error("Invalid origin. Application halted.");
            }
        } catch (e) {
            console.error(e);
            return; // Stop execution if there's an error in the check.
            // Re-throw the error to halt script execution, as 'return' is not allowed in the global scope.
            throw e;
        }
    }
    // DOM Elements
    const termsModal = document.getElementById('terms-modal');
    const acceptTermsBtn = document.getElementById('accept-terms-btn');
    const showTermsLink = document.getElementById('show-terms-link');
    const privacyModal = document.getElementById('privacy-modal');
    const privacyModalCloseBtn = document.getElementById('privacy-modal-close-btn');
    const showPrivacyLink = document.getElementById('show-privacy-link');
    const contactModal = document.getElementById('contact-modal');
    const contactModalCloseBtn = document.getElementById('contact-modal-close-btn');
    const showContactLink = document.getElementById('show-contact-link');
    const aboutModal = document.getElementById('about-modal');
    const aboutModalCloseBtn = document.getElementById('about-modal-close-btn');
    const showAboutLink = document.getElementById('show-about-link');

    // Main application elements
    const fileInput1 = document.getElementById('file1');
    const delimiter1 = document.getElementById('delimiter1');
    const delimiterWrapper1 = document.getElementById('delimiter-wrapper-1');
    const preview1 = document.getElementById('preview1');

    const fileInput2 = document.getElementById('file2');
    const file2Label = document.getElementById('file2-label');
    const delimiter2 = document.getElementById('delimiter2');
    const delimiterWrapper2 = document.getElementById('delimiter-wrapper-2');
    const preview2 = document.getElementById('preview2');

    const runComparisonBtn = document.getElementById('run-comparison-btn');
    const clearMatchesBtn = document.getElementById('clear-matches-btn');
    const manualMatchBtn = document.getElementById('manual-match-btn');
    const resetAllBtn = document.getElementById('reset-all-btn');
    
    const valueMappingModal = document.getElementById('value-mapping-modal');
    const matchingActionsContainer = document.getElementById('matching-actions-container');


    const groupByCheckbox = document.getElementById('group-by-key');
    const resultsSection = document.getElementById('results-section');
    const resultsTableContainer = document.getElementById('results-table-container');
    const resultsSummary = document.getElementById('results-summary');
    const resultsOptions = document.getElementById('results-options');
    const hideUnmatchedKeysCheckbox = document.getElementById('hide-unmatched-keys-checkbox');
    const showDiffOnlyCheckbox = document.getElementById('show-diff-only-checkbox');
    const showDiffColumnsCheckbox = document.getElementById('show-diff-columns-checkbox');
    const numericToleranceInput = document.getElementById('numeric-tolerance-input');
    
    // Modal elements
    const modal = document.getElementById('fixed-width-modal');
    const modalCloseBtn = document.getElementById('modal-close-btn');
    const fixedWidthPreview = document.getElementById('fixed-width-preview');
    const breakLinesOverlay = document.getElementById('break-lines-overlay');
    const applyFixedWidthBtn = document.getElementById('apply-fixed-width-btn');

    const suggestFixedWidthBtn = document.getElementById('suggest-fixed-width-btn');
    // Value Mapping Modal elements
    const valueMapModalCloseBtn = document.getElementById('value-map-modal-close-btn');
    const applyValueMapBtn = document.getElementById('apply-value-map-btn');
    const unmapAllBtn = document.getElementById('unmap-all-btn');

    // Excel Options Modal elements
    const excelOptionsModal = document.getElementById('excel-options-modal');
    const excelModalCloseBtn = document.getElementById('excel-modal-close-btn');
    const worksheetSelect = document.getElementById('worksheet-select');
    const cellRangeInput = document.getElementById('cell-range-input');
    const applyExcelOptionsBtn = document.getElementById('apply-excel-options-btn');

    // Header Definition Modal elements (for inconsistent CSVs)
    const headerModal = document.getElementById('header-definition-modal');
    const headerModalCloseBtn = document.getElementById('header-modal-close-btn');
    const headerPreview = document.getElementById('header-preview-table');
    const customHeadersContainer = document.getElementById('custom-csv-headers-container');
    const applyHeaderDefinitionBtn = document.getElementById('apply-header-definition-btn');

    // Template elements
    const templateManagementContainer = document.getElementById('template-management-container');
    const templateDropdownBtn = document.getElementById('template-dropdown-btn');
    const templateDropdownMenu = document.getElementById('template-dropdown-menu');
    const loadTemplateAction = document.getElementById('load-template-action');
    const saveTemplateAction = document.getElementById('save-template-action');
    const applyTemplateInput = document.getElementById('apply-template-input');
    const saveTemplateModal = document.getElementById('save-template-modal');
    const saveTemplateModalCloseBtn = document.getElementById('save-template-modal-close-btn');
    const confirmSaveTemplateBtn = document.getElementById('confirm-save-template-btn');


    const canvas = document.getElementById('line-canvas');
    const ctx = canvas.getContext('2d');

    // State
    let fileData1 = null;
    let fullFileData1 = null; // Holds the original, unfiltered data
    let fileData2 = null;
    let fullFileData2 = null; // Holds the original, unfiltered data
    let unmatchedHeaders1 = new Set();
    let unmatchedHeaders2 = new Set(); // A pair is an object: { file1ColIndex, file2ColIndex, file1Th, file2Th, type: 'key' | 'compare' }
    let comparisonPairs = []; // New state for non-key pairs
    let columnPairs = []; // { file1ColIndex, file2ColIndex, file1Th, file2Th }
    let alignedData1 = null; // To hold comparison-ready data for file 1
    let selectedTh = null; // { el, fileIndex, colIndex }
    let rightClickSelectedTh = null; // For comparison-only pairing
    let comparisonResults = [];
    let alignedOriginalHeaders2ForRender = []; // To pass headers between functions
    let activeFileForModal = null; // { fileIndex, file }
    let breakPositions = []; // Array of character indices for breaks
    let inconsistentCsvData = null; // { fileIndex, data: string[][] }
    let keyValueMappings = new Map(); // Stores manual value mappings: file1Value -> file2Value
    let suggestedKeyValueMappings = new Set(); // Stores file1Values that were auto-suggested

    let nonSummableColumnsForDisclaimer = []; // To hold names of columns excluded from grouped sum
    let hasClearedMappings = false; // Flag to track if user has intentionally cleared mappings
    let templateApplied = false; // Flag to indicate if a template's state is active
    let primaryKeySource = 1; // 1 or 2, for grouped results display key
    let appliedFilters1 = []; // Store the currently applied filters for file 1
    let appliedFilters2 = []; // Store the currently applied filters for file 2
    let fixedWidthSettings1 = null; // { breaks, useFirstRowAsHeader, customHeaders }
    let fixedWidthSettings2 = null; // { breaks, useFirstRowAsHeader, customHeaders }
    let file1WasInconsistent = false; // Flag to track if a file was loaded as inconsistent
    let file2WasInconsistent = false; // Flag to track if a file was loaded as inconsistent
    let file1HasNoHeader = false; // Flag for "File has no header" checkbox state
    let file2HasNoHeader = false; // Flag for "File has no header" checkbox state
    let sortState = {
        columnIndex: -1,
        direction: 'asc' // 'asc' or 'desc'
    };
    let isCancellingNoHeader = false; // Flag to prevent re-opening the header modal on cancel.
    let isApplyingTemplate = false; // Flag to change behavior during template application.
    let isDragDisabled = false; // Flag to disable column dragging after an action is taken.
    let isTutorialActive = false; // Flag to track if the tutorial is running.
    // --- Context Menu ---
    let numericColumns = new Set(); // Holds headers of columns identified as purely numeric post-filtering.
    let columnTypes1 = new Map(); // { [headerName]: 'text' | 'number' | 'date' | 'auto' }
    let columnTypes2 = new Map();
    let typeMenu = null; // To hold the single instance of the type menu

    let contextMenu = null;


    // --- Utility Functions ---
    function debounce(func, delay = 250) {
        let timeoutId;
        return function(...args) {
            clearTimeout(timeoutId);
            timeoutId = setTimeout(() => {
                func.apply(this, args);
            }, delay);
        };
    }

    function levenshteinDistance(a, b) {
        if (a.length === 0) return b.length;
        if (b.length === 0) return a.length;

        const matrix = Array(b.length + 1).fill(null).map(() => Array(a.length + 1).fill(null));

        for (let i = 0; i <= a.length; i++) { matrix[0][i] = i; }
        for (let j = 0; j <= b.length; j++) { matrix[j][0] = j; }

        for (let j = 1; j <= b.length; j++) {
            for (let i = 1; i <= a.length; i++) {
                const cost = a[i - 1] === b[j - 1] ? 0 : 1;
                matrix[j][i] = Math.min(
                    matrix[j - 1][i] + 1,      // deletion
                    matrix[j][i - 1] + 1,      // insertion
                    matrix[j - 1][i - 1] + cost // substitution
                );
            }
        }
        const distance = matrix[b.length][a.length];
        const maxLength = Math.max(a.length, b.length);
        if (maxLength === 0) return 1; // Both are empty, 100% similar
        return 1 - (distance / maxLength); // Return similarity ratio
    }

    /**
     * Converts a hex color string to an rgba string with a given alpha.
     * @param {string} hex The hex color (e.g., '#RRGGBB').
     * @param {number} alpha The alpha transparency (0.0 to 1.0).
     * @returns {string} The rgba color string.
     */
    function hexToRgba(hex, alpha = 1) {
        if (!hex || !/^#([A-Fa-f0-9]{3}){1,2}$/.test(hex)) {
            // Fallback for invalid hex or if a CSS variable isn't found
            return `rgba(0, 0, 0, ${alpha})`;
        }
        let c = hex.substring(1).split('');
        if (c.length === 3) {
            c = [c[0], c[0], c[1], c[1], c[2], c[2]];
        }
        c = '0x' + c.join('');
        const r = (c >> 16) & 255;
        const g = (c >> 8) & 255;
        const b = c & 255;
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }
    /**
     * Ensures all headers in an array are unique by appending a suffix (_1, _2, etc.) to duplicates.
     * @param {string[]} headers - An array of header names.
     * @returns {string[]} A new array with unique header names.
     */
    function uniquifyHeaders(headers) {
        const counts = {};
        return headers.map(header => {
            const originalHeader = header.trim();
            if (counts[originalHeader]) {
                const newHeader = `${originalHeader}_${counts[originalHeader]++}`;
                return newHeader;
            } else {
                counts[originalHeader] = 1;
                return originalHeader;
            }
        });
    }

    /**
     * Analyzes a piece of text to detect the most likely delimiter.
     * @param {string} text - The first few lines of the file content.
     * @returns {string|null} The detected delimiter or null if not found.
     */
    function detectDelimiter(text) {
        const delimiters = [',', ';', '\t', '|'];
        // Use a small sample of lines for performance
        const lines = text.split(/\r?\n/).slice(0, 10);

        // Not enough data to make a reliable guess
        if (lines.length < 2) return null;

        for (const delimiter of delimiters) {
            const firstLineCount = (lines[0].split(delimiter).length - 1);

            // If the first line has no delimiters, it's not a good candidate
            if (firstLineCount === 0) continue;

            // Check if subsequent lines have a consistent number of delimiters
            let isConsistent = true;
            for (let i = 1; i < lines.length; i++) {
                if (lines[i].trim() === '') continue; // Skip empty lines
                if ((lines[i].split(delimiter).length - 1) !== firstLineCount) {
                    isConsistent = false;
                    break;
                }
            }
            if (isConsistent) return delimiter;
        }
        return null; // No consistent delimiter found
    }

    /**
     * Compares two values, treating them as numbers if they appear to be numeric.
     * This prevents false differences between formats like "100" and "100.00".
     * @param {*} v1 The first value.
     * @param {*} v2 The second value.
     * @returns {boolean} True if the values are different, false otherwise.
     */
    function areValuesDifferent(v1, v2, h1, h2, type1, type2) {
        const tolerance = parseFloat(numericToleranceInput.value) || 0;
        const effectiveType = type1 !== 'auto' ? type1 : (type2 !== 'auto' ? type2 : 'auto');

        if (effectiveType === 'number') {
            // Coerce both to numbers, treating null/undefined/empty-string as 0.
            const num1 = Number(v1) || 0;
            const num2 = Number(v2) || 0;
            return Math.abs(num1 - num2) > tolerance;
        }
        if (effectiveType === 'date') {
            const d1 = v1 ? new Date(v1).getTime() : 0;
            const d2 = v2 ? new Date(v2).getTime() : 0;
            return (isNaN(d1) ? 0 : d1) !== (isNaN(d2) ? 0 : d2);
        }

        // Highest priority: If at least one value is a number (from grouped sums), perform a direct numeric comparison.
        // This correctly handles number vs number, number vs null, and number vs string representations.
        if (typeof v1 === 'number' || typeof v2 === 'number') {
            const num1 = Number(v1) || 0;
            const num2 = Number(v2) || 0;
            return Math.abs(num1 - num2) > tolerance;
        }

        // For all other cases, trim and convert to string for consistent comparison.
        const val1 = (v1 ?? '').toString().trim();
        const val2 = (v2 ?? '').toString().trim();

        // If the column was pre-identified as numeric (legacy), force numeric comparison.
        const forceNumeric = (h1 && numericColumns.has(h1)) || (h2 && numericColumns.has(h2)) || effectiveType === 'number';

        // Regex to check for a string that is a valid number (integer or float).
        const isNumericRegex = /^-?\d*\.?\d+$/;

        // If forcing numeric OR if both values look like numbers, perform a numeric comparison.
        // This handles row-by-row string numbers and mixed types in grouped results (e.g., a sum vs a single string value).
        if (forceNumeric || (isNumericRegex.test(val1) && isNumericRegex.test(val2))) {
            // Use parseFloat on the trimmed strings. parseFloat handles empty strings as NaN, so we default to 0.
            const num1 = parseFloat(val1) || 0;
            const num2 = parseFloat(val2) || 0;
            return Math.abs(num1 - num2) > tolerance;
        }

        // Otherwise, perform a standard string comparison on the trimmed values.
        return val1 !== val2;
    }

    /**
     * Formats two numeric string values to have the same number of decimal places.
     * @param {string} v1 The first value.
     * @param {string} v2 The second value.
     * @param {string} h1 The header for the first value.
     * @param {string} h2 The header for the second value.
     * @returns {{f1: string, f2: string}} An object with the formatted values.
     */
    function formatNumericValues(v1, v2, h1, h2, type1, type2) {
        const s1 = (v1 ?? '').toString().trim();
        const s2 = (v2 ?? '').toString().trim();

        const effectiveType = type1 !== 'auto' ? type1 : (type2 !== 'auto' ? type2 : 'auto');

        const isNumericRegex = /^-?\d*\.?\d+$/;
        // This formatting should only apply if the column is identified as numeric.
        if (effectiveType === 'number' || (h1 && numericColumns.has(h1)) || (h2 && numericColumns.has(h2))) {
            // Parse numeric strings to numbers and convert back to strings.
            // This standardizes the format by removing trailing zeros (e.g., "100.00" becomes "100").
            const f1 = isNumericRegex.test(s1) ? parseFloat(s1).toString() : s1;
            const f2 = isNumericRegex.test(s2) ? parseFloat(s2).toString() : s2;

            return { f1, f2 };
        }

        // If not a numeric column, return original (trimmed) values.
        return { f1: s1, f2: s2 }; // Return original (trimmed) values if not numeric
    }
    // --- Terms and Conditions Modal Logic ---
    // The modal is no longer shown automatically on first load.
    // Users can view it by clicking the link in the footer.
    // if (localStorage.getItem('termsAccepted') !== 'true') { termsModal.classList.remove('hidden'); }

    acceptTermsBtn.addEventListener('click', () => {
        localStorage.setItem('termsAccepted', 'true');
        termsModal.classList.add('hidden');
    });

    showTermsLink.addEventListener('click', (e) => {
        e.preventDefault();
        termsModal.classList.remove('hidden');
    });

    showPrivacyLink.addEventListener('click', (e) => {
        e.preventDefault();
        privacyModal.classList.remove('hidden');
    });

    privacyModalCloseBtn.addEventListener('click', () => {
        privacyModal.classList.add('hidden');
    });

    showContactLink.addEventListener('click', (e) => {
        e.preventDefault();
        contactModal.classList.remove('hidden');
    });

    contactModalCloseBtn.addEventListener('click', () => {
        contactModal.classList.add('hidden');
    });

    showAboutLink.addEventListener('click', (e) => {
        e.preventDefault();
        aboutModal.classList.remove('hidden');
    });

    aboutModalCloseBtn.addEventListener('click', () => {
        aboutModal.classList.add('hidden');
    });

    // --- "No Header" Checkbox Listeners ---
    // This is the fix: Trigger a re-parse when the "no header" checkbox is changed.
    document.getElementById('no-header-checkbox-1').addEventListener('change', (e) => {
        // Only re-parse if a file is actually loaded
        if (fileInput1.files.length > 0) {
            // Update state when checkbox changes
            file1HasNoHeader = e.target.checked;
            parseAndRenderDelimitedFile(1);
        }
    });
    document.getElementById('no-header-checkbox-2').addEventListener('change', (e) => {
        if (fileInput2.files.length > 0) {
            parseAndRenderDelimitedFile(2);
        }
    });

    // --- Event Listeners ---
    // --- UI Helper Functions ---
    function showToast(message, type = 'success', duration = 7000) {
        const toastContainer = document.getElementById('toast-container');
        if (!toastContainer) return;

        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        
        const messageEl = document.createElement('div');
        messageEl.className = 'toast-message';
        messageEl.textContent = message;
        toast.appendChild(messageEl);

        toastContainer.appendChild(toast);

        setTimeout(() => toast.classList.add('show'), 100);

        setTimeout(() => {
            toast.classList.remove('show');
            toast.addEventListener('transitionend', () => toast.remove());
        }, duration);
    }

    function showPreviewLoader(fileIndex, show = true, message = 'Processing...') {
        const previewWrapper = document.getElementById(`preview-wrapper-${fileIndex}`);
        if (!previewWrapper) return;

        let loader = previewWrapper.querySelector('.preview-loader');
        if (show) {
            if (!loader) {
                loader = document.createElement('div');
                loader.className = 'preview-loader';
                loader.innerHTML = `<div class="spinner"></div><p>${message}</p>`;
                previewWrapper.style.position = 'relative'; // For positioning the loader
                previewWrapper.insertBefore(loader, previewWrapper.firstChild);
            }
            loader.querySelector('p').textContent = message;
            loader.classList.remove('hidden');
        } else {
            if (loader) {
                loader.classList.add('hidden');
            }
        }
    }

    fileInput1.addEventListener('change', () => handleFileSelection(1));

    fileInput2.addEventListener('change', () => handleFileSelection(2));

    runComparisonBtn.addEventListener('click', runComparison);
    clearMatchesBtn.addEventListener('click', clearAllMatches);
    manualMatchBtn.addEventListener('click', () => openValueMappingModal());
    resetAllBtn.addEventListener('click', resetAll);

    document.querySelector('.previews-container').addEventListener('click', (e) => {
        const target = e.target;
        const button = target.closest('[data-file-index]');
        if (!button) return;
        const fileIndex = parseInt(button.dataset.fileIndex, 10);

        if (button.matches('.add-rule-btn')) addFilterRule(fileIndex);
        if (button.matches('.apply-filters-btn')) applyFilters(fileIndex);
        if (button.matches('.toggle-filter-btn')) toggleFilterBuilder(fileIndex);
    });

    // --- Download Dropdown Logic ---
    const downloadBtn = document.getElementById('download-results-btn');
    const downloadDropdown = document.getElementById('download-dropdown-menu');
    const downloadContainer = document.querySelector('.dropdown-container');

    downloadBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        downloadDropdown.classList.toggle('hidden');
        downloadContainer.classList.toggle('open');
    });

    document.addEventListener('click', (e) => {
        if (!downloadBtn.contains(e.target) && !downloadDropdown.contains(e.target)) {
            downloadDropdown.classList.add('hidden');
            downloadContainer.classList.remove('open');
        }
        // Hide the type menu if clicking anywhere else
        if (typeMenu && !typeMenu.contains(e.target) && !e.target.classList.contains('type-menu-trigger')) {
            typeMenu.remove();
            typeMenu = null;
        }
    });

    document.getElementById('download-csv-action').addEventListener('click', downloadResultsAsCSV);
    document.getElementById('download-excel-action').addEventListener('click', downloadResultsAsExcel);

    showDiffOnlyCheckbox.addEventListener('change', () => {
        // Re-render the results table with the current filter state
        renderResultsTable(alignedOriginalHeaders2ForRender);
    });

    hideUnmatchedKeysCheckbox.addEventListener('change', () => {
        // Re-render the results table with the current filter state
        renderResultsTable();
    });

    showDiffColumnsCheckbox.addEventListener('change', () => {
        // Re-render the results table with the current filter state
        renderResultsTable();
    });

    numericToleranceInput.addEventListener('input', () => {
        renderResultsTable();
    });

    //groupByCheckbox.addEventListener('change', updateUIState);
    groupByCheckbox.addEventListener('change', () => { resultsSection.classList.add('hidden'); });

    window.addEventListener('resize', debounce(redrawLines, 100));
    
    // Modal listeners
    modalCloseBtn.addEventListener('click', () => {
        modal.classList.add('hidden');
        // If the tutorial is active, end it because the user cancelled.
        if (isTutorialActive) {
            endTutorial(); // Explicitly hide tutorial UI
            resetAll(); // Reset application state
            showToast("Tutorial ended because the modal was closed.", "warning");
        }
    });
    applyFixedWidthBtn.addEventListener('click', applyFixedWidthParsing);
    suggestFixedWidthBtn.addEventListener('click', suggestFixedWidthBreaks);

    // Value Mapping Modal listeners
    valueMapModalCloseBtn.addEventListener('click', () => valueMappingModal.classList.add('hidden'));
    applyValueMapBtn.addEventListener('click', () => {
        valueMappingModal.classList.add('hidden');
    });

    // Excel Options Modal listeners
    excelModalCloseBtn.addEventListener('click', () => {
        excelOptionsModal.classList.add('hidden');
        if (activeFileForModal) {
            // If the user closes the modal, treat it as cancelling the file selection.
            if (isTutorialActive) {
                endTutorial(); // Explicitly hide tutorial UI
                resetAll(); // Reset application state
                showToast("Tutorial ended because the modal was closed.", "warning");
            }
            resetFileUI(activeFileForModal.fileIndex);
            document.getElementById(`file${activeFileForModal.fileIndex}`).value = '';
            document.getElementById(`file${activeFileForModal.fileIndex}-name`).textContent = 'No file chosen';
        }
    });
    applyExcelOptionsBtn.addEventListener('click', applyExcelOptions);
    worksheetSelect.addEventListener('change', handleWorksheetChange);
    unmapAllBtn.addEventListener('click', () => {
        // If the tutorial is active, end it.
        if (isTutorialActive) {
            endTutorial();
            showToast("Tutorial ended because value mappings were cleared.", "warning");
        }
        keyValueMappings.clear(); // Clear all existing mappings
        suggestedKeyValueMappings.clear();
        openValueMappingModal(true, true); // Re-open/refresh the modal, skipping auto-pairing and treating as a refresh
        hasClearedMappings = true; // Set the flag
        resultsSection.classList.add('hidden'); // Hide results as they are now invalid
    });

    // Primary Key Source for Grouped Results listener
    valueMappingModal.addEventListener('change', (e) => {
        if (e.target.name === 'primary-key-source') {
            primaryKeySource = parseInt(e.target.value, 10);
        }
    });

    // Template event listeners
    templateDropdownBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        templateDropdownMenu.classList.toggle('hidden');
        templateManagementContainer.classList.toggle('open');
    });

    document.addEventListener('click', (e) => {
        if (!templateDropdownBtn.contains(e.target) && !templateDropdownMenu.contains(e.target)) {
            templateDropdownMenu.classList.add('hidden');
            templateManagementContainer.classList.remove('open');
        }
    });

    loadTemplateAction.addEventListener('click', (e) => {
        e.preventDefault();
        applyTemplateInput.click();
        templateDropdownMenu.classList.add('hidden');
    });

    saveTemplateAction.addEventListener('click', (e) => {
        e.preventDefault();
        handleSaveTemplate();
        templateDropdownMenu.classList.add('hidden');
    });
    // Header Definition Modal listeners    
    headerModalCloseBtn.addEventListener('click', () => {
        // Check if the modal was opened for a "no header" file.
        const noHeaderMode = !document.getElementById('header-definition-p-no-header').classList.contains('hidden');
        
        if (noHeaderMode && inconsistentCsvData) {
            const { fileIndex } = inconsistentCsvData;
            const checkbox = document.getElementById(`no-header-checkbox-${fileIndex}`);
            if (checkbox && checkbox.checked) {
                // User is cancelling. Set a flag, uncheck the box, and dispatch the change event.
                // The change event handler will see the flag and prevent a re-parse loop.
                isCancellingNoHeader = true;
                checkbox.checked = false;
                checkbox.dispatchEvent(new Event('change'));
            }
        }
        headerModal.classList.add('hidden');
    });
    applyHeaderDefinitionBtn.addEventListener('click', applyHeaderDefinition);
    saveTemplateModalCloseBtn.addEventListener('click', () => saveTemplateModal.classList.add('hidden'));
    confirmSaveTemplateBtn.addEventListener('click', confirmSaveTemplate);


    applyTemplateInput.addEventListener('change', handleApplyTemplate); // The label's 'for' attribute handles the click

    function processFile(file, fileIndex) {
        const fileInput = fileIndex === 1 ? fileInput1 : fileInput2;
        // If the file is invalid, reset and exit.
        if (!file) {
            return;
        }

        // Check file size
        if (isTutorialActive) { // Use the state flag
            const MAX_TUTORIAL_FILE_SIZE = 20 * 1024 * 1024; // 20 MB
            if (file.size > MAX_TUTORIAL_FILE_SIZE) {
                showToast(`For the tutorial, please select a file smaller than 20 MB.`, "warning", 10000);
                fileInput.value = ''; // Clear the invalid file from the input
                resetFileUI(fileIndex);
                return;
            }
        } else {
            // General file size limit for normal use
            const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100 MB
            if (file.size > MAX_FILE_SIZE) {
                showToast(`File size cannot exceed 100 MB. Please select a smaller file.`, "error", 10000);
                fileInput.value = ''; // Clear the invalid file from the input
                resetFileUI(fileIndex);
                return;
            }
        }

        // 1. A new file has been chosen, so reset the UI state for this file index.
        resetFileUI(fileIndex);

        // 2. Update the file name display with the new file.
        document.getElementById(`file${fileIndex}-name`).textContent = file.name;

        // Show the "Clear All" button only when file 1 is loaded.
        if (fileIndex === 1) {
            resetAllBtn.classList.remove('hidden');
        }

        // 3. Show the delimiter dropdown for the user to make a selection.
        const fileExtension = file.name.split('.').pop().toLowerCase();
        if (['xlsx', 'xlsm'].includes(fileExtension)) {
            document.getElementById(`delimiter-wrapper-${fileIndex}`).classList.add('hidden');
            openExcelOptionsModal(fileIndex, file);
        } else if (ENABLE_XML_JSON_SUPPORT && fileExtension === 'xml') {
            document.getElementById(`delimiter-wrapper-${fileIndex}`).classList.add('hidden');
            parseXmlFile(file, fileIndex);
        } else if (ENABLE_XML_JSON_SUPPORT && fileExtension === 'json') {
            document.getElementById(`delimiter-wrapper-${fileIndex}`).classList.add('hidden');
            parseJsonFile(file, fileIndex);
        } else {
            const delimiterSelect = document.getElementById(`delimiter${fileIndex}`);
            const delimiterWrapper = document.getElementById(`delimiter-wrapper-${fileIndex}`);
            delimiterWrapper.classList.remove('hidden'); // Show the delimiter dropdown area
            delimiterSelect.disabled = false;
            // Attempt to auto-detect the delimiter
            const reader = new FileReader();
            reader.onload = (e) => {
                const text = e.target.result;
                const detectedDelimiter = detectDelimiter(text);

                if (detectedDelimiter) {
                    delimiterSelect.value = detectedDelimiter;
                    // Since we found a delimiter, we can parse immediately.
                    parseAndRenderDelimitedFile(fileIndex);
                } else {
                    // Could not detect, let the user choose.
                    // The dropdown is already visible.
                }
            };
            // Read only the first ~10KB for detection to be fast.
            const blob = file.slice(0, 10240);
            reader.readAsText(blob);
        }
    }

    function handleFileSelection(fileIndex) {
        const fileInput = fileIndex === 1 ? fileInput1 : fileInput2;
        const file = fileInput.files[0];

        if (file) {
            processFile(file, fileIndex);
        } else {
            // Handle cancellation from the file dialog
            document.getElementById(`file${fileIndex}-name`).textContent = 'No file chosen';
            resetFileUI(fileIndex);
            if (isTutorialActive) {
                endTutorial();
                showToast("Tutorial ended because file selection was cancelled.", "warning");
            }
        }
    }
    
    function resetFileUI(fileIndex) {
        document.getElementById(`preview-wrapper-${fileIndex}`).classList.add('hidden');
        document.getElementById(`preview${fileIndex}`).innerHTML = ''; // Clear preview content
        document.getElementById(`delimiter${fileIndex}`).value = ''; // Reset delimiter dropdown
        const noHeaderCheckbox = document.getElementById(`no-header-checkbox-${fileIndex}`);
        // Hide the "no header" control and uncheck it
        if (noHeaderCheckbox) noHeaderCheckbox.checked = false;
        const noHeaderControl = document.querySelector(`#delimiter-wrapper-${fileIndex} .no-header-control`);
        if (noHeaderControl) noHeaderControl.style.display = 'none';
        document.getElementById(`filter-builder-${fileIndex}`).classList.add('hidden');
        document.getElementById(`delimiter${fileIndex}`).disabled = true;
        excelOptionsModal.classList.add('hidden');
        // Also remove the 'other' delimiter input if it exists
        const customInput = document.getElementById(`custom-delimiter-${fileIndex}`);
        if (customInput) {
            customInput.remove();
        }
        document.getElementById(`delimiter-wrapper-${fileIndex}`).classList.add('hidden'); // Hide the entire delimiter area
        document.querySelector(`.toggle-filter-btn[data-file-index="${fileIndex}"]`).style.display = 'none';

        // If we are resetting file 1, it triggers a full cascade reset.
        // --- BUG FIX ---
        // Explicitly clear the data state for the selected file index.
        // This prevents a previous file's preview from appearing behind the
        // "Define Headers" modal if a new, inconsistent file is selected.
        if (fileIndex === 1) { fileData1 = null; fullFileData1 = null; }
        else { fileData2 = null; fullFileData2 = null; }
            // --- BUG FIX ---
            // Also clear the inconsistent data state. If this is not cleared,
            // loading a valid file after an invalid one will cause the spinner
        // to get stuck.
            inconsistentCsvData = null;
        if (fileIndex === 1) {
            file1HasNoHeader = false;
        } else {
            file2HasNoHeader = false;
        }

            if (fileIndex === 1) file1WasInconsistent = false;
            else file2WasInconsistent = false;
        if (fileIndex === 1) {
            // Fully reset the file 2 UI and its file input.
            resetFileUI(2);
            fileInput2.value = null;
            document.getElementById('file2-name').textContent = 'No file chosen';
            fileInput2.disabled = true;
            file2Label.classList.add('btn-disabled');

            // Clear all matching and result states.
            clearAllMatches(); // This clears pairings and results
        clearComparisonPairs();
            matchingActionsContainer.classList.add('hidden');
            appliedFilters1 = []; // Clear applied filters
            appliedFilters2 = []; // Clear applied filters
            templateApplied = false; // Loading a new file invalidates template state
            keyValueMappings.clear();
            columnTypes1.clear();
            columnTypes2.clear();
            hasClearedMappings = false; // Reset flag on new file load
        } else {
            // If only file 2 is reset, we still need to clear pairings and lines
            // and also clear any existing key value mappings as they are now invalid.
            keyValueMappings.clear();
            hasClearedMappings = false;

            clearPairings();
        clearComparisonPairs();
            matchingActionsContainer.classList.add('hidden');
            // --- MEMORY MANAGEMENT ---
            // Explicitly clear data for file 2.
            fileData2 = null;
            fullFileData2 = null;
            file2WasInconsistent = false;
            columnTypes2.clear();
            appliedFilters2 = []; // Clear applied filters
        }
        resultsSection.classList.add('hidden');
        numericToleranceInput.value = '0'; // Reset numeric tolerance
    }

    function resetAll() {
        // Reset file 1 input and its UI elements
        fileInput1.value = null;
        document.getElementById('file1-name').textContent = 'No file chosen';
        document.getElementById(`preview-wrapper-1`).classList.add('hidden');
        document.getElementById(`delimiter-wrapper-1`).classList.add('hidden');
        document.getElementById(`filter-builder-1`).classList.add('hidden');
        document.querySelector(`.toggle-filter-btn[data-file-index="1"]`).style.display = 'none';

        // Reset file 2 input and its UI elements
        fileInput2.value = null;
        document.getElementById('file2-name').textContent = 'No file chosen';
        document.getElementById(`preview-wrapper-2`).classList.add('hidden');
        document.getElementById(`delimiter-wrapper-2`).classList.add('hidden');
        document.getElementById(`filter-builder-2`).classList.add('hidden');
        document.querySelector(`.toggle-filter-btn[data-file-index="2"]`).style.display = 'none';

        // This will clear all state variables, pairings, results, etc. by cascading from file 1.
        resetFileUI(1);

        isDragDisabled = false;
        fixedWidthSettings1 = null;
        fixedWidthSettings2 = null;
        file1WasInconsistent = false;
        file2WasInconsistent = false;
        file1HasNoHeader = false;
        file2HasNoHeader = false;
        // Hide the "Clear All" button since everything is reset.
        resetAllBtn.classList.add('hidden');
    groupByCheckbox.checked = false;
    }

    // --- File Handling and Parsing ---
    function parseAndRenderDelimitedFile(fileIndex) {
        const isFile1 = fileIndex === 1;

        // If this function was called as part of cancelling the "no header" modal,
        // we just want to revert the checkbox state, not re-parse and re-open the modal.
        if (isCancellingNoHeader) {
            isCancellingNoHeader = false; // Reset the flag
            return; // Stop execution
        }
        const fileInput = isFile1 ? fileInput1 : fileInput2;
        const delimiterSelect = isFile1 ? delimiter1 : delimiter2;
        const preview = isFile1 ? preview1 : preview2;
        const noHeaderCheckbox = document.getElementById(`no-header-checkbox-${fileIndex}`);
        const file = fileInput.files[0];

        let delimiter = delimiterSelect.value;
        if (delimiter === 'other') {
            // For 'other', we need to create and manage a custom input.
            // This part of the logic is handled by the 'change' event listener on the delimiter select.
            // Here, we just need to read from the dynamically created input.
            const customInput = document.getElementById(`custom-delimiter-${fileIndex}`);
            delimiter = customInput ? customInput.value : '';
        }

        // --- CRITICAL BUG FIX ---
        // Add a guard clause to prevent execution if fileIndex is invalid.
        // This can happen during certain UI reset sequences (e.g., closing the Excel modal).
        if (!fileIndex) {
            return;
        }


        if (!file || !delimiter) {
            return;
        }

        // --- CRITICAL FIX ---
        // Make the preview wrapper visible *before* showing the loader.
        // This ensures the loader has a rendered parent to attach to.
        document.getElementById(`preview-wrapper-${fileIndex}`).classList.remove('hidden');

        const reader = new FileReader();
        reader.onload = (e) => {
            // Defer the heavy parsing to allow the loader to render.
            showPreviewLoader(fileIndex, true, 'Parsing file...'); // Update message
            setTimeout(() => {
                try {
                    const text = e.target.result;
                    const hasNoHeader = noHeaderCheckbox.checked;

                    if (hasNoHeader) {
                        // --- BUG FIX ---
                        // If the user is switching to "no header" mode, all existing column
                        // pairings are now invalid and must be cleared from the state.
                        clearPairings();
                        clearComparisonPairs();

                        // Set state flag
                        if (fileIndex === 1) file1HasNoHeader = true;
                        else file2HasNoHeader = true; // If user specified no header, parse as raw data and open the definition modal
                        const rawParseResult = Papa.parse(text, { delimiter: delimiter, header: false, skipEmptyLines: true, newline: "" });
                        const rawData = rawParseResult.data.filter(row => row.some(cell => cell && cell.trim() !== ''));
                        inconsistentCsvData = { fileIndex, data: rawData };
                        showPreviewLoader(fileIndex, false); // Hide loader before opening modal
                        openHeaderDefinitionModal(true); // Pass flag to indicate no header
                        return;
                    } else {
                        showPreviewLoader(fileIndex, false);
                    }

                    const parseResult = Papa.parse(text, {
                        delimiter: delimiter,
                        header: true,
                        skipEmptyLines: 'greedy',
                        newline: "" // Auto-detect newline character
                    });

                    if (parseResult.errors.length > 0) { // If there are parsing errors
                        if (isTutorialActive) {
                            showToast("This file has an inconsistent structure. Please choose a different, well-formatted file for the tutorial.", "error", 10000);
                            const fileInput = fileIndex === 1 ? fileInput1 : fileInput2;
                            fileInput.value = '';
                            document.getElementById(`file${fileIndex}-name`).textContent = 'No file chosen';
                            resetFileUI(fileIndex);
                            showPreviewLoader(fileIndex, false); // Hide loader here for tutorial exit
                            return;
                        }
                        // This is the second heavy parse. The spinner must remain visible.
                        const rawParseResult = Papa.parse(text, { delimiter: delimiter, header: false, skipEmptyLines: true, newline: "" });
                        const rawData = rawParseResult.data.filter(row => row.some(cell => cell != null && cell.toString().trim() !== ''));
                        inconsistentCsvData = { fileIndex, data: rawData };
                        showPreviewLoader(fileIndex, false); // Hide loader before opening modal
                        openHeaderDefinitionModal(); // This will eventually hide the loader
                        return;
                    }

                    const headers = (parseResult.meta.fields || []).map(h => h.trim());
                    let allRows = parseResult.data.map(obj => headers.map(h => obj[h] ?? ""));
                    allRows = allRows.filter(row => row.some(cell => cell != null && cell.toString().trim() !== ''));
                    const fullData = { headers, rows: allRows };
                    if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(fullData)); else fullFileData2 = JSON.parse(JSON.stringify(fullData));
                    processAndRenderParsedData(fullData, fileIndex);
                } catch (error) {
                    console.error("Error during file parsing:", error);
                    showToast(`Error parsing file: ${error.message}`, 'error');
                    preview.innerHTML = '<p class="error-message">Could not parse file. Please check the delimiter or file content.</p>';
                    if (fileIndex === 1) {
                        file2Label.classList.add('btn-disabled');
                        fileInput2.disabled = true;
                    }
                } finally {
                    // The loader is now hidden in the functions that are called from here,
                    // ensuring it stays visible during all parsing steps.
                    if (!inconsistentCsvData) { // Only hide if not opening the modal
                        showPreviewLoader(fileIndex, false);
                    }
                }
            }, 10);
        };
        reader.onerror = () => {
            showPreviewLoader(fileIndex, false);
            showToast('Error reading file.', 'error');
        };

        // Use a nested setTimeout to give the browser a guaranteed render cycle for the spinner
        // before starting the potentially blocking file read operation.
        setTimeout(() => {
            showPreviewLoader(fileIndex, true, 'Reading file...');
            setTimeout(() => {
                // This is the call that can block the UI on very large files
                reader.readAsText(file);
            }, 0); // A delay of 0 is sufficient to push this to the next event loop cycle
        }, 0);
    }

    function openHeaderDefinitionModal(noHeader = false) {
        const { data } = inconsistentCsvData;
        const { fileIndex } = inconsistentCsvData;

        if (!data || data.length === 0) {
            showToast("The selected CSV file appears to be empty or unreadable.", "error");
            return;
        }

        // --- BUG FIX ---
        // Add a guard clause to ensure the modal's internal containers exist before proceeding.
        if (!headerPreview || !customHeadersContainer) {
            showToast("Cannot open header definition modal: UI elements are missing.", "error");
            return;
        }

        // Hide the loader for the preview now that this modal is about to show.
        // This is the correct place to ensure the loader is hidden as the modal appears.
        headerModal.classList.remove('hidden');

        headerPreview.innerHTML = '';
        customHeadersContainer.innerHTML = '';

        const modalSubHeader = headerModal.querySelector('.modal-sub-header');
        const option1Section = headerModal.querySelector('.header-option-section:first-of-type');
        const option2Header = headerModal.querySelector('.header-option-section:last-of-type h3');
        const pInconsistent = document.getElementById('header-definition-p-inconsistent');
        const pNoHeader = document.getElementById('header-definition-p-no-header');
        const modalContent = headerModal.querySelector('.modal-content');

        if (noHeader) {
            modalSubHeader.classList.add('hidden');
            option1Section.classList.add('hidden');
            option2Header.classList.add('hidden');
            pInconsistent.classList.add('hidden');
            pNoHeader.classList.remove('hidden');
            modalContent.classList.add('modal-content-auto');
        } else {
            modalSubHeader.classList.remove('hidden');
            option1Section.classList.remove('hidden');
            option2Header.classList.remove('hidden');
            pInconsistent.classList.remove('hidden');
            pNoHeader.classList.add('hidden');
            modalContent.classList.remove('modal-content-auto');
        }

        const table = document.createElement('table');
        const tbody = document.createElement('tbody');

        // Find the maximum number of columns in the preview data
        const maxCols = data.slice(0, 20).reduce((max, row) => Math.max(max, row.length), 0);

        // Generate custom header inputs
        for (let i = 0; i < maxCols; i++) {
            const wrapper = document.createElement('div');
            wrapper.className = 'custom-header-input-wrapper';
            const defaultHeaderName = `Column_${i + 1}`;
            wrapper.innerHTML = `<label>Col ${i + 1}</label><input type="text" class="custom-csv-header-input" data-col-index="${i}" placeholder="${defaultHeaderName}" value="${noHeader ? defaultHeaderName : ''}">`;
            customHeadersContainer.appendChild(wrapper);
        }

        if (noHeader) {
            customHeadersContainer.querySelector('input')?.focus();
        }
        // Populate preview table
        data.slice(0, 20).forEach((rowData, rowIndex) => {
            const tr = document.createElement('tr');
            const radioTd = document.createElement('td');
            radioTd.innerHTML = `<input type="radio" name="header-row-select" value="${rowIndex}" title="Use row ${rowIndex + 1} as header">`;
            tr.appendChild(radioTd);

            for (let i = 0; i < maxCols; i++) {
                const td = document.createElement('td');
                td.textContent = rowData[i] || '';
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        headerPreview.appendChild(table);
    }
    
    // Add event listeners inside the modal to handle the logic between Option 1 and Option 2
    customHeadersContainer.addEventListener('input', (e) => {
        if (e.target.classList.contains('custom-csv-header-input')) {
            // If the user types in any custom header input, unselect the radio button from Option 1.
            const selectedRadio = headerPreview.querySelector('input[name="header-row-select"]:checked');
            if (selectedRadio) {
                selectedRadio.checked = false;
            }
        }
    });
    headerPreview.addEventListener('click', (e) => {
        if (e.target.type === 'radio' && e.target.name === 'header-row-select') {
            // If user selects a radio button, clear all custom header inputs to avoid ambiguity.
            customHeadersContainer.querySelectorAll('.custom-csv-header-input').forEach(input => input.value = '');
        }
    });

    function applyHeaderDefinition() {
        let loader = headerModal.querySelector('.modal-loader');
        const content = headerModal.querySelector('.modal-content');
        // Get specific elements to hide
        const modalHeader = headerModal.querySelector('.modal-header');
        const modalSubHeader = headerModal.querySelector('.modal-sub-header');
        const headerOptions = headerModal.querySelector('.header-definition-options');
        const modalActions = headerModal.querySelector('.modal-actions');

        // --- BUG FIX ---
        // Add a guard clause to ensure the modal's body container exists.
        // If it doesn't, we cannot append the loader and must exit gracefully.
        if (!content) {
            showToast("Cannot apply headers: Modal structure is missing.", "error");
            return;
        }

        // --- CRITICAL FIX ---
        // If the loader element doesn't exist in the modal, create and inject it.
        // This ensures the spinner can be shown, resolving the original request.
        if (!loader) {
            loader = document.createElement('div');
            loader.className = 'modal-loader';
            loader.innerHTML = `<div class="spinner"></div><p>Applying Headers...</p>`;
            content.appendChild(loader); // Append it to the modal body
        }

        // Show loader and hide all other content within the modal
        loader.classList.remove('hidden');
        if (modalHeader) modalHeader.style.visibility = 'hidden';
        if (modalSubHeader) modalSubHeader.style.visibility = 'hidden';
        if (headerOptions) headerOptions.style.visibility = 'hidden';
        if (modalActions) modalActions.style.visibility = 'hidden';

        setTimeout(() => {
            try {
                const { fileIndex, data } = inconsistentCsvData;
                const selectedRadio = document.querySelector('input[name="header-row-select"]:checked');
                const noHeaderMode = headerModal.querySelector('.header-option-section:first-of-type').classList.contains('hidden');
                let headers = [];
                let rows = [];

                const customInputs = Array.from(document.querySelectorAll('.custom-csv-header-input'));
                const anyCustomInputHasValue = customInputs.some(input => input.value.trim() !== '');

                if (selectedRadio) {
                    const headerRowIndex = parseInt(selectedRadio.value, 10); // Default to Column_N if header is blank
                    headers = uniquifyHeaders(data[headerRowIndex].map((h, i) => (h || '').trim() || `Column_${i + 1}`));
                    rows = data.filter((_, index) => index !== headerRowIndex);
                } else if (anyCustomInputHasValue) {
                    const allCustomInputsAreFilled = customInputs.every(input => input.value.trim() !== '');
                    if (noHeaderMode && !allCustomInputsAreFilled) {
                        showToast("Please provide a name for all custom headers before applying.", "error");
                        return; // Return without hiding loader, user needs to fix input
                    }
                    const rawHeaders = customInputs.map(input => (input.value || '').trim() || `Column_${parseInt(input.dataset.colIndex, 10) + 1}`);
                    headers = uniquifyHeaders(rawHeaders);

                    rows = data;
                } else {
                    showToast("Please select a header row or define custom headers.", "error");
                    return; // Return without hiding loader
                }
                const formattedData = { headers, rows };

                if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(formattedData));
                else fullFileData2 = JSON.parse(JSON.stringify(formattedData));

                // Set a flag indicating this file originated from an inconsistent state
                if (fileIndex === 1) file1WasInconsistent = true;
                else file2WasInconsistent = true;

                // This is the crucial fix: Directly process the data instead of re-parsing.
                // The `parseAndRenderDelimitedFile` function is not needed here and was
                // causing an error because it expects a file input.
                processAndRenderParsedData(formattedData, fileIndex);

                // --- NEW LOGIC ---
                // If both files are loaded, we need to re-evaluate the unmatched headers
                // for the *other* file as well, since the headers for the current file have changed.
                if (fileData1 && fileData2) {
                    processAndRenderParsedData(fileIndex === 1 ? fileData2 : fileData1, fileIndex === 1 ? 2 : 1);
                }

                headerModal.classList.add('hidden');
            } catch (error) {
                showToast(`Error applying headers: ${error.message}`, 'error');
            } finally {
                // Ensure loader is hidden and content is visible if an error occurs or validation fails
                loader.classList.add('hidden');
                if (modalHeader) modalHeader.style.visibility = 'visible'; // Restore visibility
                if (modalSubHeader) modalSubHeader.style.visibility = 'visible';
                if (headerOptions) headerOptions.style.visibility = 'visible';
                if (modalActions) modalActions.style.visibility = 'visible';
            }
        }, 10);
    }

    function processAndRenderParsedData(formattedData, fileIndex) {
        const isFile1 = fileIndex === 1;
        const preview = isFile1 ? preview1 : preview2;
        const previewWrapper = document.getElementById(`preview-wrapper-${fileIndex}`);
        const toggleFilterBtn = document.querySelector(`.toggle-filter-btn[data-file-index="${fileIndex}"]`);

        detectAndSetColumnTypes(formattedData, fileIndex);
        const { headers, rows } = formattedData;

        if (headers.length === 0) {
            throw new Error('No headers found or defined. Check file or parsing configuration.');
        }

        if (isFile1) {
            fileData1 = formattedData;
            // fullFileData1 is now set in the parsing function
            initFilterBuilder(1, headers);
            renderPreviewTable(fileData1, preview, 1);
            toggleFilterBtn.style.display = 'inline-flex';
            previewWrapper.classList.remove('hidden');
            file2Label.classList.remove('btn-disabled');
            fileInput2.disabled = false; // Enable file 2 input
        } else if (fileData1) { // Ensure file 1 is loaded before processing file 2
            // --- Schema Difference Detection (No Reordering) ---
            const headers1 = fileData1.headers;
            const originalHeaders2 = formattedData.headers;
            const headerSet1 = new Set(headers1);
            const headerSet2 = new Set(originalHeaders2);

            // The file data is used as-is, without reordering columns.
            fileData2 = formattedData;

            // Identify which headers are unique to each file.
            const uniqueHeaders1 = headers1.filter(h => !headerSet2.has(h));
            const uniqueHeaders2 = originalHeaders2.filter(h => !headerSet1.has(h));
            unmatchedHeaders1 = new Set(uniqueHeaders1);
            unmatchedHeaders2 = new Set(uniqueHeaders2);

            updateUnmatchedStyles(1, unmatchedHeaders1);

            initFilterBuilder(2, fileData2.headers);
            toggleFilterBtn.style.display = 'inline-flex';

            renderPreviewTable(fileData2, preview2, 2, uniqueHeaders2);
            previewWrapper.classList.remove('hidden');
            matchingActionsContainer.classList.remove('hidden');
        }

        // Show the "no header" checkbox now that the preview is loaded for the current file
        const delimiterWrapper = document.getElementById(`delimiter-wrapper-${fileIndex}`);
        const noHeaderControl = delimiterWrapper.querySelector('.no-header-control');
        if (noHeaderControl) noHeaderControl.style.display = 'flex';

        updateUIState();
    }

    function detectAndSetColumnTypes(data, fileIndex) {
        const typeMap = fileIndex === 1 ? columnTypes1 : columnTypes2;
        typeMap.clear(); // Reset before detection

        if (!data || !data.headers || !data.rows) return;

        data.headers.forEach((header, index) => {
            // Default to 'auto' which will behave like text unless specified otherwise
            typeMap.set(header, 'auto');

            // Use a sample of rows for performance
            const sampleRows = data.rows.slice(0, 100);
            if (sampleRows.length === 0) return;

            let isPotentiallyNumeric = true;
            let isPotentiallyDate = true;
            let hasValues = false;

            for (const row of sampleRows) {
                const value = (row[index] ?? '').toString().trim();
                if (value === '') continue; // Ignore empty cells for detection

                hasValues = true;

                // Check for numeric
                if (isPotentiallyNumeric && isNaN(Number(value))) {
                    isPotentiallyNumeric = false;
                }

                // Check for date
                if (isPotentiallyDate) {
                    const date = new Date(value);
                    // A simple check: if it's a valid date and not just a number being interpreted as a date
                    if (isNaN(date.getTime()) || /^\d+$/.test(value)) {
                        isPotentiallyDate = false;
                    }
                }

                if (!isPotentiallyNumeric && !isPotentiallyDate) break; // No need to check further
            }

            if (hasValues) {
                if (isPotentiallyNumeric) typeMap.set(header, 'number');
                else if (isPotentiallyDate) typeMap.set(header, 'date');
                // else it remains 'auto' (effectively text)
            }
        });
    }

    function reEvaluateUnmatchedHeaders(templateSettings = null) {
        if (!fileData1 || !fileData2) return;

        const headers1 = new Set(fileData1.headers);
        const headers2 = new Set(fileData2.headers);

        const templateHeaders1 = new Set(templateSettings?.file1 ? Object.keys(templateSettings.file1) : []);
        const templateHeaders2 = new Set(templateSettings?.file2 ? Object.keys(templateSettings.file2) : []);

        // Find headers that are unique to file 1 AND were not mentioned in the template's settings for file 1.
        // This prevents the function from overriding a template's decision to "include" a unique column.
        const uniqueInFile1 = fileData1.headers.filter(h => !headers2.has(h) && !templateHeaders1.has(h));
        uniqueInFile1.forEach(h => unmatchedHeaders1.add(h));

        // Find headers that are unique to file 2 AND were not mentioned in the template's settings for file 2.
        const uniqueInFile2 = fileData2.headers.filter(h => !headers1.has(h) && !templateHeaders2.has(h));
        uniqueInFile2.forEach(h => unmatchedHeaders2.add(h));

        updateUnmatchedStyles(1, unmatchedHeaders1);
        updateUnmatchedStyles(2, unmatchedHeaders2);
    }

    // --- Excel File Handling ---
    function openExcelOptionsModal(fileIndex, file) {
        activeFileForModal = { fileIndex, file };
        const loader = excelOptionsModal.querySelector('.modal-loader');
        const form = document.getElementById('excel-options-form');

        excelOptionsModal.classList.remove('hidden');
        loader.classList.remove('hidden');
        form.classList.add('hidden');

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                activeFileForModal.workbook = workbook; // Store workbook for later use

                worksheetSelect.innerHTML = '';
                workbook.SheetNames.forEach(name => {
                    const option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;
                    worksheetSelect.appendChild(option);
                });

                // Auto-detect range for the first sheet
                handleWorksheetChange();

                loader.classList.add('hidden');
                form.classList.remove('hidden');
            } catch (error) {
                console.error("Error reading Excel file:", error);
                alert("Could not read the Excel file. It might be corrupted or in an unsupported format.");
                excelOptionsModal.classList.add('hidden');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleWorksheetChange() {
        const sheetName = worksheetSelect.value;
        const workbook = activeFileForModal.workbook;
        if (!sheetName || !workbook) return;

        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet || !worksheet['!ref']) {
            cellRangeInput.value = '';
            return;
        }

        // Find the first continuous block of data
        let firstRow = -1, lastRow = -1, firstCol = -1, lastCol = -1;
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        for (let R = range.s.r; R <= range.e.r; ++R) {
            let rowHasData = false;
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell_address = { c: C, r: R };
                const cell_ref = XLSX.utils.encode_cell(cell_address);
                if (worksheet[cell_ref]) { rowHasData = true; break; }
            }
            if (rowHasData) {
                if (firstRow === -1) firstRow = R; // Found the first row with data
                lastRow = R;
            } else {
                if (firstRow !== -1) break; // Stop after the first block of data
            }
        }

        if (firstRow !== -1) {
            // Now that we have the row bounds, find the column bounds within those rows.
            firstCol = range.e.c; // Start high
            lastCol = range.s.c;  // Start low

            for (let R = firstRow; R <= lastRow; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
                    if (worksheet[cell_ref]) {
                        if (C < firstCol) firstCol = C;
                        if (C > lastCol) lastCol = C;
                    }
                }
            }
            const detectedRange = { s: { r: firstRow, c: firstCol }, e: { r: lastRow, c: lastCol } };
            cellRangeInput.value = XLSX.utils.encode_range(detectedRange);
        } else {
            cellRangeInput.value = worksheet['!ref'] || ''; // Fallback to full sheet range
        }
    }

    function applyExcelOptions() {
        const { fileIndex, workbook } = activeFileForModal;
        const sheetName = worksheetSelect.value;
        const range = cellRangeInput.value;

        if (!sheetName) {
            alert("Please select a worksheet.");
            return;
        }

        excelOptionsModal.classList.add('hidden');

        const worksheet = workbook.Sheets[sheetName];
        // Parse to array of arrays, `raw: false` ensures we get formatted text.
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: range || undefined, raw: false, defval: null });
        
        // --- BUG FIX ---
        // Add a guard clause. If the sheet is completely empty, `data` can be undefined.
        if (!data) {
            showToast("The selected worksheet or range appears to be completely empty.", "warning");
            resetFileUI(fileIndex);
            return;
        }
        const nonEmptyRows = data.filter(row => row.some(cell => cell !== null && cell !== '' && cell.trim() !== ''));

        if (nonEmptyRows.length === 0) {
            showToast("The selected range in the Excel file appears to be empty.", "warning");
            resetFileUI(fileIndex);
            return;
        }

        // --- Improved Inconsistency Heuristic for Excel ---
        // With `defval: null`, `data` is a rectangular grid. We can now use a more robust check.
        // 1. Find the actual data width by checking the last non-null cell in a sample of rows.
        const sampleRows = nonEmptyRows.slice(0, 100);
        let maxColumnIndex = 0;
        sampleRows.forEach(row => {
            for (let i = row.length - 1; i >= 0; i--) {
                if (row[i] !== null) {
                    if (i > maxColumnIndex) maxColumnIndex = i;
                    break; // Move to the next row
                }
            }
        });
        const dataWidth = maxColumnIndex + 1;

        // 2. Check if the first row looks like a title (e.g., one cell) while the rest of the data is wider.
        let isLikelyInconsistent = false;
        if (nonEmptyRows.length > 1) {
            const firstRowCellCount = nonEmptyRows[0].slice(0, dataWidth).filter(c => c !== null).length;
            if (firstRowCellCount < dataWidth / 2 && dataWidth > 1) {
                isLikelyInconsistent = true;
            }
        }

        if (isLikelyInconsistent && isTutorialActive) {
            showToast("This file has an inconsistent structure. Please choose a different, well-formatted file for the tutorial.", "error", 10000);
            resetFileUI(fileIndex);
            return;
        }

        if (isLikelyInconsistent) {
            inconsistentCsvData = { fileIndex, data: nonEmptyRows };
            openHeaderDefinitionModal();
        } else {
            // Proceed with the standard assumption: first row is the header.
            const headers = nonEmptyRows.length > 0 ? nonEmptyRows[0].map(h => String(h || '').trim()) : [];
            const rows = nonEmptyRows.length > 0 ? nonEmptyRows.slice(1) : [];
            const fullData = { headers, rows };
            if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(fullData)); else fullFileData2 = JSON.parse(JSON.stringify(fullData));
            processAndRenderParsedData(fullData, fileIndex);
        }
    }

    /**
     * Parses a JSON file, assuming it's an array of objects.
     * @param {File} file The JSON file object.
     * @param {number} fileIndex The index of the file (1 or 2).
     */
    function parseJsonFile(file, fileIndex) {
        document.getElementById(`preview-wrapper-${fileIndex}`).classList.remove('hidden');
        showPreviewLoader(fileIndex, true, 'Parsing JSON...');

        const reader = new FileReader();
        reader.onload = (e) => {
            setTimeout(() => {
                try {
                    let jsonData = JSON.parse(e.target.result);

                    // Handle common wrapper object case (e.g., { "data": [...] })
                    if (typeof jsonData === 'object' && !Array.isArray(jsonData)) {
                        const keys = Object.keys(jsonData);
                        const arrayProperty = keys.find(k => Array.isArray(jsonData[k]));
                        if (arrayProperty) {
                            jsonData = jsonData[arrayProperty];
                        } else {
                            throw new Error("JSON is not an array of objects and no array property was found.");
                        }
                    }

                    if (!Array.isArray(jsonData) || jsonData.length === 0) {
                        throw new Error("JSON file must be a non-empty array of objects.");
                    }

                    const headers = Object.keys(jsonData[0]);
                    const rows = jsonData.map(obj => headers.map(h => obj[h] ?? ""));

                    const fullData = { headers, rows };
                    if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(fullData));
                    else fullFileData2 = JSON.parse(JSON.stringify(fullData));

                    processAndRenderParsedData(fullData, fileIndex);

                } catch (error) {
                    console.error("Error parsing JSON file:", error);
                    showToast(`Error parsing JSON: ${error.message}`, 'error');
                    resetFileUI(fileIndex);
                } finally {
                    showPreviewLoader(fileIndex, false);
                }
            }, 10);
        };
        reader.onerror = () => {
            showPreviewLoader(fileIndex, false);
            showToast('Error reading file.', 'error');
        };
        reader.readAsText(file);
    }

    /**
     * Parses an XML file, assuming a simple table-like structure.
     * @param {File} file The XML file object.
     * @param {number} fileIndex The index of the file (1 or 2).
     */
    function parseXmlFile(file, fileIndex) {
        document.getElementById(`preview-wrapper-${fileIndex}`).classList.remove('hidden');
        showPreviewLoader(fileIndex, true, 'Parsing XML...');

        const reader = new FileReader();
        reader.onload = (e) => {
            setTimeout(() => {
                try {
                    const parser = new DOMParser();
                    const xmlDoc = parser.parseFromString(e.target.result, "application/xml");

                    if (xmlDoc.getElementsByTagName("parsererror").length) {
                        throw new Error("XML file contains parsing errors.");
                    }

                    const root = xmlDoc.documentElement;
                    if (!root || !root.children.length) {
                        throw new Error("XML file is empty or has no child elements under the root.");
                    }

                    // Assume the direct children of the root are the "rows"
                    const rowNodes = Array.from(root.children);
                    const firstRow = rowNodes[0];
                    if (!firstRow || !firstRow.children.length) {
                        throw new Error("First record in XML has no child elements to use as headers.");
                    }

                    const headers = Array.from(firstRow.children).map(child => child.tagName);
                    const rows = rowNodes.map(rowNode => {
                        return headers.map(header => rowNode.querySelector(header)?.textContent ?? "");
                    });

                    const fullData = { headers, rows };
                    if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(fullData));
                    else fullFileData2 = JSON.parse(JSON.stringify(fullData));

                    processAndRenderParsedData(fullData, fileIndex);

                } catch (error) {
                    console.error("Error parsing XML file:", error);
                    showToast(`Error parsing XML: ${error.message}`, 'error');
                    resetFileUI(fileIndex);
                } finally {
                    showPreviewLoader(fileIndex, false);
                }
            }, 10);
        };
        reader.onerror = () => {
            showPreviewLoader(fileIndex, false);
            showToast('Error reading file.', 'error');
        };
        reader.readAsText(file);
    }
    /**
     * Replaces tab characters in a string with a fixed number of spaces.
     * This is crucial for aligning visual representation with parsing logic.
     * @param {string} text The input string.
     * @param {number} tabWidth The number of spaces to replace a tab with.
     * @returns {string} The string with tabs replaced by spaces.
     */
    function normalizeTabs(text, tabWidth = 8) {
        // Correctly handle tab stops. A tab advances to the next multiple of tabWidth.
        const lines = text.split(/\r?\n/);
        const processedLines = lines.map(line => {
            let processedLine = '';
            let currentPos = 0;
            const parts = line.split('\t');

            parts.forEach((part, index) => {
                processedLine += part;
                currentPos += part.length;

                if (index < parts.length - 1) { // If it's not the last part, there was a tab after it
                    const spacesToAdd = tabWidth - (currentPos % tabWidth) || tabWidth;
                    processedLine += ' '.repeat(spacesToAdd);
                    currentPos += spacesToAdd;
                }
            });
            return processedLine;
        });
        return processedLines.join('\n');
    }
    function parseFixedWidthFile(fileContent, breaks, useFirstRowAsHeader, customHeaders) {
        const lines = normalizeTabs(fileContent).split(/\r?\n/);
        
        // Corrected Logic: The 'breaks' define the end of each column.
        // We create a 'segments' array that represents the start and end of each column.
        // It starts implicitly at 0.
        const sortedBreaks = [...new Set(breaks)].sort((a, b) => a - b);
        const segments = [0, ...sortedBreaks];
    
        let headers = [];
        const dataRows = [];
    
        const firstDataLine = lines.find(line => line.trim() !== '');
        if (!firstDataLine) return { headers: [], rows: [] }; // Empty file
    
        if (useFirstRowAsHeader) {
            // Create N+1 headers for N breaks.
            for (let i = 0; i < segments.length; i++) {
                const start = segments[i];
                // For the last segment, substring goes to the end of the line.
                const end = (i === segments.length - 1) ? undefined : segments[i + 1];
                headers.push(firstDataLine.substring(start, end).trim() || `Column_${i + 1}`);
            }
        } else {
            headers = customHeaders;
        }
    
        const dataLinesToParse = useFirstRowAsHeader ? lines.slice(1) : lines;
    
        dataLinesToParse.forEach(line => {
            if (line.trim() === '') return;
    
            const row = [];
            // Create N+1 columns for N breaks.
            for (let i = 0; i < segments.length; i++) {
                const start = segments[i];
                const end = (i === segments.length - 1) ? undefined : segments[i + 1];
                row.push(line.substring(start, end).trim());
            }
            dataRows.push(row);
        });
    
        return { headers, rows: dataRows };
    }
    
    // Add event listeners to delimiter dropdowns to handle the 'other' option
    [delimiter1, delimiter2].forEach((delimiterSelect, index) => {
        const fileIndex = index + 1;
        delimiterSelect.addEventListener('change', () => {
            // CRITICAL FIX: When the delimiter changes, all existing pairings for that file
            // (and any subsequent files) become invalid. We must clear them before re-parsing.
            if (fileIndex === 1) {
                clearAllMatches();
            } else {
                clearPairings();
            }

            const fileInput = document.getElementById(`file${fileIndex}`);
            let customInput = document.getElementById(`custom-delimiter-${fileIndex}`);
            const selectedOption = delimiterSelect.value;

            if (selectedOption === 'fixed') {
                if (customInput) customInput.remove();
                if (fileInput.files.length > 0) {
                    openFixedWidthEditor(fileIndex, fileInput.files[0]);
                } else {
                    alert(`Please select File ${fileIndex} before choosing Fixed-width.`);
                    delimiterSelect.value = ''; // Reset dropdown
                }
            } else if (selectedOption === 'other') {
                if (!customInput) {
                    customInput = document.createElement('input');
                    customInput.type = 'text';
                    customInput.id = `custom-delimiter-${fileIndex}`;
                    customInput.placeholder = '...';
                    customInput.maxLength = 1;
                    customInput.style.width = '40px';
                    customInput.style.textAlign = 'center';
                    customInput.addEventListener('input', () => parseAndRenderDelimitedFile(fileIndex));
                    delimiterSelect.insertAdjacentElement('afterend', customInput);
                }
                customInput.focus();
            } else {
                // This block now handles standard delimiters (comma, tab, etc.)
                if (customInput) customInput.remove();
                parseAndRenderDelimitedFile(fileIndex); // Re-parse with standard delimiter
            }
        });
    });

    function reorderFileData(data, newHeaderOrder) {
        const oldHeaderIndexMap = new Map(data.headers.map((h, i) => [h, i]));
        const newRows = data.rows.map(oldRow => {
            const newRow = [];
            newHeaderOrder.forEach(header => {
                const oldIndex = oldHeaderIndexMap.get(header);
                newRow.push(oldRow[oldIndex]);
            });
            return newRow;
        });
        return { headers: newHeaderOrder, rows: newRows };
    }

    function toggleFilterBuilder(fileIndex) {
        const builder = document.getElementById(`filter-builder-${fileIndex}`);
        builder.classList.toggle('hidden');
        redrawLines();
    }

    // --- Fixed-width Modal Logic ---
    function openFixedWidthEditor(fileIndex, file) {
        activeFileForModal = { fileIndex, file };
        breakPositions = []; // Reset breaks

        // Clear previous header options to ensure a clean state for the new file.
        document.getElementById('header-options').innerHTML = '';

        const reader = new FileReader();
        reader.onload = (e) => {
            const content = e.target.result;
            const first20Lines = content.split(/\r?\n/).slice(0, 20).join('\n');
            fixedWidthPreview.textContent = normalizeTabs(first20Lines);
            modal.classList.remove('hidden');            
            renderBreaks();
            renderHeaderOptions();
        };
        reader.readAsText(file);
    }

    function renderBreaks() {
        const firstLine = fixedWidthPreview.textContent.split('\n')[0];
        if (!firstLine) {
            breakLinesOverlay.innerHTML = '';
            return;
        }
        const charBoundaries = getCharacterBoundaries(firstLine);

        breakLinesOverlay.innerHTML = '';
        breakPositions.forEach(pos => {
            if (pos >= charBoundaries.length) return; // Safety check
            const lineEl = document.createElement('div');
            lineEl.className = 'break-line';
            lineEl.style.left = `${charBoundaries[pos]}px`;
            lineEl.dataset.pos = pos;
            breakLinesOverlay.appendChild(lineEl);
        });
    }

    function getCharacterBoundaries(lineText) {
        const boundaries = [0]; // The first boundary is always at pixel 0
        const span = document.createElement('span');
        span.style.font = window.getComputedStyle(fixedWidthPreview).font;
        span.style.whiteSpace = 'pre';
        span.style.visibility = 'hidden';
        span.style.position = 'absolute';
        document.body.appendChild(span);
    
        for (let i = 1; i <= lineText.length; i++) {
            span.textContent = lineText.substring(0, i);
            boundaries.push(span.offsetWidth);
        }
    
        document.body.removeChild(span);
        return boundaries;
    }

    breakLinesOverlay.addEventListener('click', (e) => {
        if (e.target !== breakLinesOverlay) return; // Ignore clicks on lines themselves
        const firstLine = fixedWidthPreview.textContent.split('\n')[0];
        if (!firstLine) return;

        const charBoundaries = getCharacterBoundaries(firstLine);
        const rect = fixedWidthPreview.getBoundingClientRect();
        const previewStyle = window.getComputedStyle(fixedWidthPreview);
        const paddingLeft = parseFloat(previewStyle.paddingLeft);
        const clickX = e.clientX - rect.left - paddingLeft;

        // Find the character cell the click falls into.
        // We iterate through the boundaries to find the first one that is to the right of the click.
        // The index of that boundary is the character index for the break.
        let charIndex = 0;
        for (let i = 0; i < charBoundaries.length; i++) {
            if (clickX < charBoundaries[i]) {
                charIndex = i;
                break;
            }
            // If the click is past the last boundary, it belongs to the last character cell
            charIndex = i;
        }

        if (!breakPositions.includes(charIndex)) {
            breakPositions.push(charIndex);
            breakPositions.sort((a, b) => a - b);
            renderBreaks();
            renderHeaderOptions();
        }
    });

    // --- Fixed-width Line Dragging Logic ---
    let draggedLine = null;
    let originalBreakPos = -1;

    // --- Mouse Dragging Logic ---
    breakLinesOverlay.addEventListener('mousedown', (e) => {
        if (e.target.classList.contains('break-line')) {
            e.preventDefault();
            draggedLine = e.target;
            originalBreakPos = parseInt(draggedLine.dataset.pos, 10);
            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        }
    });

    function onMouseMove(e) {
        if (!draggedLine) return;
        const firstLine = fixedWidthPreview.textContent.split('\n')[0];
        if (!firstLine) return;
        const charBoundaries = getCharacterBoundaries(firstLine);

        const overlayRect = breakLinesOverlay.getBoundingClientRect();
        const previewStyle = window.getComputedStyle(fixedWidthPreview);
        const paddingLeft = parseFloat(previewStyle.paddingLeft);
        const mouseX = e.clientX - overlayRect.left - paddingLeft;

        // Find the character cell the mouse is over.
        let charIndex = 0;
        for (let i = 0; i < charBoundaries.length; i++) {
            if (mouseX < charBoundaries[i]) {
                charIndex = i;
                break;
            }
            charIndex = i;
        }

        // Update the visual position of the line being dragged
        if (charIndex >= charBoundaries.length) charIndex = charBoundaries.length - 1; // Clamp to max
        draggedLine.style.left = `${charBoundaries[charIndex]}px`;
    }

    function onMouseUp(e) {
        if (!draggedLine) return;
        const firstLine = fixedWidthPreview.textContent.split('\n')[0];
        if (firstLine) {
            const charBoundaries = getCharacterBoundaries(firstLine);
            const rect = fixedWidthPreview.getBoundingClientRect();
            const previewStyle = window.getComputedStyle(fixedWidthPreview);
            const paddingLeft = parseFloat(previewStyle.paddingLeft);
            const mouseX = e.clientX - rect.left - paddingLeft;

            // Find the character cell the mouse was released over.
            let newCharIndex = 0;
            for (let i = 0; i < charBoundaries.length; i++) {
                if (mouseX < charBoundaries[i]) {
                    newCharIndex = i;
                    break;
                }
                newCharIndex = i;
            }

            // Update the state: remove old, add new
            if (newCharIndex >= charBoundaries.length) newCharIndex = charBoundaries.length - 1; // Clamp to max
            breakPositions = breakPositions.filter(p => p !== originalBreakPos);
            if (!breakPositions.includes(newCharIndex)) {
                breakPositions.push(newCharIndex);
            }
            breakPositions.sort((a, b) => a - b);
            renderBreaks(); // Re-render all lines from the updated state
            renderHeaderOptions();
        }

        // Cleanup
        draggedLine = null;
        originalBreakPos = -1;
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
    }

    breakLinesOverlay.addEventListener('contextmenu', (e) => {
        e.preventDefault(); // Prevent the default right-click menu
        if (e.target.classList.contains('break-line')) {
            const posToRemove = parseInt(e.target.dataset.pos, 10);
            breakPositions = breakPositions.filter(p => p !== posToRemove);
            renderBreaks();
            renderHeaderOptions();
        }
    });

    // --- Touch Event Logic for Mobile ---
    let touchTimer = null;
    let touchStartX = 0;
    let touchStartY = 0;
    const LONG_PRESS_DURATION = 500; // 500ms for a long press

    breakLinesOverlay.addEventListener('touchstart', (e) => {
        const touch = e.touches[0];
        touchStartX = touch.clientX;
        touchStartY = touch.clientY;

        // Check if the touch is on a break line for potential removal
        if (e.target.classList.contains('break-line')) {
            e.preventDefault(); // Prevent scrolling while holding a line
            // Prepare for a potential drag, same as mousedown
            draggedLine = e.target;
            originalBreakPos = parseInt(draggedLine.dataset.pos, 10);

            // Start a timer for long press (remove)
            touchTimer = setTimeout(() => {
                const posToRemove = parseInt(e.target.dataset.pos, 10);
                breakPositions = breakPositions.filter(p => p !== posToRemove);
                renderBreaks();
                renderHeaderOptions();
                showToast('Break line removed.', 'success', 2000);
                touchTimer = null; // Reset timer
                draggedLine = null; // Cancel drag if it was a long press
                originalBreakPos = -1;
            }, LONG_PRESS_DURATION);
        }
    }, { passive: false });

    breakLinesOverlay.addEventListener('touchmove', (e) => {
        // If the finger moves, it's not a long press, so cancel the timer.
        if (touchTimer) {
            const touch = e.touches[0];
            // Check for significant movement to differentiate from accidental jitter
            if (Math.abs(touch.clientX - touchStartX) > 10 || Math.abs(touch.clientY - touchStartY) > 10) {
                clearTimeout(touchTimer);
                touchTimer = null;
            }
        }

        // --- Touch Dragging Logic ---
        if (draggedLine) {
            e.preventDefault(); // Prevent page scroll while dragging a line
            const firstLine = fixedWidthPreview.textContent.split('\n')[0];
            if (!firstLine) return;
            const charBoundaries = getCharacterBoundaries(firstLine);

            const overlayRect = breakLinesOverlay.getBoundingClientRect();
            const previewStyle = window.getComputedStyle(fixedWidthPreview);
            const paddingLeft = parseFloat(previewStyle.paddingLeft);
            const touchX = e.touches[0].clientX - overlayRect.left - paddingLeft;

            // Find the character cell the finger is over.
            let charIndex = 0;
            for (let i = 0; i < charBoundaries.length; i++) {
                if (touchX < charBoundaries[i]) {
                    charIndex = i;
                    break;
                }
                charIndex = i;
            }

            // Update the visual position of the line being dragged
            if (charIndex >= charBoundaries.length) charIndex = charBoundaries.length - 1; // Clamp to max
            draggedLine.style.left = `${charBoundaries[charIndex]}px`;
        }
    }, { passive: false });

    breakLinesOverlay.addEventListener('touchend', (e) => {
        // If the timer is still active when the touch ends, it was a short tap, not a long press.
        if (touchTimer) {
            clearTimeout(touchTimer);
            touchTimer = null;

            // If it was a short tap on a line, we don't want to do anything else.
            // So we reset draggedLine and exit.
            if (draggedLine) {
                draggedLine = null;
                originalBreakPos = -1;
                renderBreaks(); // Snap the line back to its original position
                return;
            }
        }

        // --- Finalize Touch Drag ---
        if (draggedLine) {
            const firstLine = fixedWidthPreview.textContent.split('\n')[0];
            if (firstLine) {
                const charBoundaries = getCharacterBoundaries(firstLine);
                const rect = fixedWidthPreview.getBoundingClientRect();
                const previewStyle = window.getComputedStyle(fixedWidthPreview);
                const paddingLeft = parseFloat(previewStyle.paddingLeft);
                const touchX = e.changedTouches[0].clientX - rect.left - paddingLeft;

                // Find the character cell the finger was released over.
                let newCharIndex = 0;
                for (let i = 0; i < charBoundaries.length; i++) {
                    if (touchX < charBoundaries[i]) {
                        newCharIndex = i;
                        break;
                    }
                    newCharIndex = i;
                }

                // Update the state: remove old, add new
                if (newCharIndex >= charBoundaries.length) newCharIndex = charBoundaries.length - 1; // Clamp to max
                breakPositions = breakPositions.filter(p => p !== originalBreakPos);
                if (!breakPositions.includes(newCharIndex)) {
                    breakPositions.push(newCharIndex);
                }
                breakPositions.sort((a, b) => a - b);
                renderBreaks(); // Re-render all lines from the updated state
                renderHeaderOptions();
            }

            // Cleanup
            draggedLine = null;
            originalBreakPos = -1;
            return; // End here to prevent tap-to-add logic from firing
        }

        // --- Tap to Add Logic ---
        // Only add a break if the tap was on the overlay itself, not on an existing line.
        if (e.target === breakLinesOverlay) {
            const firstLine = fixedWidthPreview.textContent.split('\n')[0];
            if (!firstLine) return;

            const charBoundaries = getCharacterBoundaries(firstLine);
            const rect = fixedWidthPreview.getBoundingClientRect();
            const previewStyle = window.getComputedStyle(fixedWidthPreview);
            const paddingLeft = parseFloat(previewStyle.paddingLeft);
            const touchX = e.changedTouches[0].clientX - rect.left - paddingLeft;

            let charIndex = 0;
            for (let i = 0; i < charBoundaries.length; i++) {
                if (touchX < charBoundaries[i]) { charIndex = i; break; }
                charIndex = i;
            }

            if (!breakPositions.includes(charIndex)) {
                breakPositions.push(charIndex);
                breakPositions.sort((a, b) => a - b);
                renderBreaks();
                renderHeaderOptions();
            }
        }
    });

    function renderHeaderOptions() {
        // Check if the checkbox element already exists in the DOM.
        const checkboxElement = document.getElementById('use-first-row-header-checkbox');

        // Determine the state of the checkbox.
        // If the checkbox already exists, respect its current state (the user is toggling it).
        // If it doesn't exist, this is the initial render for the modal, so default to checked.
        const useHeader = checkboxElement ? checkboxElement.checked : true;

        const container = document.getElementById('header-options');

        container.innerHTML = `
            <div>
                <input type="checkbox" id="use-first-row-header-checkbox" ${useHeader ? 'checked' : ''}>
                <label for="use-first-row-header-checkbox">First row contains headers</label>
            </div>
            <div id="custom-headers-container" class="${useHeader ? 'hidden' : ''}"></div>
        `;

        container.querySelector('#use-first-row-header-checkbox').addEventListener('change', (e) => {
            document.getElementById('custom-headers-container').classList.toggle('hidden', e.target.checked);
            renderHeaderOptions(); // Re-render to populate inputs if needed
        });

        if (!useHeader) {
            const customHeadersContainer = document.getElementById('custom-headers-container');
            // Determine the maximum line length from the preview content for accurate width calculation.
            const previewContent = fixedWidthPreview.textContent;
            const maxLineLength = previewContent.split('\n').reduce((max, line) => Math.max(max, line.length), 0);

            // Use the max line length as the "end" for the last column's width calculation.
            // The 10000 is still used for parsing but not for UI rendering.
            const allBreaksForUI = [0, ...breakPositions.sort((a, b) => a - b), maxLineLength];

            for (let i = 0; i < allBreaksForUI.length - 1; i++) {
                if (i >= 20) break; // Limit to a reasonable number of columns
                const wrapper = document.createElement('div');
                wrapper.className = 'custom-header-input-wrapper';
                // Calculate width based on the UI-specific breaks array.
                wrapper.innerHTML = `<label>Col ${i + 1}</label><input type="text" class="custom-header-input" data-col-index="${i}" style="width: ${((allBreaksForUI[i+1] - allBreaksForUI[i])*0.9)}ch">`;
                customHeadersContainer.appendChild(wrapper);
            }
        }
    }

    function applyFixedWidthParsing() {
        const { fileIndex, file } = activeFileForModal;
        const useFirstRowAsHeader = document.getElementById('use-first-row-header-checkbox').checked;
        let customHeaders = [];

        if (!useFirstRowAsHeader) {
            const rawHeaders = Array.from(document.querySelectorAll('.custom-header-input')).map(input => input.value || `Column_${parseInt(input.dataset.colIndex, 10) + 1}`);
            customHeaders = uniquifyHeaders(rawHeaders);
            if (customHeaders.length === 0 && breakPositions.length > 0) {
                alert("Please define custom headers or check 'First row contains headers'.");
                return;
            }
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const content = e.target.result;
            try {
                // When applying fixed-width parsing, store the settings
                if (fileIndex === 1) {
                    fixedWidthSettings1 = { breaks: breakPositions, useFirstRowAsHeader, customHeaders };
                } else {
                    fixedWidthSettings2 = { breaks: breakPositions, useFirstRowAsHeader, customHeaders };
                }
                const parsedData = parseFixedWidthFile(content, breakPositions, useFirstRowAsHeader, customHeaders);

                // This is the crucial fix: Store the complete parsed data in the 'fullFileData' variables.
                // The value mapping modal relies on this unfiltered data.
                if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(parsedData));
                else fullFileData2 = JSON.parse(JSON.stringify(parsedData));

                processAndRenderParsedData(parsedData, fileIndex);
                modal.classList.add('hidden');
            } catch (error) {
                alert(`Error parsing fixed-width file: ${error.message}`);
            }
        };
        reader.readAsText(file);
    }

    /**
     * Analyzes the fixed-width file preview to suggest column breaks.
     * This advanced version analyzes multiple patterns with weighted scores:
     * 1. Data-to-Space transitions (alignment gaps).
     * 2. Space-to-Data transitions (start of aligned data).
     * 3. Transitions in data type (e.g., text to number).
     * 4. Word breaks in the header row (e.g., space or CamelCase).
     */
    function suggestFixedWidthBreaks() {
        const content = fixedWidthPreview.textContent;
        if (!content) {
            showToast("No content to analyze.", "warning");
            return;
        }
    
        const lines = content.split('\n').filter(line => line.trim() !== '');
        if (lines.length < 1) { // A header-only file can still be analyzed
            showToast("Not enough data to suggest breaks.", "warning");
            return;
        }
    
        const maxLineLength = lines.reduce((max, line) => Math.max(max, line.length), 0);
        const breakScores = new Array(maxLineLength).fill(0);
        const isNumeric = char => char >= '0' && char <= '9';
        const isAlpha = char => (char >= 'a' && char <= 'z') || (char >= 'A' && char <= 'Z');
        const isSpace = char => char === ' ';
    
        // --- Scoring Algorithm ---
    
        // 1. Analyze each column position across all lines
        for (let i = 1; i < maxLineLength; i++) {
            let dataTypeChanges = 0;
            let dataToSpaceTransitions = 0;
            let spaceToDataTransitions = 0;
    
            lines.forEach(line => {
                const prevChar = line[i - 1] || ' ';
                const currentChar = line[i] || ' ';
    
                // A. Score transitions in data type (e.g., number to text)
                const prevIsNum = isNumeric(prevChar);
                const currentIsNum = isNumeric(currentChar);
                const prevIsAlpha = isAlpha(prevChar);
                const currentIsAlpha = isAlpha(currentChar);
    
                if ((prevIsNum && currentIsAlpha) || (prevIsAlpha && currentIsNum)) {
                    dataTypeChanges++;
                }
    
                // B. Score alignment gaps (data ending and space beginning)
                if (!isSpace(prevChar) && isSpace(currentChar)) {
                    dataToSpaceTransitions++;
                }
    
                // C. Score space ending and data beginning
                if (isSpace(prevChar) && !isSpace(currentChar)) {
                    spaceToDataTransitions++;
                }
            });
    
            // Normalize and add weighted scores for this position
            const lineCount = lines.length;
            breakScores[i] += (dataTypeChanges / lineCount) * 2.0; // Strong indicator
            breakScores[i] += (dataToSpaceTransitions / lineCount) * 1.5;
            breakScores[i] += (spaceToDataTransitions / lineCount) * 1.5;
        }
    
        // 2. Add bonus for word breaks in the header
        const headerLine = lines[0];
        if (headerLine) {
            for (let i = 1; i < maxLineLength; i++) {
                const prevChar = headerLine[i - 1] || ' ';
                const currentChar = headerLine[i] || ' ';
    
                // Bonus for space between words (e.g., "Item Number")
                if (!isSpace(prevChar) && isSpace(currentChar)) {
                    breakScores[i] += 0.8;
                }
                // Bonus for CamelCase breaks (e.g., "NumberUnit")
                if (isAlpha(prevChar) && isAlpha(currentChar) && currentChar === currentChar.toUpperCase() && prevChar !== ' ') {
                    breakScores[i] += 1.2; // Higher weight for this pattern
                }
            }
        }
    
        // 3. Identify break points from scores by finding local peaks
        const potentialBreaks = new Set();
        let lastBreak = 0;
        const MIN_COLUMN_WIDTH = 3; // Avoid creating tiny, nonsensical columns
        const SCORE_THRESHOLD = 0.9; // Minimum score to be considered a potential break
    
        for (let i = 1; i < maxLineLength; i++) {
            const currentScore = breakScores[i];
            const prevScore = breakScores[i - 1] || 0;
            const nextScore = breakScores[i + 1] || 0;
    
            // A point is a peak if it's higher than its neighbors and above the threshold
            if (currentScore > SCORE_THRESHOLD && currentScore > prevScore && currentScore >= nextScore && (i - lastBreak) >= MIN_COLUMN_WIDTH) {
                potentialBreaks.add(i);
                lastBreak = i;
            }
        }
    
        if (potentialBreaks.size === 0) {
            showToast("Could not automatically determine column breaks.", "info");
            return;
        }
    
        // 4. Apply the suggestions
        breakPositions = [...potentialBreaks].sort((a, b) => a - b);
        renderBreaks();
        renderHeaderOptions();
        showToast(`Suggested ${breakPositions.length} column breaks.`, "success");
    }
    // --- Filter Builder Logic ---
    function initFilterBuilder(fileIndex, headers = []) {
        const builderId = `filter-builder-${fileIndex}`;
        const container = document.getElementById(builderId);
        if (!container) return;

        const rulesContainer = container.querySelector('.filter-rules-container');
        if (rulesContainer) rulesContainer.innerHTML = ''; // Clear existing rules
        addFilterRule(fileIndex, headers); // Add one initial rule
    }

    function addFilterRule(fileIndex, headers) {
        const fileHeaders = headers || (fileIndex === 1 ? fullFileData1?.headers : fullFileData2?.headers) || [];
        const rulesContainer = document.querySelector(`#filter-builder-${fileIndex} .filter-rules-container`);
        if (!rulesContainer) return;

        const ruleDiv = document.createElement('div');
        ruleDiv.className = 'filter-rule';

        // Add AND/OR selector for all but the first rule
        const logicSelector = rulesContainer.children.length > 0
            ? `<select class="filter-rule-logic">
                   <option value="AND" selected>AND</option>
                   <option value="OR">OR</option>
               </select>`
            : `<div class="filter-rule-logic" style="flex: 1;">Where</div>`; // Placeholder for alignment

        // Security Fix: Build elements programmatically to prevent XSS from file headers.
        ruleDiv.insertAdjacentHTML('beforeend', logicSelector);

        const columnSelect = document.createElement('select');
        columnSelect.className = 'filter-column';
        fileHeaders.forEach(h => {
            const option = document.createElement('option');
            option.value = h;
            option.textContent = h; // Use textContent to safely render header names
            columnSelect.appendChild(option);
        });
        ruleDiv.appendChild(columnSelect);

        const operatorSelect = document.createElement('select');
        operatorSelect.className = 'filter-operator';
        operatorSelect.innerHTML = `
            <option value="==">is equal to</option>
            <option value="!=">is not equal to</option>
            <option value=">">is greater than</option>
            <option value="<">is less than</option>
            <option value=">=">is greater than or equal to</option>
            <option value="<=">is less than or equal to</option>
            <option value="contains">contains</option>
            <option value="not_contains">does not contain</option>
            <option value="starts_with">starts with</option>
            <option value="ends_with">ends with</option>
            <option value="is_empty">is empty</option>
            <option value="is_not_empty">is not empty</option>
            <option value="regex">matches regex</option>
        `;
        ruleDiv.appendChild(operatorSelect);

        ruleDiv.insertAdjacentHTML('beforeend', '<input type="text" class="filter-value" placeholder="Value">');
        ruleDiv.insertAdjacentHTML('beforeend', '<button class="btn btn-icon remove-rule-btn" title="Remove rule">&times;</button>');

        ruleDiv.querySelector('.remove-rule-btn').addEventListener('click', () => {
            ruleDiv.remove();
            applyFilters(fileIndex); // Automatically re-apply filters on removal
        });
        rulesContainer.appendChild(ruleDiv);

        // Redraw lines as adding a rule changes the layout
        redrawLines();
    }

    function reapplyPairings() {
        // This function re-applies the 'paired' class and updates the `th` element
        // references in columnPairs after a table re-render.
        const allTh1 = Array.from(document.querySelectorAll('#preview1 th'));
        const allTh2 = Array.from(document.querySelectorAll('#preview2 th'));

        columnPairs.forEach(pair => {
            const newTh1 = allTh1.find(th => parseInt(th.dataset.colIndex, 10) === pair.file1ColIndex);
            const newTh2 = allTh2.find(th => parseInt(th.dataset.colIndex, 10) === pair.file2ColIndex);

            if (newTh1 && newTh2) {
                newTh1.classList.add('paired');
                newTh2.classList.add('paired');
                pair.file1Th = newTh1; // Update reference
                pair.file2Th = newTh2; // Update reference
            }
        });

        // --- NEW LOGIC ---
        // Also re-apply comparison pairs (dotted lines)
        comparisonPairs.forEach(pair => {
            const newTh1 = allTh1.find(th => parseInt(th.dataset.colIndex, 10) === pair.file1ColIndex);
            const newTh2 = allTh2.find(th => parseInt(th.dataset.colIndex, 10) === pair.file2ColIndex);

            if (newTh1 && newTh2) {
                newTh1.classList.add('paired-compare');
                newTh2.classList.add('paired-compare');
                pair.file1Th = newTh1; // Update reference
                pair.file2Th = newTh2; // Update reference
            }
        });
        redrawLines();
    }

    function applyFilters(fileIndex, silent = false) { // This function will now return a Promise
        if (!silent) showPreviewLoader(fileIndex, true, 'Applying filters...');

        isDragDisabled = true;
        // Defer the heavy filtering to allow the loader to render.
        setTimeout(() => {
            const fullData = fileIndex === 1 ? fullFileData1 : fullFileData2;
            const preview = fileIndex === 1 ? preview1 : preview2;
            const unmatched = fileIndex === 1 ? Array.from(unmatchedHeaders1) : Array.from(unmatchedHeaders2);

            if (!fullData) {
                if (!silent) showPreviewLoader(fileIndex, false);
                return;
            }
            resultsSection.classList.add('hidden'); // Hide results as they are now invalid

            let filteredData = JSON.parse(JSON.stringify(fullData)); // Start with a fresh copy
            const builder = document.getElementById(`filter-builder-${fileIndex}`);
            const rules = Array.from(builder.querySelectorAll('.filter-rule'));

            const currentFilters = rules.map(ruleEl => ({
                logic: ruleEl.querySelector('.filter-rule-logic')?.value || 'AND',
                column: ruleEl.querySelector('.filter-column').value,
                operator: ruleEl.querySelector('.filter-operator').value,
                value: ruleEl.querySelector('.filter-value').value
            }));

            if (fileIndex === 1) appliedFilters1 = currentFilters;
            else appliedFilters2 = currentFilters;

            if (rules.length === 0) {
                if (fileIndex === 1) fileData1 = filteredData; else fileData2 = filteredData;
                renderPreviewTable(filteredData, preview, fileIndex, unmatched);
                if (!silent) reapplyPairings();
                if (!silent) showPreviewLoader(fileIndex, false);
                return;
            }

            const headerIndexMap = new Map(fullData.headers.map((h, i) => [h, i]));

            const filteredRows = fullData.rows.filter(row => {
                if (rules.length === 0) return true;
                let finalResult;
                for (let i = 0; i < rules.length; i++) {
                    const ruleEl = rules[i];
                    const column = ruleEl.querySelector('.filter-column').value;
                    const operator = ruleEl.querySelector('.filter-operator').value;
                    const value = ruleEl.querySelector('.filter-value').value;
                    const ruleResult = checkCondition(row[headerIndexMap.get(column)], operator, value);
                    if (i === 0) {
                        finalResult = ruleResult;
                    } else {
                        const logic = ruleEl.querySelector('.filter-rule-logic').value;
                        if (logic === 'AND') finalResult = finalResult && ruleResult; else finalResult = finalResult || ruleResult;
                    }
                }
                return finalResult;
            });

            filteredData.rows = filteredRows;
            if (fileIndex === 1) fileData1 = filteredData; else fileData2 = filteredData;
            renderPreviewTable(filteredData, preview, fileIndex, unmatched);
            if (!silent) {
                // This is the fix. Re-applying pairings will now also handle
                // the comparison pairs and redraw all lines correctly.
                reapplyPairings();
            }
            if (!silent) showPreviewLoader(fileIndex, false);
        }, 10); // The function returns before this timeout completes
    }

    function checkCondition(cellValue, operator, filterValue) {
        const cell = (cellValue || '').toString().trim();
        const filter = filterValue.toString().trim();
        const cellNum = parseFloat(cell);
        const filterNum = parseFloat(filter);

        switch (operator) {
            case '==': return cell === filter;
            case '!=': return cell !== filter;
            case '>': return !isNaN(cellNum) && !isNaN(filterNum) && cellNum > filterNum;
            case '<': return !isNaN(cellNum) && !isNaN(filterNum) && cellNum < filterNum;
            case '>=': return !isNaN(cellNum) && !isNaN(filterNum) && cellNum >= filterNum;
            case '<=': return !isNaN(cellNum) && !isNaN(filterNum) && cellNum <= filterNum;
            case 'contains': return cell.toLowerCase().includes(filter.toLowerCase());
            case 'not_contains': return !cell.toLowerCase().includes(filter.toLowerCase());
            case 'starts_with': return cell.toLowerCase().startsWith(filter.toLowerCase());
            case 'ends_with': return cell.toLowerCase().endsWith(filter.toLowerCase());
            case 'is_empty': return cell === '';
            case 'is_not_empty': return cell !== '';
            case 'regex':
                try {
                    // Check if the filter string is a valid regex pattern (e.g., /pattern/flags)
                    const match = filter.match(/^\/(.*)\/([gimyus]*)$/);
                    if (match) {
                        // It's a full regex literal, use its parts
                        return new RegExp(match[1], match[2]).test(cell);
                    }
                    return new RegExp(filter).test(cell);
                } catch (e) {
                    return false; // Invalid regex pattern
                }
            default: return false;
        }
    }

    // --- UI Rendering ---
    function renderPreviewTable(data, container, fileIndex) {
        if (!data) {
            container.innerHTML = '<p>No data to display.</p>';
            return;
        }
        const unmatchedSet = fileIndex === 1 ? unmatchedHeaders1 : unmatchedHeaders2;

        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');
        const headerRow = document.createElement('tr');

        const keyColumnIndices = new Set(
            columnPairs.map(p => fileIndex === 1 ? p.file1ColIndex : p.file2ColIndex)
        );

        data.headers.forEach((header, index) => {
            const th = document.createElement('th');
            th.textContent = header;
            th.dataset.fileIndex = fileIndex;
            th.dataset.colIndex = index;
            
            // Columns are draggable by default, but disabled if any action has been taken.
            th.draggable = !isDragDisabled;

            const isExcluded = unmatchedSet.has(header);

            const lockIcon = document.createElement('span');
            lockIcon.className = 'lock-icon';
            lockIcon.addEventListener('click', (e) => {
                e.stopPropagation();
                toggleColumnLock(th);
            });

            if (isExcluded) {
                // This column is explicitly excluded by the user (or was unique).
                th.classList.add('unmatched');
                th.dataset.unmatched = 'true';
                lockIcon.textContent = '🔓';
                lockIcon.title = 'Include this column in matching';
            } else {
                // This is a common, non-key column, included by default.
                th.dataset.unmatched = 'false'; // This column is common, not unique
                lockIcon.textContent = '🔒'; // Default to included (locked)
                lockIcon.title = 'Exclude this column from matching'; // Correct title
            }
            th.addEventListener('dragstart', handleDragStart);
            th.addEventListener('dragover', handleDragOver);
            th.addEventListener('drop', handleDrop);

            // Right-click handler for context menu
            th.addEventListener('contextmenu', handleColumnRightClick);
            th.appendChild(lockIcon);

            // Add the type menu trigger
            const typeTrigger = document.createElement('div');
            typeTrigger.className = 'type-menu-trigger';
            typeTrigger.title = 'Change column data type';
            typeTrigger.addEventListener('click', (e) => {
                e.stopPropagation(); showColumnTypeMenu(th, fileIndex, index, header);
            });
            th.appendChild(typeTrigger);

            // Click handler for pairing
            th.addEventListener('click', handleColumnClick);

            headerRow.appendChild(th);
        });

        thead.appendChild(headerRow);

        data.rows.slice(0, 20).forEach(rowData => {
            const tr = document.createElement('tr');
            rowData.forEach(cellData => {
                const td = document.createElement('td');
                td.textContent = cellData;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        table.appendChild(thead);
        table.appendChild(tbody);
        container.innerHTML = '';
        container.appendChild(table);

        // --- CRITICAL FIX ---
        // After rendering the table, find its container and attach the scroll listener.
        // This ensures that scrolling the table content correctly triggers a redraw of the lines.
        container.addEventListener('scroll', redrawLines);
    }

    function showColumnTypeMenu(th, fileIndex, colIndex, header) {
        // Remove any existing menu
        if (typeMenu) {
            typeMenu.remove();
        }

        const typeMap = fileIndex === 1 ? columnTypes1 : columnTypes2;
        const currentType = typeMap.get(header) || 'auto';

        typeMenu = document.createElement('div');
        typeMenu.className = 'column-type-menu';
        typeMenu.innerHTML = `
            <label>Data Type</label>
            <select>
                <option value="auto" ${currentType === 'auto' ? 'selected' : ''}>Auto-Detect</option>
                <option value="text" ${currentType === 'text' ? 'selected' : ''}>Text</option>
                <option value="number" ${currentType === 'number' ? 'selected' : ''}>Number</option>
                <option value="date" ${currentType === 'date' ? 'selected' : ''}>Date</option>
            </select>
        `;

        document.body.appendChild(typeMenu);

        const rect = th.getBoundingClientRect();
        typeMenu.style.top = `${window.scrollY + rect.bottom + 5}px`;
        typeMenu.style.left = `${window.scrollX + rect.left}px`;

        const select = typeMenu.querySelector('select');
        select.addEventListener('change', (e) => {
            const newType = e.target.value;
            typeMap.set(header, newType);
            reprocessColumn(fileIndex, colIndex, header, newType);
            typeMenu.remove();
            typeMenu = null;
        });
    }

    function reprocessColumn(fileIndex, colIndex, header, newType) {
        const fullData = fileIndex === 1 ? fullFileData1 : fullFileData2;
        const currentData = fileIndex === 1 ? fileData1 : fileData2;
        const preview = fileIndex === 1 ? preview1 : preview2;

        if (!fullData || !currentData) return;

        showPreviewLoader(fileIndex, true, `Applying type '${newType}'...`);

        setTimeout(() => {
            // Find the index in the potentially filtered data
            const currentDataColIndex = currentData.headers.indexOf(header);
            if (currentDataColIndex === -1) {
                showPreviewLoader(fileIndex, false);
                return; // Column not in current view
            }

            // Reprocess the values in the *current* view
            currentData.rows.forEach(row => {
                const originalValue = row[currentDataColIndex];
                if (newType === 'number') {
                    const num = Number(originalValue);
                    // Only change if it's a valid number, otherwise keep original
                    if (!isNaN(num) && originalValue.trim() !== '') {
                        row[currentDataColIndex] = num;
                    }
                } else if (newType === 'date') {
                    // For dates, we just keep the string and handle it in comparison.
                    // No in-place conversion needed.
                } else { // text or auto
                    // Ensure it's a string
                    row[currentDataColIndex] = (originalValue ?? '').toString();
                }
            });

            renderPreviewTable(currentData, preview, fileIndex);
            reapplyPairings();
            showPreviewLoader(fileIndex, false);
        }, 10);
    }

    function updateUnmatchedStyles(fileIndex, unmatchedSet) {
        const previewContainer = document.getElementById(`preview${fileIndex}`);
        const thElements = previewContainer.querySelectorAll('th');

        thElements.forEach(th => {
            const headerText = th.textContent.replace(/🔓|🔒/g, '').trim();
            const isPaired = th.classList.contains('paired');
            let lockIcon = th.querySelector('.lock-icon');

            if (unmatchedSet.has(headerText)) {
                th.classList.add('unmatched');
                th.dataset.unmatched = 'true';
                th.draggable = false;
                if (lockIcon) {
                    lockIcon.textContent = '🔓';
                    lockIcon.title = 'Include this column in matching';
                } else {
                    lockIcon = document.createElement('span');
                    lockIcon.className = 'lock-icon';
                    lockIcon.addEventListener('click', (e) => { e.stopPropagation(); toggleColumnLock(th); });
                    th.appendChild(lockIcon);
                    lockIcon.textContent = '🔓';
                    lockIcon.title = 'Include this column in matching';
                }
            } else {
                th.classList.remove('unmatched');
                th.dataset.unmatched = 'false';

                if (!isPaired) {
                    if (lockIcon) {
                        lockIcon.textContent = '🔒';
                        lockIcon.title = 'Exclude this column from matching';
                    } else {
                        // This is a common, non-key column. It should have a lock icon.
                        lockIcon = document.createElement('span');
                        lockIcon.className = 'lock-icon';
                        lockIcon.addEventListener('click', (e) => { e.stopPropagation(); toggleColumnLock(th); });
                        th.appendChild(lockIcon);
                        lockIcon.textContent = '🔒';
                        lockIcon.title = 'Exclude this column from matching';
                    }
                }
            }
        });
        redrawLines();
    }

    function toggleColumnLock(th) {
        const isCurrentlyExcluded = th.dataset.unmatched === 'true';
        const fileIndex = parseInt(th.dataset.fileIndex, 10); // The header text is read from the element, so it must be clean.
        const icon = th.querySelector('.lock-icon');
        const headerText = th.textContent.replace(icon.textContent, '').trim();

        if (isCurrentlyExcluded) {
            // --- INCLUDING (Locking) ---
            // User is clicking '🔓' to include the column.
            th.classList.remove('unmatched');
            th.dataset.unmatched = 'false';
            icon.textContent = '🔒'; // Lock emoji
            icon.title = 'Exclude this column from matching';
            // Remove from the exclusion list
            if (fileIndex === 1) {
                unmatchedHeaders1.delete(headerText);
            } else {
                unmatchedHeaders2.delete(headerText);
            }
        } else {
            // --- EXCLUDING (Unlocking) ---
            // User is clicking '🔒' to exclude the column.
            th.classList.add('unmatched');
            th.dataset.unmatched = 'true';
            icon.textContent = '🔓'; // Unlock emoji
            icon.title = 'Include this column in matching';
            // Add to the exclusion list
            if (fileIndex === 1) {
                unmatchedHeaders1.add(headerText);
            } else {
                unmatchedHeaders2.add(headerText);
            }

            // --- NEW LOGIC ---
            // If the excluded column is part of a comparison pair, break that pairing too.
            const comparePairIndex = comparisonPairs.findIndex(p =>
                (fileIndex === 1 && p.file1ColIndex === parseInt(th.dataset.colIndex, 10)) ||
                (fileIndex === 2 && p.file2ColIndex === parseInt(th.dataset.colIndex, 10))
            );
            if (comparePairIndex > -1) unpairColumn(th);

            // Check if the excluded column is part of a pair.
            const colIndex = parseInt(th.dataset.colIndex, 10);
            const pairIndex = columnPairs.findIndex(p => 
                (fileIndex === 1 && p.file1ColIndex === colIndex) ||
                (fileIndex === 2 && p.file2ColIndex === colIndex)
            );

            if (pairIndex > -1) {
                // This column was part of a pair, so we need to break that specific pairing.
                const removedPair = columnPairs.splice(pairIndex, 1)[0];
                
                // Remove the 'paired' class from both headers of the now-broken pair.
                removedPair.file1Th.classList.remove('paired');
                removedPair.file2Th.classList.remove('paired');

                // Redraw lines to reflect the broken connection.
                redrawLines();
            }

            // --- NEW LOGIC ---
            // If the column being excluded is the one currently selected for pairing, unselect it.
            if (selectedTh && selectedTh.el === th) {
                th.classList.remove('selected');
                selectedTh = null;
            }
        }
        resultsSection.classList.add('hidden');
    }

    function updateUIState() {
        const bothFilesLoaded = !!(fileData1 && fileData2);

        // Comparison buttons
        manualMatchBtn.disabled = !(bothFilesLoaded && columnPairs.length > 0);
        // Template management visibility and button states (client-side)
        if (bothFilesLoaded) {
            templateManagementContainer.classList.remove('hidden');
        } else {
            templateManagementContainer.classList.add('hidden');
        }
        // Check if there are any meaningful filters applied.
        // A rule is meaningful if it's not the single default empty rule.
        const hasMeaningfulFilters = Array.from(document.querySelectorAll('.filter-rule')).some(rule => {
            const value = rule.querySelector('.filter-value')?.value;
            const operator = rule.querySelector('.filter-operator')?.value;
            // A rule is meaningful if it has a value, or if the operator doesn't require a value.
            return value || ['is_empty', 'is_not_empty'].includes(operator);
        });

        const canSaveTemplate = bothFilesLoaded && (columnPairs.length > 0 || keyValueMappings.size > 0 || hasMeaningfulFilters);
        runComparisonBtn.disabled = !(bothFilesLoaded && columnPairs.length > 0); // Enable template button if files are loaded
        templateDropdownBtn.disabled = !bothFilesLoaded;
        // Disable save action if there's nothing to save
        saveTemplateAction.classList.toggle('disabled', !canSaveTemplate);

        redrawLines();
    }

    // --- Column Drag & Drop ---
    let draggedColumnIndex = null;
    let draggedFileIndex = null;

    function handleDragStart(e) {
        if (isDragDisabled) {
            e.preventDefault();
            return;
        }
        draggedColumnIndex = parseInt(e.target.dataset.colIndex, 10);
        draggedFileIndex = parseInt(e.target.dataset.fileIndex, 10);
        e.dataTransfer.effectAllowed = 'move';
    }

    function handleDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
    }

    function handleDrop(e) {
        e.preventDefault();
        const targetTh = e.target.closest('th');
        if (!targetTh) return;

        if (isDragDisabled) return; // Prevent dragging if state is locked
        const targetColumnIndex = parseInt(targetTh.dataset.colIndex, 10);
        const targetFileIndex = parseInt(targetTh.dataset.fileIndex, 10);

        if (draggedFileIndex !== targetFileIndex) return; // Prevent dragging between tables

        // --- BUG FIX ---
        // The reordering must happen on BOTH the filtered data (fileData) and the
        // original full data (fullFileData). The value mapping modal uses fullFileData,
        // so if it's not reordered, it will pull values from the wrong columns.
        const dataToReorder = (draggedFileIndex === 1) ? [fileData1, fullFileData1] : [fileData2, fullFileData2];
        const container = (draggedFileIndex === 1) ? preview1 : preview2;

        dataToReorder.forEach(data => {
            if (!data) return; // In case one of the datasets is null
            
            // Reorder headers
            const [draggedHeader] = data.headers.splice(draggedColumnIndex, 1);
            data.headers.splice(targetColumnIndex, 0, draggedHeader);
    
            // Reorder rows
            data.rows.forEach(row => {
                const [draggedCell] = row.splice(draggedColumnIndex, 1);
                row.splice(targetColumnIndex, 0, draggedCell);
            });
        });

        clearAllMatches();
        renderPreviewTable(dataToReorder[0], container, draggedFileIndex); // Render using the filtered data
    }

    // --- Column Pairing and Line Drawing ---
    function handleColumnClick(e) {
        // Ensure we're targeting the TH element, not a child like the checkbox
        const th = e.target.closest('th');
        if (!th) return;
        // Do not allow any action on an already paired (key or compare) or unmatched column.
        if (th.classList.contains('paired') || th.classList.contains('paired-compare') || th.classList.contains('unmatched')) {
            return;
        }
        // If a right-click selection is active, a left-click should cancel it.
        if (rightClickSelectedTh) {
            cancelRightClickSelection();
        }

        // --- BUG FIX ---
        // Before handling a new click, ensure the `th` references in the existing
        // columnPairs array are up-to-date with the current DOM. This prevents
        // template-applied pairs from losing their visual lines when a new
        // column is selected for pairing.
        if (columnPairs.length > 0) {
            const allTh1 = Array.from(document.querySelectorAll('#preview1 th'));
            const allTh2 = Array.from(document.querySelectorAll('#preview2 th'));
            columnPairs.forEach(pair => {
                pair.file1Th = allTh1.find(th => parseInt(th.dataset.colIndex, 10) === pair.file1ColIndex);
                pair.file2Th = allTh2.find(th => parseInt(th.dataset.colIndex, 10) === pair.file2ColIndex);
            });
        }
        if (comparisonPairs.length > 0) {
            const allTh1 = Array.from(document.querySelectorAll('#preview1 th'));
            const allTh2 = Array.from(document.querySelectorAll('#preview2 th'));
            comparisonPairs.forEach(pair => {
                pair.file1Th = allTh1.find(th => parseInt(th.dataset.colIndex, 10) === pair.file1ColIndex);
                pair.file2Th = allTh2.find(th => parseInt(th.dataset.colIndex, 10) === pair.file2ColIndex);
            });
        }

        const fileIndex = parseInt(th.dataset.fileIndex, 10);
        const colIndex = parseInt(th.dataset.colIndex, 10);

        if (selectedTh) {
            // A column is already selected. Check for the new cases.
            if (selectedTh.el === th) {
                // Case 1: The user clicked the *same* column again. Unselect it.
                th.classList.remove('selected');
                selectedTh = null;
            } else if (selectedTh.fileIndex !== fileIndex) {
                // Case 2: The user clicked a column in the other file. Form a pair.
                const isComparePair = th.classList.contains('paired-compare') || selectedTh.el.classList.contains('paired-compare');

                const pair = {
                    file1ColIndex: fileIndex === 1 ? colIndex : selectedTh.colIndex,
                    file2ColIndex: fileIndex === 2 ? colIndex : selectedTh.colIndex,
                    file1Th: fileIndex === 1 ? th : selectedTh.el,
                    file2Th: fileIndex === 2 ? th : selectedTh.el,
                };
                pair.type = 'key';
                columnPairs.push(pair);
                selectedTh.el.classList.remove('selected');
                selectedTh.el.classList.add('paired');
                th.classList.add('paired');
                selectedTh = null;
                isDragDisabled = true;
            } else {
                // Case 3: The user clicked a different column in the same file. Switch selection.
                selectedTh.el.classList.remove('selected');
                th.classList.add('selected');
                selectedTh = { el: th, fileIndex, colIndex };
            }
        } else {
            // No column was selected. This is the first selection.
            th.classList.add('selected');
            selectedTh = { el: th, fileIndex, colIndex };
        }
        updateUIState();
    }

    function handleColumnRightClick(e) {
        e.preventDefault(); // Prevent the default browser context menu
        const th = e.target.closest('th');
        if (!th) return;

        // Do not allow any action on an already paired (key or compare) or unmatched column.
        if (th.classList.contains('paired') || th.classList.contains('paired-compare') || th.classList.contains('unmatched')) {
            return;
        }

        // A right-click should cancel any active left-click selection.
        if (selectedTh) {
            selectedTh.el.classList.remove('selected');
            selectedTh = null;
        }

        const fileIndex = parseInt(th.dataset.fileIndex, 10);
        const colIndex = parseInt(th.dataset.colIndex, 10);

        if (rightClickSelectedTh) {
            // A column is already selected for comparison.
            if (rightClickSelectedTh.el === th) {
                // Case 1: Right-clicked the same column again. Unselect it.
                th.classList.remove('selected-compare');
                rightClickSelectedTh = null;
            } else if (rightClickSelectedTh.fileIndex !== fileIndex) {
                // Case 2: Right-clicked a column in the other file. Form a comparison pair.
                const pair = {
                    file1ColIndex: fileIndex === 1 ? colIndex : rightClickSelectedTh.colIndex,
                    file2ColIndex: fileIndex === 2 ? colIndex : rightClickSelectedTh.colIndex,
                    file1Th: fileIndex === 1 ? th : rightClickSelectedTh.el,
                    file2Th: fileIndex === 2 ? th : rightClickSelectedTh.el,
                    type: 'compare'
                };
                comparisonPairs.push(pair);
                rightClickSelectedTh.el.classList.remove('selected-compare');
                rightClickSelectedTh.el.classList.add('paired-compare');
                th.classList.add('paired-compare');
                rightClickSelectedTh = null;
                isDragDisabled = true; // An action was taken
            } else {
                // Case 3: Right-clicked a different column in the same file. Switch selection.
                rightClickSelectedTh.el.classList.remove('selected-compare');
                th.classList.add('selected-compare');
                rightClickSelectedTh = { el: th, fileIndex, colIndex };
            }
        } else {
            // No column was selected for comparison. This is the first selection.
            th.classList.add('selected-compare');
            rightClickSelectedTh = { el: th, fileIndex, colIndex };
        }
        updateUIState();
    }

    function clearPairings() {
        columnPairs = [];
        if (selectedTh) {
            selectedTh.el.classList.remove('selected');
            selectedTh.el.classList.remove('selected-compare');
            selectedTh = null;
        }
        document.querySelectorAll('th.paired').forEach(th => th.classList.remove('paired'));
        redrawLines(); // Ensure canvas is cleared immediately after pairings are removed
        isDragDisabled = false;
        resultsSection.classList.add('hidden');
        updateUIState();
    }

    function clearComparisonPairs() {
        comparisonPairs = [];
        cancelRightClickSelection();
        document.querySelectorAll('th.paired-compare').forEach(th => th.classList.remove('paired-compare'));
        redrawLines(); // Ensure canvas is cleared immediately after pairings are removed
        isDragDisabled = false;
        resultsSection.classList.add('hidden');
        updateUIState();
    }

    function cancelRightClickSelection() {
        if (rightClickSelectedTh) {
            rightClickSelectedTh.el.classList.remove('selected-compare');
            rightClickSelectedTh = null;
        }
    }

    function clearAllMatches() {
        // Clear pairings
        columnPairs = [];
        if (selectedTh) {
            templateApplied = false; // Clearing matches invalidates template state
            selectedTh.el.classList.remove('selected');
            selectedTh = null;
        }

        document.querySelectorAll('th.paired').forEach(th => th.classList.remove('paired'));
        redrawLines(); // Ensure canvas is cleared immediately after pairings are removed
        
        // Clear filters and reset data to original state
        // Do NOT clear keyValueMappings, filters, or unmatchedHeaders here.
        // The user only wants to clear the visual column pairings.
        // If they want to clear other things, they can use the dedicated buttons or reload files.
        templateApplied = false; // Clearing matches invalidates template state, as pairings are a core part.

        // --- BUG FIX ---
        // Clearing column matches invalidates key value mappings. They must be cleared as well.
        keyValueMappings.clear();
        suggestedKeyValueMappings.clear();
        clearPairings();
        clearComparisonPairs();

        // Re-render both tables. The render function will correctly style headers
        // based on whether they are in a pair or in the (preserved) unmatchedHeaders sets.
        if (fileData1) {
            // Pass the complete, correct set of unmatched headers from the state.
            // --- MEMORY MANAGEMENT ---
            // Explicitly clear large result objects when matches are cleared.
            comparisonResults = [];
            alignedData1 = null;
            alignedOriginalHeaders2ForRender = [];

            renderPreviewTable(fileData1, preview1, 1);
        }
        if (fileData2) {
            // Pass the complete, correct set of unmatched headers from the state.
            renderPreviewTable(fileData2, preview2, 2);
        }

        // Hide the results section
        resultsSection.classList.add('hidden');

        // Clear comparison-only pairs as well
        clearComparisonPairs();

        isDragDisabled = false;

        updateUIState();
    }

    function redrawLines() {
        const containerRect = document.querySelector('.previews-container').getBoundingClientRect();
        const preview1Rect = preview1.getBoundingClientRect();
        const preview2Rect = preview2.getBoundingClientRect();

        // --- CRITICAL FIX for SCROLLING ---
        // Get the scroll containers to account for their scroll position.
        // This ensures lines stay anchored when individual tables are scrolled.
        const scrollContainer1 = preview1.querySelector('.table-container');
        const scrollContainer2 = preview2.querySelector('.table-container');

        canvas.width = containerRect.width;
        canvas.height = containerRect.height;
        ctx.clearRect(0, 0, canvas.width, canvas.height);

        // Get computed theme colors and make them semi-transparent
        const computedStyle = getComputedStyle(document.documentElement);
        const primaryColor = computedStyle.getPropertyValue('--primary-color').trim();
        const secondaryColor = computedStyle.getPropertyValue('--secondary-color').trim();
        const LINE_ALPHA = 0.5;
        
        // Draw Key Pairs (thicker, primary color)
        columnPairs.forEach(pair => {
            const rect1 = pair.file1Th.getBoundingClientRect();
            const rect2 = pair.file2Th.getBoundingClientRect();

            ctx.strokeStyle = hexToRgba(primaryColor, LINE_ALPHA);
            ctx.lineWidth = 2.5;
            ctx.setLineDash([]); // Solid line for key pairs

            let startX = rect1.left - containerRect.left + rect1.width / 2;
            let startY = rect1.top - containerRect.top + rect1.height;
            let endX = rect2.left - containerRect.left + rect2.width / 2;
            let endY = rect2.top - containerRect.top;

            // Get the boundaries of the preview containers relative to the canvas
            const p1Left = preview1Rect.left - containerRect.left;
            const p1Right = preview1Rect.right - containerRect.left;
            const p2Left = preview2Rect.left - containerRect.left;
            const p2Right = preview2Rect.right - containerRect.left;

            // Clip the start and end points to their respective container boundaries
            startX = Math.max(p1Left, Math.min(startX, p1Right));
            endX = Math.max(p2Left, Math.min(endX, p2Right));

            ctx.beginPath();
            ctx.moveTo(startX, startY);
            ctx.bezierCurveTo(startX, startY + 50, endX, endY - 50, endX, endY);
            ctx.stroke();
        });

        // Draw Comparison Pairs (thinner, secondary color)
        ctx.strokeStyle = hexToRgba(secondaryColor, LINE_ALPHA);
        ctx.lineWidth = 1.5;
        ctx.setLineDash([4, 4]); // Dashed line for comparison pairs

        comparisonPairs.forEach(pair => {
            const rect1 = pair.file1Th.getBoundingClientRect();
            const rect2 = pair.file2Th.getBoundingClientRect();

            let startX = rect1.left - containerRect.left + rect1.width / 2;
            let startY = rect1.top - containerRect.top + rect1.height;
            let endX = rect2.left - containerRect.left + rect2.width / 2;
            let endY = rect2.top - containerRect.top;

            const p1Left = preview1Rect.left - containerRect.left;
            const p1Right = preview1Rect.right - containerRect.left;
            const p2Left = preview2Rect.left - containerRect.left;
            const p2Right = preview2Rect.right - containerRect.left;

            startX = Math.max(p1Left, Math.min(startX, p1Right));
            endX = Math.max(p2Left, Math.min(endX, p2Right));

            ctx.beginPath();
            ctx.moveTo(startX, startY);
            ctx.bezierCurveTo(startX, startY + 40, endX, endY - 40, endX, endY);
            ctx.stroke();
        });
        ctx.setLineDash([]); // Reset line dash
    }

    // --- Comparison Logic ---
    function runComparison() {
        const btnText = runComparisonBtn.querySelector('.btn-text');
        const btnLoader = runComparisonBtn.querySelector('.btn-loader');

        if (!fileData1 || !fileData2 || columnPairs.length === 0) {
            alert("Please match at least one pair of columns to use as a key.");
            return;
        }

        isDragDisabled = true;
        // --- Start Loading State ---
        runComparisonBtn.disabled = true;
        runComparisonBtn.classList.add('loading');

        comparisonResults = [];

        // Defer the heavy work to allow the UI to update and show the spinner
        setTimeout(() => {
            performComparison();
        }, 10);
    }

    function performComparison() {
        const groupBy = groupByCheckbox.checked;
        const btnLoader = runComparisonBtn.querySelector('.btn-loader');
        btnLoader.classList.remove('hidden');

        try {
            if (!fullFileData1 || !fullFileData2) {
                throw new Error("Please ensure both files are loaded correctly.");
            }
            if (columnPairs.length === 0) {
                throw new Error("At least one column pair must be designated as a 'Key' for matching rows.");
            }

            // Use the filtered data as the starting point for comparison
            const data1 = fileData1 || fullFileData1;
            const data2 = fileData2 || fullFileData2;

            // Analyze columns for numeric content *after* filters have been applied.
            analyzeNumericColumns(data1, data2);

            if (groupBy) {
                performGroupedComparison(data1, data2);
            } else {
                performRowByRowComparison(data1, data2);
            }

            showDiffOnlyCheckbox.checked = false;
            showDiffColumnsCheckbox.checked = false;
            renderResultsTable();

        } catch (error) {
            showToast(error.message, "error");
            console.error("Comparison failed:", error);
        } finally {
            // --- End Loading State ---
            runComparisonBtn.disabled = false;
            runComparisonBtn.classList.remove('loading');
            btnLoader.classList.add('hidden');
        }
    }

    // --- Results Rendering and Download ---
    function renderResultsTable() {
        sortState = { columnIndex: -1, direction: 'asc' }; // Reset sort on new comparison
        resultsSection.classList.remove('hidden');
        resultsTableContainer.innerHTML = ''; // Clear previous results
        resultsSummary.innerHTML = ''; // Clear previous summary
        if (comparisonResults.length === 0) {
            resultsTableContainer.innerHTML = '<p>No differences or matching rows found based on selected key columns.</p>';
            return;
        }

        // Update summary message
        const totalRows = comparisonResults.length;
        if (totalRows > 20) {
            resultsSummary.textContent = `Displaying top 20 of ${totalRows} total result rows. Download for the full results.`;
        } else {
            resultsSummary.textContent = `Displaying all ${totalRows} result rows.`;
        }

        // Check if results exceed Excel's row limit
        const MAX_EXCEL_ROWS = 1048576;
        const downloadExcelAction = document.getElementById('download-excel-action');
        if (totalRows > MAX_EXCEL_ROWS) {
            downloadExcelAction.classList.add('disabled');
            downloadExcelAction.title = `Results (${totalRows.toLocaleString()}) exceed Excel's row limit of ${MAX_EXCEL_ROWS.toLocaleString()}.`;
        } else {
            downloadExcelAction.classList.remove('disabled');
            downloadExcelAction.title = '';
        }


        // Conditionally show the 'Diff' column checkbox
        const diffColumnOption = document.getElementById('diff-column-option');
        const hasDiffs = comparisonResults.length > 0 && comparisonResults[0].diffs && Object.keys(comparisonResults[0].diffs).length > 0;
        if (hasDiffs) {
            diffColumnOption.classList.remove('hidden');
        } else {
            diffColumnOption.classList.add('hidden');
        }

        // Conditionally show the numeric tolerance option
        const numericToleranceOption = document.getElementById('numeric-tolerance-option');
        const hasNumericPairs = [...comparisonPairs, ...columnPairs].some(p => (columnTypes1.get(fullFileData1.headers[p.file1ColIndex]) === 'number' || columnTypes2.get(fullFileData2.headers[p.file2ColIndex]) === 'number'));
        if (hasNumericPairs) {
            numericToleranceOption.classList.remove('hidden');
        } else {
            numericToleranceOption.classList.add('hidden');
        }
        resultsOptions.classList.remove('hidden'); // Make checkbox visible for both views
        if (groupByCheckbox.checked) {
            renderGroupedResultsTable();
        } else {
            renderRowByRowResultsTable();
        }
        resultsSection.scrollIntoView({ behavior: 'smooth' });
    }

    function renderRowByRowResultsTable() {
        const showDiffOnly = showDiffOnlyCheckbox.checked;
        const hideUnmatched = hideUnmatchedKeysCheckbox.checked;
        resultsTableContainer.innerHTML = ''; // Clear previous table before re-rendering
        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');
        const headerRow = document.createElement('tr');
        headerRow.style.cursor = 'pointer';

        alignedData1.headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = `File 1: ${header}`;
            th.dataset.sortIndex = headerRow.children.length;
            headerRow.appendChild(th);
        });

        alignedOriginalHeaders2ForRender.forEach(header => { // Use aligned original headers
            const th = document.createElement('th');
            th.textContent = `File 2: ${header}`;
            th.dataset.sortIndex = headerRow.children.length;
            headerRow.appendChild(th);
        });

        // Add headers for the new Diff columns at the end
        if (showDiffColumnsCheckbox.checked && comparisonResults.length > 0 && comparisonResults[0].diffs) {
            Object.keys(comparisonResults[0].diffs).forEach(diffKey => {
                const th = document.createElement('th');
                th.textContent = diffKey;
                th.dataset.sortIndex = headerRow.children.length;
                headerRow.appendChild(th);
            });
        }
        thead.appendChild(headerRow);

        headerRow.addEventListener('click', (e) => {
            const th = e.target.closest('th');
            if (th && th.dataset.sortIndex) {
                const index = parseInt(th.dataset.sortIndex, 10);
                sortRowByRowResults(index, alignedData1.headers.length, alignedOriginalHeaders2ForRender.length, showDiffColumnsCheckbox.checked ? Object.keys(comparisonResults[0]?.diffs || {}).length : 0);
                // Re-render the table with sorted data
                renderRowByRowResultsTable();
                // Update header indicators
                updateSortIndicators(headerRow, index);
            }
        });

        updateSortIndicators(headerRow, sortState.columnIndex);

        let filteredResults = showDiffOnly
            ? comparisonResults.filter(result => {
                // A row has a difference if any cell in row1 doesn't match the corresponding cell in row2
                // This requires finding the corresponding columns if they are paired.
                const allPairs = [...columnPairs, ...comparisonPairs];
                const headers1 = alignedData1.headers;
                const headers2 = alignedOriginalHeaders2ForRender;

                // Check for differences in paired/common columns
                for (let i = 0; i < headers1.length; i++) {
                    const h1 = headers1[i];
                    const pair = allPairs.find(p => fullFileData1.headers[p.file1ColIndex] === h1);
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h1;
                    const j = headers2.indexOf(h2);

                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';

                    if (j !== -1 && areValuesDifferent(result.row1[i], result.row2[j], h1, h2, type1, type2)) return true;
                }
                return false; // No differences found in comparable columns
            })
            : comparisonResults;

        if (hideUnmatched) {
            filteredResults = filteredResults.filter(result => {
                // A row is considered "matched" if it has data from both files.
                // We check if either row is a "blank" row (all nulls).
                const isRow1Blank = result.row1.every(cell => cell === null);
                const isRow2Blank = result.row2.every(cell => cell === null);
                return !isRow1Blank && !isRow2Blank;
            });
        }

        // --- Implicit Pairing Logic for Row-by-Row ---
        let implicitPair = null;
        const nonKeyNumericHeaders1 = alignedData1.headers.filter(h => numericColumns.has(h) && !columnPairs.some(p => fullFileData1.headers[p.file1ColIndex] === h));
        const nonKeyNumericHeaders2 = alignedOriginalHeaders2ForRender.filter(h => numericColumns.has(h) && !columnPairs.some(p => fullFileData2.headers[p.file2ColIndex] === h));

        if (nonKeyNumericHeaders1.length === 1 && nonKeyNumericHeaders2.length === 1 && nonKeyNumericHeaders1[0] !== nonKeyNumericHeaders2[0]) {
            const h1 = nonKeyNumericHeaders1[0];
            const h2 = nonKeyNumericHeaders2[0];
            // Check if they are not already part of a comparison pair
            const isAlreadyPaired = comparisonPairs.some(p => fullFileData1.headers[p.file1ColIndex] === h1 && fullFileData2.headers[p.file2ColIndex] === h2);
            if (!isAlreadyPaired) {
                implicitPair = { file1Header: h1, file2Header: h2 };
            }
        }

        filteredResults.slice(0, 20).forEach(result => {
            const tr = document.createElement('tr');

            // Create all cells first to handle formatting between pairs
            const cells1 = result.row1.map((cell, index) => {
                const td = document.createElement('td');
                // Find corresponding cell in row2 to check for diff
                const h1 = alignedData1.headers[index];
                let pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData1.headers[p.file1ColIndex] === h1);
                if (!pair && implicitPair && implicitPair.file1Header === h1) {
                    pair = { file2ColIndex: fullFileData2.headers.indexOf(implicitPair.file2Header) };
                }
                const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h1;
                const index2 = alignedOriginalHeaders2ForRender.indexOf(h2);
                const val2 = result.row2[index2];
                
                const type1 = columnTypes1.get(h1) || 'auto';
                const type2 = columnTypes2.get(h2) || 'auto';

                const formatted = formatNumericValues(cell, val2, h1, h2, type1, type2);
                td.textContent = formatted.f1;
                if (index2 === -1 || areValuesDifferent(cell, result.row2[index2], h1, h2, type1, type2)) {
                    td.classList.add('diff');
                }
                return { td, h1, h2, index2 }; // Pass info for the next loop
            });

            const cells2 = result.row2.map((cell, index) => {
                const td = document.createElement('td');
                // Find corresponding cell in row1 to check for diff
                const h2 = alignedOriginalHeaders2ForRender[index];
                let pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData2.headers[p.file2ColIndex] === h2);
                if (!pair && implicitPair && implicitPair.file2Header === h2) {
                    pair = { file1ColIndex: fullFileData1.headers.indexOf(implicitPair.file1Header) };
                }
                const h1 = pair ? fullFileData1.headers[pair.file1ColIndex] : h2;
                const index1 = alignedData1.headers.indexOf(h1);
                const val1 = result.row1[index1]; 

                const type1 = columnTypes1.get(h1) || 'auto';
                const type2 = columnTypes2.get(h2) || 'auto';
 
                const formatted = formatNumericValues(val1, cell, h1, h2, type1, type2);
                td.textContent = formatted.f2;
 
                if (index1 === -1 || areValuesDifferent(val1, cell, h1, h2, type1, type2)) {
                    td.classList.add('diff');
                }
                return td;
            });
 
            // Now append all created and formatted cells to the row
            cells1.forEach(c => tr.appendChild(c.td));
            cells2.forEach(td => tr.appendChild(td));

            // Create and append cells for diff columns at the end
            if (showDiffColumnsCheckbox.checked && result.diffs) {
                Object.values(result.diffs).forEach(diffValue => {
                    const diffTd = document.createElement('td');
                    diffTd.textContent = diffValue !== null ? diffValue : '';
                    tr.appendChild(diffTd);
                });
            }

            tbody.appendChild(tr); // Append the completed row
        });

        // Update summary based on filtered results
        const displayedCount = Math.min(filteredResults.length, 20);
        const totalRows = filteredResults.length;
        const originalTotal = comparisonResults.length;
        const showing = showDiffOnly ? `${totalRows} rows with differences (out of ${originalTotal})` : `all ${totalRows} result rows`;

        table.appendChild(thead);
        table.appendChild(tbody);
        resultsTableContainer.appendChild(table);

        resultsSummary.textContent = `Displaying top ${displayedCount} of ${showing}. Download for the full results.`;
    }

    function sortRowByRowResults(columnIndex, file1HeaderCount, file2HeaderCount, diffColCount) {
        if (sortState.columnIndex === columnIndex) {
            sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
        } else {
            sortState.columnIndex = columnIndex;
            sortState.direction = 'asc';
        }
        const direction = sortState.direction === 'asc' ? 1 : -1;
        const totalHeadersBeforeDiff = file1HeaderCount + file2HeaderCount;

        comparisonResults.sort((a, b) => {
            let valA, valB;
            if (columnIndex < file1HeaderCount) { // File 1 columns
                valA = a.row1[columnIndex] || '';
                valB = b.row1[columnIndex] || '';
            } else if (columnIndex < totalHeadersBeforeDiff) { // File 2 columns
                valA = a.row2[columnIndex - file1HeaderCount] || '';
                valB = b.row2[columnIndex - file1HeaderCount] || '';
            } else { // Diff columns
                const diffKey = Object.keys(a.diffs)[columnIndex - totalHeadersBeforeDiff];
                valA = a.diffs[diffKey] ?? -Infinity;
                valB = b.diffs[diffKey] ?? -Infinity;
                return (valA - valB) * direction;
            }
            return valA.toString().localeCompare(valB.toString(), undefined, { numeric: true }) * direction;
        });
    }

    function renderGroupedResultsTable() {
        const showDiffOnly = showDiffOnlyCheckbox.checked;
        const hideUnmatched = hideUnmatchedKeysCheckbox.checked;
        resultsTableContainer.innerHTML = ''; // Clear previous table before re-rendering
        const table = document.createElement('table');
        const thead = document.createElement('thead');

        const allGroupedHeaders1 = new Set();
        const allGroupedHeaders2 = new Set();
        comparisonResults.forEach(result => {
            Object.keys(result.file1Sum).forEach(h => allGroupedHeaders1.add(h));
            Object.keys(result.file2Sum).forEach(h => allGroupedHeaders2.add(h));
        });
        const sortedHeaders1 = [...allGroupedHeaders1].sort();
        const sortedHeaders2 = [...allGroupedHeaders2].sort();
        // --- FIX: Define ordered headers to maintain original column order ---
        const orderedHeaders1 = fullFileData1.headers.filter(h => allGroupedHeaders1.has(h));
        const orderedHeaders2 = fullFileData2.headers.filter(h => allGroupedHeaders2.has(h));

        const tbody = document.createElement('tbody');

        // --- Implicit Pairing Logic ---
        // If there is exactly one non-key, included, numeric column in each file,
        // and they are not already paired, treat them as an implicit pair for this comparison.
        let implicitPair = null;
        if (sortedHeaders1.length === 1 && sortedHeaders2.length === 1 && sortedHeaders1[0] !== sortedHeaders2[0]) {
            const h1 = sortedHeaders1[0];
            const h2 = sortedHeaders2[0];
            const isAlreadyPaired = [...columnPairs, ...comparisonPairs].some(p => {
                const pairH1 = fullFileData1.headers[p.file1ColIndex];
                const pairH2 = fullFileData2.headers[p.file2ColIndex];
                return (pairH1 === h1 && pairH2 === h2) || (pairH1 === h2 && pairH2 === h1);
            });

            if (!isAlreadyPaired) {
                implicitPair = { file1Header: h1, file2Header: h2 };
            }
        }

        const headerRow = document.createElement('tr');
        headerRow.style.cursor = 'pointer';

        const keyTh = document.createElement('th');
        keyTh.textContent = 'Key';
        keyTh.dataset.sortIndex = 0;
        headerRow.appendChild(keyTh);

        orderedHeaders1.forEach(h => {
            const th1 = document.createElement('th');
            th1.textContent = `File 1: ${h}`;
            th1.dataset.sortIndex = headerRow.children.length;
            th1.dataset.sortHeader = h;
            th1.dataset.sortFile = '1';
            headerRow.appendChild(th1);
        });
        orderedHeaders2.forEach(h => {
            const th2 = document.createElement('th');
            th2.textContent = `File 2: ${h}`;
            th2.dataset.sortIndex = headerRow.children.length;
            th2.dataset.sortHeader = h; // Keep original header for sorting
            th2.dataset.sortFile = '2';
            headerRow.appendChild(th2);
        });

        // Add headers for the new Diff columns at the end
        if (showDiffColumnsCheckbox.checked && comparisonResults.length > 0 && comparisonResults[0].diffs) {
            Object.keys(comparisonResults[0].diffs).forEach(diffKey => {
                const th = document.createElement('th');
                th.textContent = diffKey;
                th.dataset.sortIndex = headerRow.children.length;
                th.dataset.sortHeader = diffKey;
                headerRow.appendChild(th);
            });
        }

        thead.appendChild(headerRow);

        headerRow.addEventListener('click', (e) => {
            const th = e.target.closest('th');
            if (th && th.dataset.sortIndex) {
                const index = parseInt(th.dataset.sortIndex, 10);
                const header = th.dataset.sortHeader;
                const file = th.dataset.sortFile; // Check if the column is numeric for sorting purposes
                const isNumeric = !nonSummableColumnsForDisclaimer.includes(header);
                sortGroupedResults(index, header, file, isNumeric, showDiffColumnsCheckbox.checked ? Object.keys(comparisonResults[0]?.diffs || {}).length : 0, orderedHeaders1, orderedHeaders2);
                renderGroupedResultsTable();
                updateSortIndicators(headerRow, index);
            }
        });
        updateSortIndicators(headerRow, sortState.columnIndex);

        let filteredResults = showDiffOnly
            ? comparisonResults.filter(result => {
                const allHeaders = new Set([...Object.keys(result.file1Sum), ...Object.keys(result.file2Sum)]);
                return [...allHeaders].some(h => {
                    const pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData1.headers[p.file1ColIndex] === h || fullFileData2.headers[p.file2ColIndex] === h);
                    const h1 = pair ? fullFileData1.headers[pair.file1ColIndex] : h;
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h;
                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';
                    return areValuesDifferent(result.file1Sum[h1], result.file2Sum[h2], h1, h2, type1, type2);
                });
            })
            : comparisonResults;

        if (hideUnmatched) {
            filteredResults = filteredResults.filter(result => {
                // A key is considered "matched" if it has summed values from both files.
                // We check if either of the sum objects is empty.
                const hasFile1Data = Object.keys(result.file1Sum).length > 0;
                const hasFile2Data = Object.keys(result.file2Sum).length > 0;
                return hasFile1Data && hasFile2Data;
            });
        }
        filteredResults.slice(0, 20).forEach(result => {
            const tr = document.createElement('tr');
            const keyTd = document.createElement('td');
            keyTd.textContent = result.key;
            tr.appendChild(keyTd);

            orderedHeaders1.forEach(h => {
                const td = document.createElement('td');
                // Find the corresponding header in file 2, if paired.
                let pair = [...columnPairs, ...comparisonPairs].find(pair => fullFileData1.headers[pair.file1ColIndex] === h);
                // Check for an implicit pair if no explicit one was found.
                if (!pair && implicitPair && implicitPair.file1Header === h) {
                    pair = { file2ColIndex: fullFileData2.headers.indexOf(implicitPair.file2Header) };
                }
                const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h;

                const val1 = result.file1Sum[h] ?? null;
                const val2 = result.file2Sum[h2] ?? null; // Use the paired header name for lookup
                td.textContent = val1 ?? ''; // Display empty string for null
                
                const type1 = columnTypes1.get(h) || 'auto';
                const type2 = columnTypes2.get(h2) || 'auto';
                if (areValuesDifferent(val1, val2, h, h2, type1, type2)) {
                    td.classList.add('diff');
                }
                tr.appendChild(td);
            });

            orderedHeaders2.forEach(h => {
                const td = document.createElement('td');
                // Find the corresponding header in file 1, if paired.
                let pair = [...columnPairs, ...comparisonPairs].find(pair => fullFileData2.headers[pair.file2ColIndex] === h);
                // Check for an implicit pair if no explicit one was found.
                if (!pair && implicitPair && implicitPair.file2Header === h) {
                    pair = { file1ColIndex: fullFileData1.headers.indexOf(implicitPair.file1Header) };
                }
                const h1 = pair ? fullFileData1.headers[pair.file1ColIndex] : h;

                const val1 = result.file1Sum[h1] ?? null; // Use the paired header name for lookup
                const val2 = result.file2Sum[h] ?? null;
                td.textContent = val2 ?? ''; // Display empty string for null

                const type1 = columnTypes1.get(h1) || 'auto';
                const type2 = columnTypes2.get(h) || 'auto';
                if (areValuesDifferent(val1, val2, h1, h, type1, type2)) {
                    td.classList.add('diff');
                }
                tr.appendChild(td);
            });

            // Add cells for diff columns at the end
            if (showDiffColumnsCheckbox.checked && result.diffs) {
                Object.values(result.diffs).forEach(diffValue => {
                    const diffTd = document.createElement('td');
                    diffTd.textContent = diffValue !== null ? diffValue : '';
                    tr.appendChild(diffTd);
                });
            }
            tbody.appendChild(tr);
        });
        table.appendChild(thead);
        table.appendChild(tbody);
        resultsTableContainer.appendChild(table);

        // Update summary based on filtered results
        const displayedCount = Math.min(filteredResults.length, 20);
        const totalRows = filteredResults.length;
        const originalTotal = comparisonResults.length;
        const showing = showDiffOnly ? `${totalRows} keys with differences (out of ${originalTotal})` : `all ${totalRows} result keys`;

        resultsSummary.textContent = `Displaying top ${displayedCount} of ${showing}. Download for the full results.`; // Add disclaimer for non-summable columns if any were excluded

        // Add disclaimer for non-summable columns if any were excluded
        if (nonSummableColumnsForDisclaimer.length > 0) {
            const disclaimer = document.createElement('p');
            disclaimer.className = 'summary-disclaimer';
            disclaimer.textContent = `Note: The following columns contained non-numeric data and were not included in the grouped sums: ${nonSummableColumnsForDisclaimer.join(', ')}.`;
            resultsSummary.appendChild(disclaimer);
        }
    }

    function analyzeNumericColumns(data1, data2) {
        numericColumns.clear();
        const numericRegex = /^-?\d*\.?\d+$/;

        const analyze = (data) => {
            if (!data || !data.headers || !data.rows) return;
            data.headers.forEach((header, index) => {
                // If the column is already marked as numeric, skip it.
                if (numericColumns.has(header)) return;

                // Check if all non-empty values in this column are numeric.
                const isNumeric = data.rows.every(row => {
                    const value = (row[index] ?? '').toString().trim();
                    return value === '' || numericRegex.test(value);
                });

                if (isNumeric) {
                    numericColumns.add(header);
                }
            });
        };

        analyze(data1);
        analyze(data2);
    }

    function performRowByRowComparison(data1, data2) {
        // 1. Define the headers to be included for each file separately.
        // Only "locked" (included) columns should be in the result.
        const headers1 = data1.headers.filter(h => !unmatchedHeaders1.has(h));
        const headers2 = data2.headers.filter(h => !unmatchedHeaders2.has(h));

        // 2. Create maps from header name to original index for data retrieval.
        const f1HeaderOrigIndexMap = new Map(data1.headers.map((h, i) => [h, i]));
        const f2HeaderOrigIndexMap = new Map(data2.headers.map((h, i) => [h, i]));

        // 3. Create the data structures that will be used for comparison.
        // The data is NOT realigned into a master structure; each file keeps its own columns.
        alignedData1 = {
            headers: headers1,
            rows: data1.rows.map(row => headers1.map(h => row[f1HeaderOrigIndexMap.get(h)]))
        };

        const alignedData2 = {
            headers: headers2,
            rows: data2.rows.map(row => headers2.map(h => row[f2HeaderOrigIndexMap.get(h)]))
        };

        // 4. Store the headers for File 2 for rendering. This is now straightforward.
        alignedOriginalHeaders2ForRender = headers2;

        // 5. Map rows by key and perform comparison.
        const keyIndices1 = columnPairs.map(p => p.file1ColIndex);
        const keyIndices2 = columnPairs.map(p => p.file2ColIndex);

        const file1Map = new Map();
        data1.rows.forEach((row, rowIndex) => {
            const originalKey = keyIndices1.map(i => row[i]).join('||');
            // The key for file 1's map should always be its original key.
            // The mapping is applied to file 2's keys to match them to file 1's.
            if (!file1Map.has(originalKey)) {
                file1Map.set(originalKey, { rows: [] });
            }
            // We need to find the corresponding row in the *newly created* alignedData1
            const alignedRow = alignedData1.rows[rowIndex];
            file1Map.get(originalKey).rows.push(alignedRow);
        });

        const file2Map = new Map();
        data2.rows.forEach((row, rowIndex) => {
            const originalKey = keyIndices2.map(i => row[i]).join('||');
            // Find the corresponding key from file 1 if it's a mapped value.
            const comparisonKey = [...keyValueMappings.entries()].find(([k, v]) => v === originalKey)?.[0] || originalKey;
            if (!file2Map.has(comparisonKey)) {
                file2Map.set(comparisonKey, { rows: [] });
            }
            const alignedRow = alignedData2.rows[rowIndex];
            file2Map.get(comparisonKey).rows.push(alignedRow);
        });

        const allComparisonKeys = new Set([...file1Map.keys(), ...file2Map.keys()]);
        const blankRow1 = Array(alignedData1.headers.length).fill(null);
        const blankRow2 = Array(alignedData2.headers.length).fill(null);

        // --- NEW: Identify numeric pairs for diff calculation ---
        const diffPairs = [];
        const allPairsForDiff = [...comparisonPairs];
        const keyHeaders2 = new Set(columnPairs.map(p => data2.headers[p.file2ColIndex]));
        const keyHeaders1 = new Set(columnPairs.map(p => data1.headers[p.file1ColIndex]));

        // Add implicit pair if it exists
        const nonKeyNumericHeaders1 = headers1.filter(h => (numericColumns.has(h) || columnTypes1.get(h) === 'number') && !keyHeaders1.has(h));
        const nonKeyNumericHeaders2 = headers2.filter(h => (numericColumns.has(h) || columnTypes2.get(h) === 'number') && !keyHeaders2.has(h));

        if (nonKeyNumericHeaders1.length === 1 && nonKeyNumericHeaders2.length === 1 && nonKeyNumericHeaders1[0] !== nonKeyNumericHeaders2[0]) {
            const h1 = nonKeyNumericHeaders1[0];
            const h2 = nonKeyNumericHeaders2[0];
            const isAlreadyPaired = comparisonPairs.some(p => data1.headers[p.file1ColIndex] === h1 && data2.headers[p.file2ColIndex] === h2);
            if (!isAlreadyPaired) {
                allPairsForDiff.push({
                    file1ColIndex: data1.headers.indexOf(h1),
                    file2ColIndex: data2.headers.indexOf(h2)
                });
            }
        }

        allComparisonKeys.forEach(comparisonKey => {
            const rows1 = file1Map.get(comparisonKey)?.rows;
            const rows2 = file2Map.get(comparisonKey)?.rows;
            const numRows = Math.max(rows1?.length || 0, rows2?.length || 0);

            for (let i = 0; i < numRows; i++) {
                comparisonResults.push({ row1: rows1?.[i] || blankRow1, row2: rows2?.[i] || blankRow2 });
                const result = comparisonResults[comparisonResults.length - 1];
                
                // --- MODIFIED: Initialize diffs object and populate it only for numeric pairs ---
                result.diffs = {};

                allPairsForDiff.forEach(pair => {
                    const h1 = data1.headers[pair.file1ColIndex];
                    const h2 = data2.headers[pair.file2ColIndex];
                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';

                    if ((type1 === 'number' || type2 === 'number') || (numericColumns.has(h1) && numericColumns.has(h2))) {
                        const idx1 = alignedData1.headers.indexOf(h1);
                        const idx2 = alignedOriginalHeaders2ForRender.indexOf(h2);
                        if (idx1 > -1 && idx2 > -1) {
                            const v1 = parseFloat(result.row1[idx1]) || 0;
                            const v2 = parseFloat(result.row2[idx2]) || 0;
                            const diffKey = `Diff (File 1: ${h1} vs File 2: ${h2})`;
                            result.diffs[diffKey] = parseFloat((v1 - v2).toFixed(10));
                        }
                    }
                });
            }
        });

        nonSummableColumnsForDisclaimer = []; // Not applicable for row-by-row
    }

    function performGroupedComparison(data1, data2) {
        // 1. Define which columns to process for each file (non-key, included columns)
        const keyHeaders1 = new Set(columnPairs.map(p => data1.headers[p.file1ColIndex]));
        const keyHeaders2 = new Set(columnPairs.map(p => data2.headers[p.file2ColIndex]));

        const processableCols1 = data1.headers.map((h, i) => ({ h, i })).filter(col => !keyHeaders1.has(col.h) && !unmatchedHeaders1.has(col.h));
        const processableCols2 = data2.headers.map((h, i) => ({ h, i })).filter(col => !keyHeaders2.has(col.h) && !unmatchedHeaders2.has(col.h));

        // 2. Map rows by key for each file
        const keyIndices1 = columnPairs.map(p => p.file1ColIndex);
        const file1Map = mapDataByKey(data1, keyIndices1, processableCols1, false);

        const keyIndices2 = columnPairs.map(p => p.file2ColIndex);
        const file2Map = mapDataByKey(data2, keyIndices2, processableCols2, true);

        // 3. Merge the results
        const allComparisonKeys = new Set([...file1Map.keys(), ...file2Map.keys()]);
        const allNonKeyHeaders = new Set([...processableCols1.map(c => c.h), ...processableCols2.map(c => c.h)]);
        nonSummableColumnsForDisclaimer = [];

        allComparisonKeys.forEach(comparisonKey => {
            const group1 = file1Map.get(comparisonKey);
            const group2 = file2Map.get(comparisonKey);

            const displayKey = primaryKeySource === 1
                ? (group1?.originalKey || comparisonKey)
                : (group2?.originalKey || comparisonKey);
            
            // --- MODIFIED: Initialize diffs object and populate it only for numeric pairs ---
            const diffs = {};
            const file1Sum = group1 ? group1.processedData : {};
            const file2Sum = group2 ? group2.processedData : {};

            // Calculate diffs for numeric comparison pairs
            comparisonPairs.forEach(pair => {
                const h1 = data1.headers[pair.file1ColIndex];
                const h2 = data2.headers[pair.file2ColIndex];
                const type1 = columnTypes1.get(h1) || 'auto';
                const type2 = columnTypes2.get(h2) || 'auto';

                // Check if at least one of the columns was summable (is a number) or if both are null.
                // This ensures a diff is calculated even if one side has no data for the key.
                if (typeof file1Sum[h1] === 'number' || typeof file2Sum[h2] === 'number' || (file1Sum[h1] === null && file2Sum[h2] === null)) {
                    const v1 = parseFloat(file1Sum[h1]) || 0; // Treat null/non-numeric as 0
                    const v2 = parseFloat(file2Sum[h2]) || 0; // Treat null/non-numeric as 0
                    const diffKey = `Diff (File 1: ${h1} vs File 2: ${h2})`;
                    diffs[diffKey] = parseFloat((v1 - v2).toFixed(10));
                }
            });

            if (group1) nonSummableColumnsForDisclaimer.push(...group1.nonSummable);
            if (group2) nonSummableColumnsForDisclaimer.push(...group2.nonSummable);

            comparisonResults.push({
                key: displayKey,
                file1Sum,
                file2Sum,
                diffs
            });
        });

        nonSummableColumnsForDisclaimer = [...new Set(nonSummableColumnsForDisclaimer)];
    }

    function mapDataByKey(data, keyIndices, processableCols, isSourceFile2) {
        const map = new Map();
        const numericRegex = /^-?\d*\.?\d+$/;
        const typeMap = isSourceFile2 ? columnTypes2 : columnTypes1;

        data.rows.forEach(row => {
            const originalKey = keyIndices.map(i => row[i]).join('||');
            // When mapping data, the comparison key should be derived from the mapping if it's file 2.
            const comparisonKey = isSourceFile2 ? ([...keyValueMappings.entries()].find(([k, v]) => v === originalKey)?.[0] || originalKey) : originalKey;
            
            if (!map.has(comparisonKey)) {
                // --- BUG FIX ---
                // When processing file 2, the originalKey for the group should be the one from the mapping,
                // not the one from file 2, so that the display key is correct in the results.
                map.set(comparisonKey, { originalKey: isSourceFile2 ? comparisonKey : originalKey, rows: [], nonSummable: new Set() });
            }
            map.get(comparisonKey).rows.push(row);
        });

        // Process each group
        for (const [key, group] of map.entries()) {
            const processedData = {};
            const nonSummableInGroup = new Set();

            processableCols.forEach(col => {
                // Trim values before processing to correctly identify numeric content.
                const values = group.rows.map(r => (r[col.i] ?? '').toString().trim()).filter(v => v !== '');

                if (values.length === 0) {
                    processedData[col.h] = null;
                    return;
                }

                const columnType = typeMap.get(col.h) || 'auto';
                const allValuesAreNumeric = values.every(v => numericRegex.test(v));

                // Determine if we should sum this column for this group.
                // We sum if:
                // 1. The user explicitly set the type to 'number' AND all values are numeric.
                // 2. The type is 'auto' AND all values are numeric.
                const shouldSum = (columnType === 'number' || columnType === 'auto') && allValuesAreNumeric;

                if (shouldSum) {
                    // Sum the parsed float values.
                    const sum = values.reduce((s, v) => s + parseFloat(v), 0);
                    // Round to a safe number of decimal places to avoid floating point inaccuracies.
                    processedData[col.h] = parseFloat(sum.toFixed(10));
                } else {
                    // This column is either explicitly 'text' or contains non-numeric data.
                    const uniqueValues = new Set(values);
                    if (uniqueValues.size === 1) {
                        const singleValue = uniqueValues.values().next().value;
                        // If the type is 'number' but it couldn't be summed (e.g., one value was text),
                        // we still try to treat the single value as a number for comparison.
                        processedData[col.h] = (columnType === 'number' && numericRegex.test(singleValue)) ? parseFloat(singleValue) : singleValue;
                    } else {
                        // Display a semicolon-separated list of unique values up to a limit.
                        const UNIQUE_VALUES_LIMIT = 5;
                        const uniqueValuesArray = Array.from(uniqueValues);
                        if (uniqueValuesArray.length > UNIQUE_VALUES_LIMIT) {
                            processedData[col.h] = uniqueValuesArray.slice(0, UNIQUE_VALUES_LIMIT).join(';') + ', ...';
                        } else {
                            processedData[col.h] = uniqueValuesArray.join(';');
                        }
                    }
                    nonSummableInGroup.add(col.h);
                }
            });
            group.processedData = processedData;
            group.nonSummable = Array.from(nonSummableInGroup);
        }
        return map;
    }
    
    function sortGroupedResults(columnIndex, header, file, isNumeric, diffColCount, sortedHeaders1, sortedHeaders2) {
        if (sortState.columnIndex === columnIndex) {
            sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
        } else {
            sortState.columnIndex = columnIndex;
            sortState.direction = 'asc';
        }

        const direction = sortState.direction === 'asc' ? 1 : -1;

        const keyColCount = 1;
        const file1HeaderCount = sortedHeaders1.length;
        const file2HeaderCount = sortedHeaders2.length;
        const totalHeadersBeforeDiff = keyColCount + file1HeaderCount + file2HeaderCount;
    
        comparisonResults.sort((a, b) => {
            let valA, valB;
            if (columnIndex === 0) { // Key column
                valA = a.key;
                valB = b.key;
            } else if (file === '1') { // File 1 columns
                valA = a.file1Sum[header] ?? (isNumeric ? -Infinity : '');
                valB = b.file1Sum[header] ?? (isNumeric ? -Infinity : '');
                if (isNumeric) return (valA - valB) * direction;
            } else if (file === '2') { // File 2 columns
                valA = a.file2Sum[header] ?? (isNumeric ? -Infinity : '');
                valB = b.file2Sum[header] ?? (isNumeric ? -Infinity : '');
                if (isNumeric) return (valA - valB) * direction;
            } else { // Diff columns
                valA = a.diffs[header] ?? -Infinity;
                valB = b.diffs[header] ?? -Infinity;
                return (valA - valB) * direction;
            }
            // Use localeCompare for keys and text values
            return valA.toString().localeCompare(valB.toString(), undefined, { numeric: true }) * direction;
        });
    }

    function updateSortIndicators(headerRow, activeIndex) {
        headerRow.querySelectorAll('th').forEach((th, index) => {
            // Use a more robust way to handle text content to avoid removing parts of header names
            const indicatorSpan = th.querySelector('.sort-indicator');
            if (indicatorSpan) indicatorSpan.remove();
            if (index === activeIndex) {
                th.innerHTML += `<span class="sort-indicator"> ${sortState.direction === 'asc' ? '↑' : '↓'}</span>`;
            }
        });
    }
    
    function setDownloadButtonLoading(isLoading) {
        const downloadBtn = document.getElementById('download-results-btn');
        const btnLoader = downloadBtn.querySelector('.btn-loader');
        const dropdownArrow = downloadBtn.querySelector('.dropdown-arrow');

        downloadBtn.disabled = isLoading;
        if (isLoading) {
            downloadBtn.classList.add('loading');
            btnLoader?.classList.remove('hidden');
            dropdownArrow?.classList.add('hidden');
        } else {
            downloadBtn.classList.remove('loading');
            btnLoader?.classList.add('hidden');
            dropdownArrow?.classList.remove('hidden');
        }
    }

    function downloadResultsAsCSV(e) {
        e.preventDefault();
        downloadDropdown.classList.add('hidden');
        if (comparisonResults.length === 0) return;
        const groupBy = groupByCheckbox.checked;
        const showDiffOnly = showDiffOnlyCheckbox.checked;
        const hideUnmatched = hideUnmatchedKeysCheckbox.checked;

        // Apply the same filters as the UI to the data being downloaded
        let resultsToDownload = [...comparisonResults];

        if (showDiffOnly) {
            if (groupBy) {
                resultsToDownload = resultsToDownload.filter(result => {
                    const allHeaders = new Set([...Object.keys(result.file1Sum), ...Object.keys(result.file2Sum)]);
                    return [...allHeaders].some(h => {
                        const pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData1.headers[p.file1ColIndex] === h || fullFileData2.headers[p.file2ColIndex] === h);
                        const h1 = pair ? fullFileData1.headers[pair.file1ColIndex] : h;
                        const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h;
                        const type1 = columnTypes1.get(h1) || 'auto';
                        const type2 = columnTypes2.get(h2) || 'auto';
                        return areValuesDifferent(result.file1Sum[h1], result.file2Sum[h2], h1, h2, type1, type2);
                    });
                });
            } else {
                resultsToDownload = resultsToDownload.filter(result => result.row1.some((cell, index) => {
                    const h1 = alignedData1.headers[index];
                    const pair = [...columnPairs, ...comparisonPairs].find(p => p && fullFileData1.headers[p.file1ColIndex] === h1);
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h1;
                    const index2 = alignedOriginalHeaders2ForRender.indexOf(h2);
                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';
                    return index2 !== -1 && areValuesDifferent(cell, result.row2[index2], h1, h2, type1, type2);
                }));
            }
        }

        if (hideUnmatched) {
            resultsToDownload = resultsToDownload.filter(result =>
                (groupBy && Object.keys(result.file1Sum).length > 0 && Object.keys(result.file2Sum).length > 0) ||
                (!groupBy && !result.row1.every(c => c === null) && !result.row2.every(c => c === null))
            );
        }
        let csvContent = '';

        setDownloadButtonLoading(true);

        if (groupBy) {
            const allHeaders1 = new Set();
            const allHeaders2 = new Set();
            resultsToDownload.forEach(result => {
                Object.keys(result.file1Sum).forEach(h => allHeaders1.add(h));
                Object.keys(result.file2Sum).forEach(h => allHeaders2.add(h));
            });
            const sortedHeaders1 = [...allHeaders1].sort();
            const sortedHeaders2 = [...allHeaders2].sort();
            
            const diffHeaders = showDiffColumnsCheckbox.checked && resultsToDownload.length > 0 ? Object.keys(resultsToDownload[0].diffs || {}) : [];
            const headers = ['Difference', 'Key', ...sortedHeaders1.map(h => `File1_${h}`), ...sortedHeaders2.map(h => `File2_${h}`), ...diffHeaders];

            const rows = comparisonResults.map(result => {
                const key = `"${result.key.replace(/"/g, '""')}"`;
                const sums1 = sortedHeaders1.map(h => `"${(result.file1Sum[h] ?? '').toString().replace(/"/g, '""')}"`);
                const sums2 = sortedHeaders2.map(h => `"${(result.file2Sum[h] ?? '').toString().replace(/"/g, '""')}"`);
                const diffValues = diffHeaders.map(h => `"${(result.diffs[h] ?? '').toString().replace(/"/g, '""')}"`);
                
                const hasDifference = [...new Set([...sortedHeaders1, ...sortedHeaders2])].some(h => areValuesDifferent(result.file1Sum[h], result.file2Sum[h], h, h));
                const differenceFlag = hasDifference ? 'Yes' : 'No';
                return [differenceFlag, key, ...sums1, ...sums2, ...diffValues].join(',');
            });

            csvContent = [headers.join(','), ...rows].join('\n');
        } else {
            // Get headers of columns to INCLUDE in the results
            const headers1 = alignedData1.headers.map(h => `File1_${h.replace(/"/g, '""')}`);
            const headers2 = alignedOriginalHeaders2ForRender.map(h => `File2_${(h || 'N/A').replace(/"/g, '""')}`);
            const diffHeaders = showDiffColumnsCheckbox.checked && resultsToDownload.length > 0 ? Object.keys(resultsToDownload[0].diffs || {}) : [];
            const csvHeaders = ['Difference', ...headers1, ...headers2, ...diffHeaders].join(',');

            const csvRows = resultsToDownload.map(result => {
                const hasDifference = result.row1.some((cell, index) => {
                    const h1 = alignedData1.headers[index];
                    const pair = [...columnPairs, ...comparisonPairs].find(p => p && fullFileData1.headers[p.file1ColIndex] === h1);
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h1;
                    const index2 = alignedOriginalHeaders2ForRender.indexOf(h2);
                    
                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';
                    if (index2 !== -1 && areValuesDifferent(cell, result.row2[index2], h1, h2, type1, type2)) {
                        return true;
                    }
                    return false;
                });
            
                const differenceFlag = hasDifference ? 'Yes' : 'No';
                const csvRow1 = result.row1.map(cell => `"${(cell || '').toString().replace(/"/g, '""')}"`);
                const csvRow2 = result.row2.map(cell => `"${(cell || '').toString().replace(/"/g, '""')}"`);
                const diffValues = diffHeaders.map(h => `"${(result.diffs[h] ?? '').toString().replace(/"/g, '""')}"`);
                return [differenceFlag, ...csvRow1, ...csvRow2, ...diffValues].join(',');
            });

            csvContent = [csvHeaders, ...csvRows].join('\n');
        }

       
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });

        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', 'comparison_results.csv');
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        setDownloadButtonLoading(false);
    }

    async function downloadResultsAsExcel(e) {
        e.preventDefault();
        // Prevent download if the action is disabled
        if (e.currentTarget.classList.contains('disabled')) {
            showToast("Cannot download as Excel. The number of result rows exceeds Excel's limit.", 'error');
            downloadDropdown.classList.add('hidden');
            return;
        }
        downloadDropdown.classList.add('hidden');
        if (comparisonResults.length === 0) return;

        setDownloadButtonLoading(true);

        const groupBy = groupByCheckbox.checked;

        const showDiffOnly = showDiffOnlyCheckbox.checked;
        const hideUnmatched = hideUnmatchedKeysCheckbox.checked;

        // Apply the same filters as the UI to the data being downloaded
        let resultsToDownload = [...comparisonResults];

        if (showDiffOnly) {
            if (groupBy) {
                resultsToDownload = resultsToDownload.filter(result => {
                    const allHeaders = new Set([...Object.keys(result.file1Sum), ...Object.keys(result.file2Sum)]);
                    return [...allHeaders].some(h => {
                        const pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData1.headers[p.file1ColIndex] === h || fullFileData2.headers[p.file2ColIndex] === h);
                        const h1 = pair ? fullFileData1.headers[pair.file1ColIndex] : h;
                        const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h;
                        const type1 = columnTypes1.get(h1) || 'auto';
                        const type2 = columnTypes2.get(h2) || 'auto';
                        return areValuesDifferent(result.file1Sum[h1], result.file2Sum[h2], h1, h2, type1, type2);
                    });
                });
            } else {
                resultsToDownload = resultsToDownload.filter(result => result.row1.some((cell, index) => {
                    const h1 = alignedData1.headers[index];
                    const pair = [...columnPairs, ...comparisonPairs].find(p => p && fullFileData1.headers[p.file1ColIndex] === h1);
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h1;
                    const index2 = alignedOriginalHeaders2ForRender.indexOf(h2);
                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';
                    return index2 !== -1 && areValuesDifferent(cell, result.row2[index2], h1, h2, type1, type2);
                }));
            }
        }

        if (hideUnmatched) {
            resultsToDownload = resultsToDownload.filter(result =>
                (groupBy && Object.keys(result.file1Sum).length > 0 && Object.keys(result.file2Sum).length > 0) ||
                (!groupBy && !result.row1.every(c => c === null) && !result.row2.every(c => c === null))
            );
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Comparison Results');
        const diffStyle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } }; // Light yellow

        if (groupBy) {
            // --- Grouped View Logic ---
            const allHeaders1 = new Set();
            const allHeaders2 = new Set();
            resultsToDownload.forEach(result => {
                Object.keys(result.file1Sum).forEach(h => allHeaders1.add(h));
                Object.keys(result.file2Sum).forEach(h => allHeaders2.add(h));
            });
            const sortedHeaders1 = [...allHeaders1].sort();
            const sortedHeaders2 = [...allHeaders2].sort();

            // --- Implicit Pairing Logic for Excel Download ---
            let implicitPair = null;
            if (sortedHeaders1.length === 1 && sortedHeaders2.length === 1 && sortedHeaders1[0] !== sortedHeaders2[0]) {
                const h1 = sortedHeaders1[0];
                const h2 = sortedHeaders2[0];
                const isAlreadyPaired = [...columnPairs, ...comparisonPairs].some(p => {
                    const pairH1 = fullFileData1.headers[p.file1ColIndex];
                    const pairH2 = fullFileData2.headers[p.file2ColIndex];
                    return (pairH1 === h1 && pairH2 === h2) || (pairH1 === h2 && pairH2 === h1);
                });

                if (!isAlreadyPaired) {
                    implicitPair = { file1Header: h1, file2Header: h2 };
                }
            }

            const diffHeaders = showDiffColumnsCheckbox.checked && resultsToDownload.length > 0 ? Object.keys(resultsToDownload[0].diffs || {}) : [];
            worksheet.columns = [
                { header: 'Key', key: 'key', width: 30 },
                ...sortedHeaders1.map(h => ({ header: `File 1: ${h}`, key: `f1_${h}`, width: 20 })),
                ...sortedHeaders2.map(h => ({ header: `File 2: ${h}`, key: `f2_${h}`, width: 20 })),
                ...diffHeaders.map(h => ({ header: h, key: h, width: 20 }))
            ];

            resultsToDownload.forEach(result => {
                const rowData = { key: result.key };
                sortedHeaders1.forEach(h => { rowData[`f1_${h}`] = result.file1Sum[h] ?? null; });
                sortedHeaders2.forEach(h => { rowData[`f2_${h}`] = result.file2Sum[h] ?? null; });
                if (diffHeaders.length > 0) {
                    diffHeaders.forEach(h => { rowData[h] = result.diffs[h] ?? null; });
                }
                worksheet.addRow(rowData);
            });

            // Apply styles
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                if (rowNumber === 1) return; // Skip header
                const result = resultsToDownload[rowNumber - 2];
                if (!result) return;

                // Highlight File 1 columns that are different or standalone
                sortedHeaders1.forEach(h => {
                    let pair = [...columnPairs, ...comparisonPairs].find(pair => fullFileData1.headers[pair.file1ColIndex] === h);
                    if (!pair && implicitPair && implicitPair.file1Header === h) {
                        pair = { file2ColIndex: fullFileData2.headers.indexOf(implicitPair.file2Header) };
                    }
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h;

                    const val1 = result.file1Sum[h] ?? null;
                    const val2 = result.file2Sum[h2] ?? null;
                    
                    const type1 = columnTypes1.get(h) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';
                    if (areValuesDifferent(val1, val2, h, h2, type1, type2)) {
                        row.getCell(`f1_${h}`).fill = diffStyle;
                    }
                });

                // Highlight File 2 columns that are different or standalone
                sortedHeaders2.forEach(h => {
                    let pair = [...columnPairs, ...comparisonPairs].find(pair => fullFileData2.headers[pair.file2ColIndex] === h);
                    if (!pair && implicitPair && implicitPair.file2Header === h) {
                        pair = { file1ColIndex: fullFileData1.headers.indexOf(implicitPair.file1Header) };
                    }
                    const h1 = pair ? fullFileData1.headers[pair.file1ColIndex] : h;

                    const val1 = result.file1Sum[h1] ?? null;
                    const val2 = result.file2Sum[h] ?? null;
                    
                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h) || 'auto';
                    if (areValuesDifferent(val1, val2, h1, h, type1, type2)) {
                        row.getCell(`f2_${h}`).fill = diffStyle;
                    }
                });
            });

        } else {
            // --- Row-by-Row View Logic ---
            const file1ColCount = alignedData1.headers.length;
            const diffHeaders = showDiffColumnsCheckbox.checked && resultsToDownload.length > 0 ? Object.keys(resultsToDownload[0].diffs || {}) : [];
            const columns = [
                ...alignedData1.headers.map((h, i) => ({ header: `File 1: ${h || 'N/A'}`, key: `f1_${i}`, width: 20 })),
                ...alignedOriginalHeaders2ForRender.map((h, i) => ({ header: `File 2: ${h || 'N/A'}`, key: `f2_${i}`, width: 20 })),
                ...diffHeaders.map(h => ({ header: h, key: h, width: 20 }))
            ];
            worksheet.columns = columns;

            // --- Implicit Pairing Logic for Row-by-Row Excel Download ---
            let implicitPair = null;
            const nonKeyNumericHeaders1 = alignedData1.headers.filter(h => numericColumns.has(h) && !columnPairs.some(p => fullFileData1.headers[p.file1ColIndex] === h));
            const nonKeyNumericHeaders2 = alignedOriginalHeaders2ForRender.filter(h => numericColumns.has(h) && !columnPairs.some(p => fullFileData2.headers[p.file2ColIndex] === h));

            if (nonKeyNumericHeaders1.length === 1 && nonKeyNumericHeaders2.length === 1 && nonKeyNumericHeaders1[0] !== nonKeyNumericHeaders2[0]) {
                const h1 = nonKeyNumericHeaders1[0];
                const h2 = nonKeyNumericHeaders2[0];
                const isAlreadyPaired = comparisonPairs.some(p => fullFileData1.headers[p.file1ColIndex] === h1 && fullFileData2.headers[p.file2ColIndex] === h2);
                if (!isAlreadyPaired) {
                    implicitPair = { file1Header: h1, file2Header: h2 };
                }
            }

            resultsToDownload.forEach(result => {
                const rowData = {};

                // Populate rowData with formatted values
                alignedData1.headers.forEach((h1, i) => {
                    let pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData1.headers[p.file1ColIndex] === h1);
                    if (!pair && implicitPair && implicitPair.file1Header === h1) {
                        pair = { file2ColIndex: fullFileData2.headers.indexOf(implicitPair.file2Header) };
                    }
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h1;
                    const index2 = alignedOriginalHeaders2ForRender.indexOf(h2);

                    const val1 = result.row1[i];
                    const val2 = result.row2[index2];

                    const type1 = columnTypes1.get(h1) || 'auto';
                    const type2 = columnTypes2.get(h2) || 'auto';

                    const formatted = formatNumericValues(val1, val2, h1, h2, type1, type2);
                    rowData[`f1_${i}`] = formatted.f1;
                    if (index2 !== -1) {
                        rowData[`f2_${index2}`] = formatted.f2; // This was missing
                    }
                });

                if (diffHeaders.length > 0) {
                    diffHeaders.forEach(h => { rowData[h] = result.diffs[h] ?? null; });
                }
                worksheet.addRow(rowData);
            });

            // Apply styles
            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                if (rowNumber === 1) return; // Skip header
                const result = resultsToDownload[rowNumber - 2];
                if (!result) return;

                // Check for differences from File 1's perspective
                for (let i = 0; i < file1ColCount; i++) {
                    const h1 = alignedData1.headers[i];
                    let pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData1.headers[p.file1ColIndex] === h1);
                    if (!pair && implicitPair && implicitPair.file1Header === h1) {
                        pair = { file2ColIndex: fullFileData2.headers.indexOf(implicitPair.file2Header) };
                    }
                    const h2 = pair ? fullFileData2.headers[pair.file2ColIndex] : h1;
                    const index2 = alignedOriginalHeaders2ForRender.indexOf(h2);

                    if (index2 === -1) { // Column is standalone in File 1
                        row.getCell(`f1_${i}`).fill = diffStyle;
                    } else {
                        // Column is comparable, compare values
                        const v1 = result.row1[i];
                        const v2 = result.row2[index2];
                        
                        const type1 = columnTypes1.get(h1) || 'auto';
                        const type2 = columnTypes2.get(h2) || 'auto';
                        if (areValuesDifferent(v1, v2, h1, h2, type1, type2)) {
                            row.getCell(`f1_${i}`).fill = diffStyle;
                            row.getCell(`f2_${index2}`).fill = diffStyle;
                        }
                    }
                }

                // Check for differences from File 2's perspective (for standalone columns in File 2)
                for (let i = 0; i < alignedOriginalHeaders2ForRender.length; i++) {
                    const h2 = alignedOriginalHeaders2ForRender[i];
                    let pair = [...columnPairs, ...comparisonPairs].find(p => fullFileData2.headers[p.file2ColIndex] === h2);
                    if (!pair && implicitPair && implicitPair.file2Header === h2) {
                        pair = { file1ColIndex: fullFileData1.headers.indexOf(implicitPair.file1Header) };
                    }
                    const h1 = pair ? fullFileData1.headers[pair.file1ColIndex] : h2;
                    if (alignedData1.headers.indexOf(h1) === -1) { // Column is standalone in File 2
                        row.getCell(`f2_${i}`).fill = diffStyle;
                    }
                }
            });
        }

        // Make the header row bold
        worksheet.getRow(1).font = { bold: true };

        // Freeze the header row so it's always visible
        worksheet.views = [{ state: 'frozen', ySplit: 1 }];

        // Write to buffer and trigger download
        try {
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'comparison_results.xlsx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (error) {
            console.error('Error generating Excel file:', error);
            showToast('Failed to generate Excel file.', 'error');
        } finally {
            setDownloadButtonLoading(false);
        }
    }

    // --- Value Mapping Modal Logic ---

    function openValueMappingModal(skipAutoPairing = false, isRefresh = false) {
        if (columnPairs.length === 0) {
            alert("Please define at least one key column pair before mapping values.");
            return; 
        }
        // Ensure both files are loaded before proceeding.
        isDragDisabled = true;
        if (!fullFileData1 || !fullFileData2) {
            console.error("Value mapping modal opened without both files being loaded.");
            showToast("Please load both files before mapping values.", "error");
            return;
        }

        // Check if the number of unique keys exceeds the practical limit for the modal UI.
        // This check does not affect the comparison, only the modal's accessibility.
        if (!isRefresh) { // Only check when initially opening, not on internal refreshes.
            const MAX_MODAL_ITEMS = 1000;
            const keyIndices1 = columnPairs.map(p => p.file1ColIndex);
            const keyIndices2 = columnPairs.map(p => p.file2ColIndex);
            const uniqueValues1Count = new Set(fullFileData1.rows.map(row => keyIndices1.map(i => row[i]).join('||'))).size;
            const uniqueValues2Count = new Set(fullFileData2.rows.map(row => keyIndices2.map(i => row[i]).join('||'))).size;

            if (uniqueValues1Count > MAX_MODAL_ITEMS || uniqueValues2Count > MAX_MODAL_ITEMS) {
                showToast(`The number of unique key values (${Math.max(uniqueValues1Count, uniqueValues2Count).toLocaleString()}) exceeds the recommended limit of ${MAX_MODAL_ITEMS.toLocaleString()} for the interactive mapping screen. The comparison will still work with exact matches.`, 'warning', 10000);
                return; // Prevent the modal from opening.
            }
        }


    
        const loader = valueMappingModal.querySelector('.modal-loader');
        const content = valueMappingModal.querySelector('.value-mapping-container');
        const list1Container = document.getElementById('value-list-1');
        const list2Container = document.getElementById('value-list-2');
        const mappedListContainer = document.getElementById('mapped-values-list');
    
        // Show modal and loader immediately
        if (!isRefresh) {
            valueMappingModal.classList.remove('hidden');
        }
        loader.classList.remove('hidden');
        content.style.visibility = 'hidden'; // Keep content hidden while loading
    
        // Defer the heavy computation to allow the UI to update
        setTimeout(() => {
            // 1. Get unique key values from both files
            const keyIndices1 = columnPairs.map(p => p.file1ColIndex);
            const keyIndices2 = columnPairs.map(p => p.file2ColIndex);
    
            const uniqueValues1 = new Set(fullFileData1.rows.map(row => keyIndices1.map(i => row[i]).join('||')));
            const uniqueValues2 = new Set(fullFileData2.rows.map(row => keyIndices2.map(i => row[i]).join('||')));
    
            // 2. Populate the lists
            mappedListContainer.innerHTML = ''; 
            list1Container.innerHTML = '';
            list2Container.innerHTML = '';
    
            // Auto-pair identical values and initialize mappings if not skipping
            // If skipAutoPairing is false, AND there are no pre-existing mappings (from a template), then run auto-pairing.
            if (!skipAutoPairing && !hasClearedMappings && keyValueMappings.size === 0) {
                keyValueMappings.clear(); // Clear previous mappings to apply fresh defaults
                suggestedKeyValueMappings.clear(); // Clear previous suggestions
                const availableValues2 = new Set(uniqueValues2); 
    
                // Pass 1: Exact matches
                for (const val1 of uniqueValues1) {
                    if (availableValues2.has(val1)) {
                        keyValueMappings.set(val1, val1);
                        availableValues2.delete(val1); // Remove from available pool
                    }
                }
    
                // Pass 2: Case-insensitive matches for remaining values
                let unmappedValues1 = [...uniqueValues1].filter(v => !keyValueMappings.has(v));
                for (const val1 of unmappedValues1) {
                    // Find a value in the available pool that matches when both are lowercased
                    const match = [...availableValues2].find(val2 => val1.toLowerCase() === val2.toLowerCase());
                    if (match) {
                        keyValueMappings.set(val1, match);
                        suggestedKeyValueMappings.add(val1); // Mark as a suggestion
                        availableValues2.delete(match); // Remove from the pool
                    }
                }

                // Pass 3: Fuzzy (Levenshtein) matches for the rest
                const SIMILARITY_THRESHOLD = 0.8; // e.g., 80% similar
                // Get the remaining unmapped values after the case-insensitive pass
                const stillUnmappedValues1 = [...uniqueValues1].filter(v => !keyValueMappings.has(v));
    
                for (const val1 of stillUnmappedValues1) {
                    let bestMatch = null;
                    let highestSimilarity = -1;
                    
                    for (const val2 of availableValues2) {
                        // Perform a case-insensitive Levenshtein comparison for better results
                        const similarity = levenshteinDistance(val1.toLowerCase(), val2.toLowerCase());
                        if (similarity > highestSimilarity) {
                            highestSimilarity = similarity;
                            bestMatch = val2;
                        }
                    }
                    if (bestMatch && highestSimilarity >= SIMILARITY_THRESHOLD) {
                        keyValueMappings.set(val1, bestMatch);
                        suggestedKeyValueMappings.add(val1);
                        availableValues2.delete(bestMatch);
                    }
                }
            }
    
            // Create reverse map for quick lookup
            const reverseMappings = new Map([...keyValueMappings.entries()].map(([k, v]) => [v, k]));
    
            // Render Mapped List
            const sortedMappedKeys = [...keyValueMappings.keys()].sort();
            sortedMappedKeys.forEach(val1 => {
                const val2 = keyValueMappings.get(val1);
                const pairDiv = document.createElement('div');
                pairDiv.className = 'mapped-item-pair';

                const item1 = document.createElement('div');
                item1.className = 'value-item';
                item1.textContent = val1;
                item1.draggable = true;
                item1.dataset.value = val1;
                item1.dataset.fileIndex = '1';

                const item2 = document.createElement('div');
                item2.className = 'value-item';
                item2.textContent = val2;
                item2.draggable = false; // Only drag from file 1 side of mapped pair
                item2.dataset.value = val2;
                item2.dataset.fileIndex = '2';

                if (suggestedKeyValueMappings.has(val1)) {
                    item1.classList.add('suggested');
                    item2.classList.add('suggested');
                } else {
                    item1.classList.add('mapped');
                    item2.classList.add('mapped');
                }

                pairDiv.appendChild(item1);
                pairDiv.innerHTML += '<span class="arrow">↔</span>';
                pairDiv.appendChild(item2);
                mappedListContainer.appendChild(pairDiv);
            });

            // Render Unmapped Lists
            [...uniqueValues1].sort().forEach(value => {
                if (!keyValueMappings.has(value)) {
                    const item = document.createElement('div');
                    item.className = 'value-item';
                    item.textContent = value;
                    item.draggable = true;
                    item.dataset.value = value;
                    item.dataset.fileIndex = '1';
                    list1Container.appendChild(item);
                }
            });
    
            [...uniqueValues2].sort().forEach(value => {
                if (!reverseMappings.has(value)) {
                    const item = document.createElement('div');
                    item.className = 'value-item';
                    item.textContent = value;
                    item.draggable = true;
                    item.dataset.value = value;
                    item.dataset.fileIndex = '2';
                    list2Container.appendChild(item);
                    }
            });
    
            // 3. Add drag-and-drop event listeners
            addValueMappingDragDropListeners();
    
            // 4. Hide loader and show content
            loader.classList.add('hidden');
            content.style.visibility = 'visible';
    
            // 5. Align items and add listeners
            const search1 = document.getElementById('value-search-1');
            const search2 = document.getElementById('value-search-2');
            const newSearch1 = search1.cloneNode(true);
            const newSearch2 = search2.cloneNode(true);
            // Clear old listeners and add new ones to prevent duplicates
            search1.parentNode.replaceChild(newSearch1, search1);
            search2.parentNode.replaceChild(newSearch2, search2);
            newSearch1.addEventListener('input', () => filterValueList());
            newSearch2.addEventListener('input', () => filterValueList());

        }, 10); // A small timeout is enough to free the main thread
    }

    function addValueMappingDragDropListeners() {
        const items = document.querySelectorAll('.unmapped-values-container .value-item, .mapped-values-wrapper .value-item');
        let draggedItem = null;

        items.forEach(item => {
            item.addEventListener('dragstart', (e) => {
                draggedItem = e.target;
                e.dataTransfer.effectAllowed = 'move';
            });

            item.addEventListener('dragover', (e) => {
                e.preventDefault(); // Necessary to allow dropping
            });

            item.addEventListener('drop', (e) => {
                e.preventDefault();
                const droppedOnItem = e.target;
                if (!draggedItem || draggedItem === droppedOnItem) return;

                // Scenario 1: Dragging from unmapped to unmapped
                const isUnmappedDrag = !draggedItem.closest('.mapped-item-pair');
                if (isUnmappedDrag && draggedItem.dataset.fileIndex !== droppedOnItem.dataset.fileIndex) {
                     const val1 = draggedItem.dataset.fileIndex === '1' ? draggedItem.dataset.value : droppedOnItem.dataset.value;
                     const val2 = draggedItem.dataset.fileIndex === '2' ? draggedItem.dataset.value : droppedOnItem.dataset.value;
                     
                     // A simple new mapping
                     keyValueMappings.set(val1, val2);
                     suggestedKeyValueMappings.delete(val1); // Manual action
                     hasClearedMappings = false; // A new mapping was made
                     
                     openValueMappingModal(true, true); // Refresh
                     resultsSection.classList.add('hidden');
                }

                // Scenario 2: Dragging from a MAPPED pair to an UNMAPPED item
                const isMappedDrag = draggedItem.closest('.mapped-item-pair');
                const isUnmappedDrop = !droppedOnItem.closest('.mapped-item-pair');
                if (isMappedDrag && isUnmappedDrop && draggedItem.dataset.fileIndex !== droppedOnItem.dataset.fileIndex) {
                    const draggedVal = draggedItem.dataset.value; // This is always a file 1 value from a mapped pair
                    const droppedOnVal = droppedOnItem.dataset.value; // This is the new file 2 value

                    // The old file 2 value that was paired with `draggedVal` is now unmapped.
                    // We don't need to do anything special; it will just not be in the map anymore.

                    // Create the new mapping
                    keyValueMappings.set(draggedVal, droppedOnVal);
                    suggestedKeyValueMappings.delete(draggedVal); // Manual action
                    hasClearedMappings = false; // A new mapping was made

                    openValueMappingModal(true, true); // Refresh the modal
                    resultsSection.classList.add('hidden');
                }

                // Scenario 3: Dragging an UNMAPPED item onto a MAPPED item
                const isMappedDrop = droppedOnItem.closest('.mapped-item-pair');
                if (isUnmappedDrag && isMappedDrop && draggedItem.dataset.fileIndex !== droppedOnItem.dataset.fileIndex) {
                    const unmappedVal = draggedItem.dataset.value;
                    const mappedVal = droppedOnItem.dataset.value;

                    // Determine which is file 1 and which is file 2
                    const val1 = draggedItem.dataset.fileIndex === '1' ? unmappedVal : mappedVal;
                    const val2 = draggedItem.dataset.fileIndex === '2' ? unmappedVal : mappedVal;

                    // Find and break the old mapping for the value that was dropped on.
                    // We need to iterate through the map to find the key associated with the old value.
                    for (const [key, value] of keyValueMappings.entries()) {
                        if (value === mappedVal || key === mappedVal) {
                            keyValueMappings.delete(key);
                            break;
                        }
                    }
                    keyValueMappings.set(val1, val2);
                    hasClearedMappings = false; // A new mapping was made
                    openValueMappingModal(true, true); // Refresh
                    resultsSection.classList.add('hidden');
                }

                // Other scenarios (e.g., mapped to mapped) are ignored for now.
                draggedItem = null;
            });
        });
    }

    function filterValueList() {
        // This function now needs to handle filtering for both lists based on both search boxes.
        // Let's get the current values of both search boxes.
        const search1Val = document.getElementById('value-search-1').value.toLowerCase();
        const search2Val = document.getElementById('value-search-2').value.toLowerCase();

        // Filter unmapped list 1
        document.querySelectorAll('#value-list-1 .value-item').forEach(item => {
            const isMatch = item.textContent.toLowerCase().includes(search1Val);
            item.classList.toggle('hidden', !isMatch);
        });

        // Filter unmapped list 2
        document.querySelectorAll('#value-list-2 .value-item').forEach(item => {
            const isMatch = item.textContent.toLowerCase().includes(search2Val);
            item.classList.toggle('hidden', !isMatch);
        });

        // Filter the mapped list based on both search boxes
        const mappedList = document.getElementById('mapped-values-list');
        mappedList.querySelectorAll('.mapped-item-pair').forEach(pair => {
            const item1Text = pair.querySelector('.value-item[data-file-index="1"]').textContent.toLowerCase();
            const item2Text = pair.querySelector('.value-item[data-file-index="2"]').textContent.toLowerCase();

            // A pair is visible if its file 1 value matches the search 1 input AND its file 2 value matches the search 2 input.
            // An empty search box effectively matches all values for its respective column.
            const isVisible = item1Text.includes(search1Val) && item2Text.includes(search2Val);
            pair.classList.toggle('hidden', !isVisible);
        });
    }

    // --- Template Management ---

    function handleSaveTemplate() {
        // Open the modal instead of using a prompt

        // Pre-populate mappings if they haven't been generated yet to get an accurate count.
        // This ensures the count is correct even if the value mapping modal was never opened.
        if (keyValueMappings.size === 0 && fullFileData1 && fullFileData2 && columnPairs.length > 0) {
            const keyIndices1 = columnPairs.map(p => p.file1ColIndex);
            const keyIndices2 = columnPairs.map(p => p.file2ColIndex);
            const uniqueValues1 = new Set(fullFileData1.rows.map(row => keyIndices1.map(i => row[i]).join('||')));
            const uniqueValues2 = new Set(fullFileData2.rows.map(row => keyIndices2.map(i => row[i]).join('||')));

            // Just do a quick pass for exact matches to get a baseline count.
            // This avoids the expensive fuzzy matching unless the user opens the modal.
            for (const val1 of uniqueValues1) {
                if (uniqueValues2.has(val1)) {
                    keyValueMappings.set(val1, val1);
                }
            }
        }


        saveTemplateModal.classList.remove('hidden');

        const mappingsCheckbox = document.getElementById('save-mappings-checkbox');
        const mappingsWarning = document.getElementById('mappings-limit-warning');
        const MAX_MAPPINGS_TO_SAVE = 1000;
        const fixedWidthCheckbox = document.getElementById('save-fixedwidth-checkbox');
        const fixedWidthLabel = fixedWidthCheckbox.parentElement;

        if (keyValueMappings.size > MAX_MAPPINGS_TO_SAVE) {
            mappingsCheckbox.checked = false;
            mappingsCheckbox.disabled = true;
            mappingsWarning.classList.remove('hidden');
        } else {
            // Default to checked only if it's not disabled
            if (!mappingsCheckbox.disabled) {
                mappingsCheckbox.checked = true;
            }
            mappingsCheckbox.disabled = false;
            mappingsWarning.classList.add('hidden');
        }

        // Hide the fixed-width option if no fixed-width settings are present.
        if (!fixedWidthSettings1 && !fixedWidthSettings2) {
            fixedWidthCheckbox.checked = false;
            fixedWidthCheckbox.disabled = true;
            fixedWidthLabel.classList.add('hidden');
        } else {
            fixedWidthCheckbox.checked = true;
            fixedWidthCheckbox.disabled = false;
            fixedWidthLabel.classList.remove('hidden');
        }
    }

    function confirmSaveTemplate() {
        const nameInput = document.getElementById('template-name-input');
        const name = nameInput.value.trim();
        if (!name) {
            showToast("Please enter a name for the template.", "error");
            return;
        }

        // All configs are saved in the client-side version
        const configsToSave = getSelectedTemplateOptions();

        // Prepare column pairs using header names for robustness and portability
        const columnPairsByName = configsToSave.includes('columns') ? columnPairs.map(p => ({
            file1Header: p.file1Th.textContent.replace(/🔓|🔒/g, '').trim(), // Trim here for safety
            file2Header: p.file2Th.textContent.replace(/🔓|🔒/g, '').trim()  // Trim here for safety
        })) : [];

        // Prepare comparison pairs (right-clicked)
        const comparisonPairsByName = configsToSave.includes('columns') ? comparisonPairs.map(p => ({
            file1Header: p.file1Th.textContent.replace(/🔓|🔒/g, '').trim(),
            file2Header: p.file2Th.textContent.replace(/🔓|🔒/g, '').trim()
        })) : [];
        // Prepare fixed-width settings
        const fixedWidthSettings = {
            file1: configsToSave.includes('fixedWidth') ? fixedWidthSettings1 : null,
            file2: configsToSave.includes('fixedWidth') ? fixedWidthSettings2 : null,
        };

        const fileSettings = {
            file1: { noHeader: file1HasNoHeader },
            file2: { noHeader: file2HasNoHeader }
        };

        const templateData = {
            name,
            column_pairs: columnPairsByName, // This line is correct
            comparison_pairs: comparisonPairsByName,
            key_value_mappings: configsToSave.includes('mappings') ? Array.from(keyValueMappings.entries()) : [],
            filters1: configsToSave.includes('filters') ? appliedFilters1 : [],
            filters2: configsToSave.includes('filters') ? appliedFilters2 : [],
            // Only include fixed_width_settings if it has data
            ...( (fixedWidthSettings.file1 || fixedWidthSettings.file2) && { fixed_width_settings: fixedWidthSettings } ),
            file_settings: fileSettings
        };

        if (configsToSave.includes('exclusions')) {
            const getUnmatchedSettings = (fileIndex) => {
                const settings = {};
                const ths = document.querySelectorAll(`#preview${fileIndex} th`);
                ths.forEach(th => { // Trim here for safety
                    const headerName = th.textContent.replace(/🔓|🔒/g, '').trim();
                    const isUnmatched = th.dataset.unmatched === 'true';
                    settings[headerName] = isUnmatched;
                });
                return settings;
            };
            templateData.unmatched_settings = {
                file1: getUnmatchedSettings(1),
                file2: getUnmatchedSettings(2)
            };
        }

        // Create a blob and trigger download
        const blob = new Blob([JSON.stringify(templateData, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${name}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast('Template file generated.', 'success');
        saveTemplateModal.classList.add('hidden'); // Close modal after saving
    }

    function getSelectedTemplateOptions() {
        const options = [];
        if (document.getElementById('save-columns-checkbox').checked) options.push('columns');
        if (document.getElementById('save-exclusions-checkbox').checked) options.push('exclusions');
        if (document.getElementById('save-filters-checkbox').checked) options.push('filters');
        const fixedWidthCheckbox = document.getElementById('save-fixedwidth-checkbox');
        if (fixedWidthCheckbox.checked && !fixedWidthCheckbox.disabled) {
            options.push('fixedWidth');
        }
        // Only include mappings if the checkbox is checked AND not disabled
        const mappingsCheckbox = document.getElementById('save-mappings-checkbox');
        if (mappingsCheckbox.checked && !mappingsCheckbox.disabled) {
            options.push('mappings');
        }
        return options;
    }

    function handleApplyTemplate(event) { // This is now an event handler
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                // Security Hardening: Use a reviver to prevent prototype pollution and other attacks from malicious JSON.
                const jsonReviver = (key, value) => {
                    if (key === '__proto__') {
                        return undefined; // Block prototype pollution
                    }
                    if (key === 'constructor' && value && typeof value === 'object' && 'prototype' in value) {
                        return undefined; // Block constructor-based attacks
                    }
                    return value;
                };
                const template = JSON.parse(e.target.result, jsonReviver);
                await applyTemplateObject(template);
            } catch (error) {
                showToast(`Error reading template file: ${error.message}`, 'error');
            } finally {
                // Reset the input so the user can load the same file again
                applyTemplateInput.value = '';
            }
        };
        reader.readAsText(file);
    }

    async function applyTemplateObject(template) {
        // Show loaders immediately
        showPreviewLoader(1, true, 'Applying template...');
        showPreviewLoader(2, true, 'Applying template...');

        // Defer the heavy work to allow the loader to render.
        setTimeout(async () => {
            try {
                // 2. Clear existing configurations
                clearAllMatches();
                document.querySelector('#filter-builder-1 .filter-rules-container').innerHTML = '';
                appliedFilters1 = [];
                document.querySelector('#filter-builder-2 .filter-rules-container').innerHTML = '';
                appliedFilters2 = [];
                applyFilters(1, true); // Reset data to full
                applyFilters(2, true); // Reset data to full

                const report = {
                    columns: { applied: 0, total: template.column_pairs?.length || 0 },
                    comparisons: { applied: 0, total: template.comparison_pairs?.length || 0 },
                    mappings: { applied: 0, total: template.key_value_mappings?.length || 0 },
                    filters1: { applied: 0, total: template.filters1?.length || 0 },
                    filters2: { applied: 0, total: template.filters2?.length || 0 },
                    fixedWidth: { applied: 0, total: 0 },
                    exclusions: { applied: 0, total: 0 }
                };

                // Apply Fixed-width settings FIRST
                const applyFixedWidthSettings = (fileIndex, settings) => { // This function now returns a Promise
                    return new Promise((resolve, reject) => {
                        if (!settings) return resolve(); // Nothing to do

                        const delimiterSelect = document.getElementById(`delimiter${fileIndex}`);
                        const currentDelimiter = delimiterSelect.value;

                        if (currentDelimiter && currentDelimiter !== 'fixed') {
                            return resolve();
                        }

                        report.fixedWidth.total++;
                        delimiterSelect.value = 'fixed'; // Set the dropdown to 'fixed'
                        const { breaks, useFirstRowAsHeader, customHeaders } = settings;
                        const file = document.getElementById(`file${fileIndex}`).files[0];
                        if (!file) return resolve(); // No file loaded, can't apply

                        const reader = new FileReader();
                        reader.onload = (e) => {
                            try {
                                const parsedData = parseFixedWidthFile(e.target.result, breaks, useFirstRowAsHeader, customHeaders);
                                if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(parsedData));
                                else fullFileData2 = JSON.parse(JSON.stringify(parsedData));
                                processAndRenderParsedData(parsedData, fileIndex);
                                report.fixedWidth.applied++;
                                resolve(); // Resolve the promise on success
                            } catch (error) {
                                reject(error); // Reject on error
                            }
                        };
                        reader.onerror = () => reject(new Error('File could not be read.'));
                        // Defer the potentially blocking read operation to allow UI to update
                        setTimeout(() => {
                            reader.readAsText(file);
                        }, 0);
                    });
                };
                
                // --- NEW "NO HEADER" AND DATA REBASING LOGIC ---
                const rebaseAndPrepareData = async (fileIndex) => {
                    const isNoHeader = template.file_settings?.[`file${fileIndex}`]?.noHeader;
                    const wasInconsistent = fileIndex === 1 ? file1WasInconsistent : file2WasInconsistent;
                    const unmatchedSettings = template.unmatched_settings?.[`file${fileIndex}`];

                    if (isNoHeader) {
                        // 1. Update state and UI
                        document.getElementById(`no-header-checkbox-${fileIndex}`).checked = true;
                        if (fileIndex === 1) file1HasNoHeader = true; else file2HasNoHeader = true;

                        // 2. Re-parse the file as raw data, treating every row as data.
                        const file = document.getElementById(`file${fileIndex}`).files[0];
                        if (!file) return;

                        const delimiter = document.getElementById(`delimiter${fileIndex}`).value;
                        const fileContent = await file.text();
                        const rawParseResult = Papa.parse(fileContent, { delimiter, header: false, skipEmptyLines: 'greedy' });
                        const rawData = rawParseResult.data.filter(row => row.some(cell => cell && cell.trim() !== ''));

                        // 3. Immediately apply headers from the template.
                        const newHeaders = unmatchedSettings ? Object.keys(unmatchedSettings) : [];
                        const rebasedData = { headers: newHeaders, rows: rawData };

                        if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(rebasedData));
                        else fullFileData2 = JSON.parse(JSON.stringify(rebasedData));

                    } else if (wasInconsistent && unmatchedSettings) {
                        // Handle files that were already loaded as inconsistent (but not "no header")
                        const fileData = fileIndex === 1 ? fullFileData1 : fullFileData2;
                        const newHeaders = Object.keys(unmatchedSettings);
                        const rebasedData = { headers: newHeaders, rows: fileData.rows };
                        if (fileIndex === 1) fullFileData1 = JSON.parse(JSON.stringify(rebasedData));
                        else fullFileData2 = JSON.parse(JSON.stringify(rebasedData));
                    }
                };

                // Prepare data for both files before applying other settings
                await rebaseAndPrepareData(1);
                await rebaseAndPrepareData(2);

                // Apply Fixed-width settings (if any)
                if (template.fixed_width_settings) {
                    await Promise.all([
                        applyFixedWidthSettings(1, template.fixed_width_settings.file1),
                        applyFixedWidthSettings(2, template.fixed_width_settings.file2)
                    ]).catch(error => showToast(`Error applying fixed-width settings: ${error.message}`, 'error'));
                }

                // Apply Filters
                const applyFilterRules = async (fileIndex, rules) => {
                    return new Promise(resolve => {
                        const fileHeaders = fileIndex === 1 ? fullFileData1?.headers : fullFileData2?.headers;
                        if (!rules || rules.length === 0 || !fileHeaders) return resolve();
                        
                        const rulesContainer = document.querySelector(`#filter-builder-${fileIndex} .filter-rules-container`);
                        rulesContainer.innerHTML = ''; // Clear any default rule

                        const currentHeaders = new Set(fileIndex === 1 ? fullFileData1?.headers : fullFileData2?.headers);

                        rules.forEach((rule, index) => {
                            if (currentHeaders.has(rule.column)) {
                            addFilterRule(fileIndex, fileHeaders);
                            const ruleEl = rulesContainer.lastElementChild;
                            if (ruleEl) {
                                if (index > 0) {
                                    const logicSelect = ruleEl.querySelector('.filter-rule-logic');
                                    if (logicSelect) logicSelect.value = rule.logic;
                                }
                                ruleEl.querySelector('.filter-column').value = rule.column;
                                ruleEl.querySelector('.filter-operator').value = rule.operator;
                                ruleEl.querySelector('.filter-value').value = rule.value;
                                report[`filters${fileIndex}`].applied++;
                            }
                            }
                        });

                        if (rules.length > 0) {
                            // Update the loader message to give more specific feedback
                            // before starting the potentially long filtering operation.
                            showPreviewLoader(fileIndex, true, 'Applying filters...');

                            document.getElementById(`filter-builder-${fileIndex}`).classList.remove('hidden');
                            
                            const sourceData = fileIndex === 1 ? fullFileData1 : fullFileData2;
                            const headerIndexMap = new Map(sourceData.headers.map((h, i) => [h, i]));

                            const filteredRows = sourceData.rows.filter(row => {
                                let finalResult;
                                for (let i = 0; i < rules.length; i++) {
                                    const rule = rules[i];
                                    const ruleResult = checkCondition(row[headerIndexMap.get(rule.column)], rule.operator, rule.value);
                                    if (i === 0) finalResult = ruleResult;
                                    else {
                                        const logic = rule.logic || 'AND';
                                        if (logic === 'AND') finalResult = finalResult && ruleResult;
                                        else finalResult = finalResult || ruleResult;
                                    }
                                }
                                return finalResult;
                            });

                            // Update the working data, not the full data
                            if (fileIndex === 1) fileData1 = { headers: sourceData.headers, rows: filteredRows };
                            else fileData2 = { headers: sourceData.headers, rows: filteredRows };
                        }
                        resolve(); // Resolve immediately, filtering happens in the background
                    });
                };

                await Promise.all([applyFilterRules(1, template.filters1), applyFilterRules(2, template.filters2)]);
                if (!fileData1 && fullFileData1) fileData1 = JSON.parse(JSON.stringify(fullFileData1));
                if (!fileData2 && fullFileData2) fileData2 = JSON.parse(JSON.stringify(fullFileData2));

                // --- RENDER ONCE ---
                // After all data manipulation (re-basing, filtering) is done, render the tables.
                if (fileData1) renderPreviewTable(fileData1, preview1, 1);
                if (fileData2) renderPreviewTable(fileData2, preview2, 2);

                // 3.5 Apply Column Exclusions (after filters, before pairings)
                if (template.unmatched_settings) {
                    const file1Exclusions = template.unmatched_settings.file1 ? Object.values(template.unmatched_settings.file1).filter(v => v === true).length : 0;
                    const file2Exclusions = template.unmatched_settings.file2 ? Object.values(template.unmatched_settings.file2).filter(v => v === true).length : 0; // If the file was just rebased, the unmatched state is already defined

                    const applyExclusions = (fileIndex, exclusionSettings) => {
                        if (!exclusionSettings) return;
                        const ths = document.querySelectorAll(`#preview${fileIndex} th`);
                        const targetSet = fileIndex === 1 ? unmatchedHeaders1 : unmatchedHeaders2;
                        targetSet.clear(); // Clear and rebuild from template
                        let appliedCountForThisFile = 0;
                        
                        // --- BUG FIX ---
                        // The previous logic only iterated over `ths` in the DOM. If a header from the template
                        // was not found, its setting was ignored. The correct logic is to iterate over the
                        // settings in the template and find the corresponding `th` element.
                        for (const headerName in exclusionSettings) {
                            const th = Array.from(ths).find(t => t.textContent.replace(/🔓|🔒/g, '').trim() === headerName);
                            if (!th) continue; // Header from template not found in current file.
                            
                            appliedCountForThisFile++;
                            const isExcluded = exclusionSettings[headerName];
                            
                            if (isExcluded) {
                                targetSet.add(headerName);
                            }
                            
                            // Update the UI for the found header
                            const icon = th.querySelector('.lock-icon');
                            if (icon) {
                                th.dataset.unmatched = isExcluded.toString();
                                th.classList.toggle('unmatched', isExcluded);
                                icon.textContent = isExcluded ? '🔓' : '🔒';
                                icon.title = isExcluded ? 'Include this column in matching' : 'Exclude this column from matching';
                            }
                        }
                        report.exclusions.applied += appliedCountForThisFile;
                        report.exclusions.total += Object.keys(exclusionSettings).length; // Total remains the number of settings in the template.
                    };
                    applyExclusions(1, template.unmatched_settings.file1);
                    applyExclusions(2, template.unmatched_settings.file2);
                }

                // 5. Apply Column Pairings *after* tables are rendered by filters
                if (template.column_pairs && template.column_pairs.length > 0) {
                    const allTh1 = Array.from(document.querySelectorAll('#preview1 th'));
                    const allTh2 = Array.from(document.querySelectorAll('#preview2 th'));
                    template.column_pairs.forEach(pair => {
                        const th1 = allTh1.find(th => th.textContent.replace(/🔓|🔒/g, '').trim() === pair.file1Header);
                        const th2 = allTh2.find(th => th.textContent.replace(/🔓|🔒/g, '').trim() === pair.file2Header);
                        if (th1 && th2 && !th1.classList.contains('paired') && !th2.classList.contains('paired')) {
                            columnPairs.push({
                                file1ColIndex: parseInt(th1.dataset.colIndex, 10),
                                file2ColIndex: parseInt(th2.dataset.colIndex, 10),
                                file1Th: th1,
                                file2Th: th2,
                            });
                            th1.classList.add('paired');
                            th2.classList.add('paired');
                            report.columns.applied++;
                        }
                    });
                }

                // Apply Comparison Pairings
                if (template.comparison_pairs && template.comparison_pairs.length > 0) {
                    const allTh1 = Array.from(document.querySelectorAll('#preview1 th'));
                    const allTh2 = Array.from(document.querySelectorAll('#preview2 th'));
                    template.comparison_pairs.forEach(pair => {
                        const th1 = allTh1.find(th => th.textContent.replace(/🔓|🔒/g, '').trim() === pair.file1Header);
                        const th2 = allTh2.find(th => th.textContent.replace(/🔓|🔒/g, '').trim() === pair.file2Header);
                        if (th1 && th2 && !th1.classList.contains('paired-compare') && !th2.classList.contains('paired-compare')) {
                            comparisonPairs.push({
                                file1ColIndex: parseInt(th1.dataset.colIndex, 10),
                                file2ColIndex: parseInt(th2.dataset.colIndex, 10),
                                file1Th: th1,
                                file2Th: th2,
                                type: 'compare'
                            });
                            th1.classList.add('paired-compare');
                            th2.classList.add('paired-compare');
                            report.comparisons.applied++;
                        }
                    });
                }

                // 4. Apply Key Value Mappings *after* clearing state
                if (template.key_value_mappings?.length > 0) {
                    const keyIndices1 = columnPairs.map(p => p.file1ColIndex);
                    const keyIndices2 = columnPairs.map(p => p.file2ColIndex);
                    const uniqueValues1 = new Set(fullFileData1.rows.map(row => keyIndices1.map(i => row[i]).join('||')));
                    const uniqueValues2 = new Set(fullFileData2.rows.map(row => keyIndices2.map(i => row[i]).join('||')));
                    const validMappings = template.key_value_mappings.filter(([val1, val2]) => uniqueValues1.has(val1) && uniqueValues2.has(val2));
                    if (validMappings.length > 0) {
                        keyValueMappings = new Map(validMappings);
                        report.mappings.applied = keyValueMappings.size;
                    }
                }

                // Defer the summary toast to ensure it appears after all rendering is complete.
                setTimeout(() => {
                    let summaryMessage = `Template "${template.name}" applied.\n\n`;
                    let appliedSomething = false;
                    let partialApply = false;
                    if (report.columns.total > 0) {
                        summaryMessage += `- Column Pairs: ${report.columns.applied} of ${report.columns.total} applied.\n`;
                        if (report.columns.applied > 0) appliedSomething = true;
                        if (report.columns.applied < report.columns.total) partialApply = true;
                    }
                    if (report.comparisons.total > 0) {
                        summaryMessage += `- Comparison Pairs: ${report.comparisons.applied} of ${report.comparisons.total} applied.\n`;
                        if (report.comparisons.applied > 0) appliedSomething = true;
                        if (report.comparisons.applied < report.comparisons.total) partialApply = true;
                    }
                    if (report.mappings.total > 0) {
                        summaryMessage += `- Key Mappings: ${report.mappings.applied} of ${report.mappings.total} applied.\n`;
                        if (report.mappings.applied > 0) appliedSomething = true;
                    }
                    if (report.filters1.total > 0 || report.filters2.total > 0) {
                        summaryMessage += `- Filters: ${report.filters1.applied + report.filters2.applied} of ${report.filters1.total + report.filters2.total} applied.\n`;
                        if (report.filters1.applied > 0 || report.filters2.applied > 0) appliedSomething = true;
                    }
                    if (report.fixedWidth.total > 0) {
                        summaryMessage += `- Fixed-width: ${report.fixedWidth.applied} of ${report.fixedWidth.total} settings applied.\n`;
                        if (report.fixedWidth.applied > 0) appliedSomething = true;
                    }
                    if (report.exclusions.total > 0) {
                        summaryMessage += `- Exclusions: ${report.exclusions.applied} of ${report.exclusions.total} settings applied.\n`;
                        if (report.exclusions.applied > 0) appliedSomething = true;
                    }
                    if (!appliedSomething) showToast(`Could not apply template "${template.name}".\nNo matching columns or filters found in the current files.`, 'error');
                    else if (partialApply) showToast(`Partially applied template "${template.name}".\nSome columns or filters were not found.\n\n` + summaryMessage, 'warning');
                    else showToast(summaryMessage, 'success');
                }, 100);

                // --- CRITICAL BUG FIX ---
                // After applying all template settings, including exclusions,
                // re-evaluate the unmatched headers to correctly update the UI.
                reEvaluateUnmatchedHeaders(template.unmatched_settings);

                // --- FINAL OVERRIDE ---
                // After all other logic, ensure any columns in a comparison_pair are locked.
                // This must happen LAST to prevent reEvaluateUnmatchedHeaders from unlocking them.
                if (template.comparison_pairs && template.comparison_pairs.length > 0) {
                    const allTh1 = Array.from(document.querySelectorAll('#preview1 th'));
                    const allTh2 = Array.from(document.querySelectorAll('#preview2 th'));
                    template.comparison_pairs.forEach(pair => {
                        const th1 = allTh1.find(th => th.textContent.replace(/🔓|🔒/g, '').trim() === pair.file1Header);
                        const th2 = allTh2.find(th => th.textContent.replace(/🔓|🔒/g, '').trim() === pair.file2Header);
                        
                        if (th1 && th1.dataset.unmatched === 'true') {
                            // This column is part of a comparison pair but is currently unlocked. Force it to be locked.
                            toggleColumnLock(th1);
                        }
                        if (th2 && th2.dataset.unmatched === 'true') {
                            // This column is part of a comparison pair but is currently unlocked. Force it to be locked.
                            toggleColumnLock(th2);
                        }
                    });
                }


                updateUIState();
                templateApplied = true;

            } catch (error) {
                showToast(`Error applying template: ${error.message}`, 'error');
            } finally {
                showPreviewLoader(1, false);
                showPreviewLoader(2, false);
            }
        }, 10);
    }

    // --- Tutorial Logic ---
    const startTutorialBtn = document.getElementById('start-interactive-tutorial-action');
    const tutorialOverlay = document.getElementById('tutorial-overlay');
    const tutorialPopover = document.getElementById('tutorial-popover');
    const tutorialContent = document.getElementById('tutorial-content');
    const tutorialNextBtn = document.getElementById('tutorial-next');
    const tutorialLoadSampleBtn = document.getElementById('tutorial-load-sample-btn');
    const tutorialEndBtn = document.getElementById('tutorial-end');
    const tutorialStepIndicator = document.getElementById('tutorial-step-indicator');

    // --- Tutorial Dropdown Logic ---
    const tutorialDropdownBtn = document.getElementById('tutorial-dropdown-btn');
    const tutorialDropdownMenu = document.getElementById('tutorial-dropdown-menu');
    const tutorialDropdownContainer = tutorialDropdownBtn.parentElement;

    tutorialDropdownBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        tutorialDropdownMenu.classList.toggle('hidden');
        tutorialDropdownContainer.classList.toggle('open');
    });

    document.addEventListener('click', (e) => {
        if (tutorialDropdownContainer && !tutorialDropdownBtn.contains(e.target) && !tutorialDropdownMenu.contains(e.target)) {
            tutorialDropdownMenu.classList.add('hidden');
            tutorialDropdownContainer.classList.remove('open');
        }
    });
    // Add a listener to the header definition modal to end the tutorial if it's closed.
    headerModalCloseBtn.addEventListener('click', () => {
        if (isTutorialActive) {
            endTutorial();
        }
    });
    let currentTutorialStep = 0;

    // --- NEW: Sample data and function for tutorial ---
    const SAMPLE_FILE_1_CONTENT = `OrderID,Product_SKU,UnitPrice,Quantity,Discount_Applied,OrderDate
10248,11,14,12,0,2025-07-04
10248,42,9.8,10,0,2025-07-04
10249,14,18.6,9,0,2025-07-05
10250,51,42.4,10,0.1,2025-07-08
10251,22,16.8,6,0.05,2025-07-08
10252,60,27.2,15,0.05,2025-07-09
10253,31,10,12,0,2025-07-10
10254,55,19.6,21,0.15,2025-07-11`; // Common: OrderID, Quantity. Different: Product_SKU, UnitPrice, Discount_Applied, OrderDate

    const SAMPLE_FILE_2_CONTENT = `OrderID,Product_Code,Item_Price,Quantity,Discount_Percent,Transaction_Date
10248,11,14,24,0,2025-08-03
10248,42,9.8,22,0,2025-08-05
10249,14,18.6,6,0,2025-08-12
10250,51,42.4,16,10%,2025-08-08
10251,22,16.8,17,5%,2025-08-11
10252,60,27.2,11,5%,2025-08-09
10253,31,10,12,0,2025-08-10
10255,16,13.9,2,0,2025-08-12`; // Common: OrderID, Quantity. Different: Product_Code, Item_Price, Discount_Percent, Transaction_Date


    const tutorialSteps = [
        {
            element: 'body', // Target the body for positioning
            content: "Welcome to the Files Match! This interactive guide will walk you through the main features. Click 'Next' to begin.",
            position: 'center'
        },
        {
            element: '#file-control-1',
            content: "First, click 'Choose File' to select your primary file, or use the button below to load a sample file. The next step will unlock after a file is loaded.",
            position: 'bottom',
            action: () => {
                return new Promise(resolve => {
                    // This action needs to resolve when a file is processed, which can happen
                    // directly (delimited) or after a modal (Excel/Fixed-width).
                    const observer = new MutationObserver(() => {
                        const previewVisible = !document.getElementById('preview-wrapper-1').classList.contains('hidden');
                        const excelModalVisible = !document.getElementById('excel-options-modal').classList.contains('hidden');
                        const fixedWidthModalVisible = !document.getElementById('fixed-width-modal').classList.contains('hidden');

                        if (previewVisible || excelModalVisible || fixedWidthModalVisible) {
                            observer.disconnect();
                            // Resolve with the type of UI that appeared.
                            if (excelModalVisible) resolve('excel');
                            else if (fixedWidthModalVisible) resolve('fixed-width');
                            else resolve('preview');
                        }
                    });
                    // Observe all potential outcomes.
                    observer.observe(document.getElementById('preview-wrapper-1'), { attributes: true });
                    observer.observe(document.getElementById('excel-options-modal'), { attributes: true });
                    observer.observe(document.getElementById('fixed-width-modal'), { attributes: true });
                });
            }
        },
        {
            element: '#excel-options-modal .modal-content',
            content: "For Excel files, first select the correct worksheet. You can also specify a cell range or leave it blank to auto-detect.",
            position: 'bottom',
            action: () => new Promise(resolve => document.getElementById('apply-excel-options-btn').addEventListener('click', resolve, { once: true }))
        },
        {
            element: '#fixed-width-modal .modal-content',
            content: "For Fixed-width files, click in the preview to create column breaks. Then, click 'Apply Columns' to continue.",
            position: 'bottom-right-overlap',
            action: () => new Promise(resolve => document.getElementById('apply-fixed-width-btn').addEventListener('click', resolve, { once: true }))
        },
        {
            element: '#file-control-2',
            content: "Great! Now, select the second file you want to compare against, or load the second sample file.",
            position: 'bottom',
            action: () => {
                return new Promise(resolve => {
                    const observer = new MutationObserver(() => {
                        const previewVisible = !document.getElementById('preview-wrapper-2').classList.contains('hidden');
                        const excelModalVisible = !document.getElementById('excel-options-modal').classList.contains('hidden');
                        const fixedWidthModalVisible = !document.getElementById('fixed-width-modal').classList.contains('hidden');

                        if (previewVisible || excelModalVisible || fixedWidthModalVisible) {
                            observer.disconnect();
                            if (excelModalVisible) resolve('excel');
                            else if (fixedWidthModalVisible) resolve('fixed-width');
                            else resolve('preview');
                        }
                    });
                    observer.observe(document.getElementById('preview-wrapper-2'), { attributes: true });
                    observer.observe(document.getElementById('excel-options-modal'), { attributes: true });
                    observer.observe(document.getElementById('fixed-width-modal'), { attributes: true });
                });
            }
        },
        {
            element: '#excel-options-modal .modal-content',
            content: "Select the worksheet for the second file and click 'Apply & Import'.",
            position: 'bottom',
            action: () => new Promise(resolve => document.getElementById('apply-excel-options-btn').addEventListener('click', resolve, { once: true }))
        },
        {
            element: '#fixed-width-modal .modal-content',
            content: "Define the columns for the second file and click 'Apply Columns'.",
            position: 'bottom-right-overlap',
            action: () => new Promise(resolve => document.getElementById('apply-fixed-width-btn').addEventListener('click', resolve, { once: true }))
        },
        {
            element: '#preview-wrapper-1 thead',
            content: "To create a comparison key, click on a column header in File 1...(Unlock a header if needed)",
            position: 'bottom',
            action: () => {
                return new Promise(resolve => {
                    const listener = (e) => {
                        // Ensure the click is on a header in the correct table.
                        const th = e.target.closest('th');
                        if (th && th.closest('#preview-wrapper-1') && th.classList.contains('selected')) {
                            document.querySelector('.previews-container').removeEventListener('click', listener);
                            resolve();
                        }
                    };
                    document.querySelector('.previews-container').addEventListener('click', listener);
                });
            },
            onBefore: () => {
                // This step should only run if both previews are visible.
                if (document.getElementById('preview-wrapper-1').classList.contains('hidden') || document.getElementById('preview-wrapper-2').classList.contains('hidden')) {
                    return false; // Skip this step if previews aren't ready
                }
                return true;
            },
        },
        {
            element: '#preview-wrapper-2 thead',
            content: "...and now click the matching header in File 2 to form a pair. A line will connect them (Unlock a header if needed)",
            position: 'bottom',
            action: () => {
                return new Promise(resolve => {
                    const checkPairs = () => {
                        if (columnPairs.length > 0) {
                            resolve();
                        } else {
                            setTimeout(checkPairs, 100);
                        }
                    };
                    checkPairs();
                });
            }
        },
        {
            element: '#preview-wrapper-1 thead',
            content: "It looks like there are no other columns with matching names. To compare columns with different names (e.g., 'Sales 2024' vs 'Sales 2025'), lock the column, then right-click a header in File 1...",
            position: 'bottom',
            onBefore: () => {
                if (!fileData1 || !fileData2 || columnPairs.length === 0) return false;

                const keyPairHeaders1 = new Set(columnPairs.map(p => fileData1.headers[p.file1ColIndex]));
                const nonKeyHeaders1 = fileData1.headers.filter(h => !keyPairHeaders1.has(h) && !unmatchedHeaders1.has(h));
                const nonKeyHeaders2 = new Set(fileData2.headers.filter(h => !unmatchedHeaders2.has(h)));

                const commonNonKeyHeaders = nonKeyHeaders1.filter(h => nonKeyHeaders2.has(h));
                
                // Activate this step only if there are no common, non-key, included headers.
                return commonNonKeyHeaders.length === 0;
            },
            action: () => {
                return new Promise(resolve => {
                    const listener = () => {
                        if (rightClickSelectedTh) {
                            document.querySelector('.previews-container').removeEventListener('contextmenu', listener);
                            resolve();
                        }
                    };
                    document.querySelector('.previews-container').addEventListener('contextmenu', listener);
                });
            }
        },
        {
            element: '#preview-wrapper-2 thead',
            content: "...and then lock the column in File 2, then right-click the corresponding header in File 2 to create a comparison-only pair (dotted line).",
            position: 'bottom',
            onBefore: () => rightClickSelectedTh !== null, // Only show if a right-click selection is active
            action: () => {
                return new Promise(resolve => {
                    const checkPairs = () => {
                        if (comparisonPairs.length > 0) resolve();
                        else setTimeout(checkPairs, 100);
                    };
                    checkPairs();
                });
            }
        },
        {
            element: '#manual-match-btn',
            content: "Optional: If your key columns have different values that mean the same thing (e.g., 'USA' vs 'United States'), use this button to define custom mappings. You can try this after the tutorial. Click 'Next' to continue.",
            position: 'bottom',
            disableInteraction: true
        },
        {
            element: '#template-management-container',
            content: "You can also save your current configuration as a template or load a previous one. You can try this after the tutorial. Click 'Next' to continue.",
            position: 'bottom',
            disableInteraction: true
        },
        {
            element: '.comparison-controls',
            content: "The 'Group Result by Key' option is useful for summarizing data. We've checked it for you. Now, click 'Run Comparison' to see the grouped results.",
            position: 'bottom',
            onBefore: () => {
                document.getElementById('group-by-key').checked = true; // Auto-check the box
                return true; // Allow step to proceed
            },
            action: () => {
                return new Promise(resolve => {
                    runComparisonBtn.addEventListener('click', resolve, { once: true });
                });
            }
        },
        {
            element: '#results-section',
            content: "The results preview appears here, highlighting any differences found. You've learned the basics! Click 'Finish' to end the tour and explore the results.",
            position: 'bottom',
            onBefore: () => {
                // This step is shown after the comparison runs, so no pre-check needed anymore.
                return true; 
            }
        }
    ];

    function loadSampleFileForTutorial(fileIndex) {
        const content = fileIndex === 1 ? SAMPLE_FILE_1_CONTENT : SAMPLE_FILE_2_CONTENT;
        const fileName = fileIndex === 1 ? '@File1.csv' : '@File2.csv';
        const file = new File([content], fileName, { type: 'text/csv' });

        // Manually update the file input's files property. This is crucial for
        // ensuring that subsequent actions (like changing a delimiter) can
        // access the file as if it were selected via the dialog.
        const dataTransfer = new DataTransfer();
        dataTransfer.items.add(file);
        document.getElementById(`file${fileIndex}`).files = dataTransfer.files;

        processFile(file, fileIndex);
    }

    function startTutorial() {
        resetAll(); // Start with a clean slate
        isTutorialActive = true; // Set the flag when the tutorial starts
        currentTutorialStep = 0;
        document.body.appendChild(tutorialPopover); // Move popover to body to ensure correct positioning context
        showTutorialStep(currentTutorialStep);
    }

    function endTutorial(finishedSuccessfully = false) {
        tutorialOverlay.classList.add('hidden');
        isTutorialActive = false; // Unset the flag when the tutorial ends
        tutorialPopover.classList.add('hidden');
        tutorialNextBtn.disabled = false; // Re-enable for next time
        document.querySelector('main').appendChild(tutorialPopover); // Return popover to its original place
        if (!finishedSuccessfully) {
            resetAllBtn.classList.add('hidden'); // Hide the Clear All button on any tutorial exit.
        }
        document.querySelectorAll('.tutorial-highlight').forEach(el => el.classList.remove('tutorial-highlight'));
    }
    // Add resetAll to the global scope so it can be called from modal close handlers

    function repositionTutorialPopover() { // If the tutorial is not active, do nothing.
        if (!isTutorialActive) return;
        if (tutorialPopover.classList.contains('hidden') || currentTutorialStep >= tutorialSteps.length) {
            return;
        }
    
        const step = tutorialSteps[currentTutorialStep];
        const targetElement = document.querySelector(step.element);
    
        if (!targetElement) return;
    
        const targetRect = targetElement.getBoundingClientRect();
        const popoverHeight = tutorialPopover.offsetHeight;
        const popoverWidth = tutorialPopover.offsetWidth;

        let finalPosition = step.position || 'bottom';
        const spaceBelow = window.innerHeight - targetRect.bottom;
        const spaceAbove = targetRect.top;
    
        // Only flip between top/bottom if it's not a special position.
        const specialPositions = ['bottom-right-overlap', 'center'];
        if (!specialPositions.includes(finalPosition)) {
            if (finalPosition === 'bottom' && spaceBelow < (popoverHeight + 20) && spaceAbove > (popoverHeight + 20)) {
                finalPosition = 'top';
            } else if (finalPosition === 'top' && spaceAbove < popoverHeight + 20 && spaceBelow > popoverHeight + 20) {
                finalPosition = 'bottom';
            }
        }
    
        let topPos, leftPos;

        if (finalPosition === 'bottom-right-overlap') {            
            // Use fixed positioning to anchor to the viewport's bottom-right corner.
            // This is the most reliable way to ensure it's visible without scrolling.
            // Position from the top of the viewport to ensure it's always visible.
            tutorialPopover.style.position = 'fixed';
            tutorialPopover.style.top = '10px'; // 10px from the top of the visible screen
            tutorialPopover.style.left = 'auto';
            tutorialPopover.style.bottom = 'auto';
            tutorialPopover.style.right = '20px'; // 20px from the right edge
        } else if (finalPosition === 'center') {
            tutorialPopover.style.position = 'fixed';
            tutorialPopover.style.top = '50%';
            tutorialPopover.style.left = '50%';
            tutorialPopover.style.transform = 'translate(-50%, -50%)';
            tutorialPopover.style.bottom = 'auto';
            tutorialPopover.style.right = 'auto';
        } else {
            // For all other steps, ensure we are using absolute positioning relative to the document.
            tutorialPopover.style.position = 'absolute';
            tutorialPopover.style.bottom = '';
            tutorialPopover.style.right = '';

            topPos = (finalPosition === 'bottom') 
                ? window.scrollY + targetRect.bottom + 10 
                : window.scrollY + targetRect.top - popoverHeight - 10;
            
            // Calculate the ideal centered left position.
            leftPos = window.scrollX + targetRect.left + (targetRect.width / 2) - (popoverWidth / 2);

            // --- FIX: Ensure the popover does not overflow the viewport horizontally ---
            // Clamp the left position to stay within the viewport boundaries.
            leftPos = Math.max(5, Math.min(leftPos, window.innerWidth - popoverWidth - 5));

            tutorialPopover.style.top = `${topPos}px`;
            tutorialPopover.style.left = `${leftPos}px`;
            tutorialPopover.style.transform = ''; // Reset transform
        }
    }

    // This is a simple way to handle function ordering in a single-file script.
    window.resetAll = resetAll;

    async function showTutorialStep(stepIndex) {
        let layoutObserver = null; // To hold a MutationObserver for layout changes

        // --- CRITICAL FIX ---
        // If the tutorial is no longer active, stop processing any further steps.
        // This prevents pending actions from re-triggering the tutorial UI.
        if (!isTutorialActive) return;

        if (stepIndex < 0 || stepIndex >= tutorialSteps.length) { // This means the tutorial is finished
            endTutorial(true); // Pass true to indicate successful completion, which keeps the "Clear All" button visible.
            return;
        }

        const step = tutorialSteps[stepIndex];
        currentTutorialStep = stepIndex; // Update global step index

        // Run pre-condition check if it exists
        if (step.onBefore && !step.onBefore()) {
            endTutorial();
            return;
        }

        // Disconnect any previous observers when moving to a new step
        if (layoutObserver) {
            layoutObserver.disconnect();
        }

        document.querySelectorAll('.tutorial-highlight').forEach(el => el.classList.remove('tutorial-highlight'));
        document.querySelectorAll('[data-tutorial-disabled]').forEach(el => el.removeAttribute('data-tutorial-disabled'));

        const targetElement = document.querySelector(step.element);
        if (!targetElement) {
            // This can happen when skipping irrelevant steps (e.g., a modal step for a CSV file), which is fine.
            endTutorial();
            return;
        }

        const isModalStep = step.element.includes('-modal');
        const disableInteraction = step.disableInteraction || false;

        // For modal steps or steps with no highlight, we don't want the dark overlay.
        if (isModalStep) {
            tutorialOverlay.classList.add('hidden');
        } else { // This is the default case
            tutorialOverlay.classList.remove('hidden');
            // Only highlight if the target is not the body itself
            if (step.element !== 'body') targetElement.classList.add('tutorial-highlight');
        }

        tutorialPopover.classList.remove('hidden');
        if (disableInteraction) {
            targetElement.setAttribute('data-tutorial-disabled', 'true');
        }
        targetElement.scrollIntoView({ behavior: 'smooth', block: 'center' });

        tutorialContent.textContent = step.content;
        tutorialStepIndicator.textContent = `${stepIndex + 1} / ${tutorialSteps.length}`;        

        // Hide the "End Tour" button on the final step to guide the user to click "Finish".
        if (stepIndex === tutorialSteps.length - 1) {
            tutorialEndBtn.classList.add('hidden');
        } else {
            tutorialEndBtn.classList.remove('hidden');
        }

        // If the delimiter dropdown appears for a file, the layout will change.
        // We need to reposition the popover to avoid overlap.
        if (step.element === '#file-control-1' || step.element === '#file-control-2') {
            new MutationObserver(() => repositionTutorialPopover()).observe(targetElement, { childList: true, subtree: true });
        }
        tutorialPopover.style.transition = 'none'; // Remove transition for initial placement
        repositionTutorialPopover(); // Use the new centralized function for positioning

        tutorialNextBtn.textContent = stepIndex === tutorialSteps.length - 1 ? 'Finish' : 'Next';

        // Handle interactive steps
        if (step.action) {
            // --- FIX ---
            // Show the "Load Sample File" button for the file selection steps *before* awaiting the action.
            if (step.element === '#file-control-1' || step.element === '#file-control-2') {
                tutorialLoadSampleBtn.classList.remove('hidden');
                const fileIndex = step.element === '#file-control-1' ? 1 : 2;
                tutorialLoadSampleBtn.onclick = () => loadSampleFileForTutorial(fileIndex);
            }

            tutorialNextBtn.classList.add('hidden');
            if (layoutObserver) layoutObserver.disconnect(); // Stop observing once action starts
            const actionResult = await step.action();

            // After the action is complete, hide the sample file button again
            tutorialLoadSampleBtn.classList.add('hidden');
            tutorialLoadSampleBtn.onclick = null; // Remove the specific listener

            // If the action was loading a sample file, we need to resolve the promise
            // that the original `step.action` was waiting for. Since we've now taken
            // over, we can manually determine the next step.
            const previewVisible = !document.getElementById(`preview-wrapper-${currentTutorialStep === 1 ? 1 : 2}`).classList.contains('hidden');
            if (previewVisible) {
                // This simulates the 'preview' result from the original action promise
                // and ensures the tutorial advances correctly.
            }


            // Logic to advance to the correct next step based on the action's outcome
            let nextStepIndex;
            if (actionResult === 'excel') {
                // Find the next step that is for the excel modal
                nextStepIndex = tutorialSteps.findIndex((s, i) => i > currentTutorialStep && s.element.includes('excel-options-modal'));
            } else if (actionResult === 'fixed-width') {
                // Find the next step for the fixed-width modal
                nextStepIndex = tutorialSteps.findIndex((s, i) => i > currentTutorialStep && s.element.includes('fixed-width-modal'));
            } else { // Covers 'preview' and undefined (e.g., after a modal action)
                // After an action, check if the next step has a condition. If the condition is false, skip it.
                nextStepIndex = currentTutorialStep + 1;
                while (nextStepIndex < tutorialSteps.length && (tutorialSteps[nextStepIndex].element.includes('-modal') || (tutorialSteps[nextStepIndex].onBefore && !tutorialSteps[nextStepIndex].onBefore()))) {
                    nextStepIndex++;
                }
            }

            setTimeout(() => {
                showTutorialStep(nextStepIndex);
            }, 500); // A small delay for user to see the result of their action
        } else {
            tutorialNextBtn.classList.remove('hidden');
            tutorialLoadSampleBtn.classList.add('hidden'); // Hide sample button for non-action steps
        }
    }

    startTutorialBtn.addEventListener('click', startTutorial);
    tutorialEndBtn.addEventListener('click', () => {
        tutorialLoadSampleBtn.classList.add('hidden'); // Ensure it's hidden on exit
        tutorialLoadSampleBtn.onclick = null;
        // Only reset the application if the user ends the tutorial *before* the final step.
        if (currentTutorialStep < tutorialSteps.length - 1) {
            resetAll();
            showToast("Tutorial ended and application has been reset.", "info");
        } else {
            showToast("Tutorial finished. You can now explore the results.", "success");
        }
        endTutorial(); // This is for premature ending, so it correctly hides the button.
    });
    tutorialOverlay.addEventListener('click', () => {
                // If the current step is one where interaction is disabled,
        // a click on the overlay was likely a "click-through" on the
        // highlighted element. In this case, we should not end the tour.
        const step = tutorialSteps[currentTutorialStep];
        if (step && step.disableInteraction) {
            return;
        }
        tutorialLoadSampleBtn.classList.add('hidden'); // Ensure it's hidden on exit
        tutorialLoadSampleBtn.onclick = null;
        endTutorial();
    });
    tutorialNextBtn.addEventListener('click', () => {
        currentTutorialStep++;
        // Skip modal steps on manual "Next" click
        // Also skip conditional steps whose conditions are not met
        while (currentTutorialStep < tutorialSteps.length && (tutorialSteps[currentTutorialStep].element.includes('-modal') || (tutorialSteps[currentTutorialStep].onBefore && !tutorialSteps[currentTutorialStep].onBefore()))) {
            currentTutorialStep++;
        }

        while (currentTutorialStep < tutorialSteps.length && tutorialSteps[currentTutorialStep].element.includes('-modal')) {
            currentTutorialStep++;
        }
        showTutorialStep(currentTutorialStep);
    });

    // --- Drag and Drop File Logic ---
    [1, 2].forEach(fileIndex => {
        const fileControl = document.getElementById(`file-control-${fileIndex}`);

        fileControl.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.stopPropagation();
            // Do not show drop zone for file 2 if it's disabled
            if (fileIndex === 2 && fileInput2.disabled) return;
            fileControl.classList.add('drag-over');
        });

        fileControl.addEventListener('dragleave', (e) => {
            e.preventDefault();
            e.stopPropagation();
            fileControl.classList.remove('drag-over');
        });

        fileControl.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            fileControl.classList.remove('drag-over');

            if (fileIndex === 2 && fileInput2.disabled) return;

            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                if (e.dataTransfer.files.length > 1) {
                    showToast("Please drop only one file at a time.", "warning");
                    return;
                }
                const droppedFile = e.dataTransfer.files[0];

                // --- BUG FIX ---
                // Manually update the file input's files property. This is crucial for
                // ensuring that subsequent actions (like changing a delimiter) can
                // access the file as if it were selected via the dialog.
                const dataTransfer = new DataTransfer();
                dataTransfer.items.add(droppedFile);
                document.getElementById(`file${fileIndex}`).files = dataTransfer.files;

                processFile(droppedFile, fileIndex);
            }
        });
    });
});