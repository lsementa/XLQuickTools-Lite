// --- Modal Control Functions ---
// Modal.js

Office.onReady(() => {
    // On Ready
});


// Get user selected range and update the input field
function getUserRange(element) {
    Excel.run(async (context) => {
        try {
            const range = context.workbook.getSelectedRange();
            const effectiveRange = await getEffectiveRangeForSelection(context, range);

            if (effectiveRange) {
                effectiveRange.load("address");
                await context.sync();
                element.value = effectiveRange.address;
            } else {
                element.value = "";
            }
        } catch (error) {
            element.value = "";
        }
    }).catch((error) => {
        console.error("Excel.run error in getUserRange:", error);
    });
}

function updateFormatOptions(locale) {
    const convertFormat = document.getElementById('convertFormat');
    if (!convertFormat) return;

    // Clear existing options
    convertFormat.innerHTML = '';

    // Define format sets
    const formats = {
        US: [
            "yyyy-MM-dd",
            "M/d/yyyy",
            "MM/dd/yyyy",
            "MMM dd, yyyy",
            "MMMM dd, yyyy",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy MMM dd"
        ],
        Other: [
            "yyyy-MM-dd",
            "dd/MM/yyyy",
            "d/M/yyyy",
            "dd MMM yyyy",
            "dd MMMM yyyy",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyy MMM dd"
        ]
    };

    // Add new options
    formats[locale].forEach(fmt => {
        const option = document.createElement('option');
        option.value = fmt;
        option.textContent = fmt;
        convertFormat.appendChild(option);
    });

    // Set default format
    convertFormat.value = 'yyyy-MM-dd';
}


// Show modal message to user
function showModalMessage(title, message, showCancel = false) {
    const modal = document.getElementById('message');
    const modalTitle = modal.querySelector('h2');
    const modalBody = modal.querySelector('.modal-body');
    const okButton = document.getElementById('OkButton');
    const cancelButton = document.getElementById('CancelButton');
    modalTitle.textContent = title;
    modalBody.innerHTML = message;

    // Show/hide cancel button based on parameter
    cancelButton.style.display = showCancel ? 'inline-block' : 'none';

    // Add the 'show-modal' class to make it visible and trigger transitions
    modal.classList.add('show-modal');

    return new Promise((resolve) => {
        const handleOk = () => {
            modal.classList.remove('show-modal');
            okButton.removeEventListener('click', handleOk);
            if (showCancel) {
                cancelButton.removeEventListener('click', handleCancel);
            }
            resolve(true);
        };

        const handleCancel = () => {
            modal.classList.remove('show-modal');
            cancelButton.removeEventListener('click', handleCancel);
            okButton.removeEventListener('click', handleOk);
            resolve(false);
        };

        okButton.addEventListener('click', handleOk);
        if (showCancel) {
            cancelButton.addEventListener('click', handleCancel);
        }
    });
}

// Show the "Add Leading/Trailing Text" modal
function showAddLeadTrailModal() {
    const modal = document.getElementById('addLeadTrailModal');
    const leading = document.getElementById('leadingText');
    const trailing = document.getElementById('trailingText');

    if (modal) {
        modal.classList.add('show-modal');
        // Clear inputs and focus when showing
        leading.value = '';
        trailing.value = '';
        // Focus on the first input
        leading.focus();
    }
}

// Hide the "Add Leading/Trailing Text" modal
function hideAddLeadTrailModal() {
    const modal = document.getElementById('addLeadTrailModal');
    if (modal) {
        modal.classList.remove('show-modal');
    }
}

// OK button clicked on the "Add Leading/Trailing Text" modal
async function onAddLeadTrailOk() {
    const leadingText = document.getElementById('leadingText').value;
    const trailingText = document.getElementById('trailingText').value;

    // console.log(`Modal OK clicked. Leading: "${leadingText}", Trailing: "${trailingText}"`);
    hideAddLeadTrailModal();

    // Run
    try {
        await addLeaadTrail(leadingText, trailingText);
    } catch (error) {
        console.error("Error applying leading/trailing text:", error);
    }
}

// Show the "Add Hyperlinks" modal
function showAddHyperlinksModal() {
    const modal = document.getElementById('addHyperlinksModal');
    const url = document.getElementById('url');
    const urllabel = document.getElementById('urllabel');
    const headers = document.getElementById('URLheaders');
    const cellurls = document.getElementById('URLcells');

    if (modal) {
        modal.classList.add('show-modal');
        // Clear inputs and focus when showing
        headers.checked = true;
        cellurls.checked = false;
        url.value = '';
        urllabel.classList.add('show');
        url.classList.add('show');
        url.disabled = false;
        url.focus();
    }
}

// Hide the "Add Hyperlinks" modal
function hideAddHyperlinksModal() {
    const modal = document.getElementById('addHyperlinksModal');
    if (modal) {
        modal.classList.remove('show-modal');
    }
}

// OK button clicked on the "Add Hyperlinks" modal
async function onAddHyperlinksOk() {
    const url = document.getElementById('url').value;
    const headers = document.getElementById('URLheaders').checked;
    const cellurls = document.getElementById('URLcells').checked;

    // console.log(`Modal OK clicked. URL: "${url}", Display Text: "${displayText}"`);
    hideAddHyperlinksModal();

    // Run
    try {
        if (url) {
            await addHyperlinks(url, headers, cellurls);
        }
    } catch (error) {
        console.error("Error adding URLs:", error);
    }
}

// Hide the "Selection Plus" modal
function hideSelectionPlusModal() {
    const modal = document.getElementById('selectionPlusModal');
    if (modal) {
        modal.classList.remove('show-modal');
        const customDelimiterInput = document.getElementById('customDelimiterInput');
        if (customDelimiterInput) {
            customDelimiterInput.classList.remove('show');
            customDelimiterInput.disabled = true;
        }
    }
}

// Shows the "Selection Plus"" modal and resets its input fields
function showSelectionPlusModal() {
    const modal = document.getElementById('selectionPlusModal');
    const delimiterSelect = document.getElementById('delimiterSelect');
    const customDelimiterInput = document.getElementById('customDelimiterInput');
    const leadingTextInput = document.getElementById('selectionPlusLeadingText');
    const trailingTextInput = document.getElementById('selectionPlusTrailingText');

    if (modal && delimiterSelect && customDelimiterInput && leadingTextInput && trailingTextInput) {
        // Reset fields to default values when showing
        delimiterSelect.value = "custom"; // Default to custom
        customDelimiterInput.value = ",";
        leadingTextInput.value = "";
        trailingTextInput.value = "";
        // Focus on first textbox
        leadingTextInput.focus();

        if (delimiterSelect.value === 'custom') {
            customDelimiterInput.classList.add('show');
            customDelimiterInput.disabled = false;
        } else {
            customDelimiterInput.classList.remove('show');
            customDelimiterInput.disabled = true;
        }

        modal.classList.add('show-modal'); // Show the modal
    }
}

// Ok button clicked on "selection plus" modal
async function selectionPlusOkButtonHandler() {
    const delimiterSelect = document.getElementById('delimiterSelect');
    const customDelimiterInput = document.getElementById('customDelimiterInput');
    const leadingTextInput = document.getElementById('selectionPlusLeadingText');
    const trailingTextInput = document.getElementById('selectionPlusTrailingText');

    // Basic validation to ensure elements exist before trying to read values
    if (!delimiterSelect || !customDelimiterInput || !leadingTextInput || !trailingTextInput) {
        console.error("Selection Plus modal inputs not found. Cannot proceed with OK action.");
        return;
    }

    // Get Delimiter
    const selectedDelimiterOption = delimiterSelect.value;
    let delimiter = selectedDelimiterOption;

    if (selectedDelimiterOption === 'custom') {
        delimiter = customDelimiterInput.value;
    } else {
        // Convert common escape sequences to actual characters
        switch (selectedDelimiterOption) {
            case '\\t': delimiter = '\t'; break;
            case '\\r': delimiter = '\r'; break;
            case '\\n': delimiter = '\n'; break;
            case '\\v': delimiter = '\v'; break;
            case '\\f': delimiter = '\f'; break;
            case '\\r\\n': delimiter = '\r\n'; break;
            case '\\u00A0': delimiter = '\u00A0'; break;
        }
    }

    const leadingText = leadingTextInput.value;
    const trailingText = trailingTextInput.value;

    // console.log(`Selection Plus OK clicked. Delimiter: "${delimiter}", Leading: "${leadingText}", Trailing: "${trailingText}"`);
    hideSelectionPlusModal();

    // Run selection plus
    try {
        await selectionPlus(leadingText, trailingText, delimiter);
    } catch (error) {
        console.error("Error during Selection Plus Excel operation:", error);
    }
}

// Show the "Split to Rows" modal
function showSplitToRowsModal() {
    const modal = document.getElementById('splitToRowsModal');
    const delimiterSelect = document.getElementById('SRdelimiterSelect');
    const customDelimiterInput = document.getElementById('SRcustomDelimiterInput');
    const headers = document.getElementById('SRheaders');

    if (modal && delimiterSelect && customDelimiterInput) {
        modal.classList.add('show-modal');
        // Clear inputs and focus when showing
        customDelimiterInput.value = '';
        headers.checked = true;
        // Focus on the input and add defaults
        delimiterSelect.value = "custom"; // Default to custom
        if (delimiterSelect.value === 'custom') {
            customDelimiterInput.classList.add('show');
            customDelimiterInput.disabled = false;
            customDelimiterInput.focus();
        } else {
            customDelimiterInput.classList.remove('show');
            customDelimiterInput.disabled = true;
        }

    }
}

// Hide the "Split to Rows" modal
function hideSplitToRowsModal() {
    const modal = document.getElementById('splitToRowsModal');
    const customDelimiterInput = document.getElementById('SRcustomDelimiterInput');

    if (modal) {
        modal.classList.remove('show-modal');
        if (customDelimiterInput) {
            customDelimiterInput.classList.remove('show');
            customDelimiterInput.disabled = true;
        }
    }
}

// Ok button clicked on "Split to Rows" modal
async function splitToRowsOkButtonHandler() {
    const delimiterSelect = document.getElementById('SRdelimiterSelect');
    const customDelimiterInput = document.getElementById('SRcustomDelimiterInput');
    const headers = document.getElementById('SRheaders').checked;

    // Basic validation to ensure elements exist before trying to read values
    if (!delimiterSelect || !customDelimiterInput) {
        console.error("Split to Rows modal inputs not found. Cannot proceed with OK action.");
        return;
    }
    const selectedDelimiterOption = delimiterSelect.value;
    let delimiter = selectedDelimiterOption;

    if (selectedDelimiterOption === 'custom') {
        delimiter = customDelimiterInput.value;
    } else {
        // Convert common escape sequences to actual characters
        switch (selectedDelimiterOption) {
            case '\\t': delimiter = '\t'; break;
            case '\\r': delimiter = '\r'; break;
            case '\\n': delimiter = '\n'; break;
            case '\\v': delimiter = '\v'; break;
            case '\\f': delimiter = '\f'; break;
            case '\\r\\n': delimiter = '\r\n'; break;
            case '\\u00A0': delimiter = '\u00A0'; break;
        }
    }

    // console.log(`Split to Rows OK clicked. Delimiter: "${delimiter}", Headers: "${headers}"`);
    hideSplitToRowsModal();

    // If no delimiter exit early
    if (!delimiter) {
        return;
    }

    // Run Split to Rows
    try {
        await splitToRows(headers, delimiter);
    } catch (error) {
        console.error("Error during Split to Rows Excel operation:", error);
    }
}

// Show the "Find Missing" modal
function showFindMissingModal() {
    const modal = document.getElementById('findMissingModal');
    const range1 = document.getElementById('range1');
    const range2 = document.getElementById('range2');
    const highlight = document.getElementById('FMHighlight');

    if (modal && range1 && range2) {
        modal.classList.add('show-modal');
        // Clear inputs and enable
        range1.value = '';
        range2.value = '';
        range1.disabled = false;
        range2.disabled = false;
        highlight.checked = false;
    } else {
        console.error("Modal element 'findMissingModal' not found.");
    }
}

// Hide the "Find Missing" modal
function hideFindMissingModal() {
    const modal = document.getElementById('findMissingModal');
    if (modal) {
        modal.classList.remove('show-modal');
        const customDelimiterInput = document.getElementById('SRcustomDelimiterInput');
        if (customDelimiterInput) {
            customDelimiterInput.classList.remove('show');
            customDelimiterInput.disabled = true;
        }
    }
}

// Ok button clicked on "Find Missing" modal
async function findMissingOkButtonHandler() {
    const range1 = document.getElementById('range1');
    const range2 = document.getElementById('range2');
    const highlight = document.getElementById('FMHighlight').checked;

    // Basic validation to ensure elements exist before trying to read values
    if (!range1 || !range2) {
        console.error("Find Missing modal inputs not found. Cannot proceed with OK action.");
        return;
    }

    // console.log(`Find Missing OK clicked. Range1: "${range1}", Range2: "${range2}"`);
    hideFindMissingModal();

    // Run Find Missing
    try {
        await findMissingData(highlight, range1.value, range2.value);
    } catch (error) {
        console.error("Error during Find Missing Excel operation:", error);
    }
}

// Function to populate dropdowns with worksheet names
async function populateWorksheetSelect(selectElement) {
    await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name"); // Load the names of all worksheets

        await context.sync();

        // Clear existing options
        selectElement.innerHTML = '';

        // Add a default "Please select..." option
        const defaultOption = document.createElement('option');
        defaultOption.value = "";
        defaultOption.textContent = "-- Select a Worksheet --";
        defaultOption.disabled = true;
        defaultOption.selected = true;
        selectElement.appendChild(defaultOption);

        // Populate the select with worksheet names
        for (const sheet of worksheets.items) {
            const option = document.createElement('option');
            option.value = sheet.name;
            option.textContent = sheet.name;
            selectElement.appendChild(option);
        }
    }).catch(function (error) {
        console.error("Error populating worksheet select:", error);
        showModalMessage("Compare Worksheets", "Could not load worksheet list.", true);
    });
}

// Show the "Compare Worksheets" modal
function showCompareSheetsModal() {
    const modal = document.getElementById('compareSheetsModal');
    const worksheet1 = document.getElementById('compareSheetsWS1Select');
    const worksheet2 = document.getElementById('compareSheetsWS2Select');
    const highlight = document.getElementById('CSHighlight');

    if (modal && worksheet1 && worksheet2) {
        modal.classList.add('show-modal');
        // Clear inputs and enable
        worksheet1.value = '';
        worksheet2.value = '';
        highlight.checked = false;

        // Populate the select elements with worksheets
        populateWorksheetSelect(worksheet1);
        populateWorksheetSelect(worksheet2);

    } else {
        console.error("Modal element 'compareSheetsModal' not found.");
    }
}

// Hide the "Compare Worksheets" modal
function hideCompareSheetsModal() {
    const modal = document.getElementById('compareSheetsModal');
    if (modal) {
        modal.classList.remove('show-modal');
    }
}

// Ok button clicked on "Compare Worksheets" modal
async function compareSheetsOkButtonHandler() {
    const worksheet1 = document.getElementById('compareSheetsWS1Select');
    const worksheet2 = document.getElementById('compareSheetsWS2Select');
    const highlight = document.getElementById('CSHighlight').checked;

    // Basic validation to ensure elements exist before trying to read values
    if (!worksheet1 || !worksheet2) {
        console.error("Compare Sheets modal inputs not found. Cannot proceed with OK action.");
        return;
    }

    // console.log(`Compare Worksheets OK clicked. Worksheet1: "${worksheet1.value}", Worksheet2: "${worksheet2.value}"`);
    hideCompareSheetsModal();

    // Run Compare Worksheets
    try {
        await compareWorksheets(highlight, worksheet1.value, worksheet2.value);
    } catch (error) {
        console.error("Error during Compare Worksheets Excel operation:", error);
    }
}

// Show the "Date/Text Converter" modal
function showDateTextModal() {
    const modal = document.getElementById('dateTextModal');
    const currentLocale = document.getElementById('currentLocale');
    const convertLocale = document.getElementById('convertLocale');
    const convertFormat = document.getElementById('convertFormat');
    const convertType = document.getElementById('convertType');

    if (modal && currentLocale && convertLocale && convertFormat && convertType) {
        modal.classList.add('show-modal');

        // Get the user's current locale
        let userLocale = navigator.language || navigator.languages[0] || 'en-US'; // Fallback to 'en-US'

        // Extract the language code (e.g., "en" from "en-US")
        let localeCode = userLocale.split('-')[0].toUpperCase();

        // Determine the locale to set
        let initialLocale;
        // Check for 'US' code or full 'en-US'
        if (localeCode === 'US' || userLocale.includes('en-US')) {
            initialLocale = 'US';
        } else {
            // If not US, set to Other
            initialLocale = 'Other';
        }

        // Set defaults
        currentLocale.value = initialLocale;
        convertLocale.value = initialLocale;
        convertType.value = 'text';

        // Populate convertFormat based on US locale
        updateFormatOptions(initialLocale);

    } else {
        console.error("Modal element 'dateTextModal' not found.");
    }
}

// Hide the "Date/Text Converter" modal
function hideDateTextModal() {
    const modal = document.getElementById('dateTextModal');
    if (modal) {
        modal.classList.remove('show-modal');
    }
}

// Ok button clicked on "Date/Text Converter" modal
async function dateTextOkButtonHandler() {
    const currentLocale = document.getElementById('currentLocale');
    const convertLocale = document.getElementById('convertLocale');
    const convertFormat = document.getElementById('convertFormat');
    const convertType = document.getElementById('convertType');

    // Basic validation to ensure elements exist before trying to read values
    if (!currentLocale || !convertLocale || !convertFormat || !convertType) {
        console.error("Date/Text Converter modal inputs not found. Cannot proceed with OK action.");
        return;
    }

    // console.log(`Date/Text Converter OK clicked. Locale: "${currentLocale.value}", To: "${convertLocale.value}", Format: "${convertFormat.value}", Type: "${convertType.value}"`);
    hideDateTextModal();

    // Run Date/Text Converter
    try {

        await convertSelectedDates(currentLocale.value, convertLocale.value, convertFormat.value, convertType.value);
    } catch (error) {
        console.error("Error during Date/Text Converter Excel operation:", error);
    }
}


// --- Event Listeners for Modal Buttons (run after DOM is fully loaded) ---------------------------------------------------->

// Event Listeners
document.addEventListener('DOMContentLoaded', () => {

    // --- Listeners for 'Add Lead/Trail' Modal ---

    const addLeadTrailOkButton = document.getElementById('addLeadTrailOkButton');
    const addLeadTrailCancelButton = document.getElementById('addLeadTrailCancelButton');
    const addLeadTrailModalOverlay = document.getElementById('addLeadTrailModal');

    if (addLeadTrailOkButton) {
        addLeadTrailOkButton.addEventListener('click', onAddLeadTrailOk);
    } else {
        console.error("Add Lead/Trail OK button not found.");
    }

    if (addLeadTrailCancelButton) {
        addLeadTrailCancelButton.addEventListener('click', hideAddLeadTrailModal);
    } else {
        console.error("Add Lead/Trail Cancel button not found.");
    }

    // Close modal if clicking on the overlay (outside the content box)
    if (addLeadTrailModalOverlay) {
        addLeadTrailModalOverlay.addEventListener('click', (event) => {
            if (event.target === addLeadTrailModalOverlay) {
                hideAddLeadTrailModal();
            }
        });
    } else {
        console.error("Add Lead/Trail Modal Overlay not found.");
    }

    // --- Listeners for 'Add Hyperlinks' Modal ---

    const addHyperlinksOkButton = document.getElementById('addHyperlinksOkButton');
    const addHyperlinksCancelButton = document.getElementById('addHyperlinksCancelButton');
    const addHyperlinksModalOverlay = document.getElementById('addHyperlinksModal');
    const addHyperlinksURL = document.getElementById('url');
    const addHyperlinksURLlabel = document.getElementById('urllabel');
    const addHyperlinksCellURL = document.getElementById('URLcells');

    if (addHyperlinksOkButton) {
        addHyperlinksOkButton.addEventListener('click', onAddHyperlinksOk);
    } else {
        console.error("Add Hyperlinks OK button not found.");
    }

    if (addHyperlinksCancelButton) {
        addHyperlinksCancelButton.addEventListener('click', hideAddHyperlinksModal);
    } else {
        console.error("Add Hyperlinks Cancel button not found.");
    }

    // Listener for checkbox Cell URLs
    if (addHyperlinksCellURL) {
        const handleCellURLChange = () => {
            if (!addHyperlinksCellURL.checked) {
                addHyperlinksURLlabel.classList.add('show');
                addHyperlinksURL.classList.add('show');
                addHyperlinksURL.disabled = false;
                addHyperlinksURL.focus();
            } else {
                addHyperlinksURL.value = '';
                addHyperlinksURLlabel.classList.remove('show');
                addHyperlinksURL.classList.remove('show');
                addHyperlinksURL.disabled = true;
            }
        };
        addHyperlinksCellURL.addEventListener('change', handleCellURLChange);
    } else {
        console.error("Hyperlinks cell URLs not found for checkbox logic.");
    }

    // Close modal if clicking on the overlay (outside the content box)
    if (addHyperlinksModalOverlay) {
        addHyperlinksModalOverlay.addEventListener('click', (event) => {
            if (event.target === addHyperlinksModalOverlay) {
                hideAddHyperlinksModal();
            }
        });
    } else {
        console.error("Add Hyperlinks Modal Overlay not found.");
    }

    // --- Listeners for 'Selection Plus' Modal ---

    const selectionPlusOkButton = document.getElementById('selectionPlusOkButton');
    const selectionPlusCancelButton = document.getElementById('selectionPlusCancelButton');
    const customDelimiterInput = document.getElementById('customDelimiterInput');
    const delimiterSelect = document.getElementById('delimiterSelect');
    const selectionPlusModalOverlay = document.getElementById('selectionPlusModal');

    // Listener for the OK button
    if (selectionPlusOkButton) {
        selectionPlusOkButton.addEventListener('click', selectionPlusOkButtonHandler);
    } else {
        console.error("Selection Plus OK button (inside modal) not found.");
    }

    // Listener for the Cancel button
    if (selectionPlusCancelButton) {
        selectionPlusCancelButton.addEventListener('click', hideSelectionPlusModal);
    } else {
        console.error("Selection Plus Cancel button (inside modal) not found.");
    }

    // Listener for the custom delimiter input toggle
    if (delimiterSelect && customDelimiterInput) {
        const handleDelimiterChange = () => {
            if (delimiterSelect.value === 'custom') {
                customDelimiterInput.classList.add('show');
                customDelimiterInput.disabled = false;
                customDelimiterInput.focus();
            } else {
                customDelimiterInput.classList.remove('show');
                customDelimiterInput.disabled = true;
            }
        };
        delimiterSelect.addEventListener('change', handleDelimiterChange);
    } else {
        console.error("Delimiter select or custom input not found for toggle logic.");
    }

    // Close modal if clicking on the overlay (outside the content box)
    if (selectionPlusModalOverlay) {
        selectionPlusModalOverlay.addEventListener('click', (event) => {
            if (event.target === selectionPlusModalOverlay) {
                hideSelectionPlusModal();
            }
        });
    } else {
        console.error("Selection Plus Modal Overlay not found");
    }

    // --- Listeners for 'Split To Rows' Modal ---

    const splitToRowsOkButton = document.getElementById('splitToRowsOkButton');
    const splitToRowsCancelButton = document.getElementById('splitToRowsCancelButton');
    const SRcustomDelimiterInput = document.getElementById('SRcustomDelimiterInput');
    const SRdelimiterSelect = document.getElementById('SRdelimiterSelect');
    const splitToRowsModalOverlay = document.getElementById('splitToRowsModal');

    // Listener for the OK button
    if (splitToRowsOkButton) {
        splitToRowsOkButton.addEventListener('click', splitToRowsOkButtonHandler);
    } else {
        console.error("Split to Rows OK button (inside modal) not found.");
    }

    // Listener for the Cancel button
    if (splitToRowsCancelButton) {
        splitToRowsCancelButton.addEventListener('click', hideSplitToRowsModal);
    } else {
        console.error("Split to Rows Cancel button (inside modal) not found.");
    }

    // Listener for the custom delimiter input toggle
    if (SRdelimiterSelect && SRcustomDelimiterInput) {
        const handleSRDelimiterChange = () => {
            if (SRdelimiterSelect.value === 'custom') {
                SRcustomDelimiterInput.classList.add('show');
                SRcustomDelimiterInput.disabled = false;
                SRcustomDelimiterInput.focus();
            } else {
                SRcustomDelimiterInput.classList.remove('show');
                SRcustomDelimiterInput.disabled = true;
            }
        };
        SRdelimiterSelect.addEventListener('change', handleSRDelimiterChange);
    } else {
        console.error("Delimiter select or custom input not found for toggle logic.");
    }

    // Close modal if clicking on the overlay (outside the content box)
    if (splitToRowsModalOverlay) {
        splitToRowsModalOverlay.addEventListener('click', (event) => {
            if (event.target === splitToRowsModalOverlay) {
                hideSplitToRowsModal();
            }
        });
    } else {
        console.error("Split to Rows Modal Overlay not found");
    }

    // --- Listeners for 'Find Missing' Modal ---

    const findMissingOkButton = document.getElementById('findMissingOkButton');
    const findMissingCancelButton = document.getElementById('findMissingCancelButton');
    const findMissingModalOverlay = document.getElementById('findMissingModal');

    if (findMissingOkButton) {
        findMissingOkButton.addEventListener('click', findMissingOkButtonHandler);
    } else {
        console.error("Find Missing OK button not found.");
    }

    if (findMissingCancelButton) {
        findMissingCancelButton.addEventListener('click', hideFindMissingModal);
    } else {
        console.error("Find Missing Cancel button not found.");
    }

    // Close modal if clicking on the overlay (outside the content box)
    if (findMissingModalOverlay) {
        findMissingModalOverlay.addEventListener('click', (event) => {
            if (event.target === findMissingModalOverlay) {
                hideFindMissingModal();
            }
        });
    } else {
        console.error("Find Missing Modal Overlay not found.");
    }

    // --- Listeners for 'Compare Worksheets' Modal ---

    const compareSheetsOkButton = document.getElementById('compareSheetsOkButton');
    const compareSheetsCancelButton = document.getElementById('compareSheetsCancelButton');
    const compareSheetsModalOverlay = document.getElementById('compareSheetsModal');

    if (compareSheetsOkButton) {
        compareSheetsOkButton.addEventListener('click', compareSheetsOkButtonHandler);
    } else {
        console.error("Compare Worksheets OK button not found.");
    }

    if (compareSheetsCancelButton) {
        compareSheetsCancelButton.addEventListener('click', hideCompareSheetsModal);
    } else {
        console.error("Compare Worksheets Cancel button not found.");
    }

    // Close modal if clicking on the overlay (outside the content box)
    if (compareSheetsModalOverlay) {
        compareSheetsModalOverlay.addEventListener('click', (event) => {
            if (event.target === compareSheetsModalOverlay) {
                hideCompareSheetsModal();
            }
        });
    } else {
        console.error("Compare Worksheets Modal Overlay not found.");
    }

    // --- Listeners for 'Date/Text Converter' Modal ---

    const dateTextOkButton = document.getElementById('dateTextOkButton');
    const dateTextCancelButton = document.getElementById('dateTextCancelButton');
    const dateTextModalOverlay = document.getElementById('dateTextModal');

    if (dateTextOkButton) {
        dateTextOkButton.addEventListener('click', dateTextOkButtonHandler);
    } else {
        console.error("Date/Text Converter OK button not found.");
    }

    if (dateTextCancelButton) {
        dateTextCancelButton.addEventListener('click', hideDateTextModal);
    } else {
        console.error("Date/Text Converter Cancel button not found.");
    }

    // Locale change event
    const convertLocale = document.getElementById('convertLocale');
    if (convertLocale) {
        convertLocale.addEventListener('change', function () {
            updateFormatOptions(this.value);
        });
    } else {
        console.error("Convert Locale select not found.");
    }

    // Close modal if clicking on the overlay (outside the content box)
    if (dateTextModalOverlay) {
        dateTextModalOverlay.addEventListener('click', (event) => {
            if (event.target === dateTextModalOverlay) {
                hideDateTextModal();
            }
        });
    } else {
        console.error("Date/Text Converter Modal Overlay not found.");
    }


});