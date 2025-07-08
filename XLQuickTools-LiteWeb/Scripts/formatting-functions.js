// --- Formatting Functions ---
// formatting-functions.js


Office.onReady(() => {
    // On Ready
});

// Define constants for the options
const TextTransformOption = {
    UPPERCASE: 'UPPERCASE',
    LOWERCASE: 'LPERCASE',
    PROPERCASE: 'PROPERCASE',
    REMOVE_LETTERS: 'REMOVE_LETTERS',
    REMOVE_NUMBERS: 'REMOVE_NUMBERS',
    REMOVE_SPECIAL: 'REMOVE_SPECIAL'
};

// Clean string
function cleanString(inputString) {
    if (typeof inputString !== 'string') {
        return inputString;
    }
    return inputString.replace(/[\x00-\x1F\x7F]/g, '');
}

// Trim and Clean
function trimAndClean(input) {
    if (input === null || input === undefined) {
        return input;
    }
    if (typeof input !== 'string') {
        input = String(input);
    }
    let processedInput = input;
    processedInput = cleanString(processedInput);
    processedInput = processedInput.replace(/\s{2,}/g, ' ');
    processedInput = processedInput.trim();
    if (processedInput.length === 0) {
        processedInput = '';
    }
    return processedInput;
}

// Text options
function transformText(input, option) {
    if (typeof input !== 'string') {
        // Convert non-string inputs to string before processing
        input = String(input);
    }

    switch (option) {
        case TextTransformOption.UPPERCASE:
            return input.toUpperCase();
        case TextTransformOption.LOWERCASE:
            return input.toLowerCase();
        case TextTransformOption.PROPERCASE:
            return input.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
        case TextTransformOption.REMOVE_LETTERS:
            return input.replace(/\p{L}/gu, '');
        case TextTransformOption.REMOVE_NUMBERS:
            return input.replace(/[0-9]/g, '');
        case TextTransformOption.REMOVE_SPECIAL:
            return input.replace(/[^a-zA-Z0-9\u00C0-\u024F\s]/gu, '');
        default:
            console.warn(`Unknown transformation option: ${option}. Returning original input.`);
            return input;
    }
}

// Trim and Clean process values
function processValues(values) {
    if (!values || values.length === 0) {
        return [];
    }
    const newValues = [];
    for (let i = 0; i < values.length; i++) {
        newValues[i] = [];
        for (let j = 0; j < values[i].length; j++) {
            newValues[i][j] = trimAndClean(values[i][j]);
        }
    }
    return newValues;
}

// Process range for add/lead, trim
async function processExcelRange(context, range, processorFunction) {
    range.load("values");
    await context.sync();

    const oldValues = range.values;
    const newValues = [];
    for (let i = 0; i < oldValues.length; i++) {
        newValues[i] = [];
        for (let j = 0; j < oldValues[i].length; j++) {
            newValues[i][j] = processorFunction(oldValues[i][j]);
        }
    }

    range.values = newValues;
    await context.sync();
}

// Function to fill blank cells with values from above for any given range
async function fillBlanksInRange(context, fillRange, showMessages = false) {
    try {
        // Validate the input range
        if (!fillRange) {
            if (showMessages) {
                showModalMessage("Fill Down", "No valid effective range found for filling blanks.", false);
            }
            return 0;
        }

        // Load necessary properties
        fillRange.load("values, rowCount, columnCount, address");
        await context.sync();

        if (fillRange.rowCount === 0 || fillRange.columnCount === 0) {
            if (showMessages) {
                showModalMessage("Fill Down", "Selected range is empty.", false);
            }
            return 0;
        }

        const values = fillRange.values;
        let cellsUpdated = 0;

        // Create a copy of values to modify
        const newValues = values.map(row => [...row]);

        // Iterate through each cell in the range (skip first row since it can't fill from above)
        for (let r = 1; r < fillRange.rowCount; r++) {
            for (let c = 0; c < fillRange.columnCount; c++) {
                const cellValue = values[r][c];
                // Check if the cell is blank
                if (isBlank(cellValue)) {
                    // Get the value from the row above
                    const valueFromAbove = newValues[r - 1][c];
                    // Only fill if the value above is not blank
                    if (!isBlank(valueFromAbove)) {
                        newValues[r][c] = valueFromAbove;
                        cellsUpdated++;
                    }
                }
            }
        }

        if (cellsUpdated === 0) {
            if (showMessages) {
                showModalMessage("Fill Down", "No blank cells found to fill in the selected range.", false);
            }
            return 0;
        }

        // Update the entire range at once with the new values
        fillRange.values = newValues;
        await context.sync();

        if (showMessages) {
            showModalMessage("Fill Down", `Updates made to ${cellsUpdated} cells.`, false);
        }

        return cellsUpdated;

    } catch (error) {
        console.error("Error in fillBlanksInRange:", error);
        if (showMessages) {
            showModalMessage("Fill Down", "An error occurred while filling blanks. Please try again.", false);
        }
        throw error;
    }
}

// Fill in blanks from above
async function fillBlanksFromAbove() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            const fillRange = await getEffectiveRangeForSelection(context, selectedRange);

            // Use the shared helper function with messages enabled
            await fillBlanksInRange(context, fillRange, true);
        });
    } catch (error) {
        console.error("Error filling blanks:", error);
        showModalMessage("Fill Down (Blanks)", "An error occurred while filling blanks. Please try again.", false);
    }
}

// Helper function to check if a cell value is blank
function isBlank(value) {
    return value === null || value === undefined || value === "" ||
        (typeof value === 'string' && value.trim() === "");
}

// Split to Rows
async function splitToRows(headers, delimiter) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();

        // Load all necessary properties upfront
        usedRange.load(['values', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);

        await context.sync();

        // Check if usedRange is valid
        if (!usedRange || usedRange.rowCount === 0) {
            console.log("No data found in worksheet.");
            return;
        }

        let startRow = usedRange.rowIndex;
        let endRow = startRow + usedRange.rowCount - 1;
        let startCol = usedRange.columnIndex;
        let endCol = startCol + usedRange.columnCount - 1;

        // Adjust startRow if headers are present
        if (headers) {
            startRow++;
            if (startRow > endRow) {
                console.log("No data rows to process after skipping headers.");
                return;
            }
        }

        let changeCount = 0;

        // Process from bottom to top to avoid index shifting
        for (let iLoop = endRow; iLoop >= startRow; iLoop--) {
            let maxSplitsInRow = 0;

            // First pass: determine splits needed
            for (let i = startCol; i <= endCol; i++) {
                const cellValue = usedRange.values[iLoop - usedRange.rowIndex][i - usedRange.columnIndex];

                if (cellValue !== null && cellValue !== undefined && String(cellValue).length > 0) {
                    const parts = String(cellValue).split(delimiter);
                    const splitsCount = parts.length - 1;

                    if (splitsCount > maxSplitsInRow) {
                        maxSplitsInRow = splitsCount;
                        changeCount++;
                    }
                }
            }

            // Insert rows if needed
            if (maxSplitsInRow > 0) {
                try {
                    // Create a range for the entire row to insert
                    const insertRange = sheet.getRangeByIndexes(
                        iLoop + 1,      // Row to insert at
                        0,              // Start at column 0 (entire row)
                        maxSplitsInRow, // Number of rows to insert
                        16384           // Excel's maximum columns
                    );

                    insertRange.insert(Excel.InsertShiftDirection.down);
                    await context.sync();
                } catch (insertError) {
                    console.error(`Error inserting rows at ${iLoop + 1}:`, insertError);
                    // Try with just the used range columns as fallback
                    const insertRange = sheet.getRangeByIndexes(
                        iLoop + 1,
                        startCol,
                        maxSplitsInRow,
                        usedRange.columnCount
                    );
                    insertRange.insert(Excel.InsertShiftDirection.down);
                    await context.sync();
                }
            }

            // Second pass: populate cells
            for (let i = startCol; i <= endCol; i++) {
                const cellValue = usedRange.values[iLoop - usedRange.rowIndex][i - usedRange.columnIndex];

                if (cellValue !== null && cellValue !== undefined && String(cellValue).length > 0) {
                    const parts = String(cellValue).split(delimiter);

                    for (let jLoop = 0; jLoop < parts.length; jLoop++) {
                        const cleanedValue = cleanString(parts[jLoop].trim());
                        const targetCell = sheet.getCell(iLoop + jLoop, i);
                        targetCell.values = [[cleanedValue]];
                    }
                } else {
                    const targetCell = sheet.getCell(iLoop, i);
                    targetCell.values = [['']];
                }
            }

            await context.sync();
        }

        // Final sync after the loop to ensure all operations are applied
        await context.sync();
        // Fill in the blanks
        const currentUsedRange = sheet.getUsedRange();
        await fillBlanksInRange(context, currentUsedRange);

    }).catch(function (error) {
        console.error("Error in splitToRows:", error);
    });
}

// Selection Plus
async function selectionPlus(leading, trailing, delimiter) {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const rangeToProcess = await getEffectiveRangeForSelection(context, selectedRange);

        if (!rangeToProcess) {
            console.warn("No valid range found to process for Selection Plus.");
            return;
        }

        // Load the text (display string for each cell), rowCount, and columnCount from the range
        rangeToProcess.load("text, rowCount, columnCount");
        await context.sync();

        const displayedTextValues = rangeToProcess.text;
        const concatenatedStringParts = [];

        // Iterate through the 2D array of values
        for (let r = 0; r < rangeToProcess.rowCount; r++) {
            for (let c = 0; c < rangeToProcess.columnCount; c++) {
                let cellDisplayedValue = displayedTextValues[r][c];

                // Convert to string and trim whitespace, handling null/undefined
                if (cellDisplayedValue !== null && cellDisplayedValue !== undefined) {
                    cellDisplayedValue = String(cellDisplayedValue).trim();
                } else {
                    cellDisplayedValue = ""; // Treat null/undefined as empty string
                }

                // Only append if cellDisplayedValue is not empty
                if (cellDisplayedValue !== "") {
                    // Append delimiter only if this is not the first value
                    if (concatenatedStringParts.length > 0) {
                        concatenatedStringParts.push(delimiter);

                        // Avoid space for actual newlines if delimiter is a newline
                        if (delimiter !== "\n" && delimiter !== "\r" && delimiter !== "\r\n") {
                            concatenatedStringParts.push(" ");
                        }
                    }
                    // Append leading and trailing character
                    concatenatedStringParts.push(leading + cellDisplayedValue + trailing);
                }
            }
        }

        const finalString = concatenatedStringParts.join('');

        // Copy to clipboard if any non-empty values were found
        if (finalString.length > 0) {
            await navigator.clipboard.writeText(finalString);
        }
    });
}

// Delete empty rows
async function deleteEmptyRows() {
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = worksheet.getUsedRange();

            if (!usedRange) {
                showModalMessage("Delete Empty Rows", "No data found in worksheet!", false);
                return;
            }

            usedRange.load("values, rowCount");
            await context.sync();

            if (usedRange.rowCount <= 1) {
                return;
            }

            const values = usedRange.values;
            const emptyRowIndices = [];

            // Identify empty rows
            for (let i = values.length - 1; i >= 0; i--) {
                const isEmptyRow = values[i].every(cell =>
                    cell === null ||
                    cell === undefined ||
                    cell === "" ||
                    (typeof cell === "string" && cell.trim() === "")
                );

                if (isEmptyRow) {
                    emptyRowIndices.push(i);
                }
            }

            // Delete empty rows in batches (from bottom to top)
            for (const rowIndex of emptyRowIndices) {
                const row = usedRange.getRow(rowIndex);
                const entireRow = row.getEntireRow();
                entireRow.delete(Excel.DeleteShiftDirection.up);
            }

            await context.sync();
            showModalMessage("Delete Empty Rows", `Deleted ${emptyRowIndices.length} empty rows.`, false);
        });
    } catch (error) {
        console.error("Error deleting empty rows:", error);
    }
}

// Delete empty columns
async function deleteEmptyColumns() {
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = worksheet.getUsedRange();

            if (!usedRange) {
                showModalMessage("Delete Empty Columns", "No data found in worksheet!", false);
                return;
            }

            usedRange.load("values, columnCount");
            await context.sync();

            if (usedRange.columnCount <= 1) {
                return;
            }

            const values = usedRange.values;
            const emptyColumnIndices = [];

            // Identify empty columns
            for (let col = values[0].length - 1; col >= 0; col--) {
                const isEmptyColumn = values.every(row =>
                    row[col] === null ||
                    row[col] === undefined ||
                    row[col] === "" ||
                    (typeof row[col] === "string" && row[col].trim() === "")
                );

                if (isEmptyColumn) {
                    emptyColumnIndices.push(col);
                }
            }

            // Delete empty columns in batches (from right to left)
            for (const columnIndex of emptyColumnIndices) {
                const column = usedRange.getColumn(columnIndex);
                const entireColumn = column.getEntireColumn();
                entireColumn.delete(Excel.DeleteShiftDirection.left);
            }

            await context.sync();
            showModalMessage("Delete Empty Columns", `Deleted ${emptyColumnIndices.length} empty columns`, false);
        });
    } catch (error) {
        console.error("Error deleting empty columns:", error);
    }
}

// Trim and clean selected
async function trimCleanSelected() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            const rangeToProcess = await getEffectiveRangeForSelection(context, selectedRange);

            if (rangeToProcess) {
                await processExcelRange(context, rangeToProcess, trimAndClean);
            }
        });
    } catch (error) {
        console.error("Error during Trim & Clean (Selected):", error);
    }
}

// Trim and clean worksheet
async function trimCleanSheet() {
    try {
        await Excel.run(async (context) => {
            const activeSheetUsedRange = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
            activeSheetUsedRange.load("address, cellCount");
            await context.sync();

            if (activeSheetUsedRange.cellCount === 0) {
                return;
            }

            await processExcelRange(context, activeSheetUsedRange, trimAndClean);
            showModalMessage("Trim and Clean", "All cells in the active worksheet have been trimmed and cleaned.", false);
        });
    } catch (error) {
        console.error("Error during Trim & Clean (Sheet):", error);
    }
}

// Trim and clean workbook
async function trimCleanWorkbook() {
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items");
            await context.sync();

            for (const sheet of worksheets.items) {
                const sheetUsedRange = sheet.getUsedRange();
                sheetUsedRange.load("values");
                await context.sync();

                if (sheetUsedRange.values && sheetUsedRange.values.length > 0) {
                    const sheetValues = sheetUsedRange.values;
                    const newSheetValues = processValues(sheetValues);
                    sheetUsedRange.values = newSheetValues;
                    await context.sync();
                }
            }
            showModalMessage("Trim and Clean", "All cells in the active workbook have been trimmed and cleaned.", false);
        });
    } catch (error) {
        console.error("Error during Trim & Clean (Workbook):", error);
    }
}

// Get text option
async function getTextOptions(option) {
    if (typeof Excel === 'undefined' || !Excel.run) {
        console.error("Excel object or Excel.run is not available.");
        return;
    }

    try {
        await Excel.run(async (context) => {
            await applyTextOptionsToSelection(context, option);
        });
    } catch (error) {
        console.error("Error during Excel operation for text transformation:", error);
    }
}

// Apply text options
async function applyTextOptionsToSelection(context, option) {
    try {
        const selectedRange = context.workbook.getSelectedRange();
        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);

        if (!effectiveRange) {
            return;
        }

        await processExcelRange(context, effectiveRange, (cellValue) => transformText(cellValue, option));

    } catch (error) {
        console.error("Error applying text transformation:", error);
    }
}

// Add leading and trailing text to selected range
async function addLeaadTrail(leadingText, trailingText) {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);

        if (!effectiveRange) {
            showModalMessage("", "No effective range found for the current selection.", false);
            return;
        }

        await processExcelRange(context, effectiveRange, (cellValue) => {
            // Ensure cellValue is treated as a string before concatenation
            let processedValue = String(cellValue);
            return leadingText + processedValue + trailingText;
        });

    });
}


/*// Convert dates
async function convertSelectedDates(currentLocale, convertLocale, format, type) {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);

        if (!effectiveRange) {
            showModalMessage("", "No effective range found for the current selection.", false);
            return;
        }

        // Load values and numberFormat
        effectiveRange.load("values, numberFormat");
        await context.sync();

        const values = effectiveRange.values;
        const numberFormats = effectiveRange.numberFormat;
        const newValues = [];
        const newFormats = [];

        for (let row = 0; row < values.length; row++) {
            const newRow = [];
            const formatRow = [];

            for (let col = 0; col < values[row].length; col++) {
                const cellValue = values[row][col];
                let convertedResult;

                // Call the Excel built-in function conversion methods
                if (type === "text") {
                    convertedResult = await convertToText(context, cellValue, format, currentLocale);
                } else {
                    convertedResult = await convertToSerial(context, cellValue, format, currentLocale);
                }

                newRow.push(convertedResult.value);
                formatRow.push(convertedResult.format);
            }
            newValues.push(newRow);
            newFormats.push(formatRow);
        }

        // Apply new number formats first.
        effectiveRange.numberFormat = newFormats;
        await context.sync();

        // Then apply the new values.
        effectiveRange.values = newValues;
        await context.sync();
    });
}

// Format text using Excels built in function
async function convertToText(context, value, format, currentLocale) {
    if (value == null || value === undefined || value === "") {
        return { value: value, format: "@" }; // Return original empty value, format as text
    }

    let serialDate;
    if (typeof value === "number") {
        // If it's already a serial number, use it directly
        serialDate = value;
    } else if (typeof value === "string") {
        // If it's a string, try to parse it into an Excel serial date
        try {
            const dateValueResult = context.workbook.functions.datevalue(value);
            dateValueResult.load('value');
            await context.sync();
            serialDate = dateValueResult.value;

            // Check for DATEVALUE error (e.g., #VALUE! error in Excel)
            if (typeof serialDate !== 'number' || isNaN(serialDate)) {
                console.warn(`Could not convert string "${value}" to date serial using DATEVALUE.`);
                return { value: value, format: "@" };
            }
        } catch (error) {
            console.error(`Error parsing string "${value}" to serial date:`, error);
            return { value: value, format: "@" }; // On error, return original value as text
        }
    } else {
        return { value: value, format: "@" };
    }

    // Now format the serial date as text using the TEXT function
    const formattedTextResult = context.workbook.functions.text(serialDate, format);
    formattedTextResult.load('value');
    await context.sync();

    // The TEXT function also returns an error string (e.g., #VALUE!) if input is invalid.
    if (typeof formattedTextResult.value !== 'string' || formattedTextResult.value.startsWith('#')) {
        console.warn(`TEXT function failed to format serial ${serialDate} with format "${format}". Result: ${formattedTextResult.value}`);
        return { value: value, format: "@" }; // Return original value if formatting fails
    }

    // console.log(`Excel Serial: ${serialDate}, Formatted to: ${formattedTextResult.value}`);
    return { value: formattedTextResult.value, format: "@" };
}


// Convert to Excel serial date using built in function
async function convertToSerial(context, value, format, currentLocale) {
    if (value == null || value === undefined || value === "") {
        return { value: value, format: "General" }; // Return original empty value, general format
    }

    if (typeof value === "number") {
        // If it's already a serial number, return it with the desired date format
        return { value: value, format: `${format}` };
    } else if (typeof value === "string") {
        try {
            // Use DATEVALUE to convert the string to a serial number.
            const dateValueResult = context.workbook.functions.datevalue(value);
            dateValueResult.load('value');
            await context.sync();

            const serialDate = dateValueResult.value;

            // Check for DATEVALUE error (e.g., #VALUE! error in Excel)
            if (typeof serialDate !== 'number' || isNaN(serialDate)) {
                console.warn(`Could not convert string "${value}" to date serial using DATEVALUE.`);
                return { value: value, format: "General" }; // Return original value, general format
            }

            // console.log(`String "${value}" converted to Excel Serial: ${serialDate}`);
            // Return the serial number with a standard Excel date format
            return { value: serialDate, format: `${format}` };
        } catch (error) {
            console.error(`Error parsing string "${value}" to serial date:`, error);
            return { value: value, format: "General" }; // On error, return original value, general format
        }
    } else {
        return { value: value, format: "General" }; // Handle other types as general
    }
}*/




// Convert dates
/*async function convertSelectedDates(currentLocale, convertLocale, format, type) {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);

        if (!effectiveRange) {
            showModalMessage("", "No effective range found for the current selection.", false);
            return;
        }

        // Load values and numberFormat in one go
        effectiveRange.load("values"); // We only need values, numberFormat will be set based on conversion
        await context.sync();

        const values = effectiveRange.values;
        const newValues = [];
        const newFormats = [];

        // Prepare a batch of Excel function calls (DATEVALUE or TEXT)
        const functionsToExecute = [];
        const functionCallMap = new Map(); // To link results back to original cell positions

        for (let row = 0; row < values.length; row++) {
            const newRow = [];
            const formatRow = [];

            for (let col = 0; col < values[row].length; col++) {
                const cellValue = values[row][col];

                if (cellValue == null || cellValue === undefined || cellValue === "") {
                    // Handle empty cells directly
                    newRow.push(cellValue);
                    formatRow.push(type === "text" ? "@" : "General");
                    continue;
                }

                if (type === "text") {
                    // For converting to text
                    if (typeof cellValue === "number") {
                        // If already a serial number, prepare TEXT function
                        const textResult = context.workbook.functions.text(cellValue, format);
                        functionsToExecute.push(textResult);
                        functionCallMap.set(textResult, { row, col, originalValue: cellValue, type: "text" });
                    } else if (typeof cellValue === "string") {
                        // If string, first DATEVALUE, then TEXT
                        const dateValueResult = context.workbook.functions.datevalue(cellValue);
                        functionsToExecute.push(dateValueResult);
                        functionCallMap.set(dateValueResult, { row, col, originalValue: cellValue, type: "text" });
                    } else {
                        // Other types, just push original value as text
                        newRow.push(cellValue);
                        formatRow.push("@");
                    }
                } else { // type === "serial"
                    // For converting to serial
                    if (typeof cellValue === "number") {
                        // If already a serial number, no function call needed
                        newRow.push(cellValue);
                        formatRow.push(format);
                    } else if (typeof cellValue === "string") {
                        // If string, prepare DATEVALUE function
                        const dateValueResult = context.workbook.functions.datevalue(cellValue);
                        functionsToExecute.push(dateValueResult);
                        functionCallMap.set(dateValueResult, { row, col, originalValue: cellValue, type: "serial" });
                    } else {
                        // Other types, just push original value as general
                        newRow.push(cellValue);
                        formatRow.push("General");
                    }
                }
            }
            newValues.push(newRow); // Temporarily push empty rows, will fill later
            newFormats.push(formatRow);
        }

        // Load all batched function results in one go
        functionsToExecute.forEach(func => func.load('value'));
        await context.sync();

        // Process the results from the batched function calls
        for (const funcResult of functionsToExecute) {
            const { row, col, originalValue, type: conversionType } = functionCallMap.get(funcResult);
            let convertedValue = funcResult.value;
            let finalFormat;

            if (conversionType === "text") {
                if (typeof originalValue === "string" && !isNaN(convertedValue) && typeof convertedValue === 'number') {
                    // This was a DATEVALUE result for a string, now we need to TEXT format it
                    const textResult = context.workbook.functions.text(convertedValue, format);
                    // We need another sync to get this result. This is a potential point of optimization if many strings.
                    // For now, doing it individually for clarity, but can be batched again.
                    textResult.load('value');
                    await context.sync(); // **** One of the few syncs within the loop now ****
                    convertedValue = textResult.value;
                }

                if (typeof convertedValue !== 'string' || String(convertedValue).startsWith('#')) {
                    // If TEXT/DATEVALUE failed, revert to original value and format as text
                    console.warn(`Conversion to text failed for "${originalValue}". Result: ${convertedValue}`);
                    newValues[row][col] = originalValue;
                    newFormats[row][col] = "@";
                } else {
                    newValues[row][col] = convertedValue;
                    newFormats[row][col] = "@";
                }
            } else { // conversionType === "serial"
                if (typeof convertedValue !== 'number' || isNaN(convertedValue) || String(convertedValue).startsWith('#')) {
                    // If DATEVALUE failed, revert to original value and format as general
                    console.warn(`Conversion to serial failed for "${originalValue}". Result: ${convertedValue}`);
                    newValues[row][col] = originalValue;
                    newFormats[row][col] = "General";
                } else {
                    newValues[row][col] = convertedValue;
                    newFormats[row][col] = format;
                }
            }
        }

        // Apply new number formats and values in two separate, batched calls
        effectiveRange.numberFormat = newFormats;
        effectiveRange.values = newValues;
        await context.sync(); // Final sync to write all changes to Excel

        showModalMessage("", "Date conversion complete!", true);
    });
}*/

// ALTERNATIVE APPROACH: Use native JavaScript date parsing for even better performance
// This approach avoids Excel functions entirely for better speed
/*async function convertSelectedDates(currentLocale, convertLocale, format, type) {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);

        if (!effectiveRange) {
            showModalMessage("", "No effective range found for the current selection.", false);
            return;
        }

        effectiveRange.load("values, numberFormat");
        await context.sync();

        const values = effectiveRange.values;
        const newValues = [];
        const newFormats = [];

        for (let row = 0; row < values.length; row++) {
            const newRow = [];
            const formatRow = [];

            for (let col = 0; col < values[row].length; col++) {
                const cellValue = values[row][col];
                let convertedResult;

                if (type === "text") {
                    convertedResult = convertToText(cellValue, format);
                } else {
                    convertedResult = convertToSerial(cellValue, format);
                }

                newRow.push(convertedResult.value);
                formatRow.push(convertedResult.format);
            }
            newValues.push(newRow);
            newFormats.push(formatRow);
        }

        // Single batch update
        effectiveRange.numberFormat = newFormats;
        effectiveRange.values = newValues;
        await context.sync();

        showModalMessage("Date/Text Converter", "Date conversion complete!", false);
    });
}

// Native JavaScript date conversion functions with improved parsing
function convertToText(value, format) {
    if (value == null || value === undefined || value === "") {
        return { value: value, format: "@" };
    }

    let date;
    if (typeof value === "number") {
        // Convert Excel serial date to JavaScript Date using your method
        const msSinceEpoch = (value - 25569) * 86400000;
        date = new Date(msSinceEpoch);
        // Adjust to UTC midnight to avoid timezone discrepancies
        date = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    } else if (typeof value === "string") {
        // Parse as date string using your method
        date = new Date(value);
        if (isNaN(date.getTime())) {
            return { value: value, format: "@" };
        }
        date = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    } else {
        return { value: value, format: "@" };
    }

    // Format the date using your formatting logic
    const formattedDate = formatDateToString(date, format);
    return { value: formattedDate, format: "@" };
}

function convertToSerial(value, format) {
    if (value == null || value === undefined || value === "") {
        return { value: value, format: "General" };
    }

    if (typeof value === "number") {
        return { value: value, format: format };
    } else if (typeof value === "string") {
        // Parse as date string using your method
        const date = new Date(value);
        if (isNaN(date.getTime())) {
            return { value: value, format: "General" };
        }
        // Adjust to UTC midnight to avoid timezone discrepancies
        const adjustedDate = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
        const jsSerial = adjustedDate.getTime() / 86400000;
        const excelSerial = Math.round(jsSerial + 25569);
        return { value: excelSerial, format: format };
    } else {
        return { value: value, format: "General" };
    }
}

function formatDateToString(date, format) {
    const year = date.getUTCFullYear();
    const month = date.getUTCMonth() + 1;
    const day = date.getUTCDate();

    // Month names
    const monthNames = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    const monthNamesShort = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ];

    // Format mapping based on your format options
    switch (format) {
        case "yyyy-MM-dd":
            return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        case "M/d/yyyy":
            return `${month}/${day}/${year}`;
        case "MM/dd/yyyy":
            return `${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}/${year}`;
        case "dd/MM/yyyy":
            return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${year}`;
        case "d/M/yyyy":
            return `${day}/${month}/${year}`;
        case "MMM dd, yyyy":
            return `${monthNamesShort[month - 1]} ${String(day).padStart(2, '0')}, ${year}`;
        case "MMMM dd, yyyy":
            return `${monthNames[month - 1]} ${String(day).padStart(2, '0')}, ${year}`;
        case "dd MMM yyyy":
            return `${String(day).padStart(2, '0')} ${monthNamesShort[month - 1]} ${year}`;
        case "dd MMMM yyyy":
            return `${String(day).padStart(2, '0')} ${monthNames[month - 1]} ${year}`;
        case "yyyy/MM/dd":
            return `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
        case "yyyy.MM.dd":
            return `${year}.${String(month).padStart(2, '0')}.${String(day).padStart(2, '0')}`;
        case "yyyy MMM dd":
            return `${year} ${monthNamesShort[month - 1]} ${String(day).padStart(2, '0')}`;
        default:
            // Default to yyyy-MM-dd if format not recognized
            return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
}*/