// --- Main Functions ---
// App.js


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

// Function to convert column index to Excel column letter
function getColumnLetter(columnIndex) {
    let result = '';
    while (columnIndex >= 0) {
        result = String.fromCharCode(65 + (columnIndex % 26)) + result;
        columnIndex = Math.floor(columnIndex / 26) - 1;
    }
    return result;
}

// Get the range for selection
async function getEffectiveRangeForSelection(context, selectedRange) {
    try {
        selectedRange.load("address, cellCount, isEntireColumn, isEntireRow, columnCount");
        await context.sync();

        if (selectedRange.isEntireColumn || selectedRange.isEntireRow) {
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await context.sync();

            let usedRangeInSheet;
            try {
                usedRangeInSheet = activeSheet.getUsedRange();
                usedRangeInSheet.load("address");
                await context.sync();
            } catch (error) {
                console.warn("No used range found in worksheet:", error);
                return null;
            }

            if (usedRangeInSheet && usedRangeInSheet.address) {
                try {
                    const intersectionRange = selectedRange.getIntersection(usedRangeInSheet);
                    // Ensure all necessary properties are loaded for intersectionRange
                    intersectionRange.load("address, cellCount, values, rowCount, columnIndex, rowIndex, columnCount");
                    await context.sync();

                    if (intersectionRange.cellCount === 0) {
                        return null;
                    }
                    return intersectionRange;
                } catch (error) {
                    console.warn("Failed to get intersection range:", error);
                    return null;
                }
            } else {
                return null;
            }
        } else {
            // Ensure all necessary properties are loaded for selectedRange itself
            selectedRange.load("address, cellCount, values, rowCount, columnIndex, rowIndex, columnCount");
            await context.sync();

            if (selectedRange.cellCount === 0) {
                return null;
            }
            return selectedRange;
        }
    } catch (error) {
        console.error("Error getting effective range:", error);
        return null;
    }
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
// Column Information
async function countUniqueValuesInColumn() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();

            // Check if entire column is selected
            selectedRange.load("isEntireColumn, address, columnIndex");
            await context.sync();

            if (!selectedRange.isEntireColumn) {
                showModalMessage("Column Information", "Please select an entire column to use.", false);
                return;
            }

            const fillRange = await getEffectiveRangeForSelection(context, selectedRange);

            // Validate the effective range
            if (!fillRange) {
                showModalMessage("Column Information", "No data found in the selected column.", false);
                return;
            }

            // Load necessary properties
            fillRange.load("values, rowCount, columnCount, address");
            await context.sync();

            if (fillRange.rowCount === 0 || fillRange.columnCount === 0) {
                showModalMessage("Column Information", "No data found in the selected column.", false);
                return;
            }

            const values = fillRange.values;
            const uniqueValues = new Set();

            // Iterate through all cells in the column
            for (let r = 0; r < fillRange.rowCount; r++) {
                for (let c = 0; c < fillRange.columnCount; c++) {
                    const cellValue = values[r][c];

                    // Only count non-blank values
                    if (!isBlank(cellValue)) {
                        // Convert to string to ensure proper comparison
                        const stringValue = String(cellValue).trim();
                        if (stringValue !== "") {
                            uniqueValues.add(stringValue);
                        }
                    }
                }
            }

            const uniqueCount = uniqueValues.size;
            const totalNonBlankCells = Array.from(values.flat()).filter(val => !isBlank(val)).length;
            const totalBlankCells = Array.from(values.flat()).filter(val => isBlank(val)).length;

            // Html to make message look visually better
            showModalMessage(
                "Column Information",
                `Column ${getColumnLetter(selectedRange.columnIndex)} contains:<br>
                <table style="width:100%; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 3px 0; padding-left: 30px;">Unique Values:</td>
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${uniqueCount}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 0; padding-left: 30px;">Total Non-Blank Cells:</td>
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${totalNonBlankCells}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 0; padding-left: 30px;">Total Blank Cells:</td>
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${totalBlankCells}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 0; padding-left: 30px;">Total Rows with Data:</td>
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${fillRange.rowCount}</td>
                    </tr>
                </table>`,
                false
            );
        });
    } catch (error) {
        console.error("Error counting unique values:", error);
        showModalMessage("Column Information", "An error occurred while grabbing information. Please try again.", false);
    }
}

// Check for duplicates
async function findAndCountDuplicatesInColumn() {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();

        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);

        if (!effectiveRange) {
            showModalMessage("Check for Duplicates", "Please select a column with data.", false);
            return;
        }

        // Load all necessary properties for effectiveRange, including 'worksheet'
        effectiveRange.load("address, columnCount, rowCount, columnIndex, rowIndex, values, cellCount, worksheet");
        await context.sync();

        // Validate the worksheet object
        const worksheet = effectiveRange.worksheet;
        if (!worksheet || !(worksheet instanceof Excel.Worksheet)) {
            showModalMessage("Error", "Could not access the active worksheet. Please try again.", false);
            return;
        }

        // Check if it's a single column and contains more than one cell
        if (effectiveRange.columnCount !== 1 || effectiveRange.cellCount <= 1) {
            showModalMessage("Check for Duplicates", "Please select a single column that contains more than one cell of data to check for duplicates.", false);
            return;
        }

        // Check if effectiveRange.values is not null/undefined/empty
        if (!effectiveRange.values || effectiveRange.values.length === 0) {
            showModalMessage("Check for Duplicates", "The selected column does not contain any data to check.", false);
            return;
        }

        // Determine if we have a header (starting at row 0)
        const hasHeader = effectiveRange.rowIndex === 0;

        // Get column values, excluding header if present
        const columnValues = effectiveRange.values.map(row => row[0]);
        const dataValues = hasHeader ? columnValues.slice(1) : columnValues; // Skip first row if it's a header

        const counts = {};
        let hasDuplicates = false;

        // Only count duplicates in data values (excluding header)
        dataValues.forEach(value => {
            const stringValue = String(value);
            counts[stringValue] = (counts[stringValue] || 0) + 1;
            if (counts[stringValue] > 1) {
                hasDuplicates = true;
            }
        });

        // Create count array that matches original data structure (including header position)
        const countValues = columnValues.map((value, index) => {
            if (hasHeader && index === 0) {
                return [0]; // Header gets 0 count
            } else {
                return [counts[String(value)] || 0];
            }
        });

        if (hasDuplicates) {
            const targetColumnIndex = effectiveRange.columnIndex + 1;

            try {
                // Get the entire column range where we want to insert
                const insertRange = worksheet.getRange(`${getColumnLetter(targetColumnIndex)}:${getColumnLetter(targetColumnIndex)}`);
                insertRange.insert(Excel.InsertShiftDirection.right);
                await context.sync();

            } catch (error) {
                showModalMessage("Error", `Failed to insert a new column. Error: ${error.message}. Please ensure there is enough space.`, false);
                return; // Stop execution if column insertion fails
            }

            // Determine the starting row for the count data
            let countDataStartRow = effectiveRange.rowIndex;
            let countDataRowCount = effectiveRange.rowCount;

            // Handle header placement and data shifting
            if (effectiveRange.rowIndex === 0) {
                const headerCell = worksheet.getCell(0, targetColumnIndex);
                headerCell.values = [['Count']];
                countDataStartRow = 1; // Data starts from the next row
                countDataRowCount = effectiveRange.rowCount - 1; // Exclude header row from data count
            } else {
                const headerCell = worksheet.getCell(effectiveRange.rowIndex - 1, targetColumnIndex);
                headerCell.values = [['Count']];
            }

            // Set the count values for the data rows only
            if (countDataRowCount > 0) {
                const countColumnRange = worksheet.getRangeByIndexes(countDataStartRow, targetColumnIndex, countDataRowCount, 1);
                const dataCountValues = hasHeader ? countValues.slice(1) : countValues; // Skip header count if present
                countColumnRange.values = dataCountValues;
                await context.sync(); // Sync after setting values
            }

            // --- FILTER APPLICATION ---
            let filterRangeStartRow = effectiveRange.rowIndex;
            if (effectiveRange.rowIndex > 0) {
                filterRangeStartRow = effectiveRange.rowIndex - 1;
            } else {
                filterRangeStartRow = 0;
            }

            const totalRowCountForFilter = effectiveRange.rowCount + (effectiveRange.rowIndex > 0 ? 1 : 0);
            const numberOfColumnsForFilter = targetColumnIndex - effectiveRange.columnIndex + 1;

            const filterRange = worksheet.getRangeByIndexes(
                filterRangeStartRow,
                effectiveRange.columnIndex,
                totalRowCountForFilter,
                numberOfColumnsForFilter
            );

            // Load the filterRange address property before using it
            filterRange.load("address");
            await context.sync();

            // Clear any existing AutoFilter first to avoid conflicts
            if (worksheet.autoFilter) {
                worksheet.autoFilter.remove();
                await context.sync();
            }

            try {
                // Enable AutoFilter to used range
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const usedRange = sheet.getUsedRange();
                sheet.autoFilter.apply(usedRange);

                //filterRange.worksheet.autoFilter.apply(filterRange);
                await context.sync();

            } catch (filterError) {
                console.error("Error enabling filter dropdowns:", filterError);
                console.log("Continuing without filter dropdowns...");
            }

            showModalMessage("Check for Duplicates", "Duplicates found. Count column added.", false);

        } else {
            showModalMessage("Check for Duplicates", "No duplicates found in the selected column.", false);
        }
    });
}

// Mimic the autofilter so the user has another place they can disable/enable it
async function toggleAutoFilter() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Load the 'enabled' property of the autoFilter object to check its current state.
        sheet.autoFilter.load("enabled");

        // Get the used range and load its cellCount.
        const usedRange = sheet.getUsedRange();
        usedRange.load("cellCount");

        // Synchronize the context to get the loaded property values.
        await context.sync();

        // Check if the AutoFilter is currently enabled.
        if (sheet.autoFilter.enabled) {
            // If enabled, remove the AutoFilter.
            sheet.autoFilter.remove();
        } else {
            if (usedRange.cellCount > 1) {
                // Apply the AutoFilter to the entire used range of the sheet.
                sheet.autoFilter.apply(usedRange);
            }
        }

        // Synchronize the context again to apply the changes to the worksheet.
        await context.sync();
    }).catch((error) => {
        console.error("An unexpected error occurred while toggling AutoFilter:", error);
    });
}

// Highlight missing data cells
async function highlightMissingCells(parentRange, missingValues) {
    await Excel.run(async (context) => {
        parentRange.load("values");
        await context.sync();

        const values = parentRange.values;
        const missingSet = new Set(missingValues);

        // Iterate through the values to find and highlight matching cells individually
        for (let i = 0; i < values.length; i++) {
            for (let j = 0; j < values[i].length; j++) {
                const cellValue = values[i][j];
                if (cellValue !== null && cellValue !== undefined && cellValue !== "" && missingSet.has(cellValue)) {
                    // Get the specific cell and highlight it
                    const cell = parentRange.getCell(i, j);
                    cell.format.fill.color = "yellow";
                }
            }
        }

        await context.sync(); // Sync after highlighting
    });
}

// Find Missing Data
async function findMissingData(highlight, range1Address, range2Address) {
    await Excel.run(async (context) => {
        try {
            const workbook = context.workbook;

            // Get ranges from addresses
            const myRange1 = workbook.worksheets.getItem(range1Address.split('!')[0]).getRange(range1Address.split('!')[1]);
            const myRange2 = workbook.worksheets.getItem(range2Address.split('!')[0]).getRange(range2Address.split('!')[1]);

            // Load values for comparison
            myRange1.load("values");
            myRange2.load("values");
            await context.sync();

            // Check if ranges are valid and contain data
            if (!myRange1.values || !myRange2.values) {
                showModalMessage("Find Missing Data", "One of the ranges does not contain valid data.", true);
                return;
            }

            // Flatten arrays and filter out nulls for Set creation
            const set1 = new Set(myRange1.values.flat().filter(item => item !== null));
            const set2 = new Set(myRange2.values.flat().filter(item => item !== null));

            const missingInRange2 = [];
            for (const item of set1) {
                if (!set2.has(item)) {
                    missingInRange2.push(item);
                }
            }

            const missingInRange1 = [];
            for (const item of set2) {
                if (!set1.has(item)) {
                    missingInRange1.push(item);
                }
            }

            if (missingInRange1.length === 0 && missingInRange2.length === 0) {
                showModalMessage("Find Missing Data", "No missing data found between ranges.", false);
                return;
            }

            if (highlight) {
                highlightMissingCells(myRange1, missingInRange2);
                highlightMissingCells(myRange2, missingInRange1);
            }

            // Create the missing data report
            let missingReportSheet;
            const sheetNameReport = "Missing Data Report";

            try {
                // If worksheet already exists grab it
                missingReportSheet = workbook.worksheets.getItem(sheetNameReport);
                missingReportSheet.load('name');
                await context.sync();
            } catch (e) {
                if (e.code === 'ItemNotFound') {
                    // If worksheet doesn't exist then create it
                    try {
                        missingReportSheet = workbook.worksheets.add(sheetNameReport);
                        missingReportSheet.load('name'); // Load name to confirm it was added
                        await context.sync();
                    } catch (addError) {
                        console.error(`Error adding worksheet '${sheetNameReport}':`, addError);
                        showModalMessage("Error", `Failed to add report sheet: ${addError.message || addError}`, true);
                        return; // Exit if sheet cannot be added
                    }
                } else {
                    console.error(`Unexpected error when trying to get worksheet '${sheetNameReport}':`, e);
                    showModalMessage("Error", `An unexpected error occurred with the report sheet: ${e.message || e}`, true);
                    return; // Exit on unexpected error
                }
            }

            // Ensure the sheet is activated after it's either found or added
            try {
                missingReportSheet.activate();
                await context.sync();
            } catch (activateError) {
                console.error(`Error activating worksheet '${sheetNameReport}':`, activateError);
                return; // Exit if sheet cannot be activated
            }

            try {
                // Clear any existing content first
                const usedRange = missingReportSheet.getUsedRange(true); // true = valuesOnly
                if (usedRange) {
                    usedRange.clear();
                    await context.sync();
                }
            } catch (clearError) {
                // Ignore error if no used range exists (empty sheet)
                console.log("No existing content to clear");
            }

            // Add headers
            const headerRange = missingReportSheet.getRange("A1:B1");
            headerRange.values = [[`Missing in ${range1Address.replace(/\$/g, "")}`, `Missing in ${range2Address.replace(/\$/g, "")}`]];
            headerRange.format.borders.getItem("InsideHorizontal").weight = Excel.BorderWeight.thin;
            headerRange.format.borders.getItem("InsideVertical").weight = Excel.BorderWeight.thin;
            headerRange.format.borders.getItem("EdgeTop").weight = Excel.BorderWeight.thin;
            headerRange.format.borders.getItem("EdgeBottom").weight = Excel.BorderWeight.thin;
            headerRange.format.borders.getItem("EdgeLeft").weight = Excel.BorderWeight.thin;
            headerRange.format.borders.getItem("EdgeRight").weight = Excel.BorderWeight.thin;

            // Determine the maximum number of rows needed
            const maxRows = Math.max(missingInRange1.length, missingInRange2.length);

            // Output missing values - only if there are missing values
            if (maxRows > 0) {
                // Prepare data arrays with proper dimensions
                const outputData = [];
                for (let i = 0; i < maxRows; i++) {
                    const row = [
                        i < missingInRange1.length ? missingInRange1[i] : "",
                        i < missingInRange2.length ? missingInRange2[i] : ""
                    ];
                    outputData.push(row);
                }

                // Write all data at once
                const dataRange = missingReportSheet.getRange(`A2:B${maxRows + 1}`);
                dataRange.values = outputData;
            }

            // Apply formatting to the used range
            try {
                const finalUsedRange = missingReportSheet.getUsedRange();
                if (finalUsedRange) {
                    finalUsedRange.format.autofitColumns();
                    finalUsedRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
                }
            } catch (formatError) {
                console.log("Error formatting used range:", formatError);
            }

            await context.sync();

        } catch (error) {
            console.error("Error in findMissingData:", error);
            showModalMessage("Error", `Error comparing ranges: ${error.message || error}`, true);
        }
    });
}

// Compare two worksheets and highlight differences
async function compareWorksheets(highlight, sheet1Name, sheet2Name) {
    await Excel.run(async (context) => {
        try {
            const workbook = context.workbook;

            // Get the worksheets
            const sheet1 = workbook.worksheets.getItem(sheet1Name);
            const sheet2 = workbook.worksheets.getItem(sheet2Name);

            // Get used ranges from both sheets
            let range1, range2;
            try {
                range1 = sheet1.getUsedRange();
                range2 = sheet2.getUsedRange();
            } catch (error) {
                showModalMessage("Compare Sheets", "One or both sheets are empty.", true);
                return;
            }

            // Load values for comparison
            range1.load("values, rowCount, columnCount, address");
            range2.load("values, rowCount, columnCount, address");
            await context.sync();

            // Check if ranges are valid and contain data
            if (!range1.values || !range2.values) {
                showModalMessage("Compare Sheets", "One or both sheets are empty.", true);
                return;
            }

            const values1 = range1.values;
            const values2 = range2.values;

            const rows = Math.max(values1.length, values2.length);
            const cols = Math.max(
                Math.max(...values1.map(row => row.length)),
                Math.max(...values2.map(row => row.length))
            );

            let differencesFound = false;
            const diffList = [];
            const cellsToHighlight1 = [];
            const cellsToHighlight2 = [];

            // Compare cell by cell
            for (let row = 0; row < rows; row++) {
                for (let col = 0; col < cols; col++) {
                    const val1 = (row < values1.length && col < values1[row].length) ? values1[row][col] : null;
                    const val2 = (row < values2.length && col < values2[row].length) ? values2[row][col] : null;

                    // Convert null/undefined to empty string for comparison
                    const normalizedVal1 = (val1 === null || val1 === undefined) ? "" : val1;
                    const normalizedVal2 = (val2 === null || val2 === undefined) ? "" : val2;

                    if (normalizedVal1 !== normalizedVal2) {
                        differencesFound = true;

                        // Convert to A1 reference (row + 1, col + 1 for 1-based indexing)
                        const cellRef = convertToA1Reference(row + 1, col + 1);

                        // Add to differences list
                        diffList.push({
                            sheet1Value: normalizedVal1 || "Empty",
                            sheet1Cell: cellRef,
                            sheet2Value: normalizedVal2 || "Empty",
                            sheet2Cell: cellRef
                        });

                        // Track cells to highlight
                        if (highlight) {
                            cellsToHighlight1.push({ row, col });
                            cellsToHighlight2.push({ row, col });
                        }
                    }
                }
            }

            // Highlight differences if requested
            if (highlight && differencesFound) {
                await highlightDifferenceCells(sheet1, range1, cellsToHighlight1);
                await highlightDifferenceCells(sheet2, range2, cellsToHighlight2);
            }

            if (!differencesFound) {
                showModalMessage("Compare Sheets", "The sheets are identical!", false);
                return;
            }

            // Create comparison report
            await createComparisonReport(context, diffList, sheet1Name, sheet2Name);

            await context.sync();

        } catch (error) {
            console.error("Error in compareSheets:", error);
            showModalMessage("Error", `Error comparing sheets: ${error.message || error}`, true);
        }
    });
}

// Helper function to highlight difference cells
async function highlightDifferenceCells(sheet, usedRange, cellsToHighlight) {
    await Excel.run(async (context) => {
        for (const cellPos of cellsToHighlight) {
            const cell = usedRange.getCell(cellPos.row, cellPos.col);
            cell.format.fill.color = "yellow";
        }
        await context.sync();
    });
}

// Helper function to convert row/col to A1 reference
function convertToA1Reference(row, col) {
    let colString = "";
    while (col > 0) {
        col--;
        colString = String.fromCharCode(65 + (col % 26)) + colString;
        col = Math.floor(col / 26);
    }
    return colString + row;
}

// Create comparison report worksheet
async function createComparisonReport(context, diffList, sheet1Name, sheet2Name) {
    const workbook = context.workbook;
    let compareSheet;
    const sheetNameReport = "Compare Report";

    try {
        console.log(`Attempting to get worksheet: ${sheetNameReport}`);
        compareSheet = workbook.worksheets.getItem(sheetNameReport);
        compareSheet.load('name');
        await context.sync();
        console.log(`Worksheet '${sheetNameReport}' found.`);
    } catch (e) {
        if (e.code === 'ItemNotFound') {
            console.log(`Worksheet '${sheetNameReport}' not found, attempting to add it.`);
            try {
                compareSheet = workbook.worksheets.add(sheetNameReport);
                compareSheet.load('name');
                await context.sync();
                console.log(`Worksheet '${sheetNameReport}' successfully added.`);
            } catch (addError) {
                console.error(`Error adding worksheet '${sheetNameReport}':`, addError);
                showModalMessage("Error", `Failed to add report sheet: ${addError.message || addError}`, true);
                return;
            }
        } else {
            console.error(`Unexpected error when trying to get worksheet '${sheetNameReport}':`, e);
            showModalMessage("Error", `An unexpected error occurred with the report sheet: ${e.message || e}`, true);
            return;
        }
    }

    try {
        compareSheet.activate();
        await context.sync();
        console.log(`Worksheet '${sheetNameReport}' activated.`);
    } catch (activateError) {
        console.error(`Error activating worksheet '${sheetNameReport}':`, activateError);
        showModalMessage("Error", `Failed to activate report sheet: ${activateError.message || activateError}`, true);
        return;
    }

    // Clear any existing content
    try {
        const usedRange = compareSheet.getUsedRange(true);
        if (usedRange) {
            usedRange.clear();
            await context.sync();
        }
    } catch (clearError) {
        console.log("No existing content to clear");
    }

    // Add headers
    const headerRange = compareSheet.getRange("A1:D1");
    headerRange.values = [[
        `${sheet1Name} Cell Contains`,
        "Reference",
        `${sheet2Name} Cell Contains`,
        "Reference"
    ]];

    // Apply header formatting
    headerRange.format.borders.getItem("InsideHorizontal").weight = Excel.BorderWeight.thin;
    headerRange.format.borders.getItem("InsideVertical").weight = Excel.BorderWeight.thin;
    headerRange.format.borders.getItem("EdgeTop").weight = Excel.BorderWeight.thin;
    headerRange.format.borders.getItem("EdgeBottom").weight = Excel.BorderWeight.thin;
    headerRange.format.borders.getItem("EdgeLeft").weight = Excel.BorderWeight.thin;
    headerRange.format.borders.getItem("EdgeRight").weight = Excel.BorderWeight.thin;

    // Prepare data for output
    if (diffList.length > 0) {
        const outputData = diffList.map(diff => [
            diff.sheet1Value,
            diff.sheet1Cell,
            diff.sheet2Value,
            diff.sheet2Cell
        ]);

        // Write all data at once
        const dataRange = compareSheet.getRange(`A2:D${diffList.length + 1}`);
        dataRange.values = outputData;

        // Add hyperlinks to the reference columns (B and D)
        for (let i = 0; i < diffList.length; i++) {
            const rowNum = i + 2; // +2 because we start at row 2

            // Hyperlink for Sheet1 reference (column B)
            const sheet1RefCell = compareSheet.getCell(rowNum - 1, 1); // 0-based indexing
            sheet1RefCell.hyperlink = {
                address: `#'${sheet1Name}'!${diffList[i].sheet1Cell}`,
                textToDisplay: diffList[i].sheet1Cell
            };

            // Hyperlink for Sheet2 reference (column D)
            const sheet2RefCell = compareSheet.getCell(rowNum - 1, 3); // 0-based indexing
            sheet2RefCell.hyperlink = {
                address: `#'${sheet2Name}'!${diffList[i].sheet2Cell}`,
                textToDisplay: diffList[i].sheet2Cell
            };
        }
    }

    // Apply formatting to the used range
    try {
        const finalUsedRange = compareSheet.getUsedRange();
        if (finalUsedRange) {
            finalUsedRange.format.autofitColumns();
            finalUsedRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
        }
    } catch (formatError) {
        console.log("Error formatting used range:", formatError);
    }

    await context.sync();

    showModalMessage("Compare Sheets", `Comparison complete! Found ${diffList.length} differences. See the Compare Report sheet for details.`, false);
}

// Reset column
async function resetColumn() {
    try {
        await Excel.run(async (context) => {

            const selectedRange = context.workbook.getSelectedRange();

            // Check if entire column is selected
            selectedRange.load("isEntireColumn, address, columnIndex");
            await context.sync();

            if (!selectedRange.isEntireColumn) {
                showModalMessage("Reset Column", "Please select an entire column to reset.", false);
                return;
            }

            // Get the effective range
            const columnRange = await getEffectiveRangeForSelection(context, selectedRange);

            // Validate the effective range
            if (!columnRange) {
                showModalMessage("Reset Column", "No data found in the selected column to reset.", false);
                return;
            }

            // Load necessary properties for processing values
            columnRange.load("values, rowCount, columnCount, address");
            await context.sync();

            if (columnRange.rowCount === 0 || columnRange.columnCount === 0) {
                showModalMessage("Reset Column", "No data found in the selected column to reset.", false);
                return;
            }

            // Process values to convert text numbers to actual numbers
            const values = columnRange.values;
            const newValues = values.map(row => {
                const cellValue = row[0];
                if (cellValue === null || cellValue === undefined || cellValue === "") {
                    return [cellValue]; // Keep null/undefined/empty string as is
                }

                let stringValue = String(cellValue).trim();

                // Handle trailing minus (e.g., "123-" -> -123)
                if (stringValue.endsWith('-') && stringValue.length > 1) {
                    const numberPart = stringValue.slice(0, -1);
                    // Ensure the part before '-' is a valid number
                    if (!isNaN(numberPart) && numberPart.trim() !== '') {
                        stringValue = '-' + numberPart;
                    }
                }

                // Convert to number if possible and if it's not a valid date string that looks like a number
                if (!isNaN(stringValue) && stringValue !== '') {
                    const numValue = parseFloat(stringValue);
                    if (!isNaN(numValue) && isFinite(numValue)) {
                        // Return actual JavaScript number type
                        return [numValue];
                    }
                }

                // If not convertible to a number, return the original value (as text)
                return [cellValue];
            });

            // Clear the selected column's content and formats first
            selectedRange.clear(Excel.ClearApplyTo.contentsAndFormats);
            await context.sync(); // Sync after clearing

            // Set the new, processed values back into the effective range
            columnRange.values = newValues;
            await context.sync(); // Sync after setting values

            // Apply the "General" number format to the entire selected column
            selectedRange.numberFormat = [["General"]];
            await context.sync(); // Sync again to apply the format change

            showModalMessage("Reset Column", "Column reset successfully.", false);

        });
    } catch (error) {
        console.error("Error resetting column:", error);
        showModalMessage("Reset Column", "An error occurred: " + error.message, false);
    }
}

// Convert dates
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
}