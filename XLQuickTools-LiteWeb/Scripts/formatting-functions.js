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
    REMOVE_SPECIAL: 'REMOVE_SPECIAL',
    SUBSCRIPT_UNICODE: 'SUBSCRIPT_UNICODE'
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
        case TextTransformOption.SUBSCRIPT_UNICODE:
            const subscriptMap = {
                '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
                '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉'
            };
            return input.replace(/[0-9]/g, char => subscriptMap[char] || char);
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

        // Display message to user
        if (showMessages) {
            showModalMessage("Fill Down", `Updates made to ${cellsUpdated.toLocaleString()} cells.`, false);
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

            // Load the necessary properties from effectiveRange for undo
            fillRange.load("address");
            fillRange.worksheet.load("name");
            fillRange.load("values, numberFormat");
            await context.sync();

            // Pass the values to the undo manager
            const worksheetName = fillRange.worksheet.name;
            const rangeAddress = fillRange.address;
            const originalValues = fillRange.values;
            const originalNumberFormat = fillRange.numberFormat;

            // Store the current state BEFORE making changes
            await undoManager.copyAndStoreFormat(worksheetName, rangeAddress, originalValues, originalNumberFormat);

            // Use the shared helper function with messages enabled
            await fillBlanksInRange(context, fillRange, true);

            // Update UI: Enable undo button
            if (undoManager.canUndo()) {
                enableUndoButton();
            }
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

// Split columns to rows
async function splitToRows(headers, delimiter) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();

            // Load all necessary properties
            usedRange.load(['values', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
            await context.sync();

            // Check if usedRange is valid
            if (!usedRange || usedRange.rowCount === 0) {
                console.log("No data found in worksheet.");
                showModalMessage("Split Columns to Rows", "No data found in worksheet.", false);
                return;
            }

            const originalValues = usedRange.values;
            const originalStartRow = usedRange.rowIndex;
            const originalStartCol = usedRange.columnIndex;
            const originalRowCount = usedRange.rowCount;
            const originalColumnCount = usedRange.columnCount;

            let dataStartRowIndex = 0;
            if (headers) {
                dataStartRowIndex = 1;
                if (originalRowCount === 1) {
                    showModalMessage("Split Columns to Rows", "No data rows to process after skipping headers.", false);
                    return;
                }
            }

            const processedRows = [];
            let totalRowsAdded = 0;
            let changeMade = false;

            for (let i = dataStartRowIndex; i < originalRowCount; i++) {
                const currentRow = originalValues[i];
                let maxSplitsInRow = 0;

                // First pass: determine maximum number of delimiters used in the current row
                for (let j = 0; j < originalColumnCount; j++) {
                    const cellValue = currentRow[j];
                    if (cellValue !== null && cellValue !== undefined && String(cellValue).length > 0) {
                        const parts = String(cellValue).split(delimiter);
                        maxSplitsInRow = Math.max(maxSplitsInRow, parts.length - 1);
                    }
                }

                if (maxSplitsInRow === 0) {
                    // No splits in this row, add it as is
                    processedRows.push(currentRow);
                } else {
                    changeMade = true;
                    totalRowsAdded += maxSplitsInRow;

                    // Initialize new rows with empty strings, then populate
                    const newRowsForCurrentOriginalRow = Array.from({ length: maxSplitsInRow + 1 }, () => Array(originalColumnCount).fill(''));

                    for (let j = 0; j < originalColumnCount; j++) {
                        const cellValue = currentRow[j];
                        if (cellValue !== null && cellValue !== undefined && String(cellValue).length > 0) {
                            const parts = String(cellValue).split(delimiter).map(part => part.trim());

                            for (let k = 0; k < parts.length; k++) {
                                newRowsForCurrentOriginalRow[k][j] = parts[k];
                            }
                            // If a column has fewer splits than maxSplitsInRow,
                            // the remaining cells in newRowsForCurrentOriginalRow for this column
                            // will retain the first value from the original cell, if it was a single value.
                            // Or be empty if it was split and the parts array is shorter.
                            if (parts.length === 1 && maxSplitsInRow > 0) {
                                for (let k = 1; k <= maxSplitsInRow; k++) {
                                    // Fill down if single value
                                    newRowsForCurrentOriginalRow[k][j] = parts[0];
                                }
                            }
                        } else {
                            // If original cell was empty, fill all corresponding new cells with empty string
                            for (let k = 0; k <= maxSplitsInRow; k++) {
                                newRowsForCurrentOriginalRow[k][j] = '';
                            }
                        }
                    }
                    processedRows.push(...newRowsForCurrentOriginalRow);
                }
            }

            // If headers were present, re-add them to the beginning of the processed rows
            if (headers) {
                processedRows.unshift(originalValues[0]);
            }

            if (!changeMade) {
                showModalMessage("Split Columns to Rows", "No cells contained the delimiter. No changes made.", false);
                return;
            }

            // Determine the target range for writing the new data
            const targetRange = sheet.getRangeByIndexes(
                originalStartRow,
                originalStartCol,
                processedRows.length,
                originalColumnCount
            );

            // Clear the original used range then set new values
            sheet.getUsedRange().clear();
            targetRange.values = processedRows;

            // Apply formatting once at the end
            targetRange.load('format');
            await context.sync(); // Sync to load format object

            targetRange.format.wrapText = false;
            targetRange.format.shrinkToFit = false;
            targetRange.format.autoIndent = false;

            await context.sync(); // Final sync to apply changes

            showModalMessage("Split Columns to Rows",
                `Active worksheet's columns with the selected delimiter have been split into rows. ${totalRowsAdded.toLocaleString()} new rows added.`,
                false);

        });
    } catch (error) {
        console.error("Error in splitToRows:", error);
        showModalMessage("Split Columns to Rows", "An error occurred during processing. Check console for details.", true);
    }
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
            try {
                await navigator.clipboard.writeText(finalString);
            } catch (err) {
                console.error("Failed to copy to clipboard:", err);
                showModalMessage("Selection Plus", "Failed to copy to clipboard. If using Excel online it could be a permissions issue.", true);
            }
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

        // Load the necessary properties from effectiveRange for undo
        effectiveRange.load("address");
        effectiveRange.worksheet.load("name");
        effectiveRange.load("values, numberFormat");
        await context.sync();

        // Pass the values to the undo manager
        const worksheetName = effectiveRange.worksheet.name;
        const rangeAddress = effectiveRange.address;
        const originalValues = effectiveRange.values;
        const originalNumberFormat = effectiveRange.numberFormat;

        // Store the current state BEFORE making changes
        await undoManager.copyAndStoreFormat(worksheetName, rangeAddress, originalValues, originalNumberFormat);

        // Make changes
        await processExcelRange(context, effectiveRange, (cellValue) => transformText(cellValue, option));

        // Update UI: Enable undo button
        if (undoManager.canUndo()) {
            enableUndoButton();
        }

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

        // Load the necessary properties from effectiveRange for undo
        effectiveRange.load("address");
        effectiveRange.worksheet.load("name");
        effectiveRange.load("values, numberFormat");
        await context.sync();

        // Pass the values to the undo manager
        const worksheetName = effectiveRange.worksheet.name;
        const rangeAddress = effectiveRange.address;
        const originalValues = effectiveRange.values;
        const originalNumberFormat = effectiveRange.numberFormat;

        // Store the current state BEFORE making changes
        await undoManager.copyAndStoreFormat(worksheetName, rangeAddress, originalValues, originalNumberFormat);

        // Make changes
        await processExcelRange(context, effectiveRange, (cellValue) => {
            // Ensure cellValue is treated as a string before concatenation
            let processedValue = String(cellValue);
            return leadingText + processedValue + trailingText;
        });

        // Update UI: Enable undo button
        if (undoManager.canUndo()) {
            enableUndoButton();
        }

    });
}

// Remove excess formatting
async function removeExcess() {
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();

            // Find the actual last row and column with data (not just formatting)
            const { lastDataRow, lastDataColumn } = await findLastDataCell(worksheet, context);

            // Get worksheet dimensions
            const worksheetRange = worksheet.getRange();
            worksheetRange.load(['rowCount', 'columnCount']);
            await context.sync();

            const totalRows = worksheetRange.rowCount;
            const totalColumns = worksheetRange.columnCount;

            // Clear excess rows
            if (lastDataRow < totalRows) {
                const startRowIndex = lastDataRow; // 0-based index
                const rowCount = totalRows - lastDataRow;

                // Get range
                const rowsToClear = worksheet.getRangeByIndexes(startRowIndex, 0, rowCount, totalColumns);
                // Clear just formatting on rows
                rowsToClear.clear(Excel.ClearApplyTo.formats);
            }

            // Clear excess columns
            if (lastDataColumn < totalColumns) {
                const startColumnIndex = lastDataColumn; // 0-based index
                const columnCount = totalColumns - lastDataColumn;

                // Get range
                const columnsToClear = worksheet.getRangeByIndexes(0, startColumnIndex, totalRows, columnCount);
                // Clear just formatting on columns
                columnsToClear.clear(Excel.ClearApplyTo.formats);
            }

            await context.sync();
            if (lastDataRow >= 1) {
                showModalMessage("Remove Excess Formatting", `Cleared any excess formatting beyond Column ${getColumnLetter(lastDataColumn - 1)} and Row ${lastDataRow.toLocaleString()}.`, false);
            }
            });
    } catch (error) {
        console.error("Remove Excess:", error);
        showModalMessage("Remove Excess Formatting", "An error occurred while removing excess formatting. Please try again.", false);
    }
}

// Remove hyperlinks (both cell-based and formula)
async function removeHyperlinks() {
    try {
        await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);

            // Load all the properties that will be used
            effectiveRange.load("values, formulas");
            effectiveRange.format.load("font");
            await context.sync();

            // Remove hyperlinks by clearing formulas first, then setting values
            const formulas = effectiveRange.formulas;
            const values = effectiveRange.values;

            // Create cleaned formulas (remove HYPERLINK formulas)
            const cleanedFormulas = formulas.map(row => row.map(cell => {
                if (typeof cell === 'string' && cell.startsWith('=') && cell.toUpperCase().includes('HYPERLINK')) {
                    return ''; // Clear HYPERLINK formulas
                }
                return cell;
            }));

            // Apply changes
            effectiveRange.formulas = cleanedFormulas;
            effectiveRange.values = values;

            // Clear cell-based hyperlinks
            effectiveRange.clear(Excel.ClearApplyTo.hyperlinks);
            await context.sync();

            // Remove hyperlink formatting: reset font underline and color
            effectiveRange.format.font.underline = Excel.RangeUnderlineStyle.none;
            effectiveRange.format.font.color = "#000000"; // Set to black
            await context.sync();

            showModalMessage("Remove Hyperlinks", "All formula and cell-based hyperlinks have been removed.", false);

        });
    } catch (error) {
        console.error("Remove Hyperlinks:", error);
        showModalMessage("Remove Hyperlinks", "An error occurred while removing hyperlinks. Please try again.", false);
    }
}

// Add Hyperlinks to the selected range
async function addHyperlinks(url, headers) {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);
        if (!effectiveRange) {
            showModalMessage("", "No effective range found for the current selection.", false);
            return;
        }

        // Load the necessary properties from effectiveRange for undo and iteration
        effectiveRange.load("rowCount, columnCount, values");
        await context.sync();

        // Add Hyperlinks
        const rowCount = effectiveRange.rowCount;
        const columnCount = effectiveRange.columnCount;
        const originalValues = effectiveRange.values;

        // Build the formulas array
        const formulasArray = [];

        // Determine the starting row for iteration based on 'headers' checkbox
        const startRow = headers ? 1 : 0;

        for (let i = 0; i < rowCount; i++) {
            const row = [];
            for (let j = 0; j < columnCount; j++) {
                // If headers are present and it's the first row, just copy the original value
                if (headers && i === 0) {
                    row.push(originalValues[i][j]);
                    continue;
                }

                const cellValue = originalValues[i][j] || "";

                // Replace {ID} or {id} in the URL with the cell's value if present
                let dynamicUrl = url.replace(/{ID}|{id}/gi, cellValue);

                // Ensure the URL has a protocol for the HYPERLINK function to work reliably
                if (!dynamicUrl.startsWith("http://") && !dynamicUrl.startsWith("https://")) {
                    dynamicUrl = "https://" + dynamicUrl;
                }

                // Determine display text
                let dynamicDisplayText = cellValue;
                // Ensure display text is a string and escape double quotes for the formula
                dynamicDisplayText = String(dynamicDisplayText || "").replace(/"/g, '""');

                // Construct the HYPERLINK formula for the current cell
                const hyperlinkFormula = `=HYPERLINK("${dynamicUrl}","${dynamicDisplayText}")`;
                row.push(hyperlinkFormula);
            }
            formulasArray.push(row);
        }

        // Set all formulas at once using the formulas property
        effectiveRange.formulas = formulasArray;

        try {
            await context.sync();

            // Check if formulas were actually set (optional, good for debugging/confirmation)
            effectiveRange.load("formulas");
            await context.sync();

            showModalMessage("Add Hyperlinks", "Hyperlinks added successfully!", false);
        } catch (error) {
            showModalMessage("Error", `Failed to add hyperlinks: ${error.message}`, true);
            console.error("Failed to add hyperlinks:", error);
        }
    });
}