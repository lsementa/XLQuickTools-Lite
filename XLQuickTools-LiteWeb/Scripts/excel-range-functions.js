// --- Excel Range Functions ---
// excel-range-functions.js

Office.onReady(() => {
    // On Ready
});

// Constants for large data processing
const CHUNK_SIZE = 10000;
const BATCH_SIZE = 50;

// Function to convert column index to Excel column letter
function getColumnLetter(columnIndex) {
    let result = '';
    while (columnIndex >= 0) {
        result = String.fromCharCode(65 + (columnIndex % 26)) + result;
        columnIndex = Math.floor(columnIndex / 26) - 1;
    }
    return result;
}

// Function to find the last row and column with actual data on the active worksheet
async function findLastDataCell(worksheet, context) {
    try {
        // Get the used range of the worksheet
        const usedRange = worksheet.getUsedRange(true);

        // Load properties of the used range
        usedRange.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
        await context.sync();

        if (usedRange.isNullObject || usedRange.rowCount === 0 || usedRange.columnCount === 0) {
            // No data found
            return { lastDataRow: 0, lastDataColumn: 0 };
        }

        // Calculate the last data row and column based on properties
        const lastDataRow = usedRange.rowIndex + usedRange.rowCount;
        const lastDataColumn = usedRange.columnIndex + usedRange.columnCount;

        return { lastDataRow: lastDataRow, lastDataColumn: lastDataColumn };

    } catch (error) {
        console.error("Error finding last data cell:", error);
    }
}

// Get the effective used range. A better way of handling selected range
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