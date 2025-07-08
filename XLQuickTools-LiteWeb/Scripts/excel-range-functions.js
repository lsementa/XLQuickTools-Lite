// --- Excel Range Functions ---
// excel-range-functions.js


Office.onReady(() => {
    // On Ready
});

// Function to convert column index to Excel column letter
function getColumnLetter(columnIndex) {
    let result = '';
    while (columnIndex >= 0) {
        result = String.fromCharCode(65 + (columnIndex % 26)) + result;
        columnIndex = Math.floor(columnIndex / 26) - 1;
    }
    return result;
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