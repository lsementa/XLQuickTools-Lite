// --- Data Analysis Functions ---
// data-anlysis-functions.js


Office.onReady(() => {
    // On Ready
});

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
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${uniqueCount.toLocaleString()}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 0; padding-left: 30px;">Total Non-Blank Cells:</td>
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${totalNonBlankCells.toLocaleString()}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 0; padding-left: 30px;">Total Blank Cells:</td>
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${totalBlankCells.toLocaleString()}</td>
                    </tr>
                    <tr>
                        <td style="padding: 3px 0; padding-left: 30px;">Total Rows with Data:</td>
                        <td style="padding: 3px 0; text-align: center; padding-right: 30px">${fillRange.rowCount.toLocaleString()}</td>
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

        // Check if entire column is selected
        selectedRange.load("isEntireColumn");
        await context.sync();

        if (!selectedRange.isEntireColumn) {
            showModalMessage("Check for Duplicates", "Please select an entire column to use.", false);
            return;
        }

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

            // Flatten arrays and filter out null, undefined, and empty/whitespace-only strings for Set creation
            const set1 = new Set(myRange1.values.flat().filter(item => {
                return item !== null && item !== undefined && String(item).trim() !== '';
            }));
            const set2 = new Set(myRange2.values.flat().filter(item => {
                return item !== null && item !== undefined && String(item).trim() !== '';
            }));

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
                showModalMessage("Compare Sheets", "One or both sheets are empty.", false);
                return;
            }

            // Load values for comparison
            range1.load("values, rowCount, columnCount, address");
            range2.load("values, rowCount, columnCount, address");
            await context.sync();

            // Check if ranges are valid and contain data
            if (!range1.values || !range2.values) {
                showModalMessage("Compare Sheets", "One or both sheets are empty.", false);
                return;
            }

            const values1 = range1.values;
            const values2 = range2.values;

            const rows = Math.max(values1.length, values2.length);

            // Calculate max columns without spread operator to avoid stack overflow
            let maxCols1 = 0;
            for (const row of values1) {
                if (row.length > maxCols1) {
                    maxCols1 = row.length;
                }
            }

            let maxCols2 = 0;
            for (const row of values2) {
                if (row.length > maxCols2) {
                    maxCols2 = row.length;
                }
            }

            const cols = Math.max(maxCols1, maxCols2);

            //console.log(`Comparing ${rows} rows x ${cols} columns = ${rows * cols} cells`);

            let differencesFound = false;
            const diffList = [];
            const cellsToHighlight1 = [];
            const cellsToHighlight2 = [];

            // Progress tracking for large datasets
            const totalCells = rows * cols;
            const progressInterval = Math.max(1000, Math.floor(totalCells / 100)); // Report every 1% or 1000 cells
            let processedCells = 0;

            // Compare cell by cell with progress reporting
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

                        // Track cells to highlight (with limits for performance)
                        if (highlight && cellsToHighlight1.length < 500) { // Limit highlighting
                            cellsToHighlight1.push({ row, col });
                            cellsToHighlight2.push({ row, col });
                        }
                    }

                    processedCells++;

                    // Progress reporting
                    if (processedCells % progressInterval === 0) {
                        const progress = Math.round((processedCells / totalCells) * 100);
                        console.log(`Progress: ${progress}% (${processedCells}/${totalCells} cells) - Found ${diffList.length} differences`);
                    }
                }
            }

            //console.log(`Comparison complete. Found ${diffList.length} differences`);

            // Handle highlighting
            if (highlight && differencesFound) {
                if (cellsToHighlight1.length > 0) {
                    await highlightDifferenceCells(context, sheet1, range1, cellsToHighlight1);
                    await highlightDifferenceCells(context, sheet2, range2, cellsToHighlight2);
                }
            }

            // No differences
            if (!differencesFound) {
                showModalMessage("Compare Sheets", "The sheets are identical!", false);
                return;
            }

            // Create comparison report with chunking for large datasets
            await createComparisonReport(context, diffList, sheet1Name, sheet2Name);

           // Final sync
            await context.sync();

        } catch (error) {
            console.error("Error in compareSheets:", error);
            showModalMessage("Error", `Error comparing sheets: ${error.message || error}`, true);
        }
    });
}

// Helper function to highlight difference cells
async function highlightDifferenceCells(context, sheet, usedRange, cellsToHighlight) {

    // Process in small batches to avoid overwhelming Excel
    for (let i = 0; i < cellsToHighlight.length; i += BATCH_SIZE) {
        const batch = cellsToHighlight.slice(i, i + BATCH_SIZE);

        for (const cellPos of batch) {
            const cell = usedRange.getCell(cellPos.row, cellPos.col);
            cell.format.fill.color = "yellow";
        }

        // Sync every batch to prevent memory buildup
        await context.sync();
    }
}

// Helper function to convert row/col to A1 reference for links on report
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

    //console.log(`Creating report with ${diffList.length} differences`);

    try {
        compareSheet = workbook.worksheets.getItem(sheetNameReport);
        compareSheet.load('name');
        await context.sync();
    } catch (e) {
        if (e.code === 'ItemNotFound') {
            try {
                compareSheet = workbook.worksheets.add(sheetNameReport);
                compareSheet.load('name');
                await context.sync();
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

    // Process data in chunks for large datasets
    if (diffList.length > 0) {
        const totalChunks = Math.ceil(diffList.length / CHUNK_SIZE);

        for (let chunkIndex = 0; chunkIndex < totalChunks; chunkIndex++) {
            const startIdx = chunkIndex * CHUNK_SIZE;
            const endIdx = Math.min(startIdx + CHUNK_SIZE, diffList.length);
            const chunk = diffList.slice(startIdx, endIdx);

            //console.log(`Processing chunk ${chunkIndex + 1}/${totalChunks} (rows ${startIdx + 1}-${endIdx})`);

            // Prepare data for this chunk
            const outputData = chunk.map(diff => [
                diff.sheet1Value,
                diff.sheet1Cell,
                diff.sheet2Value,
                diff.sheet2Cell
            ]);

            // Write chunk data
            const dataRange = compareSheet.getRange(`A${startIdx + 2}:D${endIdx + 1}`);
            dataRange.values = outputData;
            await context.sync();

            // Add hyperlinks for this chunk (limit to prevent timeout)
            if (diffList.length <= LINK_LIMIT) {
                for (let i = 0; i < chunk.length; i++) {
                    const rowNum = startIdx + i + 2;
                    const diffItem = chunk[i];

                    // Hyperlink for Sheet1 reference (column B)
                    const sheet1RefCell = compareSheet.getCell(rowNum - 1, 1);
                    sheet1RefCell.hyperlink = {
                        address: `#'${sheet1Name}'!${diffItem.sheet1Cell}`,
                        textToDisplay: diffItem.sheet1Cell
                    };

                    // Hyperlink for Sheet2 reference (column D)
                    const sheet2RefCell = compareSheet.getCell(rowNum - 1, 3);
                    sheet2RefCell.hyperlink = {
                        address: `#'${sheet2Name}'!${diffItem.sheet2Cell}`,
                        textToDisplay: diffItem.sheet2Cell
                    };
                }
                await context.sync();
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

        if (diffList.length > LINK_LIMIT) {
            showModalMessage("Compare Sheets",
                `Report created with ${diffList.length} differences. Hyperlinks skipped for performance with large datasets.`,
                false);
        }
    }
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
                    return [cellValue]; // Keep as is
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