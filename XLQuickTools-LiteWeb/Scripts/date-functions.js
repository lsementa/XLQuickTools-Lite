// --- Date Functions ---
// date-functions.js


Office.onReady(() => {
    // On Ready
});

// Convert dates
// Using datevalue and text built-in excel funtions were too slowo so using JS instead:
async function convertSelectedDates(currentLocale, convertLocale, format, type) {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        const effectiveRange = await getEffectiveRangeForSelection(context, selectedRange);
        const undoRange = effectiveRange;

        if (!effectiveRange) {
            showModalMessage("", "No effective range found for the current selection.", false);
            return;
        }

        // Load the necessary properties from effectiveRange for undo
        undoRange.load("address");
        undoRange.worksheet.load("name");
        undoRange.load("values, numberFormat");
        await context.sync();

        // Pass the values to the undo manager
        const worksheetName = undoRange.worksheet.name;
        const rangeAddress = undoRange.address;
        const originalValues = undoRange.values;
        const originalNumberFormat = undoRange.numberFormat;

        // Store the current state BEFORE making changes
        await undoManager.copyAndStoreFormat(worksheetName, rangeAddress, originalValues, originalNumberFormat);

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

        // Update UI: Enable undo button
        if (undoManager.canUndo()) {
            enableUndoButton();
        }

        showModalMessage("Date/Text Converter", "Date conversion complete!", false);

    });
}

// JavaScript date conversion
function convertToText(value, format) {
    if (value == null || value === undefined || value === "") {
        return { value: value, format: "@" };
    }

    let date;
    if (typeof value === "number") {
        // Convert Excel serial date to JavaScript Date
        const msSinceEpoch = (value - 25569) * 86400000;
        date = new Date(msSinceEpoch);
        // Adjust to UTC midnight to avoid timezone discrepancies
        date = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    } else if (typeof value === "string") {
        // Parse as date string
        date = new Date(value);
        if (isNaN(date.getTime())) {
            return { value: value, format: "@" };
        }
        date = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    } else {
        return { value: value, format: "@" };
    }

    // Format the date
    const formattedDate = formatDateToString(date, format);
    return { value: formattedDate, format: "@" };
}

// Convert to Excel serial date
function convertToSerial(value, format) {
    if (value == null || value === undefined || value === "") {
        return { value: value, format: "General" };
    }

    if (typeof value === "number") {
        return { value: value, format: format };
    } else if (typeof value === "string") {
        // Parse as date string
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

// Format the date to a specified format
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

    // Format mapping based on the options the user can select
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
}

// Verion using Excel functions -----------------
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