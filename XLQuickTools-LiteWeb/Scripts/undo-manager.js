// Class to manage undo state

class UndoManager {
    constructor() {
        this.lastWorksheetName = "";
        this.lastFormattedRangeAddress = "";
        this.originalNumberFormat = null; // 2D array
        this.originalValues = null;       // 2D array
    }

    // Store the current state of the specified range (values and number format) for undo
    async copyAndStoreFormat(worksheetName, rangeAddress, values, numberFormat) {
        if (!worksheetName || !rangeAddress || values === null || numberFormat === null) {
            throw new Error("Worksheet name, range address, values, and numberFormat cannot be null or empty.");
        }

        // Store the worksheet name and range address
        this.lastWorksheetName = worksheetName;
        this.lastFormattedRangeAddress = rangeAddress.replace(/\$/g, "");

        // Store the format and values directly
        this.originalNumberFormat = numberFormat;
        this.originalValues = values;

/*        console.log("State stored for undo:", {
            worksheet: this.lastWorksheetName,
            range: this.lastFormattedRangeAddress,
            numberFormat: this.originalNumberFormat,
            values: this.originalValues
        });*/

    }

    // Restores the previously stored state (values and number format) to the range
    async undoFormatting() {
        if (!this.lastFormattedRangeAddress || !this.originalValues) {
            console.warn("No undo state available.");
            return;
        }

        await Excel.run(async (context) => {
            try {
                const workbook = context.workbook;
                const worksheet = workbook.worksheets.getItem(this.lastWorksheetName);
                const range = worksheet.getRange(this.lastFormattedRangeAddress);

                // Set the original values and number format
                range.values = this.originalValues;
                range.numberFormat = this.originalNumberFormat;
                await context.sync();

                // Select the undone range
                range.select();
                await context.sync();

                // Clear the undo state
                this.clearUndoState();

            } catch (error) {
                console.error("Error during undo operation:", error);
                showModalMessage("Undo", "Error during undo operation.", false);
            }
        });
    }

    // Clear the undo state
    clearUndoState() {
        this.lastWorksheetName = "";
        this.lastFormattedRangeAddress = "";
        this.originalNumberFormat = null;
        this.originalValues = null;
    }

    // Check if there's an undo state available
    canUndo() {
        return !!this.lastFormattedRangeAddress && this.originalValues !== null;
    }
}

// Instantiate undo manager
const undoManager = new UndoManager();

