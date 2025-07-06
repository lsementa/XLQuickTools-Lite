// The initialize function must be run each time a new page is loaded.
Office.onReady(() => {
        // If you need to initialize something you can do so here.
});

async function sampleFunction(event) { 
const values = [
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        try {
        await Excel.run(async (context) => {
                // Write sample values to a range in the active worksheet.
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("B3:D5").values = values;
                await context.sync();
        });
        } catch (error) {
        console.log(error.message);
        }
        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
}
