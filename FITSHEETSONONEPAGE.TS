/**
 * Set All Sheets to Fit on One Page
 * Purpose:
 * Applies Excel print scaling "Fit Sheet on One Page"
 * to every worksheet in the workbook.
 */

function main(workbook: ExcelScript.Workbook) {

    const sheets = workbook.getWorksheets();

    for (const sheet of sheets) {

        const pageLayout = sheet.getPageLayout();

        // Equivalent to:
        // File > Print > Scaling > Fit Sheet on One Page
        pageLayout.setZoom({
            horizontalFitToPages: 1,
            verticalFitToPages: 1
        });
    }

    console.log(
        `Completed: ${sheets.length} worksheets set to Fit Sheet on One Page.`
    );
}
