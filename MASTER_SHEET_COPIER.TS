/**
 * MASTER PRINT SETTINGS COPIER
 *
 *
 * MASTER = First worksheet in workbook.
 * Copies all print settings, layout rules, margins, scaling, and headers/footers.
 */

function main(workbook: ExcelScript.Workbook) {
    const sheets: ExcelScript.Worksheet[] = workbook.getWorksheets();

    if (sheets.length <= 1) {
        console.log("Only one worksheet found.");
        return;
    }

    // =========================================================
    // 1. READ MASTER PAGE SETTINGS
    // =========================================================
    const masterSheet: ExcelScript.Worksheet = sheets[0];
    const master: ExcelScript.PageLayout = masterSheet.getPageLayout();

    const orientation: ExcelScript.PageOrientation = master.getOrientation();
    const paperSize: ExcelScript.PaperType = master.getPaperSize();
    const centerH: boolean = master.getCenterHorizontally();
    const centerV: boolean = master.getCenterVertically();
    const printGridlines: boolean = master.getPrintGridlines();
    const printHeadings: boolean = master.getPrintHeadings();

    // Margins
    const topMargin: number = master.getTopMargin();
    const bottomMargin: number = master.getBottomMargin();
    const leftMargin: number = master.getLeftMargin();
    const rightMargin: number = master.getRightMargin();
    const headerMargin: number = master.getHeaderMargin();
    const footerMargin: number = master.getFooterMargin();

    // Zoom / Scaling
    const masterZoom: ExcelScript.PageLayoutZoomOptions = master.getZoom();

    // Header/Footer State 
    const masterHF: ExcelScript.HeaderFooterGroup = master.getHeadersFooters();
    const masterState: ExcelScript.HeaderFooterState = masterHF.getState();

    // Read Default H/F
    const defaultHF = masterHF.getDefaultForAllPages();
    const dLH = defaultHF.getLeftHeader();
    const dCH = defaultHF.getCenterHeader();
    const dRH = defaultHF.getRightHeader();
    const dLF = defaultHF.getLeftFooter();
    const dCF = defaultHF.getCenterFooter();
    const dRF = defaultHF.getRightFooter();

    // Read First Page H/F 
    const firstHF = masterHF.getFirstPage();
    const fLH = firstHF.getLeftHeader();
    const fCH = firstHF.getCenterHeader();
    const fRH = firstHF.getRightHeader();
    const fLF = firstHF.getLeftFooter();
    const fCF = firstHF.getCenterFooter();
    const fRF = firstHF.getRightFooter();

    // Read Even Page H/F 
    const evenHF = masterHF.getEvenPages();
    const eLH = evenHF.getLeftHeader();
    const eCH = evenHF.getCenterHeader();
    const eRH = evenHF.getRightHeader();
    const eLF = evenHF.getLeftFooter();
    const eCF = evenHF.getCenterFooter();
    const eRF = evenHF.getRightFooter();

    // =========================================================
    // 2. APPLY MASTER SETTINGS TO ALL OTHER SHEETS
    // =========================================================
    for (let i = 1; i < sheets.length; i++) {
        const target: ExcelScript.PageLayout = sheets[i].getPageLayout();

        // Base Page Setup
        target.setOrientation(orientation);
        target.setPaperSize(paperSize);
        target.setCenterHorizontally(centerH);
        target.setCenterVertically(centerV);
        target.setPrintGridlines(printGridlines);
        target.setPrintHeadings(printHeadings);

        // Margins
        target.setTopMargin(topMargin);
        target.setBottomMargin(bottomMargin);
        target.setLeftMargin(leftMargin);
        target.setRightMargin(rightMargin);
        target.setHeaderMargin(headerMargin);
        target.setFooterMargin(footerMargin);

        // Zoom Scaling Fix
        if (masterZoom.scale) {
            target.setZoom({ scale: masterZoom.scale });
        } else {
            target.setZoom({
                horizontalFitToPages: masterZoom.horizontalFitToPages,
                verticalFitToPages: masterZoom.verticalFitToPages
            });
        }

        // Headers & Footers Execution Group
        const targetHF = target.getHeadersFooters();

        // Set matching state first (Default, FirstPage, OddEven, or All)
        targetHF.setState(masterState);

        // Set Default / All Pages
        const targetDefault = targetHF.getDefaultForAllPages();
        targetDefault.setLeftHeader(dLH);
        targetDefault.setCenterHeader(dCH);
        targetDefault.setRightHeader(dRH);
        targetDefault.setLeftFooter(dLF);
        targetDefault.setCenterFooter(dCF);
        targetDefault.setRightFooter(dRF);

        // Set First Page 
        const targetFirst = targetHF.getFirstPage();
        targetFirst.setLeftHeader(fLH);
        targetFirst.setCenterHeader(fCH);
        targetFirst.setRightHeader(fRH);
        targetFirst.setLeftFooter(fLF);
        targetFirst.setCenterFooter(fCF);
        targetFirst.setRightFooter(fRF);

        // Set Even Pages
        const targetEven = targetHF.getEvenPages();
        targetEven.setLeftHeader(eLH);
        targetEven.setCenterHeader(eCH);
        targetEven.setRightHeader(eRH);
        targetEven.setLeftFooter(eLF);
        targetEven.setCenterFooter(eCF);
        targetEven.setRightFooter(eRF);
    }

    console.log(
        `DONE: Print settings successfully copied from master sheet '${masterSheet.getName()}' to ${sheets.length - 1} worksheets.`
    );
}
