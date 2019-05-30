var COULD_NOT_EXPORT_SHEET = "Could not export sheet";
var INVOICE_TO_DEFAULT_VALUE = "Sample text & address";
var DELIVER_TO_DEFAULT_VALUE = "Sample text";

function exportSheet(){
    if(canExport()) {
        var fileName = MAIN_SHEET.sheet.getRange(MAIN_SHEET.exportFileNameCell).getValue();
        var exportFolderId = MAIN_SHEET.sheet.getRange(MAIN_SHEET.exportFolderIdCell).getValue();

        var exportOptions = {
            exportFolderId: exportFolderId,
            sheetId: MAIN_SHEET.sheet.getSheetId(),
            exportFileName: fileName,
            range: MAIN_SHEET.exportRange
        };
        var pdfFile = ExportSpreadsheet.export(exportOptions);
        sendEmail(MAIN_SHEET.sheet, pdfFile);
    }
    else
        displayErrorPopup(getErrorMessageCellsNotValid());
}

// Create the error message containing instructions on how to fix non valid cells
function getErrorMessageCellsNotValid() {
    var hasEmptyCells = getEmptyCells().length > 0;
    var retval = "";
    if(hasEmptyCells)
        retval += "Cells " + getEmptyCells().join(", ") + " must not be empty.\n";
    if(MAIN_SHEET.sheet.getRange(MAIN_SHEET.invoiceToCell).getValue().toUpperCase() === INVOICE_TO_DEFAULT_VALUE.toUpperCase())
        retval += MAIN_SHEET.invoiceToCell + " cannot be \"" + INVOICE_TO_DEFAULT_VALUE + "\"\n";
    if(MAIN_SHEET.sheet.getRange(MAIN_SHEET.deliverToCell).getValue().toUpperCase() === DELIVER_TO_DEFAULT_VALUE.toUpperCase())
        retval += MAIN_SHEET.deliverToCell + " cannot be \"" + DELIVER_TO_DEFAULT_VALUE + "\"\n";

    retval += "\nPlease fix these errors and try again.";
    return retval;
}

// Just added to automate email when saving as PDF
function sendEmail(sheet, attachment) {
    var recipient = MAIN_SHEET.sheet.getRange(MAIN_SHEET.recipientEmailAddressCell).getValue();
    var subject = MAIN_SHEET.sheet.getRange(MAIN_SHEET.emailSubjectCell).getValue();
    var message = MAIN_SHEET.sheet.getRange(MAIN_SHEET.emailBodyCell).getValue();
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: 'Automatic packing slip mail sender'
    };
    sendEmailAux(recipient, subject, message, emailOptions);
}

// Send an email
function sendEmailAux(recipient, subject, message, emailOptions) {
    try {
        // do stuff, including send email
        MailApp.sendEmail(recipient, subject, message, emailOptions);
    } catch(e) {
        Logger.log("Error with email. Recipient " + recipient + " may not be a valid email address", e);
        displayErrorPopup("Unexpected error happened while sending the email. " +
            "Please check the recipient email address and try again.");
    }
}

// Check if the sheet can be exported
function canExport() {
    return getEmptyCells().length === 0
        && MAIN_SHEET.sheet.getRange(MAIN_SHEET.invoiceToCell).getValue().toUpperCase() !== INVOICE_TO_DEFAULT_VALUE.toUpperCase()
        && MAIN_SHEET.sheet.getRange(MAIN_SHEET.deliverToCell).getValue().toUpperCase() !== DELIVER_TO_DEFAULT_VALUE.toUpperCase()
}

// Get empty els that should n=ot be empty
function getEmptyCells() {
    var cellsToFill = [
        MAIN_SHEET.dateCell,
        MAIN_SHEET.phoneNumberCell,
        MAIN_SHEET.courriercell,
        MAIN_SHEET.conNoteNumberCell,
        MAIN_SHEET.invoiceToCell,
        MAIN_SHEET.deliverToCell,
        MAIN_SHEET.recipientEmailAddressCell,
        MAIN_SHEET.exportFolderIdCell,
        MAIN_SHEET.exportFileNameCell,
        MAIN_SHEET.emailSubjectCell,
        MAIN_SHEET.emailBodyCell
    ];
    return cellsToFill.filter(function (cell) {
        return MAIN_SHEET.sheet.getRange(cell).getValue() === "";
    });
}

// Display an error popup
function displayErrorPopup(text) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(COULD_NOT_EXPORT_SHEET, text, ui.ButtonSet.OK);
}