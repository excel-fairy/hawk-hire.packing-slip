function exportSheet(){
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
    }
}
