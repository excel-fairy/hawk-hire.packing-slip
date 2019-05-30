var MAIN_SHEET = {
    name: 'Sheet1',
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1'),
    exportRange: {
        r1: 1,
        r2: 33,
        c1: ColumnNames.letterToColumn('A'),
        c2: ColumnNames.letterToColumn('G')
    },
    dateCell: 'G2',
    phoneNumberCell: 'G3',
    courriercell: 'G4',
    conNoteNumberCell: 'G5',
    invoiceToCell: 'B12',
    deliverToCell: 'E12',
    recipientEmailAddressCell: 'J3',
    exportFolderIdCell: 'J6',
    exportFileNameCell: 'J9',
    emailSubjectCell: 'J12',
    emailBodyCell: 'J15',
    exportButtonCell: 'J20'
};
