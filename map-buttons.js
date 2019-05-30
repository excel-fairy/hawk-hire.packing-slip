function onOpen() {
    MAIN_SHEET.sheet.getRange(MAIN_SHEET.exportButtonCell).setValue(false);
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Run scripts')
        .addItem('Export as PDF and send to email', 'exportSheet')
        .addItem('Authorize scripts to access Google drive from smartphone', 'createInstallableTriggers')
        .addToUi();
}

function createInstallableTriggers(){
    deleteAllTriggers();
    ScriptApp.newTrigger('installableOnEdit')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
}

function installableOnEdit(e){
    var range = e.range;
    if(range.getSheet().getName() === MAIN_SHEET.sheet.getRange(MAIN_SHEET.exportButtonCell).getSheet().getName()
        && range.getA1Notation() === MAIN_SHEET.sheet.getRange(MAIN_SHEET.exportButtonCell).getA1Notation()
        && range.getValue() === true){
        range.setValue(false);
        exportSheet();
    }
}

function deleteAllTriggers() {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
        ScriptApp.deleteTrigger(allTriggers[i]);
    }
}
