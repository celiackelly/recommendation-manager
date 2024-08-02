function removeProtection(){
    let sheetProtections = ss.getSheetByName('bschrembs').getProtections(SpreadsheetApp.ProtectionType.SHEET)
    let rangeProtections = ss.getSheetByName('bschrembs').getProtections(SpreadsheetApp.ProtectionType.RANGE)
    let protections = sheetProtections.concat(rangeProtections)
    protections.forEach(protection => protection.remove())
}
