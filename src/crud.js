function deleteRecord(record) {
    // {deleteRecord: true, uuId: afhkwiefns342vsdfd3, position: 3}

    //map teacherCellValues onto teacher names => 'jbroccoli'
    const getTeacherName = (teacherCellValue) => {
        const regex = /[A-Za-z.]*(?=@)/ //matches all alphanumeric characters that precede '@'
    // get teacher initial + last name from email in cell value; example: 'JoMarie Broccoli (jbroccoli@nysmith.com)' => 'jbroccoli'
    return teacherCellValue.match(regex) ? teacherCellValue.match(regex)[0] : null
  }

    const mathTeacher = getTeacherName(record.mathTeacher)
    const laTeacher = getTeacherName(record.laTeacher)
    const principalRec = getTeacherName(record.principalRec)
    const thirdTeacher = getTeacherName(record.thirdTeacher)

    const teachers = [mathTeacher, laTeacher, principalRec, thirdTeacher].filter(el => el)  

    //once you have it working, redo with Sheets API (queue changes to make at once, otherwise this will be very slow)
    //but be SURE that changes are applied in the right order- otherwise this could get really messy!

    teachers.forEach(teacher => {
      const tab = ss.getSheetByName(teacher)
      const data = tab.getDataRange().getValues()
      const rowIndex = data.findIndex(row => row[teacherTabs.columnIndex.uuId] === record.uuId)
      const completionAndNotesCols = tab.getRange(rowIndex + 1, teacherTabs.columnNumbers.dateCompleted, 1, teacherTabs.columnNumbers.notes)
      completionAndNotesCols.deleteCells(SpreadsheetApp.Dimension.ROWS);    //Deleting cells and not the whole row to avoid deleting the query formula in A2
    })

    formResponsesSheet.deleteRow(record.position)  

}

function copyRecordToSheet(record, sheetName) {   
  const sheet = ss.getSheetByName(sheetName)
  const rowData = record.rowData
  const blankRow = sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowData.length)    
  blankRow.setValues([rowData])
}

function deleteQueuedRecords() {
  const data = formResponsesSheet.getDataRange().getValues()

  const queuedRecords = data
    .map((row, i) => ({
      deleteRecord: row[formResponses.columnIndex.deleteRecord],
      uuId: row[formResponses.columnIndex.uuId],
      position: i + 1,
      mathTeacher: row[formResponses.columnIndex.mathTeacher],
      laTeacher: row[formResponses.columnIndex.laTeacher],
      principalRec: row[formResponses.columnIndex.principalRec],
      thirdTeacher: row[formResponses.columnIndex.thirdTeacher],
      rowData: row
    }))
    .filter(row => row.deleteRecord === true)

  queuedRecords.forEach(record => {
    copyRecordToSheet(record, 'Deleted Requests')
    deleteRecord(record)
  })
}