function getClassFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classFolder = DriveApp.getFileById(ss.getId()).getParents().next().getParents().next().getParents().next();
  const classFolderId = classFolder.getId();
  var files = classFolder.getFiles();
  var aggSsId = null;

  while (files.hasNext()) {
    file = files.next();
    filename = file.getName();

    if (filename.toLowerCase().includes('aggregate answer analysis')) {
      aggSsId = file.getId();
    }
  }

  if (aggSsId === null) {
    const parentFolder = classFolder.getParents().next();
    const parentFiles = parentFolder.getFiles();

    Logger.log(parentFolder.getName());

    while (parentFiles.hasNext()) {
      parentFile = parentFiles.next();

      if (parentFile.getName().toLowerCase().includes('aggregate answer analysis')) {
        aggSsId = parentFile.makeCopy().moveTo(classFolder).getId();
        DriveApp.getFileById(aggSsId).setName(classFolder.getName() + ' aggregate answer analysis');
      }
    }

  }

  Logger.log(classFolderId + " " + aggSsId);

  generateClassTestAnalysis(classFolderId, aggSsId);

  return aggSsId;
}

function generateClassTestAnalysis(folderId, aggSsId) {
  var folder = DriveApp.getFolderById(folderId);
  var ssFiles = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var subFolders = DriveApp.getFolderById(folderId).getFolders();
  const aggSs = SpreadsheetApp.openById(aggSsId);
  const aggSheet = aggSs.getSheetByName('Data');
  var firstOpenAggRow = 2;

  while (ssFiles.hasNext()) {
    file = ssFiles.next();
    filename = file.getName();
    Logger.log(filename + ' ' + file.getId());

    if (filename.toLowerCase().includes('sat answer analysis')) {
      const ss = SpreadsheetApp.openById(file.getId());
      const sh = ss.getSheetByName('Practice test data');
      const studentName = filename.slice(filename.indexOf('-') + 2);

      Logger.log(studentName);

      const lastRow = sh.getLastRow();
      const allVals = sh.getRange('A1:A' + lastRow).getValues();
      const lastFilledRow = lastRow - allVals.reverse().findIndex((c) => c[0] != '');
      const numRowsToCopy = lastFilledRow - 1;
      const studentData = sh.getRange(2, 1, numRowsToCopy, 12).getValues();

      aggSheet.getRange(firstOpenAggRow, 2, numRowsToCopy, 12).setValues(studentData);
      aggSheet.getRange(firstOpenAggRow, 1, numRowsToCopy).setValue(studentName);

      firstOpenAggRow = firstOpenAggRow + numRowsToCopy;
    }
  }

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();

    generateClassTestAnalysis(subFolder.getId(), aggSsId);
  }

  const aggStudentAnswers = aggSheet.getRange(2, 12, aggSheet.getLastRow());
  const upperAggAnswers = aggStudentAnswers.getDisplayValues().map((row) => row.map((col) => (col ? col.toUpperCase() : col)));
  aggStudentAnswers.setValues(upperAggAnswers);
}