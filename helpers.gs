// Rev sheet setup functions
function getAllRowHeights(sheetName, rwIdRangeA1, mathIdRangeA1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);

  const subjectRanges = [
    {'name': 'rw', 'idRange': sh.getRange(rwIdRangeA1)},
    {'name': 'math', 'idRange': sh.getRange(mathIdRangeA1)}
  ]

  for (let subject of subjectRanges) {
    const subName = subject.name;
    const subRange = subject.idRange;
    const ids = subRange.getValues();
    const startCell = subRange.getCell(1,1);
    const idStartRow = startCell.getRow();
    const idCol = startCell.getColumn();
    const heights = [];
    let lastSetRow = idStartRow - 1;

    for (let rowOffset = 0; rowOffset < ids.length; rowOffset++) {
      const id = ids[rowOffset][0]; // Get the ID from the current row
      const height = calculateRowHeight(id, 1000, subName); // Calculate height based on ID
      heights.push([height]);

      // Batch process every 100 rows
      if ((rowOffset + 1) % 100 === 0 || rowOffset === ids.length - 1) {
        const batchStartRow = lastSetRow + 1;
        const batchEndRow = idStartRow + rowOffset;
        const slice = heights.slice(lastSetRow - idStartRow + 1); // Slice only new rows

        sh.getRange(batchStartRow, idCol + 2, slice.length).setValues(slice);
        Logger.log(`${subName} values set for rows ${batchStartRow}-${batchEndRow}`);
        lastSetRow = batchEndRow;
      }
    }
  }
}

function calculateRowHeight(questionId, containerWidth, subject) {
  var questionUrl = 'https://www.openpathtutoring.com/static/img/concepts/sat/' + subject.toLowerCase() + '/' + encodeURIComponent(questionId) + '.jpg';
  var urlOptions = { muteHttpExceptions: true };

  // Add exponential backoff retry logic
  var maxRetries = 4;
  var retryCount = 0;
  var questionImg;

  while (retryCount < maxRetries) {
    try {
      questionImg = UrlFetchApp.fetch(questionUrl, urlOptions);
      break; // Success - exit retry loop
    } catch (e) {
      retryCount++;
      if (retryCount === maxRetries) {
        SpreadsheetApp.getUi().alert('Failed to fetch image after ' + maxRetries + ' attempts: ' + e.message);
        return;
      }
      // Exponential backoff: wait 2^retryCount * 1000 milliseconds
      Utilities.sleep(Math.pow(2, retryCount) * 1000);
      Logger.log('Retry ' + retryCount);
    }
  }

  var questionBlob = questionImg.getBlob();
  var questionSize = ImgApp.getSize(questionBlob);

  if (subject.toLowerCase() === 'rw') {
    var whitespace = 40;
  } else {
    var whitespace = 160;
  }

  var rowHeight = (questionSize.height / questionSize.width) * containerWidth + whitespace;

  Logger.log(questionId + ' rowHeight: ' + rowHeight);

  return Math.round(rowHeight);
}

function getAnswerSheets(folder) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const filename = file.getName().toLowerCase();

    if (filename.includes('sat admin answer analysis')) {
      satSheetIds.admin = fileId;
    }
    else if (filename.includes('sat student answer sheet')) {
      satSheetIds.student = fileId;
    }
    else if (filename.includes('act admin answer analysis')) {
      actSheetIds.admin = fileId;
    }
    else if (filename.includes('act student answer sheet')) {
      actSheetIds.student = fileId;
    }
  }

  return [satSheetIds, actSheetIds];
}

function getLastFilledRow(sheet, col) {
  const maxRow = sheet.getMaxRows();
  const allVals = sheet.getRange(1, col, maxRow).getValues();
  const lastFilledRow = maxRow - allVals.reverse().findIndex((c) => c[0] != '');

  return lastFilledRow;
}

function getOPTPermissionsList(id) {
  var editors = DriveApp.getFileById(id)
    .getEditors()
    .map(function (e) {
      return e.getEmail();
    });
  var emails = [];

  for (var i = 0; i < editors.length; i++) {
    // Only add openpathtutoring.com emails to email list
    if (editors[i].includes('openpathtutoring.com')) {
      emails.push(editors[i]);
    }
  }

  return emails.join();
}

function getSubFolderIdsByFolderId(folderId, result = []) {
  let folder = DriveApp.getFolderById(folderId);
  let folders = folder.getFolders();
  if (folders && folders.hasNext()) {
    while (folders.hasNext()) {
      let f = folders.next();
      let childFolderId = f.getId();
      result.push(childFolderId);

      result = getSubFolderIdsByFolderId(childFolderId, result);
    }
  }
  return result.filter(onlyUnique);
}

function getTestCodes() {
  const practiceTestDataSheet = SpreadsheetApp.openById('1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE').getSheetByName(`Practice test data updated ${dataLatestDate}`);
  const lastFilledRow = getLastFilledRow(practiceTestDataSheet, 1);
  const testCodeCol = practiceTestDataSheet
    .getRange(2, 1, lastFilledRow - 1)
    .getValues()
    .map((row) => row[0]);
  const testCodes = testCodeCol.filter((x, i, a) => a.indexOf(x) == i);

  return testCodes;
}

function isFunctionRunning(functionName) {
  const url = "https://script.googleapis.com/v1/processes";
  const options = {
    method: "get",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken(),
    },
  };

  const response = UrlFetchApp.fetch(url, options);
  const processes = JSON.parse(response.getContentText()).processes;
  const now = new Date();

  // Check if any processes with the specified functionName were started within the past 6 minutes
  const isRunning = processes.some(process => {
    const processStartTime = new Date(process.startTime);
    const timeDifference = (now - processStartTime) / (1000 * 60); // Convert difference to minutes
    if (process.functionName === functionName && process.processStatus === "RUNNING" && timeDifference <= 6 && timeDifference > 0.1) {
      Logger.log(`process.functionName: ${process.functionName}, process.processStatus: ${process.processStatus}, process.startTime ${process.startTime}`)
    }
    return process.functionName === functionName && process.processStatus === "RUNNING" && timeDifference <= 6 && timeDifference > 0.1;
  });

  return isRunning;
}

function isDark(hex = '#b6d7a8') {
  hex = hex.substring(1); // strip #
  const rgb = parseInt(hex, 16); // convert rrggbb to decimal
  const r = (rgb >> 16) & 0xff; // extract red
  const g = (rgb >> 8) & 0xff; // extract green
  const b = (rgb >> 0) & 0xff; // extract blue
  const luma = 0.2126 * r + 0.7152 * g + 0.0722 * b; // per ITU-R BT.709
  
  if (luma < 205) {
    return true;
  } else {
    return false;
  }
}

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}

function processFolders(folders, folderFunction) {
  while (folders.hasNext()) {
    const folder = folders.next();
    folderFunction(folder);
    processFolders(folder.getFolders(), folderFunction);
  }
}

function renameFolder(folder, currentName, newName, isStudentFolder = true) {
  let folderName = folder.getName();
  let files = folder.getFiles();
  let subfolders = folder.getFolders();
  let revDataSsId, adminSsId, adminSs, revBackendSheet;

  if (folderName.includes(currentName) && !folderName.includes(newName)) {
    let newFoldername = folderName.replace(currentName, newName);
    folder.setName(newFoldername);
  }

  while (files.hasNext()) {
    let file = files.next();
    let filename = file.getName();

    if (filename.includes(currentName) && !filename.includes(newName)) {
      let newFilename = filename.replace(currentName, newName);
      file.setName(newFilename);
    }

    if (filename.toLowerCase().includes('sat admin answer analysis') && isStudentFolder) {
      adminSsId = file.getId();
      adminSs = SpreadsheetApp.openById(adminSsId);
      revBackendSheet = adminSs.getSheetByName('Rev sheet backend');
      revBackendSheet.getRange('K2').setValue(newName);
    }
  }

  while (subfolders.hasNext()) {
    let subfolder = subfolders.next();
    let subfolderName = subfolder.getName();

    if (subfolderName.includes(currentName) && !subfolderName.includes(newName)) {
      let newSubfolderName = subfolderName.replace(currentName, newName);
      subfolder.setName(newSubfolderName);
    }

    renameFolder(subfolder, currentName, newName, isStudentFolder);
  }

  if (adminSsId && isStudentFolder) {
    revDataSsId = revBackendSheet.getRange('U3');
    revDataSs = SpreadsheetApp.openById(revDataSsId);

    if (revDataSs.getSheetByName(newName)) {
      let ui = SpreadsheetApp.getUi();
      ui.alert('Rev sheet named ' + newName + ' already exists. Please update manually.');
      return;
    } else {
      revDataSs.getSheetByName(currentName).setName(newName);
    }
  }
}

const showAllSheetsExcept = (spreadsheetId='1_nRuW80ewwxEcsHLKy8U8o1nIxKNxxrih-IC-T2suJk', sheetNamesToHide = ['RW Rev sheet', 'Math Rev sheet', 'Rev sheet backend']) => {
  SpreadsheetApp.openById(spreadsheetId)
    .getSheets()
    .forEach((sh) => {
      // If sheets are meant to be hidden, leave them hidden
      if (sheetNamesToHide.includes(sh.getName())) {
        sh.hideSheet();
      }
      else {
        sh.showSheet();
      }
    });
}