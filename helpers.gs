// Rev sheet setup functions
function getAllRowHeights(sheetName, rwIdRangeA1, mathIdRangeA1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);

  const subjectRanges = [
    { name: 'rw', idRange: sh.getRange(rwIdRangeA1) },
    { name: 'math', idRange: sh.getRange(mathIdRangeA1) },
  ];

  for (let subject of subjectRanges) {
    const subName = subject.name;
    const subRange = subject.idRange;
    const ids = subRange.getValues();
    const startCell = subRange.getCell(1, 1);
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
    } else if (filename.includes('sat student answer sheet')) {
      satSheetIds.student = fileId;
    } else if (filename.includes('act admin answer analysis')) {
      actSheetIds.admin = fileId;
    } else if (filename.includes('act student answer sheet')) {
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

function getSatTestCodes() {
  const practiceTestDataSheet = SpreadsheetApp.openById('1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE').getSheetByName(`Practice test data updated ${dataLatestDate}`);
  const lastFilledRow = getLastFilledRow(practiceTestDataSheet, 1);
  const testCodeCol = practiceTestDataSheet
    .getRange(2, 1, lastFilledRow - 1)
    .getValues()
    .map((row) => row[0]);
  const testCodes = testCodeCol.filter((x, i, a) => a.indexOf(x) == i);

  return testCodes;
}

function getActTestCodes() {
  const practiceTestDataSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('actDataSheetId')).getSheetByName(`ACT Answers`);
  const lastFilledRow = getLastFilledRow(practiceTestDataSheet, 1);
  const testCodeCol = practiceTestDataSheet
    .getRange(2, 1, lastFilledRow - 1)
    .getValues()
    .map((row) => row[0]);
  const testCodes = testCodeCol.filter((x, i, a) => a.indexOf(x) == i);

  return testCodes;
}

function isFunctionRunning(functionName) {
  const url = 'https://script.googleapis.com/v1/processes';
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  };

  const response = UrlFetchApp.fetch(url, options);
  const processes = JSON.parse(response.getContentText()).processes;
  const now = new Date();

  // Check if any processes with the specified functionName were started within the past 6 minutes
  const isRunning = processes.some((process) => {
    const processStartTime = new Date(process.startTime);
    const timeDifference = (now - processStartTime) / (1000 * 60); // Convert difference to minutes
    if (process.functionName === functionName && process.processStatus === 'RUNNING' && timeDifference <= 6 && timeDifference > 0.1) {
      Logger.log(`process.functionName: ${process.functionName}, process.processStatus: ${process.processStatus}, process.startTime ${process.startTime}`);
    }
    return process.functionName === functionName && process.processStatus === 'RUNNING' && timeDifference <= 6 && timeDifference > 0.1;
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
  let revDataSsId, revBackendSheet, satAdminSsId, satAdminSs;

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
      satAdminSsId = file.getId();
      satAdminSs = SpreadsheetApp.openById(satAdminSsId);
      revBackendSheet = satAdminSs.getSheetByName('Rev sheet backend');
      revBackendSheet.getRange('K2').setValue(newName);
    } else if (filename.toLowerCase().includes('act admin answer analysis') && isStudentFolder) {
      actAdminSsId = file.getId();
      actAdminSs = SpreadsheetApp.openById(actAdminSsId);
      actAdminSs.getSheetByName('Student responses').getRange('G1').setValue(newName);
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

  if (satAdminSsId && isStudentFolder && revDataSsId) {
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

const showAllSheetsExcept = (spreadsheetId = '1_nRuW80ewwxEcsHLKy8U8o1nIxKNxxrih-IC-T2suJk', sheetNamesToHide = ['RW Rev sheet', 'Math Rev sheet', 'Rev sheet backend']) => {
  SpreadsheetApp.openById(spreadsheetId)
    .getSheets()
    .forEach((sh) => {
      // If sheets are meant to be hidden, leave them hidden
      if (sheetNamesToHide.includes(sh.getName())) {
        sh.hideSheet();
      } else {
        sh.showSheet();
      }
    });
};

function getActPageBreakRow(sheet) {
  const grandColData = sheet
    .getRange(1, 2, 111)
    .getValues()
    .map((row) => row[0]);
  const mathColData = sheet
    .getRange(1, 3, 111)
    .getValues()
    .map((row) => row[0]);

  const grandTotalIndex = grandColData.indexOf('Grand Total');
  if (0 < grandTotalIndex && grandTotalIndex < 80) {
    Logger.log(`Single page ending at ${grandTotalIndex + 1}`);
    sheet.hideRows(grandTotalIndex + 2, 111);
    SpreadsheetApp.flush();
    return 80;
  }

  const mathTotalIndex = mathColData.indexOf('Math Total');
  if (0 < mathTotalIndex && mathTotalIndex < 80) {
    Logger.log(`Page break at ${mathTotalIndex + 1}`);
    return mathTotalIndex + 1;
  } else {
    return 80;
  }
}

async function mergePDFs(fileIds, destinationFolderId, name = 'merged.pdf', attempt = 1) {
  const validFileIds = fileIds.filter(isValidPdf);

  if (validFileIds.length !== fileIds.length) {
    if (attempt > 5) {
      throw new Error('mergePDFs: Too many attempts, some files are still not valid PDFs.');
    }
    // Exponential backoff: wait 2^attempt * 1000 ms
    const waitMs = Math.pow(2, attempt) * 1000;
    Logger.log(`mergePDFs: Not all files are valid PDFs. Retrying in ${waitMs / 1000}s (attempt ${attempt})`);
    Utilities.sleep(waitMs);
    return await mergePDFs(fileIds, destinationFolderId, name, attempt + 1);
  }
  // Retrieve PDF data as byte arrays
  const data = fileIds.map((id) => new Uint8Array(DriveApp.getFileById(id).getBlob().getBytes()));

  // Load pdf-lib from CDN
  const cdnjs = 'https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js';
  eval(
    UrlFetchApp.fetch(cdnjs)
      .getContentText()
      .replace(/setTimeout\(.*?,.*?(\d*?)\)/g, 'Utilities.sleep($1);return t();')
  );

  // Merge PDFs
  const pdfDoc = await PDFLib.PDFDocument.create();
  for (let i = 0; i < data.length; i++) {
    const pdfData = await PDFLib.PDFDocument.load(data[i]);
    const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
    pages.forEach((page) => pdfDoc.addPage(page));
  }

  // Save merged PDF to Drive
  const bytes = await pdfDoc.save();
  const mergedBlob = Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, 'merged.pdf');
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const mergedFile = destinationFolder.createFile(mergedBlob).setName(name);

  // fileIds.forEach((id) => DriveApp.getFileById(id).setTrashed(true));

  return mergedFile;
}

function savePdfSheet(
  spreadsheetId,
  sheetId,
  studentName,
  margin = {
    top: '0.5',
    bottom: '0.5',
    left: '0.3',
    right: '0.3',
  }
) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();

    var url_base = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export';
    var url_ext =
      '?format=pdf' + //export as pdf
      // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
      (sheetId ? '&gid=' + sheetId : '&id=' + spreadsheetId) +
      // following parameters are optional...
      '&size=letter' + // paper size
      '&portrait=true' + // orientation, false for landscape
      '&fitw=true' + // fit to width, false for actual size
      '&fzr=true' + // repeat row headers (frozen rows) on each page
      '&top_margin=' +
      margin.top +
      '&bottom_margin=' +
      margin.bottom +
      '&left_margin=' +
      margin.left +
      '&right_margin=' +
      margin.right +
      '&printnotes=false' +
      '&sheetnames=false' +
      '&printtitle=false' +
      '&pagenumbers=false'; //hide optional headers and footers

    var options = {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
      muteHttpExceptions: true,
    };

    // Create PDF
    const pdfName = spreadsheet.getSheetById(sheetId).getName() + ' sheet for ' + studentName;
    const response = UrlFetchApp.fetch(url_base + url_ext, options);
    const blob = response.getBlob().setName(pdfName + '.pdf');
    const rootFolder = DriveApp.getRootFolder();
    const pdfSheet = rootFolder.createFile(blob);

    return pdfSheet.getId();
  } catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack);
  }
}

function isValidPdf(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  if (blob.getContentType() !== MimeType.PDF) return false;
  const bytes = blob.getBytes();
  const header = String.fromCharCode.apply(null, bytes.slice(0, 5));
  return header === '%PDF-';
}

function getStudentsSpreadsheetData(studentName) {
  const summarySheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('Summary');
  const lastFilledRow = getLastFilledRow(summarySheet, 1);
  const summaryData = summarySheet.getRange(1, 1, lastFilledRow, 26).getValues();
  const studentData = {
    name: null,
    hours: null,
    recentSessionDate: null,
  };

  for (let r = 0; r < lastFilledRow; r++) {
    if (summaryData[r][0] === studentName) {
      (studentData.name = summaryData[r][0]), (studentData.hours = summaryData[r][3]), (studentData.recentSessionDate = Utilities.formatDate(new Date(summaryData[r][16]), 'GMT', 'EEE M/d'));
      break;
    }
  }
  return studentData;
}

function updateOPTStudentFolderData() {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const teamDataSheet = clientDataSs.getSheetByName('Team OPT');
  const teamFolder = DriveApp.getFolderById('1tSKajFOa_EVUjH8SKhrQFbHSjDmmopP9');
  const tutorFolders = teamFolder.getFolders();
  let tutorIndex = 0;

  while (tutorFolders.hasNext()) {
    const tutorFolder = tutorFolders.next();
    const tutorFolderName = tutorFolder.getName();
    const tutorFolderId = tutorFolder.getId();

    const tutorStudentsStr = teamDataSheet.getRange(tutorIndex + 2, 4).getValue();
    let tutorStudents = tutorStudentsStr ? JSON.parse(tutorStudentsStr) : [];

    tutorData = {
      index: tutorIndex,
      name: tutorFolderName,
      studentsFolderId: tutorFolderId,
      studentsDataJSON: tutorStudents,
    };

    tutorStudents = createStudentFolders.findStudentFileIds(tutorData);

    teamDataSheet.getRange(tutorIndex + 2, 1, 1, 4).setValues([[tutorIndex, tutorFolderName, tutorFolderId, JSON.stringify(tutorStudents)]]);
    tutorIndex++;
  }

  const clientSheet = clientDataSs.getSheetByName('Clients');
  const myStudentsStr = clientSheet.getRange(2, 17).getValue();
  let myStudents = myStudentsStr ? JSON.parse(myStudentsStr) : [];

  const myStudentFolderData = {
    index: 0,
    name: 'Open Path Tutoring',
    studentsFolderId: clientSheet.getRange(2, 15).getValue(),
    studentsDataJSON: myStudents,
  };

  myStudents = createStudentFolders.getStudentFileIds(myStudentFolderData);
  clientSheet.getRange(2, 17).setValue(JSON.stringify(myStudents));
}


function formatDateYYYYMMDD(date) {
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const yyyy = date.getFullYear();
  return `${yyyy}-${mm}-${dd}`;
}


function addStudentDataToJson(
  studentData = {
    'name': null,
    'folderId': null,
    'satAdminSsId': null,
    'satStudentSsId': null,
    'actAdminSsId': null,
    'actStudentSsId': null,
    'homeworkSsId': null
  }
) {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const myStudentsJsonCell = clientDataSs.getSheetByName('Clients').getRange('Q2');
  let myStudentsStr = myStudentsJsonCell.getValue();
  const myStudentsJson = myStudentsStr ? JSON.parse(myStudentsStr) : [];
  myStudentsJson.push(studentData);
  myStudentsStr = JSON.stringify(myStudentsJson);
  myStudentsJsonCell.setValue(myStudentsStr);
}