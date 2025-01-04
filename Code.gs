function NewSatFolder(sourceFolderId, parentFolderId, nameOnReport=false) {
  if (sourceFolderId === undefined || parentFolderId === undefined) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    var sourceFolder = file.getParents().next();
    var sourceFolderId = sourceFolder.getId();
    var parentFolderId = sourceFolder.getParents().next().getId();
  }

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);
  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else {
    var studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();
  
  if (nameOnReport) {
    nameOnReport = studentName;
  }

  copyFolder(sourceFolderId, newFolderId, studentName, 'sat');
  linkSheets(newFolderId, nameOnReport);

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + '\'s folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'SAT folder created successfully');
}

function NewActFolder(sourceFolderId, parentFolderId, nameOnReport=false) {
  if (sourceFolderId === undefined || parentFolderId === undefined) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    var sourceFolder = file.getParents().next();
    var sourceFolderId = sourceFolder.getId();
    var parentFolderId = sourceFolder.getParents().next().getId();
  }

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);
  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else {
    var studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  if (nameOnReport) {
    nameOnReport = studentName;
  }
  Logger.log('nameOnReport: ' + nameOnReport);


  copyFolder(sourceFolderId, newFolderId, studentName, 'act');
  linkSheets(newFolderId, nameOnReport);

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + '\'s folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ACT folder created successfully');
}

function NewTestPrepFolder(sourceFolderId, parentFolderId, nameOnReport=false) {
  if (sourceFolderId === undefined || parentFolderId === undefined) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    var sourceFolder = file.getParents().next();
    var sourceFolderId = sourceFolder.getId();
    var parentFolderId = sourceFolder.getParents().next().getId();
  }

  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Student name:', ui.ButtonSet.OK_CANCEL);
  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else {
    var studentName = prompt.getResponseText();
  }

  const newFolder = DriveApp.getFolderById(parentFolderId).createFolder(studentName);
  const newFolderId = newFolder.getId();

  if (nameOnReport) {
    nameOnReport = studentName;
  }
  Logger.log('nameOnReport: ' + nameOnReport);

  copyFolder(sourceFolderId, newFolderId, studentName, 'all');
  linkSheets(newFolderId, nameOnReport);

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + studentName + '\'s folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Test prep folder created successfully');
}

function copyFolder(sourceFolderId = '1yqQx_qLsgqoNiDoKR9b63mLLeOiCoTwo', newFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF', studentName = '_Aaron S', folderType = 'sat') {

  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const newFolder = DriveApp.getFolderById(newFolderId);

  var sourceSubFolders = sourceFolder.getFolders();
  var files = sourceFolder.getFiles();

  if (folderType.toLowerCase() === 'sat') {
    var testType = 'SAT';
  }
  else if (folderType.toLowerCase() === 'act') {
    var testType = 'ACT';
  }
  else {
    var testType = 'Test';
  }

  while (files.hasNext()) {
    var file = files.next();
    let prefixFiles = ['Tutoring notes', 'ACT review sheet', 'SAT review sheet'];
    var fileName = file.getName();
    Logger.log(fileName);

    if (prefixFiles.includes(fileName)) {
      fileName = studentName + " " + fileName;
    }
    else if (fileName.toLowerCase().includes('template')) {
      rootName = fileName.slice(0, fileName.indexOf('-') + 2);
      fileName = rootName + studentName;
    }

    var newFile = file.makeCopy(fileName, newFolder);
    var newFileName = newFile.getName().toLowerCase();

    if (newFileName.includes('tutoring notes')) {
      var ssId = newFile.getId();
      var ss = SpreadsheetApp.openById(ssId);
      var sheet = ss.getSheetByName('Session notes');
      shId = sheet.getSheetId();
      sheet.getRange('G3').setValue('=hyperlink("https://docs.google.com/spreadsheets/d/' + ssId + '/edit?gid=' + shId + '#gid=' + shId + '&range=B"&match(G2,B1:B,0)-1,"Go to latest session")');
    }

    if (newFileName.includes('admin notes')) {
      DocumentApp.openById(newFile.getId()).getBody().replaceText('StudentName', studentName);
    }

    if (testType === 'SAT' && fileName.toLowerCase().includes('act') && fileName.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    }
    else if (testType === 'ACT' && fileName.toLowerCase().includes('sat') && fileName.toLowerCase().includes('answer analysis')) {
      newFile.setTrashed(true);
    }

    if (newFolder.getName().includes(folderType.toUpperCase()) && !newFolder.getName().includes(studentName)) {
      newFile.moveTo(newFolder.getParents().next());
      Logger.log("new location: " + newFile.getParents().next().getId());
      if (isEmptyFolder(newFolder.getId())) {
        newFolder.setTrashed(true);
        Logger.log(newFolder.getName() + " trashed")
      }
    }
  }

  while (sourceSubFolders.hasNext()) {
    var sourceSubFolder = sourceSubFolders.next();
    var folderName = sourceSubFolder.getName();
    Logger.log(folderName + ' ' + newFolder);

    if (folderName === 'Student') {
      var targetFolder = newFolder.createFolder(studentName + " " + testType + " prep");
    }
    else if (newFolder.getName().includes(folderType.toUpperCase()) && newFolder.getName() !== studentName + " " + testType + " prep") {
      var targetFolder = newFolder.getParents().next().createFolder(folderName);
      Logger.log(sourceSubFolder.getId() + " moved");
    }
    else {
      var targetFolder = newFolder.createFolder(folderName);
    }

    if (targetFolder.getName().includes('ACT') && folderType.toLowerCase() === 'sat') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolder.getName() + " trashed");
    }
    else if (targetFolder.getName().includes('SAT') && folderType.toLowerCase() === 'act') {
      targetFolder.setTrashed(true);
      Logger.log(targetFolder.getName() + " trashed");
    }
    else {
      copyFolder(sourceSubFolder.getId(), targetFolder.getId(), studentName, folderType);
    }
  }
}

var satSheetIds = {
  'admin': null,
  'student': null,
  'studentData': null,
  'adminData': null,
  'rev': null
}

var satSheetDataUrls = {
  'admin': null,
  'student': null,
  'rev': null
}

var actSheetIds = {
  'admin': null,
  'student': null,
  'studentData': null,
  'adminData': null
}

var actSheetDataUrls = {
  'admin': null,
  'student': null
}

function linkSheets(folderId, nameOnReport=false) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var subFolders = DriveApp.getFolderById(folderId).getFolders();

  while (files.hasNext()) {
    file = files.next();
    fileName = file.getName();
    if (fileName.includes('SAT')) {
      if (fileName.includes('student answer sheet')) {
        satSheetIds.student = file.getId();
        DriveApp.getFileById(satSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      }
      else if (fileName.includes('answer analysis')) {
        satSheetIds.admin = file.getId();

        var ss = SpreadsheetApp.openById(file.getId());
        if (nameOnReport) {
          for (i in ss.getSheets()) {
            var s = ss.getSheets()[i];
            var sName = s.getName().toLowerCase();
            if (sName.includes('analysis') || sName.includes('opportunity')) {
              s.getRange('D4').setValue('for ' + nameOnReport)
            }
          }
        }
      }
    }

    if (fileName.includes('ACT')) {
      if (fileName.toLowerCase().includes('student answer sheet')) {
        actSheetIds.student = file.getId();
        DriveApp.getFileById(actSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      }
      else if (fileName.toLowerCase().includes('answer analysis')) {
        actSheetIds.admin = file.getId();
      }
    }
  }

  if (satSheetIds.student && satSheetIds.admin) {
    let satAdminSheet = SpreadsheetApp.openById(satSheetIds.admin);
    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);

    // SpreadsheetApp.openById(satSheetIds.student).getSheetByName('Question bank data').getRange('I2').setValue('=iferror(importrange("' + satSheetIds.admin + '","Question bank data!I2:I"),"")');
    // SpreadsheetApp.openById(satSheetIds.student).getSheets()[0].getRange('D1').setValue('=importrange("' + satSheetIds.admin + '","Question bank data!V1")');
  }
  Logger.log('actSheetIds.student: ' + actSheetIds.student);
  Logger.log('actSheetIds.admin: ' + actSheetIds.admin);
  if (actSheetIds.student && actSheetIds.admin) {
    SpreadsheetApp.openById(actSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(actSheetIds.student);
  }

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    linkSheets(subFolder.getId(), nameOnReport);
  }
}

function isEmptyFolder(folderId) {
  const folders = DriveApp.getFolderById(folderId).getFolders();
  const files = DriveApp.getFolderById(folderId).getFiles();

  if (folders.hasNext() || files.hasNext()) {
    return false;
  }
  else {
    return true;
  }
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
    fileName = file.getName();
    Logger.log(fileName + " " + file.getId());

    if (fileName.toLowerCase().includes('sat answer analysis')) {
      const ss = SpreadsheetApp.openById(file.getId());
      const sh = ss.getSheetByName('Practice test data');
      const studentName = fileName.slice(fileName.indexOf('-') + 2);

      Logger.log(studentName);

      const lastRow = sh.getLastRow();
      const allVals = sh.getRange("A1:A" + lastRow).getValues();
      const lastFilledRow = lastRow - allVals.reverse().findIndex(c => c[0] != '');
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
  const upperAggAnswers = aggStudentAnswers.getDisplayValues().map(row => row.map(col => (col) ? col.toUpperCase() : col));
  aggStudentAnswers.setValues(upperAggAnswers);
}


function newClient() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Tutor or Business name:', ui.ButtonSet.OK_CANCEL);
  if(prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }
  else {
    var clientName = prompt.getResponseText();
  }

  var useCustomStyle = ui.alert(
    'Apply custom styles?',
    ui.ButtonSet.YES_NO
  );

  let primaryColor;
  let secondaryColor;
  let tertiaryColor;
  let fontColor;
  var isCustom = false;
  var imgUrl = '';

  if (useCustomStyle === ui.Button.YES) {
    primaryColor = ui.prompt('Primary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    secondaryColor = ui.prompt('Secondary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    tertiaryColor = ui.prompt('Tertiary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    fontColor = ui.prompt('Font color (leave blank to use primary color)', ui.ButtonSet.OK_CANCEL).getResponseText();
    imgUrl = ui.prompt('Image URL', ui.ButtonSet.OK_CANCEL).getResponseText();
    isCustom = true;
  }

  if (primaryColor === '') {
    primaryColor = '#1c4d65';
  }

  if (secondaryColor === '') {
    secondaryColor = '#f6b26b';
  }

  if (tertiaryColor === '') {
    tertiaryColor = '#e6e6e6';
  }

  if (fontColor === '') {
    fontColor = primaryColor;
  }

  // img = ''
  // if (imgUrl) {
  //   img = SpreadsheetApp
  //         .newCellImage()
  //         .setSourceUrl(imgUrl)
  //         .build();
  // }

  var customStyles = {
    'isActive': isCustom,
    'primaryColor': primaryColor,
    'secondaryColor': secondaryColor,
    'tertiaryColor': tertiaryColor,
    'fontColor': fontColor,
    'img': imgUrl
  };

  var sourceFolder = DriveApp.getFolderById('1YU3bVuKbx6en8tsJuLW7huEQkKdDEown')
  var destinationFolder = DriveApp.getFolderById('130wX98bJM4wW6aE6J-e6VffDNwqvgeNS');
  let newFolder = destinationFolder.createFolder(clientName);
  let newFolderId = newFolder.getId();

  createClientFolder(sourceFolder, newFolder, clientName, customStyles);
  linkSheets(newFolderId);
  setClientDataUrls(newFolderId);

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + newFolderId + '" target="_blank" onclick="google.script.host.close()">' + newFolder.getName() + '\'s folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Client folder created successfully");
}


function createClientFolder(sourceFolder, newFolder, clientName, customStyles) {
  var folders = sourceFolder.getFolders();
  var files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (fileName.toLowerCase().includes('template')) {
      const rootName = fileName.slice(0, fileName.indexOf('-') + 2);

      if (fileName.toLowerCase().includes('data - client')) {
        fileName = rootName + clientName;
      } else {
        fileName = rootName + "Template for " + clientName;
      }
    }

    const newFile = file.makeCopy(fileName, newFolder);
    // SpreadsheetApp.flush();

    if (customStyles.isActive && newFile.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
      Logger.log('Custom style true');

      styleClientSheets(SpreadsheetApp.openById(newFile.getId()), customStyles);
    }
  }

  while (folders.hasNext()) {
    var sourceSubFolder = folders.next();
    var folderName = sourceSubFolder.getName();
    var targetFolder = newFolder.createFolder(folderName);

    createClientFolder(sourceSubFolder, targetFolder, clientName, customStyles);
  }
}


function setClientDataUrls(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var subFolders = DriveApp.getFolderById(folderId).getFolders();

  var isSet = {
    'satStudentToData': false,
    'satAdminToStudent': false,
    'satAdminToData': false,
    'satAdminDataToStudentData': false,
    'satRevToAdmin': false,
    'satRevToStudent': false,
    'actStudentToData': false,
    'satAdminToStudent': false,
    'actAdminToData': false,
    'actAdminDataToStudentData': false,
  }

  while (files.hasNext()) {
    file = files.next();
    fileId = file.getId();
    fileName = file.getName().toLowerCase();

    if (fileName.includes('sat admin data')) {
      Logger.log('found sat admin data');
      satSheetIds.adminData = fileId;
      satSheetDataUrls.admin = '"https://docs.google.com/spreadsheets/d/' + satSheetIds.adminData + '/edit?usp=sharing"';
    }
    else if (fileName.includes('sat student data')) {
      Logger.log('found sat student data');
      satSheetIds.studentData = fileId;
      satSheetDataUrls.student = '"https://docs.google.com/spreadsheets/d/' + satSheetIds.studentData + '/edit?usp=sharing"';
      DriveApp.getFileById(satSheetIds.studentData).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    }
    else if (fileName.includes('sat student answer sheet')) {
      Logger.log('found sat student answer sheet');
      satSheetIds.student = fileId;
      DriveApp.getFileById(satSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    }
    else if (fileName.includes('sat admin answer analysis')) {
      Logger.log('found sat admin answer sheet');
      satSheetIds.admin = fileId;
    }
    else if (fileName.includes('rev sheet data')) {
      Logger.log('found rev sheet data');
      satSheetIds.rev = fileId;
      DriveApp.getFileById(satSheetIds.rev).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    }
    else if (fileName.includes('act admin data')) {
      Logger.log('found act admin data');
      actSheetIds.adminData = fileId;
      actSheetDataUrls.admin = '"https://docs.google.com/spreadsheets/d/' + actSheetIds.adminData + '/edit?usp=sharing"';
    }
    else if (fileName.includes('act student data')) {
      Logger.log('found act student data');
      actSheetIds.studentData = fileId;
      actSheetDataUrls.student = '"https://docs.google.com/spreadsheets/d/' + actSheetIds.studentData + '/edit?usp=sharing"';
      DriveApp.getFileById(actSheetIds.studentData).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    }
    else if (fileName.includes('act student answer sheet')) {
      Logger.log('found act student answer sheet');
      actSheetIds.student = fileId;
      DriveApp.getFileById(actSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    }
    else if (fileName.includes('act admin answer analysis')) {
      Logger.log('found act admin answer sheet');
      actSheetIds.admin = fileId;
    }
  }

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    setClientDataUrls(subFolder.getId());
  }

  if (!isSet.satAdminToStudent && satSheetIds.admin && satSheetIds.student) {
    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
  }

  if (!isSet.satStudentToData && satSheetIds.student && satSheetDataUrls.student) {
    SpreadsheetApp.openById(satSheetIds.student).getSheetByName('Question bank data').getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.student + ', "Question bank data!A1:G10000")');
    SpreadsheetApp.openById(satSheetIds.student).getSheetByName('Practice test data').getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.student + ', "Practice test data!A1:E10000")');

    isSet.satStudentToData = true;
  }
  if (!isSet.satAdminToData && satSheetIds.admin && satSheetDataUrls.admin) {
    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Question bank data').getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.admin + ', "Question bank data!A1:H10000")');
    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Practice test data').getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.admin + ', "Practice test data!A1:J10000")');
    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Reading & Writing').getRange('D1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.admin + ', "Question bank data!Q1")');

    isSet.satAdminToData = true;
  }

  if (!isSet.satAdminDataToStudentData && satSheetDataUrls.admin && satSheetIds.studentData) {
    SpreadsheetApp.openById(satSheetIds.studentData).getSheetByName('Question bank data').getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.admin + ', "Question bank data!A1:G10000")');
    SpreadsheetApp.openById(satSheetIds.studentData).getSheetByName('Practice test data').getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.admin + ', "Practice test data!A1:E10000")');

    isSet.satAdminDataToStudentData = true;
  }

  if (!isSet.satRevToAdmin && satSheetIds.admin && satSheetIds.rev) {
    let adminRevSheet = SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Rev sheet backend')
    adminRevSheet.getRange('D2').setValue(satSheetIds.rev);
    DriveApp.getFileById(satSheetIds.rev).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

    isSet.satRevToAdmin = true;
  }

  if (!isSet.actStudentToData && actSheetIds.student && actSheetDataUrls.student) {
    SpreadsheetApp.openById(actSheetIds.student).getSheetByName('Data').getRange('A1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.student + ', "Data!A1:D10000")');

    isSet.actStudentToData = true;
  }
  if (!isSet.actAdminToData && actSheetIds.admin && actSheetDataUrls.admin) {
    var ss = SpreadsheetApp.openById(actSheetIds.admin)
    ss.getSheetByName('Data').getRange('A1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.admin + ', "Data!A1:G10000")');
    ss.getSheets()[0].getRange('J1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.admin + ', "Data!Q1")');
    ss.getSheets()[0].getRange('G1:I1').mergeAcross()
      .setValue('=iferror(J1,"Click to connect data >>")')
    
    isSet.actAdminToData = true;
  }
  if (!isSet.actAdminDataToStudentData && actSheetDataUrls.admin && actSheetIds.studentData) {
    SpreadsheetApp.openById(actSheetIds.studentData).getSheetByName('Data').getRange('A1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.admin + ', "Data!A1:D10000")');

    isSet.satAdminDataToStudentData = true;
  }

  return isSet;
}


function styleClientFolder(clientFolder=null, customStyles={}) {
  var ui = SpreadsheetApp.getUi();
  if(clientFolder) {
    var clientFolderId = clientFolder.getId();
  }
  else {
    var prompt = ui.prompt('Client folder ID', ui.ButtonSet.OK_CANCEL);
    if(prompt.getSelectedButton() == ui.Button.CANCEL) {
      return;
    }
    else {
      var clientFolderId = prompt.getResponseText();
      clientFolder = DriveApp.getFolderById(clientFolderId);
    }
  }

  if (Object.keys(customStyles).length === 0) {
    var primaryColor = ui.prompt('Primary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    var secondaryColor = ui.prompt('Secondary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    var tertiaryColor = ui.prompt('Tertiary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    var fontColor = ui.prompt('Font color (leave blank to use primary color)', ui.ButtonSet.OK_CANCEL).getResponseText();
    var imgUrl = ui.prompt('Image URL', ui.ButtonSet.OK_CANCEL).getResponseText();
    customStyles.primaryColor = primaryColor;
    customStyles.secondaryColor = secondaryColor;
    customStyles.tertiaryColor = tertiaryColor;
    customStyles.fontColor = fontColor;
    customStyles.img = imgUrl;
  }
  else {
    var primaryColor = customStyles.primaryColor;
    var secondaryColor = customStyles.secondaryColor;
    var tertiaryColor = customStyles.tertiaryColor;
    var fontColor = customStyles.fontColor;
  }

  if (fontColor === '') {
    fontColor = primaryColor;
  }

  var folders = clientFolder.getFolders();
  var files = clientFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();

    if (file.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
      styleClientSheets(SpreadsheetApp.openById(file.getId()), customStyles);
    }
  }

  while (folders.hasNext()) {
    var folder = folders.next();
    styleClientFolder(folder, customStyles);
  }

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + clientFolderId + '" target="_blank" onclick="google.script.host.close()">Client folder</a>')
    .setWidth(250) //optional
    .setHeight(50); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Styling complete");
}


function styleClientSheets(
  ss = SpreadsheetApp.getActiveSpreadsheet(),
  customStyles={}) {

  if (Object.keys(customStyles).length === 0) {
    var ui = SpreadsheetApp.getUi();
    var primaryColor = ui.prompt('Primary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    var secondaryColor = ui.prompt('Secondary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    var tertiaryColor = ui.prompt('Tertiary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
    var fontColor = ui.prompt('Font color (leave blank to use primary color)', ui.ButtonSet.OK_CANCEL).getResponseText();
    var imgUrl = ui.prompt('Image URL', ui.ButtonSet.OK_CANCEL).getResponseText();
    customStyles.primaryColor = primaryColor;
    customStyles.secondaryColor = secondaryColor;
    customStyles.tertiaryColor = tertiaryColor;
    customStyles.fontColor = fontColor;
    customStyles.img = imgUrl;
  } else {
    var primaryColor = customStyles.primaryColor;
    var secondaryColor = customStyles.secondaryColor;
    var tertiaryColor = customStyles.tertiaryColor;
    var fontColor = customStyles.fontColor;
  }

  if (fontColor === '') {
    fontColor = primaryColor;
  }

  const ssName = ss.getName().toLowerCase();

  const satTestSheets = ['sat1', 'sat2', 'sat3', 'sat4', 'sat5', 'sat6', 'psat1', 'psat2']
  const satDataSheets = ['question bank data', 'practice test data', 'rev sheet backend']
  const actDataSheets = ['data', 'scoring']

  if (isDark(primaryColor)) {
    var primaryContrastColor = 'white'
  }
  else if (isDark(fontColor)) {
    primaryContrastColor = fontColor;
  }
  else {
    primaryContrastColor = 'black';
  }

  if (isDark(secondaryColor)) {
    var secondaryContrastColor = 'white';
  }
  else if (isDark(fontColor)) {
    secondaryContrastColor = fontColor;
  }
  else {
    secondaryContrastColor = 'black';
  }

  if (isDark(tertiaryColor)) {
    var tertiaryContrastColor = 'white'
  }
  else if (isDark(fontColor)) {
    tertiaryContrastColor = fontColor;
  }
  else {
    tertiaryContrastColor = 'black';
  }

  customStyles.primaryContrastColor = primaryContrastColor;
  customStyles.secondaryContrastColor = secondaryContrastColor;
  customStyles.tertiaryContrastColor = tertiaryContrastColor;

  if (ssName.includes('act admin answer analysis') || ssName.includes('act student answer sheet')) {
    for (let i in ss.getSheets()) {
      const sh = ss.getSheets()[i];
      const shRange = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
      shRange.setBackground('white');
      shRange.setFontColor(fontColor);

      let shName = sh.getName().toLowerCase();

      if (shName.endsWith("z")) {
        shName = shName.substring(0, shName.length - 1);
      }

      const isTestSheet = /^\d+$/.test(shName)

      if (isTestSheet) {
        sh.getRange('A1:P4').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);

        sh.getRangeList(['B3', 'F3', 'J3', 'N3']).setBorder(true, true, true, true, true, true, '#93c47d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

        sh.getRange('F1').setBackground('#93c47d');
      }
      else if (shName === 'test analysis' || shName === 'opportunity area analysis') {
        sh.getRange(1, 1, 8, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

        if (shName === 'test analysis') {
          var correctRange = 'F7:J7';
        }
        else {
          var correctRange = 'D7:H7';
        }
        sh.getRange(correctRange).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        
        const imgCell = sh.getRange('B3')
        if(imgUrl) {
          imgCell.setValue('=image("'+ customStyles.img + '")');
        }

        applyConditionalFormatting(sh, customStyles);
      }
      else if (actDataSheets.includes(shName)) {
        sh.getRange(1, 1, 1, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor);
      }
      else if (shName === 'student responses') {
        sh.getRange(1, 1, 3, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
      }
    }       
  }
  else if (ssName.includes('sat admin answer analysis') || ssName.includes('sat student answer sheet')) {

    for (let i in ss.getSheets()) {
      const sh = ss.getSheets()[i];
      const shRange = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()-2);
      shRange.setBackground('white');
      shRange.setFontColor(fontColor);

      shName = sh.getName().toLowerCase();

      // practice SAT answer sheets
      if (satTestSheets.includes(shName)) {
        // sh.getRangeList(['B2:L4', 'B33:L35']).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
        sh.getRangeList(['B2:L4', 'B33:L35']).setBackground(secondaryColor).setFontColor(secondaryContrastColor).setBorder(true, true, true, true, true, true, secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
        sh.getRangeList(['A5:A', 'E5:E', 'I5:I']).setFontColor('white');
      }
      // check for SAT analysis sheets after checking exact match
      else if (shName.includes('analysis') || shName.includes('opportunity')) {
        if(shName === 'rev analysis') {
          sh.getRange('A1:K7').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        }
        else if (shName === 'time series analysis') {
          sh.getRange('A1:K6').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          sh.getRange('D5:E6').setFontColor(fontColor)
        }
        else {
          sh.getRange('A1:K6').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        }
        
        const imgCell = sh.getRange('B2');
        if(imgUrl) {
          imgCell.setValue('=image("'+ customStyles.img + '")');
        }

        applyConditionalFormatting(sh, customStyles);
      }

      else if (shName === 'reading & writing') {
        styleSatWorksheets(sh, 6, 11, customStyles)
      }
      else if (shName === 'math') {
        styleSatWorksheets(sh, 9, 11, customStyles)
      }
      else if (shName === 'slt uniques') {
        styleSatWorksheets(sh, 1, 7, customStyles)
      }
      else if (satDataSheets.includes(shName)) {
        sh.getRange(1, 1, 1, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor);
      }
      else if (shName === 'student responses') {
        sh.getRange(1, 1, 3, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
      }
      else if (shName === 'rev sheets') {
        let revSheetHeaderRange;
        if (ssName.includes('sat admin answer analysis')) {
          revSheetHeaderRange = sh.getRangeList(['B2:E4', 'G2:J4']);
        }
        else {
          revSheetHeaderRange = sh.getRangeList(['B2:D4', 'F2:I4']);
        }
        // revSheetHeaderRange.setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
        sh.getRangeList(['B2:E4', 'G2:J4']).setBackground(secondaryColor).setFontColor(secondaryContrastColor).setBorder(true, true, true, true, true, true, secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  }
}

function styleSatWorksheets(
  sh=SpreadsheetApp.openById('1FW_3GIWmytdrgBdfSuIl2exy9hIAnQoG8IprF8k9uEY').getSheetByName('Math'),
  rowOffset=10,
  headerCols=11,
  customStyles={
    'primaryColor': '#134f5c',
    'primaryContrastColor': 'white',
  }
  ) {
  const cats = [
    'Area and volume',
    'Reading & Writing', // styles header in SLT Uniques
    'Boundaries',
    'Central ideas and details',
    'Circles',
    'Command of evidence',
    'Cross-text connections',
    'Distributions',
    'Equivalent expressions',
    'Form, structure, and sense',
    'Inferences',
    'Linear equations in one variable',
    'Linear equations in two variables',
    'Linear functions',
    'Linear inequalities',
    'Lines, angles, and triangles',
    'Models and scatterplots',
    'Nonlinear equations and systems',
    'Nonlinear functions',
    'Observational studies and experiments',
    'Percentages',
    'Probability',
    'Ratios, rates, proportions, and units',
    'Systems of linear equations',
    'Right triangles and trigonometry',
    'Sample statistics and margin of error',
    'Words in context',
    'Transitions',
    'Rhetorical synthesis',
    'Text, structure, and purpose',
  ]
  var conceptRows = []

  sh.getRange(1,1,sh.getMaxRows()).setFontColor('white');
  sh.getRange(1,5,sh.getMaxRows()).setFontColor('white');
  sh.getRange(1,9,sh.getMaxRows()).setFontColor('white');

  const colVals = sh.getRange(rowOffset, 2, sh.getMaxRows() - rowOffset).getValues();

  for (let x = 0; x < colVals.length; x++) {
    if(cats.includes(colVals[x][0])) {
      var row = x + rowOffset;
      conceptRows.push(row);
    }
  }
  for(r in conceptRows) {
      const highlightRange = sh.getRange(conceptRows[r], 2, 3, headerCols);
      // highlightRange.setBackground(customStyles.primaryColor).setFontColor(customStyles.primaryContrastColor).setBorder(true, true, true, true, true, true, customStyles.primaryColor, SpreadsheetApp.BorderStyle.SOLID);
      highlightRange.setBackground(customStyles.secondaryColor).setFontColor(customStyles.secondaryContrastColor).setBorder(true, true, true, true, true, true, customStyles.secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
  }
}

function applyConditionalFormatting(
  sheet=SpreadsheetApp.openById('1XoMGHjanL9w1xSqS6Q1kdvZnJHPXgLJgph4cWxFDx7A').getSheetByName('SAT3 analysis'), // SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
  customStyles={
    'isActive': true,
    'primaryColor': '#1c4d65',
    'secondaryColor': '#f6b26b',
    'tertiaryColor': '#efefef',
    'fontColor': 'black',
    'primaryContrastColor': 'white',
    'secondaryContrastColor': 'white',
    'tertiaryContrastColor': 'black'
  })
  
  {
  
  var rules = []
  var newRules = []

  for (i in sheet.getConditionalFormatRules()) {
    var condition = sheet.getConditionalFormatRules()[i]
    rules.push(condition)
  }

  for (i in rules) {
    if (rules[i].getGradientCondition()) {
      Logger.log(rules[i].getGradientCondition())
      newRule = rules[i].copy();
      newRules.push(newRule)
    }
  }

  if (sheet.getName().toLowerCase().includes('opportunity')) {
    var subtotalStart = 'B';
    var domainStart = 'C';
  }
  else {
    var subtotalStart = 'C';
    var domainStart = 'D'
  }
  var grandTotalRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$B7="Grand total"')
        .setBold(true)
        .setBackground(customStyles.primaryColor)
        .setFontColor(customStyles.primaryContrastColor)
        .setRanges([sheet.getRange('B7:I70')]);
  
  var subTotalRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=right($' + subtotalStart + '7,5)="Total"')
        .setBold(true)
        .setBackground(customStyles.secondaryColor)
        .setFontColor(customStyles.secondaryContrastColor)
        .setRanges([sheet.getRange(subtotalStart + '7:I70')]);
  
  var domainTotalRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=right($' + domainStart + '7,5)="Total"')
        .setBackground(customStyles.tertiaryColor)
        .setFontColor(customStyles.tertiaryContrastColor)
        .setRanges([sheet.getRange(domainStart + '7:I70')]);
  
  var backgroundColorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=sum($F7:$I7)>0')
        .setBackground('#f5f7f9')
        .setRanges([sheet.getRange('B7:I70')]);

  newRules.push(grandTotalRule, subTotalRule, domainTotalRule, backgroundColorRule);
  sheet.clearConditionalFormatRules();
  sheet.setConditionalFormatRules(newRules);
}

function isDark (hex='#b6d7a8') {
  hex = hex.substring(1);      // strip #
  var rgb = parseInt(hex, 16);   // convert rrggbb to decimal
  var r = (rgb >> 16) & 0xff;  // extract red
  var g = (rgb >>  8) & 0xff;  // extract green
  var b = (rgb >>  0) & 0xff;  // extract blue

  var luma = 0.2126 * r + 0.7152 * g + 0.0722 * b; // per ITU-R BT.709

  Logger.log(luma);
  if (luma < 205) {
    return true;
  }
  else {
    return false;
  }
}


function findNewScoreReports(parentFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF') {
  if (typeof parentFolderId == "object") {
    parentFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF';
  }

  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var fileList = getAnalysisFiles(parentFolder, n=3);
  Logger.log(fileList);

  // Sort by most recently updated first
  fileList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());

  analysisSsSearch(fileList);
}

function findTeamScoreReports() {
  var teamFolderId = '1tSKajFOa_EVUjH8SKhrQFbHSjDmmopP9';
  var teamFolder = DriveApp.getFolderById(teamFolderId);
  var tutorFolders = teamFolder.getFolders();

  // Check subfolders of _Team for score reports
  while (tutorFolders.hasNext()) {
    var tutorFolder = tutorFolders.next();

    findNewScoreReports(tutorFolder.getId());
  }
}

function getAnalysisFiles(folder, n=3, fileList=[]) {
  folder = folder || DriveApp.getRootFolder();
  var folderName = folder.getName().toLowerCase();
  if (!folderName.includes('archive')) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName().toLowerCase();
      if (fileName.includes('admin answer analysis') && !fileName.includes('template')) {
        fileList.push(file);
      }
    }

    if (--n == 0) return;

    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      getAnalysisFiles(subfolders.next(), n, fileList);
    }
  }

  return fileList;
}

function analysisSsSearch(fileList) {
  var testCodes = ['at1', 'at2', 'at3', 'at4', 'at5', 'at6', 'sat1', 'sat2', 'sat3', 'sat4', 'sat5', 'sat6', 'psat1', 'psat2', 'apt1', 'apt2'];
  var scoreSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('Scores');
  var lastRow = scoreSheet.getLastRow();
  var scoresNames = scoreSheet.getRange(1,1,lastRow);
  var nextOpenRow = lastRow - scoresNames.getValues().reverse().findIndex(c => c[0] != '') + 1;

  // Loop through analysis spreadsheets
  for (var i = 0; i < fileList.length; i++) {
    var ssId = fileList[i].getId();
    var ss = SpreadsheetApp.openById(ssId);
    var ssName = ss.getName();
    var studentName = ssName.slice(ssName.indexOf('-') + 2);
    var sheets = ss.getSheets();
    var scores = [];

    // Loop through sheets within analysis spreadsheet
    for (var s = 0; s < sheets.length; s++) {
      sheet = sheets[s];
      var sheetName = sheet.getName().toLowerCase();
      if (testCodes.includes(sheetName)) {
        // Check last answer for each module
        var mod1RWEnd = sheet.getRange('C31').getValues().filter(String).length;
        var mod2RWEnd = sheet.getRange('G31').getValues().filter(String).length;
        var mod3RWEnd = sheet.getRange('K31').getValues().filter(String).length;
        var mod1MathEnd = sheet.getRange('C57').getValues().filter(String).length;
        var mod2MathEnd = sheet.getRange('G57').getValues().filter(String).length;
        var mod3MathEnd = sheet.getRange('K57').getValues().filter(String).length;
        var modsComplete = mod1RWEnd + mod2RWEnd + mod3RWEnd + mod1MathEnd + mod2MathEnd + mod3MathEnd;
        var completionCheck = sheet.getRange('M1');
        var testName = sheetName.toUpperCase();

        if (sheetName.slice(0, 3) === 'apt') {
          testName = sheetName.replace('apt', 'psat').toUpperCase();
        }
        else if (sheetName.slice(0, 2) === 'at') {
          testName = sheetName.replace('at', 'sat').toUpperCase();
        }

        // If test is completed, add to scores array
        if (modsComplete === 4 && sheet.getRange('G1').getValue() >= 200 && sheet.getRange('I1').getValue() >= 200) {
          var rwScore = sheet.getRange('G1').getValue();
          var mScore = sheet.getRange('I1').getValue();
          var totalScore = sheet.getRange('L1').getValue();
          scores.push({
            'test': testName,
            'rw': rwScore,
            'm': mScore,
            'total': totalScore
          });

          // If test is newly completed, create score report
          if (completionCheck.getValue() !== '✔') {
            Logger.log(ssName + " " + testName + " score report started");
            createSatScoreReport(ssId, sheetName, scores);
            SpreadsheetApp.flush();

            completionCheck.setValue('✔');
            Logger.log(ssName + " " + testName + " score report complete");
            completionCheck.setVerticalAlignment('middle');
            completionCheck.setFontColor('#134f5c');

            var dateSubmitted = sheet.getRange('D2').getValue();
            if (dateSubmitted = '') {
              dateSubmitted =  Utilities.formatDate(new Date(new Date().getFullYear(),new Date().getMonth(),new Date().getDate()-1), 'UTC', 'MM/dd/yyyy');
            }
            var rowData = [[studentName, 'Practice', testName.toUpperCase(), dateSubmitted, totalScore, rwScore, mScore]]
            scoreSheet.getRange(nextOpenRow,1,1,7).setValues(rowData);
            nextOpenRow += 1;
          }
        }
      }
    }
  }
}

function createSatScoreReport(spreadsheetId, testCode, scores) {
  var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();

  var sheetsToPrint = [testCode.toLowerCase(), testCode.toLowerCase() + ' analysis']
  var fileName = spreadsheet.getName();
  var studentName = fileName.slice(fileName.indexOf('-') + 2);
  var analysisIndex = 1;

  showAllExcept(spreadsheetId);
  SpreadsheetApp.flush();

  /* PDF can be created from single sheet or all visible sheets. For a multi-sheet PDF, we need to hide
  unwanted sheets, save the PDF, then show all sheets again. */
  SpreadsheetApp.openById(spreadsheetId).getSheets().forEach(sh => {
    try {
      if (sheetsToPrint.includes(sh.getName().toLowerCase())) {
        sh.showSheet();
        if (sh.getName().includes('analysis')) {
          analysisIndex = sh.getIndex();
          spreadsheet.setActiveSheet(sh);
          // Move analysis sheet to first position so that it displays first in PDF
          spreadsheet.moveActiveSheet(1);
          // Hide column H if student did not omit any answers
          if (sh.getRange('H7').getValue() === '-') {
            sh.hideColumns(8);
          }
          else if (sh.getRange('H7').getValue() === 'BLANK') {
            sh.showColumns(8);
          }
        }
      }
      else {
        sh.hideSheet();
      }
    } catch (error) {
      Logger.log(error);
    }

  });

  var email = getOPTPermissionsList(spreadsheetId);
  SpreadsheetApp.flush();
  sendPdfScoreReport(spreadsheetId, email, studentName, scores);
  Logger.log(testCode.toUpperCase() + ' Score report created for ' + studentName);
  // SpreadsheetApp.flush();
  showAllExcept(spreadsheetId);
  // Move analysis sheet back to original position
  spreadsheet.moveActiveSheet(analysisIndex);
}

// Save spreadsheet as a PDF: https://gist.github.com/andrewroberts/26d460212874cdd3f645b55993942455
function sendPdfScoreReport(spreadsheetId, email, studentName, scores = []) {

  var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId()
  //var sheetId = sheetName ? spreadsheet.getSheetByName(sheetName).getSheetId() : null;
  var practiceDataSheet = spreadsheet.getSheetByName('Practice test data');

  if (practiceDataSheet.getRange('V1').getValue() === 'Score report folder ID:' && practiceDataSheet.getRange('W1').getValue() !== '') {
    var scoreReportFolderId = practiceDataSheet.getRange('W1').getValue();
  }
  else {
    var parentId = DriveApp.getFileById(spreadsheetId).getParents().next().getId();
    const subfolderIds = getSubFolderIdsByFolderId(parentId)

    for (let i in subfolderIds) {
      let subfolderId = subfolderIds[i];
      let subfolder = DriveApp.getFolderById(subfolderId)
      let subfolderName = subfolder.getName();
      if (subfolderName.toLowerCase().includes('score report')) {
        var scoreReportFolderId = subfolder.getId();
      }
    }
  }

  if (!scoreReportFolderId) {
    var scoreReportFolderId = DriveApp.getFolderById(parentId).createFolder('Score reports').getId();
  }

  practiceDataSheet.getRange('V1:W1').setValues([['Score report folder ID:', scoreReportFolderId]]);


  var url_base = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/"
  var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf
    // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
    //+ (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId))
    + '&id=' + spreadsheetId
    // following parameters are optional...
    + '&size=letter'      // paper size
    + '&portrait=true'    // orientation, false for landscape
    + '&fitw=true'        // fit to width, false for actual size
    + '&fzr=false'       // do not repeat row headers (frozen rows) on each page
    + '&top_margin=0.5'
    + '&bottom_margin=0.5'
    + '&left_margin=0.3'
    + '&right_margin=0.3'
    + '&printnotes=false'
    + '&sheetnames=false'
    + '&printtitle=false'
    + '&pagenumbers=false';  //hide optional headers and footers    

  var options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
    }
  }

  // Create PDF
  var currentScore = scores.slice(-1)[0];
  var pdfName = 'SAT answer analysis - ' + studentName + " - " + currentScore.test;
  var studentFirstName = studentName.split(" ")[0];
  const [studentHours, recentSessionDate] = getStudentHours(studentName);
  var response = UrlFetchApp.fetch(url_base + url_ext, options);
  var blob = response.getBlob().setName(pdfName + '.pdf');
  var scoreReportFolder = DriveApp.getFolderById(scoreReportFolderId);
  scoreReportFolder.createFile(blob);
  var message = "Hi PARENTNAME, please find the score report from " + studentFirstName + "'s recent practice test attached. " + currentScore.total + " overall (" + currentScore.rw + " Reading & Writing, " + currentScore.m + " Math)<br><br>As of the session on " + recentSessionDate + ", we have " + studentHours + " hours remaining on the current package. Let me know if you have any questions. Thanks!<br><br>"

  if (scores.length > 1) {
    message += "Previous scores - most recent last:<br><ul>"

    for (i = 0; i < scores.length - 1; i++) {
      message += "<li>" + scores[i].test + ": " + scores[i].total + " (" + scores[i].rw + " RW, " + scores[i].m + " M)</li>";
    }
    message += "</ul><br>";
  }

  if (email) {
    MailApp.sendEmail({
      to: email,
      subject: currentScore.test + " Score Report for " + studentFirstName,
      htmlBody: message,
      attachments: [blob.getAs(MimeType.PDF)]
    });
  }
}

function getStudentHours(studentName) {
  const summarySheet = SpreadsheetApp.openById('1M6Xs6zLR_QdPpOJYO0zaZOwJZ6dxdXsURD2PkpP2Vis').getSheetByName('Summary');
  const lastRow = summarySheet.getLastRow();
  const allVals = summarySheet.getRange("A1:A" + lastRow).getValues();
  const lastFilledRow = lastRow - allVals.reverse().findIndex(c => c[0] != '');
  var summaryData = summarySheet.getRange(1, 1, lastFilledRow, 26).getValues();

  for (let r = 0; r < lastFilledRow; r++) {
    if (summaryData[r][0] === studentName) {
      return [summaryData[r][3], Utilities.formatDate(new Date(summaryData[r][16]), "GMT", "EEE M/d")];
    }
  }
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

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}


function getOPTPermissionsList(id) {
  var editors = DriveApp.getFileById(id).getEditors().map(function (e) { return e.getEmail() });
  var emails = '';

  for (var i = 0; i < editors.length; i++) {
    // Only add openpathtutoring.com emails to email list
    if (editors[i].includes('openpathtutoring.com')) {
      emails += editors[i] + ',';
    }
  }

  return emails;
}

const showAllExcept = (spreadsheetId, hiddenSheets = []) => {
  SpreadsheetApp.openById(spreadsheetId).getSheets().forEach(sh => {
    // If sheets are meant to be hidden, leave them hidden
    if (!hiddenSheets.includes(sh.getName())) {
      sh.showSheet();
    }
  });
  // SpreadsheetApp.flush();
}

function renameStudentFolder(folder, studentCurrentName, studentFullName) {
  var folderName = folder.getName();
  var files = folder.getFiles();
  var subfolders = folder.getFolders();

  if (folderName.includes(studentCurrentName) && !folderName.includes(studentFullName)) {
    var newFoldername = folderName.replace(studentCurrentName, studentFullName);
    folder.setName(newFoldername);
  }

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    if (fileName.includes(studentCurrentName) && !fileName.includes(studentFullName)) {
      var newFileName = fileName.replace(studentCurrentName, studentFullName);
      file.setName(newFileName);
    }
  }

  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var subfolderName = subfolder.getName();

    if (subfolderName.includes(studentCurrentName) && !subfolderName.includes(studentFullName)) {
      var newSubfolderName = subfolderName.replace(studentCurrentName, studentFullName);
      subfolder.setName(newSubfolderName);
    }

    renameStudentFolder(subfolder, studentCurrentName, studentFullName);
  }
}

// Sorting an array of files by datecreated
function sortFoldersByDateCreated() {
  var folder = DriveApp.getFolderById('135fkvQMWxdhInu4OiPecJ9jS_vajAOYo');
  var contents = folder.getFolders();
  let arr = [];
  while (contents.hasNext()) {
    let folder = contents.next();
    arr.push(folder);
  }
  //sort arr by dateCreated
  arr.sort((a, b) => {
    let vA = new Date(a.getDateCreated()).valueOf();
    let vB = new Date(b.getDateCreated()).valueOf();
    return vA - vB
  });
  Logger.log(arr);
}


// Rev sheet setup functions

function getAllRowHeights() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Rev sheet backend');
  var rwIds = sh.getRange('M802:M').getValues();
  var mathIds = sh.getRange('P2:P').getValues();

  // rwHeights = [];
  // for (var r=0; r < rwIds.length; r++) {
  //   var id = rwIds[r][0];
    
  //   var height = calculateRowHeight(id, 820, 'rw');
  //   rwHeights.push([height]);
  //   if((r+1) % 100 === 0) {
  //     var slice = rwHeights.slice(r-99,r+1);
  //     sh.getRange(800+r-97,15,100).setValues(slice);
  //     Logger.log(slice);
  //   }
  // };

  mathHeights = [];
  for (var m=0; m < mathIds.length; m++) {
    var id = mathIds[m][0];
    
    var height = calculateRowHeight(id, 820, 'math');
    mathHeights.push([height]);
    if((m+1) % 100 === 0) {
      var slice = mathHeights.slice(m-99,m+1);
      sh.getRange(m-97,18,100).setValues(slice);
      Logger.log(slice);
    }
  };
}

function calculateRowHeight(questionId, containerWidth, subject) {
  var questionUrl = 'https://www.openpathtutoring.com/static/img/concepts/sat/' + subject.toLowerCase() + '/' + encodeURIComponent(questionId) + ".jpg";
  var urlOptions = {muteHttpExceptions: true};
  
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
  }
  else {
    var whitespace = 60;
  }

  var rowHeight = questionSize.height / questionSize.width * containerWidth + whitespace;

  Logger.log(questionId + ' rowHeight: ' + rowHeight);

  return Math.round(rowHeight);
}


// function getClassFolderId() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const classFolder = DriveApp.getFileById(ss.getId()).getParents().next().getParents().next().getParents().next();
//   const classFolderId = classFolder.getId();
//   var files = classFolder.getFiles();
//   var aggSsId = null;

//   while (files.hasNext()) {
//     file = files.next();
//     fileName = file.getName();

//     if (fileName.toLowerCase().includes('aggregate answer analysis')) {
//       aggSsId = file.getId();
//     }
//   }

//   if (aggSsId === null) {
//     const parentFolder = classFolder.getParents().next();
//     const parentFiles = parentFolder.getFiles();

//     Logger.log(parentFolder.getName());

//     while (parentFiles.hasNext()) {
//       parentFile = parentFiles.next();

//       if (parentFile.getName().toLowerCase().includes('aggregate answer analysis')) {
//         aggSsId = parentFile.makeCopy().moveTo(classFolder).getId();
//         DriveApp.getFileById(aggSsId).setName(classFolder.getName() + ' aggregate answer analysis');
//       }
//     }

//   }

//   Logger.log(classFolderId + " " + aggSsId);

//   generateClassTestAnalysis(classFolderId, aggSsId);

//   return aggSsId;
// }