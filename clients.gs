function newClient(clientTemplateFolderId, clientParentFolderId) {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Tutor or Business name:', ui.ButtonSet.OK_CANCEL);
  let customStyles, clientName;

  if (prompt.getSelectedButton() === ui.Button.CANCEL) {
    return;
  } //
  else {
    clientName = prompt.getResponseText();
  }

  const useCustomStyle = ui.alert('Apply custom styles?', ui.ButtonSet.YES_NO);

  if (useCustomStyle === ui.Button.YES) {
    customStyles = setCustomStyles();
  }

  var clientTemplateFolder = DriveApp.getFolderById(clientTemplateFolderId);
  var clientParentFolder = DriveApp.getFolderById(clientParentFolderId);
  let clientFolder = clientParentFolder.createFolder(clientName);
  let clientFolderId = clientFolder.getId();

  copyClientFolder(clientTemplateFolder, clientFolder, clientName);
  linkClientSheets(clientFolderId);
  setClientDataUrls(clientFolderId);
  addClientData(clientFolderId);

  if (useCustomStyle === ui.Button.YES) {
    getAnswerSheets(clientFolder);
    processFolders(clientFolder.getFolders(), getAnswerSheets);
    styleClientSheets(satSsIds, actSsIds, customStyles);
  }

  var htmlOutput = HtmlService.createHtmlOutput(
    '<a href="https://drive.google.com/drive/u/0/folders/' +
      clientFolderId +
      '" target="_blank" onclick="google.script.host.close()">' +
      clientName +
      "'s folder</a>" +
      '<p><a href="https://docs.google.com/spreadsheets/d/' +
      PropertiesService.getScriptProperties().getProperty('clientDataSsId') +
      '"target="_blank">Client data IDs</a></p>'
  )
    .setWidth(250) //optional
    .setHeight(100); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Client folder created successfully');
}

function copyClientFolder(sourceFolder, newFolder, clientName) {
  const folders = sourceFolder.getFolders();
  const files = sourceFolder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var filename = file.getName();

    if (filename.includes('template')) {
      const rootName = filename.slice(0, filename.indexOf('-') + 2);

      if (filename.toLowerCase().includes('data - client')) {
        filename = rootName + clientName;
      } //
      else {
        filename = rootName + 'Template for ' + clientName;
      }
    }

    file.makeCopy(filename, newFolder);
  }

  while (folders.hasNext()) {
    var folder = folders.next();
    var folderName = folder.getName();
    var targetFolder = newFolder.createFolder(folderName);

    copyClientFolder(folder, targetFolder, clientName);
  }
}

function linkClientSheets(folderId, testType = 'all') {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var subFolders = DriveApp.getFolderById(folderId).getFolders();

  while (files.hasNext()) {
    const file = files.next();
    const filename = file.getName();
    if (filename.includes('SAT')) {
      if (filename.includes('student answer sheet')) {
        satSsIds.student = file.getId();
        DriveApp.getFileById(satSsIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      } //
      else if (filename.includes('answer analysis')) {
        satSsIds.admin = file.getId();
      }
    }

    if (filename.includes('ACT')) {
      if (filename.toLowerCase().includes('student answer sheet')) {
        actSsIds.student = file.getId();
        DriveApp.getFileById(actSsIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      } //
      else if (filename.toLowerCase().includes('answer analysis')) {
        actSsIds.admin = file.getId();
      }
    }
  }

  while (subFolders.hasNext()) {
    if ((satSsIds.student && satSsIds.admin && testType === 'act') || (actSsIds.student && actSsIds.admin && testType === 'sat') || (satSsIds.student && satSsIds.admin && actSsIds.student && actSsIds.admin && testType === 'all')) {
      break;
    }
    const subFolder = subFolders.next();
    linkClientSheets(subFolder.getId(), testType);
  }

  if (satSsIds.student && satSsIds.admin) {
    let satAdminSheet = SpreadsheetApp.openById(satSsIds.admin);
    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSsIds.student);
  }

  if (actSsIds.student && actSsIds.admin) {
    SpreadsheetApp.openById(actSsIds.admin).getSheetByName('Student responses').getRange('B1').setValue(actSsIds.student);
  }
}

function setClientDataUrls(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var subFolders = DriveApp.getFolderById(folderId).getFolders();

  while (files.hasNext()) {
    file = files.next();
    fileId = file.getId();
    filename = file.getName().toLowerCase();

    if (filename.includes('sat admin data')) {
      Logger.log('found sat admin data');
      satSsIds.adminData = fileId;
    } //
    else if (filename.includes('sat student data')) {
      Logger.log('found sat student data');
      satSsIds.studentData = fileId;
      DriveApp.getFileById(satSsIds.studentData).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } //
    else if (filename.includes('sat student answer sheet')) {
      Logger.log('found sat student answer sheet');
      satSsIds.student = fileId;
      DriveApp.getFileById(satSsIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } //
    else if (filename.includes('sat admin answer analysis')) {
      Logger.log('found sat admin answer sheet');
      satSsIds.admin = fileId;
    } //
    else if (filename.includes('rev sheet data')) {
      Logger.log('found rev sheet data');
      satSsIds.rev = fileId;
      DriveApp.getFileById(satSsIds.rev).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } //
    else if (filename.includes('act admin data')) {
      Logger.log('found act admin data');
      actSsIds.adminData = fileId;
    } //
    else if (filename.includes('act student data')) {
      Logger.log('found act student data');
      actSsIds.studentData = fileId;
      DriveApp.getFileById(actSsIds.studentData).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } //
    else if (filename.includes('act student answer sheet')) {
      Logger.log('found act student answer sheet');
      actSsIds.student = fileId;
      DriveApp.getFileById(actSsIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } //
    else if (filename.includes('act admin answer analysis')) {
      Logger.log('found act admin answer sheet');
      actSsIds.admin = fileId;
    }
  }

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    setClientDataUrls(subFolder.getId());
  }

  if (satSsIds.admin && satSsIds.student) {
    SpreadsheetApp.openById(satSsIds.admin).getSheetByName('Student responses').getRange('B1').setValue(satSsIds.student);
  }

  if (satSsIds.student && satSsIds.studentData) {
    SpreadsheetApp.openById(satSsIds.student).getSheetByName('Question bank data').getRange('U7').setValue(satSsIds.studentData);
  }
  if (satSsIds.admin && satSsIds.adminData) {
    SpreadsheetApp.openById(satSsIds.admin).getSheetByName('Rev sheet backend').getRange('U5').setValue(satSsIds.adminData);
  }

  if (satSsIds.adminData && satSsIds.studentData) {
    const studentDataSs = SpreadsheetApp.openById(satSsIds.studentData);

    studentDataSs
      .getSheetByName('Question bank data')
      .getRange('A1')
      .setValue('=IMPORTRANGE("' + satSsIds.adminData + '", "Question bank data!A1:G10000")');
    studentDataSs
      .getSheetByName('Practice test data')
      .getRange('A1')
      .setValue('=IMPORTRANGE("' + satSsIds.adminData + '", "Practice test data!A1:E10000")');
    studentDataSs
      .getSheetByName('Links')
      .getRange('A1')
      .setValue('=IMPORTRANGE("' + satSsIds.adminData + '", "Links!A1:D50")');
  }

  if (satSsIds.admin && satSsIds.rev) {
    let adminRevSheet = SpreadsheetApp.openById(satSsIds.admin).getSheetByName('Rev sheet backend');
    adminRevSheet.getRange('U3').setValue(satSsIds.rev);
    DriveApp.getFileById(satSsIds.rev).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  }

  if (satSsIds.student && satSsIds.rev) {
    let studentSheet = SpreadsheetApp.openById(satSsIds.student).getSheetByName('Question bank data');
    studentSheet.getRange('U3').setValue(satSsIds.rev);
  }

  if (actSsIds.student && actSsIds.studentData) {
    SpreadsheetApp.openById(actSsIds.student)
      .getSheetByName('Data')
      .getRange('A1')
      .setValue('=IMPORTRANGE("' + actSsIds.studentData + '", "Data!A1:D10000")');
  }

  if (actSsIds.admin && actSsIds.adminData) {
    var ss = SpreadsheetApp.openById(actSsIds.admin);
    ss.getSheetByName('Data')
      .getRange('A1')
      .setValue('=IMPORTRANGE("' + actSsIds.adminData + '", "Data!A1:G10000")');
    ss.getSheets()[0]
      .getRange('E1')
      .setValue('=IMPORTRANGE("' + actSsIds.adminData + '", "Data!Q1")');
    ss.getSheets()[0].getRange('D1').setValue('=iferror(E1,"Click to connect data >>")');
  }

  if (actSsIds.adminData && actSsIds.studentData) {
    SpreadsheetApp.openById(actSsIds.studentData)
      .getSheetByName('Data')
      .getRange('A1')
      .setValue('=IMPORTRANGE("' + actSsIds.adminData + '", "Data!A1:D10000")');
  }

  Logger.log('setClientDataUrls complete');
}

function updateClientsList(parentClientFolderId = '130wX98bJM4wW6aE6J-e6VffDNwqvgeNS') {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const parentClientFolder = DriveApp.getFolderById(parentClientFolderId);
  const clientFolders = parentClientFolder.getFolders();
  const clientSheet = clientDataSs.getSheetByName('Clients');
  let newRow = getLastFilledRow(clientSheet, 1) + 1;

  while (clientFolders.hasNext()) {
    const clientFolder = clientFolders.next();
    const clientFolderId = clientFolder.getId();
    const clientFolderName = clientFolder.getName();

    if (!(clientFolderName.includes('_') || clientFolderName.includes('Ξ'))) {
      addClientData(clientFolderId, newRow);
      newRow++;
    }
  }
}

function addClientData(clientFolderId = '1Fd99S1DPdWuvr1VxkeEbdAZn_ZmP9PPj', newRow = null) {
  const clientFolder = DriveApp.getFolderById(clientFolderId);
  const clientName = clientFolder.getName();
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheetByName('Clients');
  const collaborators = clientFolder.getEditors();
  const emailList = [];
  collaborators.forEach((c) => {
    if (!PropertiesService.getScriptProperties().getProperty('myEmails').includes(c.getEmail())) {
      Logger.log(c.getEmail());
      // emailList.push(c.getEmail());
    }
  });

  if (!newRow) {
    newRow = getLastFilledRow(clientSheet, 1) + 1;
  }

  const savedClientFolderIds = clientSheet.getRange(2, 4, newRow).getValues();
  let clientIndex = savedClientFolderIds.findIndex((subArray) => subArray.includes(clientFolderId));
  let studentFolder, studentFolderId, studentFolderCount;

  if (clientIndex === -1) {
    clientIndex = newRow - 2;
    Logger.log(clientIndex + '. ' + clientName);

    getAnswerSheets(clientFolder);
    processFolders(clientFolder.getFolders(), getAnswerSheets);

    const clientSubfolders = clientFolder.getFolders();
    let dataFolderId,
      satAdminDataId,
      satStudentDataId,
      actAdminDataId,
      actStudentDataId,
      revDataId = '';

    while (clientSubfolders.hasNext()) {
      const clientSubfolder = clientSubfolders.next();
      studentFolderCount = 0;

      if (clientSubfolder.getName().toLowerCase().includes('students')) {
        studentFolder = clientSubfolder;
        studentFolderId = studentFolder.getId();

        studentFolderCount = getStudentFolderCount(studentFolderId);
      } //
      else if (clientSubfolder.getName().toLowerCase().includes('data')) {
        dataFolder = clientSubfolder;
        dataFolderId = clientSubfolder.getId();
        const dataFiles = dataFolder.getFiles();
        while (dataFiles.hasNext()) {
          const file = dataFiles.next();
          const filenameLower = file.getName().toLowerCase();
          if (filenameLower.includes('sat admin')) {
            satAdminDataId = file.getId();
          } //
          else if (filenameLower.includes('sat student')) {
            satStudentDataId = file.getId();
          } //
          else if (filenameLower.includes('act admin')) {
            actAdminDataId = file.getId();
          } //
          else if (filenameLower.includes('act student')) {
            actStudentDataId = file.getId();
          } //
          else if (filenameLower.includes('rev sheet data')) {
            revDataId = file.getId();
          }
        }
      }
    }

    clientDataSs
      .getSheetByName('Clients')
      .getRange(newRow, 1, 1, 16)
      .setValues([[clientIndex, clientName, emailList, clientFolder.getId(), satSsIds.admin, satSsIds.student, actSsIds.admin, actSsIds.student, dataFolderId, satAdminDataId, satStudentDataId, actAdminDataId, actStudentDataId, revDataId, studentFolderId, studentFolderCount]]);
    newRow++;
  } //
  else {
    studentFolderId = clientSheet.getRange(clientIndex + 2, 15).getValue();
    studentFolderCount = getStudentFolderCount(studentFolderId);
    // clientSheet.getRange(clientIndex + 2, 16).setValue(studentFolderCount);
    // Logger.log(clientFolder.getName() + ' present with ' + studentFolderCount + ' student folders');
  }
}

function updateClientFolders() {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheetByName('Clients');
  const lastFilledRow = getLastFilledRow(clientSheet, 2);
  const clientDataRange = clientSheet.getDataRange().getValues();
  const clients = [];

  for (let row = 1; row < lastFilledRow; row++) {
    // starts at 1
    clients.push({
      index: clientDataRange[row][0],
      name: clientDataRange[row][1],
      emails: clientDataRange[row][2],
      folderId: clientDataRange[row][3],
      satAdminSsId: clientDataRange[row][4],
      satStudentSsId: clientDataRange[row][5],
      actAdminSsId: clientDataRange[row][6],
      actStudentSsId: clientDataRange[row][7],
      dataFolderId: clientDataRange[row][8],
      satAdminDataId: clientDataRange[row][9],
      satStudentDataId: clientDataRange[row][10],
      actAdminDataId: clientDataRange[row][11],
      actStudentDataId: clientDataRange[row][12],
      revDataId: clientDataRange[row][13],
      studentsFolderId: clientDataRange[row][14],
    });
  }

  for (let client of clients) {
    const startIndex = PropertiesService.getScriptProperties().getProperty('startIndex');

    if (client.index >= startIndex /* 0 is OPT folder */) {
      const clientRow = client.index + 2;
      const studentsDataStr = clientSheet.getRange(clientRow, 17).getValue();
      client.studentsDataJSON = JSON.parse(studentsDataStr);

      const students = TestPrepAnalysis.getAllStudentFileIds(client);

      TestPrepAnalysis.ssUpdate202505(students);

      PropertiesService.getScriptProperties().setProperty('startIndex', client.index + 1);
      Logger.log(client.index + '. ' + client.name + ' complete');
    }
  }
  PropertiesService.getScriptProperties().setProperty('startIndex', 0);
  const triggers = ScriptApp.getProjectTriggers();

  for (let t = 0; t < triggers.length; t++) {
    const trigger = triggers[t];

    if (trigger.getHandlerFunction() === 'continueClientFolderUpdate') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Removed trigger for ${trigger.getHandlerFunction()}`);
    }
  }
}

function continueClientFolderUpdate() {
  const startIndex = PropertiesService.getScriptProperties().getProperty('startIndex');
  const isClientUpdateRunning = isFunctionRunning('continueClientFolderUpdate');
  Logger.log(`isClientUpdateRunning ${isClientUpdateRunning}`);

  while (!isClientUpdateRunning) {
    updateClientFolders();
  }

  // if (startIndex === 2) {
  //   PropertiesService.getScriptProperties().setProperty('startIndex', 0);

  //   const triggers = ScriptApp.getProjectTriggers();

  //   for (let t = 0; t < triggers.length; t++) {
  //     const trigger = triggers[t];

  //     if (trigger.getHandlerFunction() === 'continueClientFolderUpdate') {
  //       ScriptApp.deleteTrigger(trigger);
  //       Logger.log(`Removed trigger for ${trigger.getHandlerFunction()}`);
  //     }
  //   }
  // }
}

function getStudentFolderCount(studentsFolderId) {
  const studentFolder = DriveApp.getFolderById(studentsFolderId);
  const studentSubfolders = studentFolder.getFolders();
  let studentFolderCount = 0;

  while (studentSubfolders.hasNext()) {
    studentSubfolders.next();
    studentFolderCount++;
  }

  return studentFolderCount;
}
