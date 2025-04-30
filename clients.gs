function newClient(clientTemplateFolderId, clientParentFolderId) {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Tutor or Business name:', ui.ButtonSet.OK_CANCEL);
  let customStyles, clientName;

  if (prompt.getSelectedButton() == ui.Button.CANCEL) {
    return;
  } else {
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
  addClientData(clientFolderId);
  linkClientSheets(clientFolderId);
  setClientDataUrls(clientFolderId);

  if (useCustomStyle === ui.Button.YES) {
    getAnswerSheets(clientFolder);
    processFolders(clientFolder.getFolders(), getAnswerSheets);
    styleClientSheets(satSheetIds, actSheetIds, customStyles);
  }

  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + clientFolderId + '" target="_blank" onclick="google.script.host.close()">' + clientName + "'s folder</a>" +
      '<p><a href="https://docs.google.com/spreadsheets/d/' + PropertiesService.getScriptProperties().getProperty('clientDataSsId') + '">Client data IDs</a></p>')
    .setWidth(250) //optional
    .setHeight(50); //optional
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
      } else {
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
        satSheetIds.student = file.getId();
        DriveApp.getFileById(satSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      } else if (filename.includes('answer analysis')) {
        satSheetIds.admin = file.getId();
      }
    }

    if (filename.includes('ACT')) {
      if (filename.toLowerCase().includes('student answer sheet')) {
        actSheetIds.student = file.getId();
        DriveApp.getFileById(actSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      } else if (filename.toLowerCase().includes('answer analysis')) {
        actSheetIds.admin = file.getId();
      }
    }
  }

  while (subFolders.hasNext()) {
    if ((satSheetIds.student && satSheetIds.admin && testType === 'act') || (actSheetIds.student && actSheetIds.admin && testType === 'sat') || (satSheetIds.student && satSheetIds.admin && actSheetIds.student && actSheetIds.admin && testType === 'all')) {
      break;
    }
    const subFolder = subFolders.next();
    linkClientSheets(subFolder.getId(), testType);
  }

  if (satSheetIds.student && satSheetIds.admin) {
    let satAdminSheet = SpreadsheetApp.openById(satSheetIds.admin);
    satAdminSheet.getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
  }

  if (actSheetIds.student && actSheetIds.admin) {
    SpreadsheetApp.openById(actSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(actSheetIds.student);
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
      satSheetIds.adminData = fileId;
      satSheetDataUrls.admin = '"https://docs.google.com/spreadsheets/d/' + satSheetIds.adminData + '/edit?usp=sharing"';
    } else if (filename.includes('sat student data')) {
      Logger.log('found sat student data');
      satSheetIds.studentData = fileId;
      satSheetDataUrls.student = '"https://docs.google.com/spreadsheets/d/' + satSheetIds.studentData + '/edit?usp=sharing"';
      DriveApp.getFileById(satSheetIds.studentData).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } else if (filename.includes('sat student answer sheet')) {
      Logger.log('found sat student answer sheet');
      satSheetIds.student = fileId;
      DriveApp.getFileById(satSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } else if (filename.includes('sat admin answer analysis')) {
      Logger.log('found sat admin answer sheet');
      satSheetIds.admin = fileId;
    } else if (filename.includes('rev sheet data')) {
      Logger.log('found rev sheet data');
      satSheetIds.rev = fileId;
      DriveApp.getFileById(satSheetIds.rev).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } else if (filename.includes('act admin data')) {
      Logger.log('found act admin data');
      actSheetIds.adminData = fileId;
      actSheetDataUrls.admin = '"https://docs.google.com/spreadsheets/d/' + actSheetIds.adminData + '/edit?usp=sharing"';
    } else if (filename.includes('act student data')) {
      Logger.log('found act student data');
      actSheetIds.studentData = fileId;
      actSheetDataUrls.student = '"https://docs.google.com/spreadsheets/d/' + actSheetIds.studentData + '/edit?usp=sharing"';
      DriveApp.getFileById(actSheetIds.studentData).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } else if (filename.includes('act student answer sheet')) {
      Logger.log('found act student answer sheet');
      actSheetIds.student = fileId;
      DriveApp.getFileById(actSheetIds.student).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    } else if (filename.includes('act admin answer analysis')) {
      Logger.log('found act admin answer sheet');
      actSheetIds.admin = fileId;
    }
  }

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    setClientDataUrls(subFolder.getId());
  }

  // if (!isSet.satStudentToAdmin && satSheetIds.admin && satSheetIds.student) {
  if (satSheetIds.admin && satSheetIds.student) {
    SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Student responses').getRange('B1').setValue(satSheetIds.student);
    // isSet.satStudentToAdmin = true;
  }

  // if (!isSet.satStudentToData && satSheetIds.student && satSheetDataUrls.student) {
  if (satSheetIds.student && satSheetDataUrls.student) {
    SpreadsheetApp.openById(satSheetIds.student)
      .getSheetByName('Question bank data')
      .getRange('U7')
      .setValue(satSheetDataUrls.student);


    // isSet.satStudentToData = true;
  }
  if (satSheetIds.admin && satSheetDataUrls.admin) {
    SpreadsheetApp.openById(satSheetIds.admin)
      .getSheetByName('Rev sheet backend')
      .getRange('U5')
      .setValue(satSheetIds.adminData);
  }

  if (satSheetDataUrls.admin && satSheetIds.studentData) {
    SpreadsheetApp.openById(satSheetIds.studentData)
      .getSheetByName('Question bank data')
      .getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.admin + ', "Question bank data!A1:G10000")');
    SpreadsheetApp.openById(satSheetIds.studentData)
      .getSheetByName('Practice test data')
      .getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.admin + ', "Practice test data!A1:E10000")');
  }

  if (satSheetIds.admin && satSheetIds.rev) {
    let adminRevSheet = SpreadsheetApp.openById(satSheetIds.admin).getSheetByName('Rev sheet backend');
    adminRevSheet.getRange('U3').setValue(satSheetIds.rev);
    DriveApp.getFileById(satSheetIds.rev).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  }

  if (satSheetIds.student && satSheetIds.rev) {
    let studentSheet = SpreadsheetApp.openById(satSheetIds.student).getSheetByName('Question bank data');
    studentSheet.getRange('U3').setValue(satSheetIds.rev);
  }

  if (actSheetIds.student && actSheetDataUrls.student) {
    SpreadsheetApp.openById(actSheetIds.student)
      .getSheetByName('Data')
      .getRange('A1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.student + ', "Data!A1:D10000")');
  }

  if (actSheetIds.admin && actSheetDataUrls.admin) {
    var ss = SpreadsheetApp.openById(actSheetIds.admin);
    ss.getSheetByName('Data')
      .getRange('A1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.admin + ', "Data!A1:G10000")');
    ss.getSheets()[0]
      .getRange('J1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.admin + ', "Data!Q1")');
    ss.getSheets()[0].getRange('G1:I1').mergeAcross().setValue('=iferror(J1,"Click to connect data >>")');
  }

  if (actSheetDataUrls.admin && actSheetIds.studentData) {
    SpreadsheetApp.openById(actSheetIds.studentData)
      .getSheetByName('Data')
      .getRange('A1')
      .setValue('=IMPORTRANGE(' + actSheetDataUrls.admin + ', "Data!A1:D10000")');
  }

  Logger.log('setClientDataUrls complete');
}

function updateClientsList(parentClientFolderId='130wX98bJM4wW6aE6J-e6VffDNwqvgeNS') {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const parentClientFolder = DriveApp.getFolderById(parentClientFolderId);
  const clientFolders = parentClientFolder.getFolders();
  const clientSheet = clientDataSs.getSheetByName('Clients');
  let newRow = getLastFilledRow(clientSheet, 1) + 1;
  
  while (clientFolders.hasNext()) {
    const clientFolder = clientFolders.next();
    const clientFolderId = clientFolder.getId();
    const clientFolderName = clientFolder.getName();

    if (!(clientFolderName.includes('_')  || clientFolderName.includes('Ξ'))) {
      addClientData(clientFolderId, newRow);
      newRow++;
    }
  }
}

function addClientData(clientFolderId='1Fd99S1DPdWuvr1VxkeEbdAZn_ZmP9PPj', newRow=null) {
  const clientFolder = DriveApp.getFolderById(clientFolderId);
  const clientName = clientFolder.getName();
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheetByName('Clients');
  const collaborators = clientFolder.getEditors();
  const emailList = []
  collaborators.forEach(c => {
    if (!PropertiesService.getScriptProperties().getProperty('myEmails').includes(c.getEmail())) {
      Logger.log(c.getEmail());
      // emailList.push(c.getEmail());
    }
  });
  
  if(!newRow) {
    newRow = getLastFilledRow(clientSheet, 1) + 1;
  }
  
  const savedClientFolderIds = clientSheet.getRange(2, 4, newRow).getValues();
  let clientIndex = savedClientFolderIds.findIndex(subArray => subArray.includes(clientFolderId));
  let studentFolder, studentFolderId, studentFolderCount;
    
  if (clientIndex === -1) {
    clientIndex = newRow - 2;
    Logger.log(clientIndex + '. ' + clientName);

    getAnswerSheets(clientFolder);
    processFolders(clientFolder.getFolders(), getAnswerSheets);

    const clientSubfolders = clientFolder.getFolders();
    let dataFolderId, satAdminDataId, satStudentDataId, actAdminDataId, actStudentDataId, revDataId = '';

    while (clientSubfolders.hasNext()) {
      const clientSubfolder = clientSubfolders.next();
      studentFolderCount = 0;

      if (clientSubfolder.getName().toLowerCase().includes('students')) {
        studentFolder = clientSubfolder;
        studentFolderId = studentFolder.getId();

        studentFolderCount = getStudentFolderCount(studentFolderId);
      }
      else if (clientSubfolder.getName().toLowerCase().includes('data')) {
        dataFolder = clientSubfolder;
        dataFolderId = clientSubfolder.getId();
        const dataFiles = dataFolder.getFiles()
        while(dataFiles.hasNext()) {
          const file = dataFiles.next();
          const filenameLower = file.getName().toLowerCase();
          if (filenameLower.includes('sat admin')) {
            satAdminDataId = file.getId();
          }
          else if (filenameLower.includes('sat student')) {
            satStudentDataId = file.getId();
          }
          else if (filenameLower.includes('act admin')) {
            actAdminDataId = file.getId();
          }
          else if (filenameLower.includes('act student')) {
            actStudentDataId = file.getId();
          }
          else if (filenameLower.includes('rev sheet data')) {
            revDataId = file.getId();
          }
        }
      }
    }

    clientDataSs.getSheetByName('Clients').getRange(newRow, 1, 1, 16).setValues([[clientIndex, clientName, emailList, clientFolder.getId(), satSheetIds.admin, satSheetIds.student, actSheetIds.admin, actSheetIds.student, dataFolderId, satAdminDataId, satStudentDataId, actAdminDataId, actStudentDataId, revDataId, studentFolderId, studentFolderCount]]);
    newRow ++;
  }
  else {
    studentFolderId = clientSheet.getRange(clientIndex + 2, 15).getValue();
    studentFolderCount = getStudentFolderCount(studentFolderId);
    // clientSheet.getRange(clientIndex + 2, 16).setValue(studentFolderCount);
    // Logger.log(clientFolder.getName() + ' present with ' + studentFolderCount + ' student folders');
  }

  const sheetStudentCount = clientSheet.getRange(clientIndex + 2, 16).getValue();
  Logger.log(sheetStudentCount - studentFolderCount);
  if (sheetStudentCount !== studentFolderCount) {
    Logger.log(clientName + ' has ' + studentFolderCount + ' folders and ' + sheetStudentCount + ' SAT admin sheets');
  }
}

function updateClientFolders() {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheetByName('Clients');
  const lastFilledRow = getLastFilledRow(clientSheet, 2);
  const clientDataRange = clientSheet.getDataRange().getValues();
  const clients = [];

  for (let row = 1; row < lastFilledRow; row++) {       // starts at 1
    clients.push({
      'index': clientDataRange[row][0],
      'name': clientDataRange[row][1],
      'emails': clientDataRange[row][2],
      'folderId': clientDataRange[row][3],
      'satAdminSsId': clientDataRange[row][4],
      'satStudentSsId': clientDataRange[row][5],
      'actAdminSsId': clientDataRange[row][6],
      'actStudentSsId': clientDataRange[row][7],
      'dataFolderId': clientDataRange[row][8],
      'satAdminDataId': clientDataRange[row][9],
      'satStudentDataId': clientDataRange[row][10],
      'actAdminDataId': clientDataRange[row][11],
      'actStudentDataId': clientDataRange[row][12],
      'revDataId': clientDataRange[row][13],
      'studentsFolderId': clientDataRange[row][14],
    })
  }

  for (let client of clients) {
    const startIndex = PropertiesService.getScriptProperties().getProperty('startIndex');

    if (client.index >= startIndex /* 0 is OPT folder */ ) {  
      const clientRow = client.index + 2;
      const studentsDataCell = clientSheet.getRange(clientRow, 17);
      const students = updateStudentFolderData(client, clientSheet);

      const qbResArrayVal = 
        '=let(testCodes,\'Practice test data\'!$E$2:E, testResponses,\'Practice test data\'!$K$2:$K,\n' +
        '    worksheetRanges,vstack(\'Reading & Writing\'!A10:C,\'Reading & Writing\'!E10:G,\'Reading & Writing\'!I10:K,\n' +
        '                           Math!A13:C,Math!E13:G,Math!I13:K,\'SLT Uniques\'!A5:C,\'SLT Uniques\'!E5:G),\n' +
        '    z,counta(A2:A),\n' +
        '    map(offset(G1,1,0,z),offset(B1,1,0,z),offset(E1,1,0,z),offset(A1,1,0,z),\n' +
        '    lambda(    skillCode,       subject,         difficulty,      id,\n' +
        '           if(or(left(skillCode,3)="SAT",left(skillCode,4)="PSAT"),\n' +
        '           xlookup(skillCode,testCodes,testResponses,"not found"),\n' +
        '           vlookup(id,worksheetRanges,3,FALSE)))))'

      const sltFilterR1C1 = "=FILTER({'Question bank data'!R2C1:C1,'Question bank data'!R2C7:C7},left('Question bank data'!R2C7:C7,3)=\"SLT\",'Question bank data'!R2C2:C2=R[-3]C[1])"

      
      // iterate through all folders in Students including template folder
      for (student of students) {
        if(student.satAdminSsId) {
          for (ssId of [student.satAdminSsId, student.satStudentSsId]) {
            const ss = SpreadsheetApp.openById(ssId);

            for (sheetName of ['Reading & Writing', 'Math']) {
              const sheet = ss.getSheetByName(sheetName);

              sheet.getRange('A10:A').setFontColor('#ffffff');
              sheet.getRange('E10:E').setFontColor('#ffffff');
              sheet.getRange('I10:I').setFontColor('#ffffff');
            }
          }
        }
        
        // if (student.satAdminSsId && !student.updateComplete) {
        //   Logger.log('Starting student: ' + student['name'] + ' ' + student['satAdminSsId'] + ' ' + student['satStudentSsId']);
        //   const adminSs = SpreadsheetApp.openById(student['satAdminSsId']);
        //   adminSs.getSheetByName('Question bank data').getRange('I2').setValue(qbResArrayVal);

        //   const revBackendSheet = adminSs.getSheetByName('Rev sheet backend');
        //   if (revBackendSheet) {
        //     revBackendSheet.getRange('U5').setValue(client['satAdminDataId']);
        //   }
        //   Logger.log('Admin values updated')
        //   adminSs.getSheetByName('SLT Uniques').getRange('B5').setValue('');
        //   adminSs.getSheetByName('SLT Uniques').getRange('F5').setValue('');
        //   adminSs.getSheetByName('SLT Uniques').getRange('A5').setValue(sltFilterR1C1);
        //   adminSs.getSheetByName('SLT Uniques').getRange('E5').setValue(sltFilterR1C1);
        //   Logger.log('SLT Uniques filter fixed')

        //   const studentSs = SpreadsheetApp.openById(student['satStudentSsId']);
        //   const studentRevSheet = studentSs.getSheetByName('Rev sheets');
        //   if (studentRevSheet) {
        //     studentRevSheet.getRange('B5:B').setFontWeight('bold');
        //     studentRevSheet.getRange('F5:F').setFontWeight('bold');
        //   }
        //   modifyTestFormatRules(student['satAdminSsId']);
        //   createStudentFolders.updateConceptData(student['satAdminSsId'], student['satStudentSsId']);
        //   student.updateComplete = true;
        //   studentsDataCell.setValue(JSON.stringify(students));
        // }
        // else if (student.updateComplete) {
        //   Logger.log(`${student.name} data is updated`);
        // }
        // else if (!student.satAdminSsId) {
        //   student.updateComplete = true;
        //   studentsDataCell.setValue(JSON.stringify(students));
        //   Logger.log(`No SAT data found for ${student.name}`)
        // }

      }

      // updateSatDataSheets(client['satAdminDataId'], client['satStudentDataId']);
      
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
  Logger.log(`isClientUpdateRunning ${isClientUpdateRunning}`)
  
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

function updateStudentFolderData(
  client={
    'index': null,
    'name': null,
    'studentsFolderId': null,
  },
  dataSheet)
  {
  const clientRow = client.index + 2;
  Logger.log(client.index + '. ' + client.name + ' started');

  const studentFolders = DriveApp.getFolderById(client.studentsFolderId).getFolders();
  const studentFolderIds = [];
  const studentsDataCell = dataSheet.getRange(clientRow, 17);
  const studentsValue = studentsDataCell.getValue();
  let students = JSON.parse(studentsValue);

  while (studentFolders.hasNext()) {
    const studentFolder = studentFolders.next();
    const studentFolderId = studentFolder.getId();
    const studentFolderName = studentFolder.getName();

    studentFolderIds.push(studentFolderId);
    
    if (!studentFolderName.includes('Ξ')) {  
      const studentObj = students.find(obj => obj.folderId === studentFolderId);
      if (studentObj) {
        Logger.log(`${studentFolderName} found with folder ID ${studentFolderId}`);

        if (studentObj && studentObj.name !== studentFolderName) {
          // Update the name property
          studentObj.name = studentFolderName;
          Logger.log(`Updated name for folder ID ${studentFolderId} to ${studentFolderName}`);
        }
      }
      else {
        Logger.log(`Adding ${studentFolderName} to students data`);
        const adminFiles = studentFolder.getFiles();
        let satAdminSsId, satStudentSsId;

        while (adminFiles.hasNext()) {
          const adminFile = adminFiles.next();
          if (adminFile.getName().toLowerCase().includes('sat admin')) {
            satAdminSsId = adminFile.getId();
            break;
          }
        }

        if (satAdminSsId) {
          satStudentSsId = SpreadsheetApp.openById(satAdminSsId).getSheetByName('Student responses').getRange('B1').getValue();
        }

        students.push({
          'name': studentFolderName,
          'folderId': studentFolderId,
          'satAdminSsId': satAdminSsId,
          'satStudentSsId': satStudentSsId,
          'updateComplete': false
        })
      }
    }

    // only for clients with grouped student folders
    // const subfolders = studentFolder.getFolders();
    // while (subfolders.hasNext()) {
    //   const subfolder = subfolders.next();
    //   const subfiles = subfolder.getFiles();

    //   while (subfiles.hasNext()) {
    //     const subfile = subfiles.next();
    //     if (subfile.getName().toLowerCase().includes('sat admin')) {
    //       satAdminSsId = subfile.getId();
    //       break;
    //     }
    //   }

    //   if (satAdminSsId) {
    //     satStudentSsId = SpreadsheetApp.openById(satAdminSsId).getSheetByName('Student responses').getRange('B1').getValue();
    //   }

    //   students.push({
    //     'name': subfolder.getName(),
    //     'satAdminSsId': satAdminSsId,
    //     'satStudentSsId': satStudentSsId
    //   })
    // }
  }

  students = students.filter(student => studentFolderIds.includes(student.folderId));

  const studentFolderCount = students.length
  dataSheet.getRange(clientRow, 17).setValue(JSON.stringify(students));
  dataSheet.getRange(clientRow, 16).setValue(studentFolderCount);

  return students;
}

function getStudentFolderCount(studentsFolderId) {
  const studentFolder = DriveApp.getFolderById(studentsFolderId);
  const studentSubfolders = studentFolder.getFolders();
  let studentFolderCount = 0;

  while (studentSubfolders.hasNext()) {
    studentSubfolders.next();
    studentFolderCount ++;
  }

  return studentFolderCount;
}

function updateSatDataSheets(satAdminDataSsId, satStudentDataSsId) {
  const satAdminDataSs = SpreadsheetApp.openById(satAdminDataSsId);
  const satStudentDataSs = SpreadsheetApp.openById(satStudentDataSsId);
  let newAdminQbSheet = satAdminDataSs.getSheetByName('Question bank data updated ' + dataLatestDate);
  let newAdminPtSheet = satAdminDataSs.getSheetByName('Practice test data updated ' + dataLatestDate);
  let newStudentQbSheet = satStudentDataSs.getSheetByName('Question bank data updated ' + dataLatestDate);
  let newStudentPtSheet = satStudentDataSs.getSheetByName('Practice test data updated ' + dataLatestDate);

  if (!newAdminQbSheet) {
    newAdminQbSheet = satAdminDataSs.getSheetByName('Question bank data').copyTo(satAdminDataSs).setName('Question bank data updated ' + dataLatestDate);
  }
  if (!newAdminPtSheet) {
    newAdminPtSheet = satAdminDataSs.getSheetByName('Practice test data').copyTo(satAdminDataSs).setName('Practice test data updated ' + dataLatestDate);
  }
  if (!newStudentQbSheet) {
    newStudentQbSheet = satStudentDataSs.getSheetByName('Question bank data').copyTo(satStudentDataSs).setName('Question bank data updated ' + dataLatestDate);
  }
  if (!newStudentPtSheet) {
    newStudentPtSheet = satStudentDataSs.getSheetByName('Practice test data').copyTo(satStudentDataSs).setName('Practice test data updated ' + dataLatestDate);
  }

  newAdminQbSheet.getRange('A1').setValue('=importrange("1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE", "Question bank data updated ' + dataLatestDate + '!A1:H10000")');
  newAdminPtSheet.getRange('A1').setValue('=importrange("1XoANqHEGfOCdO1QBVnbA3GH-z7-_FMYwoy7Ft4ojulE", "Practice test data updated ' + dataLatestDate + '!A1:J10000")');
  Logger.log('sat admin data sheets updated')
  newStudentQbSheet.getRange('A1').setValue('=importrange("' + satAdminDataSsId + '", "Question bank data updated ' + dataLatestDate + '!A1:G10000")');
  newStudentPtSheet.getRange('A1').setValue('=importrange("' + satAdminDataSsId + '", "Practice test data updated ' + dataLatestDate + '!A1:E10000")');
  Logger.log('sat student data sheets updated')
}