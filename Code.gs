const satSheetIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
  rev: null,
};

const satSheetDataUrls = {
  admin: null,
  student: null,
  rev: null,
};

const actSheetIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};

const actSheetDataUrls = {
  admin: null,
  student: null,
};

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
    getStyledSheets(clientFolder);
    processFolders(clientFolder.getFolders(), getStyledSheets);
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

function styleClientFolder(clientFolder, customStyles = {}) {
  let clientFolderId;

  if (clientFolder) {
    clientFolderId = clientFolder.getId();
  } else {
    const ui = SpreadsheetApp.getUi();
    const prompt = ui.prompt('Client folder URL or ID', ui.ButtonSet.OK_CANCEL);
    clientFolderId = prompt.getResponseText();

    if (prompt.getSelectedButton() == ui.Button.CANCEL) {
      return;
    } else if (prompt.getResponseText().includes('/folders/')) {
      clientFolderId = prompt.getResponseText().split('/folders/')[1].split(/[/?]/)[0];
      Logger.log(clientFolderId);
      clientFolder = DriveApp.getFolderById(clientFolderId);
    } else {
      clientFolderId = prompt.getResponseText();
      clientFolder = DriveApp.getFolderById(clientFolderId);
    }
  }

  if (Object.keys(customStyles).length === 0) {
    customStyles = setCustomStyles();
  }

  Logger.log('Styling sheets for ' + clientFolder.getName());
  getStyledSheets(clientFolder);
  processFolders(clientFolder.getFolders(), getStyledSheets);
  styleClientSheets(satSheetIds, actSheetIds, customStyles);

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + clientFolderId + '" target="_blank" onclick="google.script.host.close()">Client folder</a>')
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Styling complete');
}

function getStyledSheets(folder) {
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

function processFolders(folders, folderFunction) {
  while (folders.hasNext()) {
    const folder = folders.next();
    folderFunction(folder);
    processFolders(folder.getFolders(), folderFunction);
  }
}

function styleClientSheets(satSheetIds, actSheetIds, customStyles) {
  for (let id of [satSheetIds.admin, satSheetIds.student, actSheetIds.admin, actSheetIds.student]) {
    const ss = SpreadsheetApp.openById(id);
    const ssName = ss.getName();
    const satTestSheets = getTestCodes();
    const satDataSheets = ['question bank data', 'practice test data', 'rev sheet backend'];
    const actDataSheets = ['data', 'scoring'];

    const primaryColor = customStyles.primaryColor;
    const primaryContrastColor = customStyles.primaryContrastColor;
    const secondaryColor = customStyles.secondaryColor;
    const secondaryContrastColor = customStyles.secondaryContrastColor;
    const fontColor = customStyles.fontColor;
    const imgUrl = customStyles.img;

    if (ssName.includes('ACT')) {
      for (let j in ss.getSheets()) {
        const sh = ss.getSheets()[j];
        const shRange = sh.getDataRange();
        shRange.setBackground('white');
        shRange.setFontColor(fontColor);

        let shName = sh.getName().toLowerCase();

        if (shName.endsWith('z')) {
          shName = shName.substring(0, shName.length - 1);
        }

        const isTestSheet = /^\d+$/.test(shName);

        if (isTestSheet) {
          sh.getRange('A1:P4').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);

          sh.getRangeList(['B3', 'F3', 'J3', 'N3']).setBorder(true, true, true, true, true, true, '#93c47d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

          sh.getRange('F1').setBackground('#93c47d');
        } else if (shName === 'test analysis' || shName === 'opportunity areas') {
          sh.getRange(1, 1, 8, sh.getMaxColumns())
            .setBackground(primaryColor)
            .setFontColor(primaryContrastColor)
            .setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID)
            .setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

          if (shName === 'test analysis') {
            var correctRange = 'F6:J6';
          } else {
            var correctRange = 'E6:I6';
          }
          sh.getRange(correctRange).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

          const imgCell = sh.getRange('B3');
          if (imgUrl) {
            imgCell.setValue('=image("' + customStyles.img + '")');
          }

          applyConditionalFormatting(sh, customStyles);
        } else if (actDataSheets.includes(shName)) {
          sh.getRange(1, 1, 1, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor);
        } else if (shName === 'student responses') {
          sh.getRange(1, 1, 3, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
        }
      }
    } else if (ssName.includes('SAT')) {
      for (let j in ss.getSheets()) {
        const sh = ss.getSheets()[j];
        const shRange = sh.getDataRange();
        const shName = sh.getName();
        const shNameLower = shName.toLowerCase();

        if (!shNameLower.includes('rev sheet')) {
          shRange.setFontColor(fontColor);
        }

        // practice SAT answer sheets
        if (satTestSheets.includes(shName)) {
          shRange.setBackground('white');
          if (customStyles.sameHeaderColor) {
            sh.getRangeList(['B2:L4', 'B33:L35']).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
          } else {
            sh.getRangeList(['B2:L4', 'B33:L35']).setBackground(secondaryColor).setFontColor(secondaryContrastColor).setBorder(true, true, true, true, true, true, secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
          }
          sh.getRangeList(['A1:A', 'E5:E', 'I5:I']).setFontColor('white');
        }
        // check for SAT analysis sheets after checking exact match
        else if (shNameLower.includes('analysis') || shNameLower.includes('opportunity')) {
          if (shNameLower === 'rev analysis' || shNameLower === 'opportunity areas') {
            sh.getRange('A1:K7').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          } else if (shNameLower === 'time series analysis') {
            sh.getRange('A1:K6').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
            sh.getRange('D5:E6').setFontColor(fontColor);
          } else {
            sh.getRange('A1:K6').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
            sh.getRange('A7:A8').setFontColor('white');
            applyConditionalFormatting(sh, customStyles);
          }

          const imgCell = sh.getRange('B2');
          if (imgUrl) {
            imgCell.setValue('=image("' + customStyles.img + '")');
          }
        } else if (shNameLower === 'reading & writing') {
          styleSatWorksheets(sh, 6, 11, customStyles);
        } else if (shNameLower === 'math') {
          styleSatWorksheets(sh, 9, 11, customStyles);
        } else if (shNameLower === 'slt uniques') {
          styleSatWorksheets(sh, 1, 7, customStyles);
        } else if (satDataSheets.includes(shNameLower)) {
          sh.getRange(1, 1, 1, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor);
        } else if (shNameLower === 'student responses') {
          sh.getRange(1, 1, 3, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
        } else if (shNameLower === 'rev sheets') {
          let revSheetHeaderRange;
          if (ssName.toLowerCase().includes('admin answer analysis')) {
            revSheetHeaderRange = sh.getRangeList(['B2:E4', 'G2:J4']);
          } else {
            revSheetHeaderRange = sh.getRangeList(['B2:D4', 'F2:I4']);
          }

          if (customStyles.sameHeaderColor) {
            revSheetHeaderRange.setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
          } else {
            revSheetHeaderRange.setBackground(secondaryColor).setFontColor(secondaryContrastColor).setBorder(true, true, true, true, true, true, secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
          }
        }
      }
    }
  }

  Logger.log('styleClientSheets complete');
}

function styleSatWorksheets(
  sh = SpreadsheetApp.openById('1FW_3GIWmytdrgBdfSuIl2exy9hIAnQoG8IprF8k9uEY').getSheetByName('Math'),
  rowOffset = 10,
  headerCols = 11,
  customStyles = {
    primaryColor: '#134f5c',
    primaryContrastColor: 'white',
  }) {

  const cats = createStudentFolders.cats
  cats.push('Reading & Writing'); // styles SLT Uniques header
  var conceptRows = [];

  sh.getRange(1, 1, sh.getMaxRows()).setFontColor('white');
  sh.getRange(1, 5, sh.getMaxRows()).setFontColor('white');
  sh.getRange(1, 9, sh.getMaxRows()).setFontColor('white');

  const colVals = sh.getRange(rowOffset, 2, sh.getMaxRows() - rowOffset).getValues();

  for (let x = 0; x < colVals.length; x++) {
    if (cats.includes(colVals[x][0])) {
      var row = x + rowOffset;
      conceptRows.push(row);
    }
  }
  for (r in conceptRows) {
    const highlightRange = sh.getRange(conceptRows[r], 2, 3, headerCols);

    if (customStyles.sameHeaderColor) {
      highlightRange.setBackground(customStyles.primaryColor).setFontColor(customStyles.primaryContrastColor).setBorder(true, true, true, true, true, true, customStyles.primaryColor, SpreadsheetApp.BorderStyle.SOLID);
    } else {
      highlightRange.setBackground(customStyles.secondaryColor).setFontColor(customStyles.secondaryContrastColor).setBorder(true, true, true, true, true, true, customStyles.secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}

function applyConditionalFormatting(sheet=SpreadsheetApp.openById('1nwG8o6Rd0ArGQMzrfUjRP16FkSw9urEIK-V7UD2ayJM'), customStyles=null) {
  var rules = [];
  var newRules = [];

  if (!customStyles) {
    customStyles = setCustomStyles();
  }

  for (i in sheet.getConditionalFormatRules()) {
    var condition = sheet.getConditionalFormatRules()[i];
    rules.push(condition);
  }

  for (i in rules) {
    if (rules[i].getGradientCondition()) {
      Logger.log(rules[i].getGradientCondition());
      newRule = rules[i].copy();
      newRules.push(newRule);
    }
  }

  if (sheet.getName().toLowerCase().includes('opportunity')) {
    var subtotalStart = 'B';
    var domainStart = 'C';
  } else {
    var subtotalStart = 'C';
    var domainStart = 'D';
  }
  var grandTotalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B7="Grand total"')
    .setBold(true)
    .setBackground(customStyles.primaryColor)
    .setFontColor(customStyles.primaryContrastColor)
    .setRanges([sheet.getRange('B7:I177')]);

  var subTotalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=right($' + subtotalStart + '7,5)="Total"')
    .setBold(true)
    .setBackground(customStyles.secondaryColor)
    .setFontColor(customStyles.secondaryContrastColor)
    .setRanges([sheet.getRange(subtotalStart + '7:I177')]);

  var domainTotalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=right($' + domainStart + '7,5)="Total"')
    .setBackground(customStyles.tertiaryColor)
    .setFontColor(customStyles.tertiaryContrastColor)
    .setRanges([sheet.getRange(domainStart + '7:I177')]);

  var backgroundColorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=sum($F7:$I7)>0')
    .setBackground('#f5f7f9')
    .setRanges([sheet.getRange('B7:I177')]);

  newRules.push(grandTotalRule, subTotalRule, domainTotalRule, backgroundColorRule);
  sheet.clearConditionalFormatRules();
  sheet.setConditionalFormatRules(newRules);

  Logger.log('applyConditionalFormatting complete for ' + sheet.getName());
}

function isDark(hex = '#b6d7a8') {
  hex = hex.substring(1); // strip #
  var rgb = parseInt(hex, 16); // convert rrggbb to decimal
  var r = (rgb >> 16) & 0xff; // extract red
  var g = (rgb >> 8) & 0xff; // extract green
  var b = (rgb >> 0) & 0xff; // extract blue

  var luma = 0.2126 * r + 0.7152 * g + 0.0722 * b; // per ITU-R BT.709

  Logger.log(luma);
  if (luma < 205) {
    return true;
  } else {
    return false;
  }
}

function setCustomStyles() {
  let ui = SpreadsheetApp.getUi();
  let primaryColor = ui.prompt('Primary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
  let secondaryColor = ui.prompt('Secondary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
  let tertiaryColor = ui.prompt('Tertiary background color', ui.ButtonSet.OK_CANCEL).getResponseText();
  let fontColor = ui.prompt('Font color (leave blank to use primary color)', ui.ButtonSet.OK_CANCEL).getResponseText();
  let imgFilename = ui.prompt('Image URL or filename', ui.ButtonSet.OK_CANCEL).getResponseText();
  let sameHeaderColor = ui.alert('Same color header?', ui.ButtonSet.YES_NO);

  if (primaryColor === '') {
    primaryColor = '#1c4d65';
  }

  if (secondaryColor === '') {
    secondaryColor = '#ffa874';
  }

  if (tertiaryColor === '') {
    tertiaryColor = '#c4f0f7';
  }

  if (fontColor === '') {
    fontColor = primaryColor;
  }

  if (isDark(primaryColor)) {
    var primaryContrastColor = 'white';
  } else if (isDark(fontColor)) {
    primaryContrastColor = fontColor;
  } else {
    primaryContrastColor = 'black';
  }

  if (isDark(secondaryColor)) {
    var secondaryContrastColor = 'white';
  } else if (isDark(fontColor)) {
    secondaryContrastColor = fontColor;
  } else {
    secondaryContrastColor = 'black';
  }

  if (isDark(tertiaryColor)) {
    var tertiaryContrastColor = 'white';
  } else if (isDark(fontColor)) {
    tertiaryContrastColor = fontColor;
  } else {
    tertiaryContrastColor = 'black';
  }

  if (sameHeaderColor === ui.Button.YES) {
    sameHeaderColor = true;
  } else {
    sameHeaderColor = false;
  }

  let imgUrl;
  if (imgFilename.toLowerCase().includes('www.') || imgFilename === '') {
    imgUrl = imgFilename;
  }
  else {
    imgUrl = 'https://www.openpathtutoring.com/static/img/orgs/' + imgFilename;
  }

  let customStyles = {
    primaryColor: primaryColor,
    primaryContrastColor: primaryContrastColor,
    secondaryColor: secondaryColor,
    secondaryContrastColor: secondaryContrastColor,
    tertiaryColor: tertiaryColor,
    tertiaryContrastColor: tertiaryContrastColor,
    fontColor: fontColor,
    img: imgUrl,
    sameHeaderColor: sameHeaderColor,
  };

  return customStyles;
}

function getTestCodes() {
  const practiceTestDataSheet = SpreadsheetApp.openById('1KidSURXg5y-dQn_gm1HgzUDzaICfLVYameXpIPacyB0').getSheetByName('Practice test data');
  const lastFilledRow = getLastFilledRow(practiceTestDataSheet, 1);
  const testCodeCol = practiceTestDataSheet
    .getRange(2, 1, lastFilledRow - 1)
    .getValues()
    .map((row) => row[0]);
  const testCodes = testCodeCol.filter((x, i, a) => a.indexOf(x) == i);

  return testCodes;
}

function getLastFilledRow(sheet, col) {
  const maxRow = sheet.getMaxRows();
  const allVals = sheet.getRange(1, col, maxRow).getValues();
  const lastFilledRow = maxRow - allVals.reverse().findIndex((c) => c[0] != '');

  return lastFilledRow;
}

function getAnalysisFiles(folder, n = 3, fileList = []) {
  folder = folder || DriveApp.getRootFolder();
  var folderName = folder.getName().toLowerCase();
  if (!folderName.includes('Ξ') || !folderName.includes('_')) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var filename = file.getName().toLowerCase();
      if (filename.includes('admin answer analysis') && !filename.includes('template')) {
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
  // const testCodes = ['at1', 'at2', 'at3', 'at4', 'at5', 'at6', 'sat1', 'sat2', 'sat3', 'sat4', 'sat5', 'sat6', 'sat7', 'sat8', 'sat9', 'sat10', 'psat1', 'psat2', 'apt1', 'apt2'];
  const testCodes = ['sat1', 'sat2', 'sat3', 'sat4', 'sat5', 'sat6', 'sat7', 'sat8', 'sat9', 'sat10', 'psat1', 'psat2'];
  const scoreSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('Scores');
  const lastRow = scoreSheet.getLastRow();
  const scoresNames = scoreSheet.getRange(1, 1, lastRow);
  const nextOpenRow =
    lastRow -
    scoresNames
      .getValues()
      .reverse()
      .findIndex((c) => c[0] != '') + 1;

  // Loop through analysis spreadsheets
  for (var i = 0; i < fileList.length; i++) {
    const ssId = fileList[i].getId();
    const ss = SpreadsheetApp.openById(ssId);
    const ssName = ss.getName();
    const studentName = ssName.slice(ssName.indexOf('-') + 2);
    const sheets = ss.getSheets();
    const scores = [];

    // Loop through sheets within analysis spreadsheet
    for (var s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const sheetData = sheet.getRange('A1:K57');
      const sheetName = sheet.getName().toLowerCase();
      const rwScore = sheetData[0][6];
      const mScore = sheetData[0][8];

      if (testCodes.includes(sheetName) && rwScore && mScore) {
        const totalScore = rwScore + mScore;
        // Check last answer for each module
        const mod1RWEnd = sheetData[30][2];
        const mod2RWEnd = sheetData[30][6];
        const mod3RWEnd = sheetData[30][10];
        const mod1MathEnd = sheetData[56][2];
        const mod2MathEnd = sheetData[56][6];
        const mod3MathEnd = sheetData[56][10];

        const values = [mod1RWEnd, mod2RWEnd, mod3RWEnd, mod1MathEnd, mod2MathEnd, mod3MathEnd];

        // Filter out blank or null values and count the remaining ones
        const modsComplete = values.filter(function(value) {
          return value !== "" && value !== null;
        }).length;

         
        const completionCheck = sheet.getRange('M1');
        const testName = sheetName.toUpperCase();

        if (sheetName.slice(0, 3) === 'apt') {
          testName = sheetName.replace('apt', 'psat').toUpperCase();
        } else if (sheetName.slice(0, 2) === 'at') {
          testName = sheetName.replace('at', 'sat').toUpperCase();
        }

        // If test is completed, add to scores array
        if (modsComplete === 4 && sheet.getRange('G1').getValue() >= 200 && sheet.getRange('I1').getValue() >= 200) {
          scores.push({
            test: testName,
            rw: rwScore,
            m: mScore,
            total: totalScore,
          });

          // If test is newly completed, create score report
          if (completionCheck.getValue() !== '✔') {
            Logger.log(ssName + ' ' + testName + ' score report started');
            createSatScoreReport(ssId, sheetName, scores);
            SpreadsheetApp.flush();

            completionCheck.setValue('✔');
            Logger.log(ssName + ' ' + testName + ' score report complete');
            completionCheck.setVerticalAlignment('middle');
            completionCheck.setFontColor('#134f5c');

            const dateSubmitted = sheet.getRange('D2').getValue();
            if ((dateSubmitted = '')) {
              dateSubmitted = Utilities.formatDate(new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - 1), 'UTC', 'MM/dd/yyyy');
            }
            const rowData = [[studentName, 'Practice', testName.toUpperCase(), dateSubmitted, totalScore, rwScore, mScore]];
            scoreSheet.getRange(nextOpenRow, 1, 1, 7).setValues(rowData);
            nextOpenRow += 1;
          }
        }
      }
    }
  }
}

function createSatScoreReport(spreadsheetId, testCode, scores) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();

    var sheetsToPrint = [testCode.toLowerCase(), testCode.toLowerCase() + ' analysis'];
    var filename = spreadsheet.getName();
    var studentName = filename.slice(filename.indexOf('-') + 2);
    var analysisIndex = 1;

    showAllExcept(spreadsheetId);
    SpreadsheetApp.flush();

    /* PDF can be created from single sheet or all visible sheets. For a multi-sheet PDF, we need to hide
    unwanted sheets, save the PDF, then show all sheets again. */
    SpreadsheetApp.openById(spreadsheetId)
      .getSheets()
      .forEach((sh) => {
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
              } else if (sh.getRange('H7').getValue() === 'BLANK') {
                sh.showColumns(8);
              }
            }
          } else {
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
  }
  catch (err) {
    Logger.log(err.message + '\n\n' + err.stack);
  }

  showAllExcept(spreadsheetId);
  // Move analysis sheet back to original position
  spreadsheet.moveActiveSheet(analysisIndex);
}

// Save spreadsheet as a PDF: https://gist.github.com/andrewroberts/26d460212874cdd3f645b55993942455
function sendPdfScoreReport(spreadsheetId, email, studentName, scores = []) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();
    //var sheetId = sheetName ? spreadsheet.getSheetByName(sheetName).getSheetId() : null;
    var practiceDataSheet = spreadsheet.getSheetByName('Practice test data');

    if (practiceDataSheet.getRange('V1').getValue() === 'Score report folder ID:' && practiceDataSheet.getRange('W1').getValue() !== '') {
      var scoreReportFolderId = practiceDataSheet.getRange('W1').getValue();
    } else {
      var parentId = DriveApp.getFileById(spreadsheetId).getParents().next().getId();
      const subfolderIds = getSubFolderIdsByFolderId(parentId);

      for (let i in subfolderIds) {
        let subfolderId = subfolderIds[i];
        let subfolder = DriveApp.getFolderById(subfolderId);
        let subfolderName = subfolder.getName();
        if (subfolderName.toLowerCase().includes('score report')) {
          var scoreReportFolderId = subfolder.getId();
        }
      }
    }

    if (!scoreReportFolderId) {
      var scoreReportFolderId = DriveApp.getFolderById(parentId).createFolder('Score reports').getId();
      practiceDataSheet.getRange('V1:W1').setValues([['Score report folder ID:', scoreReportFolderId]]);
    }

    var url_base = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/';
    var url_ext =
      'export?exportFormat=pdf&format=pdf' + //export as pdf
      // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
      //+ (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId))
      '&id=' +
      spreadsheetId +
      // following parameters are optional...
      '&size=letter' + // paper size
      '&portrait=true' + // orientation, false for landscape
      '&fitw=true' + // fit to width, false for actual size
      '&fzr=false' + // do not repeat row headers (frozen rows) on each page
      '&top_margin=0.5' +
      '&bottom_margin=0.5' +
      '&left_margin=0.3' +
      '&right_margin=0.3' +
      '&printnotes=false' +
      '&sheetnames=false' +
      '&printtitle=false' +
      '&pagenumbers=false'; //hide optional headers and footers

    var options = {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
    };

    // Create PDF
    var currentScore = scores.slice(-1)[0];
    var pdfName = 'SAT answer analysis - ' + studentName + ' - ' + currentScore.test;
    var studentFirstName = studentName.split(' ')[0];
    const studentData = getStudentHours(studentName);
    var response = UrlFetchApp.fetch(url_base + url_ext, options);
    var blob = response.getBlob().setName(pdfName + '.pdf');
    var scoreReportFolder = DriveApp.getFolderById(scoreReportFolderId);
    scoreReportFolder.createFile(blob);

    if (studentData.hours) {
      var message =
        'Hi PARENTNAME, please find the score report from ' +
        studentFirstName +
        "'s recent practice test attached. " +
        currentScore.total +
        ' overall (' +
        currentScore.rw +
        ' Reading & Writing, ' +
        currentScore.m +
        ' Math)<br><br>As of the session on ' +
        studentData.recentSessionDate +
        ', we have ' +
        studentHours +
        ' hours remaining on the current package. Let me know if you have any questions. Thanks!<br><br>';

      if (scores.length > 1) {
        message += 'Previous scores - most recent last:<br><ul>';

        for (i = 0; i < scores.length - 1; i++) {
          message += '<li>' + scores[i].test + ': ' + scores[i].total + ' (' + scores[i].rw + ' RW, ' + scores[i].m + ' M)</li>';
        }
        message += '</ul><br>';
      }
    }
    else {
      var message = 'Hi PARENTNAME, please find the score report from ' +
        studentFirstName +
        "'s recent practice test attached. " +
        currentScore.total +
        ' overall (' +
        currentScore.rw +
        ' Reading & Writing, ' +
        currentScore.m +
        ' Math)<br><br>'
    }

    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: currentScore.test + ' score report for ' + studentFirstName,
        htmlBody: message,
        attachments: [blob.getAs(MimeType.PDF)],
      });
    }
  }
  catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack)
  }
}

function getStudentHours(studentName) {
  const summarySheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('Summary');
  const maxRow = summarySheet.getMaxRows();
  const allVals = summarySheet.getRange('A1:A' + maxRow).getValues();
  const lastFilledRow = maxRow - allVals.reverse().findIndex((c) => c[0] != '');
  var summaryData = summarySheet.getRange(1, 1, lastFilledRow, 26).getValues();
  const studentData = {
    'name': null,
    'recentSessionDate': null
  };

  for (let r = 0; r < lastFilledRow; r++) {
    if (summaryData[r][0] === studentName) {
      studentData.name = summaryData[r][3],
      studentData.recentSessionDate = Utilities.formatDate(new Date(summaryData[r][16]), 'GMT', 'EEE M/d');
      break;
    }
  }
  return studentData;
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
  var editors = DriveApp.getFileById(id)
    .getEditors()
    .map(function (e) {
      return e.getEmail();
    });
  var emails = '';

  for (var i = 0; i < editors.length; i++) {
    // Only add openpathtutoring.com emails to email list
    if (editors[i].includes('openpathtutoring.com')) {
      emails += editors[i] + ',';
    }
  }

  return emails;
}

const showAllExcept = (spreadsheetId='1_nRuW80ewwxEcsHLKy8U8o1nIxKNxxrih-IC-T2suJk', hiddenSheets = []) => {
  SpreadsheetApp.openById(spreadsheetId)
    .getSheets()
    .forEach((sh) => {
      // If sheets are meant to be hidden, leave them hidden
      if (!hiddenSheets.includes(sh.getName())) {
        sh.showSheet();
      }
    });
  // SpreadsheetApp.flush();
};

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

    renameStudentFolder(subfolder, currentName, newName);
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
    return vA - vB;
  });
  Logger.log(arr);
}

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

function updateClientsList(parentClientFolderId='130wX98bJM4wW6aE6J-e6VffDNwqvgeNS') {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const parentClientFolder = DriveApp.getFolderById(parentClientFolderId);
  const clientFolders = parentClientFolder.getFolders();
  const clientSheet = clientDataSs.getSheets()[0];
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

function addClientData(clientFolderId, newRow=null) {
  const clientFolder = DriveApp.getFolderById(clientFolderId);
  const clientName = clientFolder.getName();
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheets()[0];
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

    getStyledSheets(clientFolder);
    processFolders(clientFolder.getFolders(), getStyledSheets);

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

    clientDataSs.getSheets()[0].getRange(newRow, 1, 1, 16).setValues([[clientIndex, clientName, emailList, clientFolder.getId(), satSheetIds.admin, satSheetIds.student, actSheetIds.admin, actSheetIds.student, dataFolderId, satAdminDataId, satStudentDataId, actAdminDataId, actStudentDataId, revDataId, studentFolderId, studentFolderCount]]);
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
  const clientSheet = clientDataSs.getSheets()[0];
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
      const students = updateStudentFolderData(client);

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

function updateStudentFolderData(
  client={
    'index': null,
    'name': null,
    'studentsFolderId': null,
  })
{
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheets()[0];
  const clientRow = client.index + 2;
  Logger.log(client.index + '. ' + client.name + ' started');

  const studentFolders = DriveApp.getFolderById(client.studentsFolderId).getFolders();
  const studentsDataCell = clientSheet.getRange(clientRow, 17);
  const studentsValue = studentsDataCell.getValue();
  const students = JSON.parse(studentsValue);

  while (studentFolders.hasNext()) {
    const studentFolder = studentFolders.next();
    const studentFolderId = studentFolder.getId();
    const studentFolderName = studentFolder.getName();
    
    if (!studentFolderName.includes('Ξ')) {  
      const studentDataExists = students.some(obj => obj.folderId===studentFolderId);
      if (studentDataExists) {
        Logger.log(`${studentFolderName} found with folder ID ${studentFolderId}`);
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

  const studentFolderCount = students.length
  clientSheet.getRange(clientRow, 17).setValue(JSON.stringify(students));
  clientSheet.getRange(clientRow, 16).setValue(studentFolderCount);

  return students;
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

  // const isRunning = processes.some(process => process.processStatus === "RUNNING" && process.functionName === functionName);
  return isRunning;
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

function updateStudentFolder(studentFolder) {
  const studentFolderName = studentFolder.getName();
  Logger.log(studentFolderName + ' started');
}

function updateSatDataSheets(satAdminDataSsId, satStudentDataSsId) {

  const dataLatestDate = '03/2025';
  
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

function modifyTestFormatRules(satAnswerSheetId='1FW_3GIWmytdrgBdfSuIl2exy9hIAnQoG8IprF8k9uEY') {
  const ss = SpreadsheetApp.openById(satAnswerSheetId);
  const tests = getTestCodes();

  for (test of tests) {
    const sh = ss.getSheetByName(test);
    if (sh) {
      var rules = sh.getConditionalFormatRules();
      const alertColor = '#cc0000';
      const updatedRules = [];
      
      for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var bgColor = rule.getBooleanCondition().getBackgroundObject().asRgbColor().asHexString();
        
        if (bgColor !== alertColor) {
          updatedRules.push(rule);
        }
      }

      const rwRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange('A5:A31'), sh.getRange('E5:E31'), sh.getRange('I5:I31')])
      .whenFormulaSatisfied('=A5<>$B$2&" "&B5')
      .setBackground(alertColor)
      .setFontColor('#ffffff')
      .build();

      const mathRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange('A36:A57'), sh.getRange('E36:E57'), sh.getRange('I36:I57')])
      .whenFormulaSatisfied('=A36<>$B$33&" "&B36')
      .setBackground(alertColor)
      .setFontColor('#ffffff')
      .build();
    
      updatedRules.push(rwRule, mathRule);
      sh.setConditionalFormatRules(updatedRules);
    }
  }
  Logger.log('Test sheets formatting updated');  
}


// Scheduled recurring functions

function updateOPTStudentFolderData() {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheets()[0];
  const teamDataSheet = clientDataSs.getSheetByName('Team OPT');

  const teamParentFolder = DriveApp.getFolderById('1tSKajFOa_EVUjH8SKhrQFbHSjDmmopP9');
  const teamFolders = teamParentFolder.getFolders();
  const teamData = [];
  const teamIndex = 0;

  const myStudentFolderData = {
    'index': 0,
    'name': 'Open Path Tutoring',
    'studentsFolderId': clientSheet.getRange(2, 15).getValue()
  }

  const myStudents = updateStudentFolderData(myStudentFolderData, '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF');
  clientSheet.getRange(2, 17).setValue(JSON.stringify(myStudents));
  
  while (teamFolders.hasNext()) {
    const teamFolder = teamFolders.next();
    const teamFolderName = teamFolder.getName();
    const teamFolderId = teamFolder.getId();
    teamData.push({
      'index': teamIndex,
      'name': teamFolderName,
      'studentsFolderId': teamFolderId
    })

    teamDataSheet.getRange(teamIndex + 2,1,1,4).setValues([[teamIndex, teamFolderName, teamFolderId]]);
    teamIndex ++;
  }
}


function findNewScoreReports() {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheets()[0];
  const myStudentDataValue = clientSheet.getRange(2, 17).getValue();
  const students = JSON.parse(myStudentDataValue);
  const fileList = [];

  // var parentFolder = DriveApp.getFolderById(myStudentFolderData.folderId);
  // const fileList = getAnalysisFiles(parentFolder, (n = 3));

  for (student of students) {
    if (student.satAdminSsId && !student.name.toLowerCase().includes('template')) {
      fileList.push(DriveApp.getFileById(student.satAdminSsId));
    }
  }

  // Sort by most recently updated first
  fileList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  Logger.log(fileList);

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

function trashFilesInFolder(folderId='15tJsdeOx_HucjIb6koTaafncTj-e6FO6') {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  let fileCount = 0;

  while(files.hasNext()) {
    const file = files.next();
    file.setTrashed(true);
    fileCount += 1;
  }
  Logger.log('Moved ' + fileCount + ' files to trash');
}
