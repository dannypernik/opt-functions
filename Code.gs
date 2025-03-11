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
      .getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.student + ', "Question bank data!A1:G10000")');
    SpreadsheetApp.openById(satSheetIds.student)
      .getSheetByName('Practice test data')
      .getRange('A1')
      .setValue('=IMPORTRANGE(' + satSheetDataUrls.student + ', "Practice test data!A1:E10000")');

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

const cats = [
    'Area and volume',
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
  ];

function styleSatWorksheets(
  sh = SpreadsheetApp.openById('1FW_3GIWmytdrgBdfSuIl2exy9hIAnQoG8IprF8k9uEY').getSheetByName('Math'),
  rowOffset = 10,
  headerCols = 11,
  customStyles = {
    primaryColor: '#134f5c',
    primaryContrastColor: 'white',
  }) {

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

function applyConditionalFormatting(sheet, customStyles) {
  var rules = [];
  var newRules = [];

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
  const lastRow = sheet.getLastRow();
  const allVals = sheet.getRange(1, col, lastRow).getValues();
  const lastFilledRow = lastRow - allVals.reverse().findIndex((c) => c[0] != '');

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
  var testCodes = ['at1', 'at2', 'at3', 'at4', 'at5', 'at6', 'sat1', 'sat2', 'sat3', 'sat4', 'sat5', 'sat6', 'psat1', 'psat2', 'apt1', 'apt2'];
  var scoreSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('Scores');
  var lastRow = scoreSheet.getLastRow();
  var scoresNames = scoreSheet.getRange(1, 1, lastRow);
  var nextOpenRow =
    lastRow -
    scoresNames
      .getValues()
      .reverse()
      .findIndex((c) => c[0] != '') +
    1;

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
        } else if (sheetName.slice(0, 2) === 'at') {
          testName = sheetName.replace('at', 'sat').toUpperCase();
        }

        // If test is completed, add to scores array
        if (modsComplete === 4 && sheet.getRange('G1').getValue() >= 200 && sheet.getRange('I1').getValue() >= 200) {
          var rwScore = sheet.getRange('G1').getValue();
          var mScore = sheet.getRange('I1').getValue();
          var totalScore = sheet.getRange('L1').getValue();
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

            var dateSubmitted = sheet.getRange('D2').getValue();
            if ((dateSubmitted = '')) {
              dateSubmitted = Utilities.formatDate(new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - 1), 'UTC', 'MM/dd/yyyy');
            }
            var rowData = [[studentName, 'Practice', testName.toUpperCase(), dateSubmitted, totalScore, rwScore, mScore]];
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
  const lastRow = summarySheet.getLastRow();
  const allVals = summarySheet.getRange('A1:A' + lastRow).getValues();
  const lastFilledRow = lastRow - allVals.reverse().findIndex((c) => c[0] != '');
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

const showAllExcept = (spreadsheetId, hiddenSheets = []) => {
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
  const parentClientFolder = DriveApp.getFolderById(parentClientFolderId);
  const clientFolders = parentClientFolder.getFolders();
  let newRow = getLastFilledRow(clientSheet, 1) + 1;
  
  while (clientFolders.hasNext()) {
    const clientFolder = clientFolders.next();
    const clientFolderId = clientFolder.getId();

    addClientData(clientFolderId, newRow);

    newRow++;
  }
}

function addClientData(clientFolderId, newRow=null) {
  const clientFolder = DriveApp.getFolderById(clientFolderId);
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheets()[0];
  if(!newRow) {
    newRow = getLastFilledRow(clientSheet, 1) + 1;
  }
  
  const savedClientFolderIds = clientSheet.getRange(2, 2, newRow).getValues();
  const isClientFolderListed = savedClientFolderIds.some(subArray => subArray.includes(clientFolderId))
    
  if (!isClientFolderListed) {
    const clientName = clientFolder.getName();
    Logger.log(newRow + '. ' + clientName);
    if (!clientName.includes('Ξ')) {
      getStyledSheets(clientFolder);
      processFolders(clientFolder.getFolders(), getStyledSheets);

      const clientSubfolders = clientFolder.getFolders();
      let studentFolder;
      let studentFolderId = 'not found';
      let studentFolderCount;

      while (clientSubfolders.hasNext()) {
        const clientSubfolder = clientSubfolders.next();
        studentFolderCount = 0;

        if (clientSubfolder.getName() === 'Students') {
          studentFolder = clientSubfolder;
          studentFolderId = studentFolder.getId();
          const studentSubfolders = studentFolder.getFolders();
          while (studentSubfolders.hasNext()) {
            studentSubfolders.next();
            studentFolderCount ++;
          }
          studentFolderCount -= 1;
          break;
        }
      }

      clientDataSs.getSheetById(0).getRange(newRow, 1, 1, 8).setValues([[clientName, clientFolder.getId(), satSheetIds.admin, satSheetIds.student, actSheetIds.admin, actSheetIds.student, studentFolderId, studentFolderCount]]);
      newRow ++;
    }
  }
  else {
    Logger.log(clientFolder.getName() + ' present in data')
  }
}

function updateClientFolders() {
  const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const clientSheet = clientDataSs.getSheets()[0];
  const lastFilledRow = getLastFilledRow(clientSheet, 2);
  const clientDataRange = clientSheet.getRange(1, 1, lastFilledRow, 10).getValues();
  const clients = [];

  for (let row = 2; row < lastFilledRow; row++) {       // starts at 2
    clients.push({
      'name': clientDataRange[row][0],
      'folderId': clientDataRange[row][1],
      'satAdminSsId': clientDataRange[row][2],
      'satStudentSsId': clientDataRange[row][3],
      'actAdminSsId': clientDataRange[row][4],
      'actStudentSsId': clientDataRange[row][5]
    })
  }

  for (let id = 28; id < clients.length; id++) {    // starts at 0. Change this if execution times out
    Logger.log(id + '. ' + clients[id]['name'] + ' started');

    // Update Math!E294:E296 on admin and student sheets
    const sltSh = SpreadsheetApp.openById(clients[id]['satAdminSsId']).getSheetByName('SLT Uniques');

    const rwStartCell = sltSh.getRange('D5');
    const mathStartCell = sltSh.getRange('H5');
    const rwCorrectRange = sltSh.getRange('D5:D199');
    const mathCorrectRange = sltSh.getRange('H5:H199');

    rwStartCell.setValue('=if(B5="","", let(result, xlookup(B5, \'Question bank data\'!$G$2:$G, \'Question bank data\'!$H$2:$H, "not found"), if(result=C5, "", result)))');
    mathStartCell.setValue('=if(F5="","", let(result, xlookup(F5, \'Question bank data\'!$G$2:$G, \'Question bank data\'!$H$2:$H, "not found"), if(result=G5, "", result)))');

    rwStartCell.autoFill(rwCorrectRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    mathStartCell.autoFill(mathCorrectRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    // const studentFolders = DriveApp.getFolderById(clients[id]['folderId']).getFoldersByName('Students').next().getFolders();
    // while (studentFolders.hasNext()) {
    //   const studentFolder = studentFolders.next();

    //   if (!studentFolder.getName().toLowerCase().includes('template')) {
    //     updateStudentFolder(studentFolder);
    //   }
    // }
  }
}

function updateStudentFolder(studentFolder) {
  const studentFolderName = studentFolder.getName();
  Logger.log(studentFolderName + ' started');
}

function addConceptRows(answerSsId = '1sdnVpuX8mVkpTdrqZgwz7zph1NdFpueX6CP45JHiNP8') {
  const ss = SpreadsheetApp.openById(answerSsId);
  const qbDataSh = ss.getSheetByName('Question bank data');
  const qbDataVals = qbDataSh.getRange(1,1, getLastFilledRow(qbDataSh, 1), 15).getValues();

  const subjectData = [
    {
      'sub': 'rw',
      'name': 'Reading & Writing',
      'rowOffset': 7,
    },
    {
      'sub': 'm',
      'name': 'Math',
      'rowOffset': 10
    }
  ]

  for (subject of subjectData) {
    const sh = ss.getSheetByName(subject['name']);
    const conceptColVals = sh.getRange(subject['rowOffset'], 2, sh.getMaxRows() - subject['rowOffset']).getValues();
    const conceptData = [];
    const modifications = [];
    let id = 1;

    for (let x = 0; x < conceptColVals.length; x++) {
      if (cats.includes(conceptColVals[x][0])) {
        var row = x + subject['rowOffset'];
        conceptData.push({
          'name': conceptColVals[x][0],
          'row': row,
          'id': id
        });
        id ++;
      }
    }

    for (concept of conceptData) {
      for (let level = 1; level < 4; level ++) {
        let count = 0;
        for (let r = 0; r < qbDataVals.length; r++) {
          if (qbDataVals[r][3].toLowerCase() === concept['name'].toLowerCase() && Number(qbDataVals[r][4]) === level && qbDataVals[r][6].slice(0,3) !== 'SAT' && qbDataVals[r][6].slice(0,4) !== 'PSAT' && qbDataVals[r][6].slice(0,3) !== 'SLT') {
            count ++;
          }
        }
        concept['level' + level] = count;
      }
    }
    
    for (concept of conceptData) {
      const rowsNeeded = Math.max(concept['level1'], concept['level2'], concept['level3']) + 4;
      const nextConcept = conceptData.find(c => c.id === concept['id'] + 1);
      let rowsToAdd, endRow;

      if (nextConcept) {
        endRow = nextConcept.row;
      }
      else {
        Logger.log(concept['name'] + ' is last concept');
        endRow = sh.getLastRow() + 1;
      }
        
      rowsToAdd = concept['row'] + rowsNeeded - endRow;

      if (rowsToAdd > 0) {
        modifications.push({
          'position': endRow - 1,
          'rows': rowsToAdd
        });
        Logger.log('Adding ' + rowsToAdd + ' to ' + concept['name'])
      }
      else if (rowsToAdd < 0) {
        modifications.push({
          'position': concept['row'] + rowsNeeded - 1,
          'rows': rowsToAdd
        });
        Logger.log('Removing ' + -1 * rowsToAdd + ' from ' + concept['name'])
      }
    }
    modifyRowsAtPositions(sh, modifications);

    const shNewRange = sh.getRange(subject['rowOffset'], 1, sh.getMaxRows() - subject['rowOffset'], sh.getMaxColumns());
    shNewRange.setNumberFormat('@STRING@');
    const shNewVals = shNewRange.getValues();
    const newConceptRows = shNewVals.map(row => row[1]);

    for (let level = 1; level < 4; level++) {
      const levelStartCol = (level - 1) * 4;

      for (concept of conceptData) {
        concept['row'] = newConceptRows.indexOf(concept['name']);

        for (qNum = 1; qNum <= concept['level' + level]; qNum++) {
          const qRow = concept['row'] + 2 + qNum;
          const dataRow = qbDataVals.find(row => row[3].toLowerCase() === concept['name'].toLowerCase() && Number(row[4]) === level && Number(row[5]) === qNum);
          shNewVals[qRow] = []
          shNewVals[qRow][levelStartCol] = dataRow[0];
          shNewVals[qRow][levelStartCol + 1] = level + '.' + qNum;
        }
      }

      const outputValues = [];
      for (let i = 0; i < shNewVals.length; i++) {
        outputValues.push([
          shNewVals[i][levelStartCol], 
          shNewVals[i][levelStartCol + 1]
        ]);
      }

      sh.getRange(subject['rowOffset'], levelStartCol + 1, outputValues.length, 2).setValues(outputValues);
    }

    const correctedFormulaR1C1 = '=let(worksheetNum,R[0]C[-2], if(worksheetNum="","", if(left(worksheetNum,5)="Level","Corrected", if(iserror(search(".",worksheetNum)), "", let(id,R[0]C[-3], result,xlookup(id,\'Question bank data\'!R2C1:C1,\'Question bank data\'!R2C8:C8,"not found"), if(result=R[0]C[-1],"",result))))))'
    
    for (let level = 1; level < 4; level++) {
      const correctedCol = 4 * (level - 1) + 4;
      const correctedRange = sh.getRange(subject['rowOffset'] + 3, correctedCol, sh.getLastRow() - subject['rowOffset'] - 1);
      correctedRange.setFormulaR1C1(correctedFormulaR1C1);
    }

    modifyIdFormatRule(sh);
  }
}

function modifyRowsAtPositions(sheet, modifications) {
  // Sort modifications in descending order of positions to avoid shifting issues
  modifications.sort((a, b) => b.position - a.position);

  // Apply each modification
  modifications.forEach(mod => {
    if (mod.rows > 0) {
      // Insert rows if `rows` is positive
      sheet.insertRows(mod.position, mod.rows);
    } else if (mod.rows < 0) {
      // Delete rows if `rows` is negative
      sheet.deleteRows(mod.position, Math.abs(mod.rows));
    }
  });
  Logger.log('Row modifications complete')
}

function modifyIdFormatRule(sheet=SpreadsheetApp.openById('1sdnVpuX8mVkpTdrqZgwz7zph1NdFpueX6CP45JHiNP8').getSheetByName('Reading & Writing')) {
  const alertColor = '#cc0000';
  const darkGreen = '#6aa84f';
  const lightGreen = '#b7e1cd';
  const darkRed = '#e06666';
  const lightRed = '#f4c7c3';
  const grey = '#f3f3f3';
  const yellow = '#fff2cc';
  // Get all existing conditional formatting rules
  var rules = sheet.getConditionalFormatRules();
  
  // Create an array to store updated rules
  var updatedRules = [];
  
  // Iterate through each rule
  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var bgColor = rule.getBooleanCondition().getBackgroundObject().asRgbColor().asHexString();
    
    if (bgColor === alertColor) {
      // Modify the rule
      var newRule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges(sheet.getRangeList(['A10:A, E10:E, I10:I']))
        .whenFormulaSatisfied('=and(len(A10)<>8,B10<>"",B9<>"")')
        .setBackground(alertColor)
        .setFontColor('#ffffff')
        .build();
      
      updatedRules.push(newRule); // Add the modified rule to the list
    }
    else if (bgColor === darkGreen) {
      var newRule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges(sheet.getRangeList(['C10:C, G10:G, K10:K']))
        .whenFormulaSatisfied('=and(C10<>"",D10="",isformula(C10))')
        .setBackground(darkGreen)
        .setFontColor('#ffffff')
        .build();
      
      updatedRules.push(newRule); 
    }
    else if (bgColor === lightGreen) {
      var newRule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges(sheet.getRangeList(['C10:C, G10:G, K10:K']))
        .whenFormulaSatisfied('=and(C10<>"",D10="",C10<>"Answer")')
        .setBackground(lightGreen)
        .build();
      
      updatedRules.push(newRule); 
    }
    else if (bgColor === darkRed) {
      var newRule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges(sheet.getRangeList(['C10:C, G10:G, K10:K']))
        .whenFormulaSatisfied('=and(or(C10="",C10="-"),B10<>"",B9<>"")')
        .setBackground(darkRed)
        .setFontColor('#ffffff')
        .build();
      
      updatedRules.push(newRule); 
    }
    else if (bgColor === lightRed) {
      var newRule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges(sheet.getRangeList(['C10:C, G10:G, K10:K']))
        .whenFormulaSatisfied('=and(C10<>"",D10<>"",C10<>"Answer")')
        .setBackground(lightRed)
        .build();
      
      updatedRules.push(newRule); 
    }
    else if (bgColor === yellow) {
      var newRule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges(sheet.getRangeList(['C10:C, G10:G, K10:K']))
        .whenFormulaSatisfied('=and(or(C10="",C10="-"),B10<>"",B9<>"")')
        .setBackground(yellow)
        .build();
      
      updatedRules.push(newRule); 
    }
    else if (bgColor === grey) {
      var newRule = SpreadsheetApp.newConditionalFormatRule()
        .setRanges(sheet.getRangeList(['C10:C, G10:G, K10:K']))
        .whenFormulaSatisfied('=and(not(isformula(C10)),C10="",B10<>"",B9<>"")')
        .setBackground(grey)
        .build();
      
      updatedRules.push(newRule); 
    }
    else {
      updatedRules.push(rule); // Keep unmodified rules as-is
    }
  }
  
  // Reapply all rules to the sheet
  sheet.setConditionalFormatRules(updatedRules);
}





// Scheduled recurring functions

function findNewScoreReports(parentFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF') {
  if (typeof parentFolderId == 'object') {
    parentFolderId = '1_qQNYnGPFAePo8UE5NfX72irNtZGF5kF';
  }

  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var fileList = getAnalysisFiles(parentFolder, (n = 3));
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
