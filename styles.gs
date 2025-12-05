function styleClientFolder(clientFolder, customStyles = {}) {
  let clientFolderId;
  const techSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tech');
  const progressRange = techSheet.getRange('I1:K1');
  let progressValue = progressRange.getValues()[0][0];
  const styleDataStr = progressRange.getValues()[0][1] || '[]';
  const ssIdsStr = progressRange.getValues()[0][2] || '[]';
  let ssIds = JSON.parse(ssIdsStr);

  
  if (Object.keys(customStyles).length === 0) {
    customStyles = JSON.parse(styleDataStr);

    if (Object.keys(customStyles).length === 0) {
      customStyles = setCustomStyles();
    }
  }

  if (progressValue) {
    clientFolderId = customStyles.clientFolderId;
    clientFolder = DriveApp.getFolderById(clientFolderId);
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<p>Continuing ${clientFolder.getName()}</p>`).setHeight(50), 'Folder in progress');
    satSsIds = ssIds.sat;
    actSsIds = ssIds.act;
  } //
  else if (clientFolder) {
    clientFolderId = clientFolder.getId();
  } //
  else {
    const ui = SpreadsheetApp.getUi();
    const prompt = ui.prompt('Client folder URL or ID', ui.ButtonSet.OK_CANCEL);
    response = prompt.getResponseText();

    if (prompt.getSelectedButton() === ui.Button.CANCEL) {
      return;
    } //
    else {
      clientFolderId = TestPrepAnalysis.getIdFromDriveUrl(response);
      clientFolder = DriveApp.getFolderById(clientFolderId);
    }
  }

  if (progressValue !== 1) {
    customStyles.clientFolderId = clientFolderId;
    customStyles.clientName = clientFolder.getName();

    Logger.log('Styling sheets for ' + customStyles.clientName);
    getAnswerSheets(clientFolder);
    processFolders(clientFolder.getFolders(), getAnswerSheets);

    ssIds = {
      sat: satSsIds,
      act: actSsIds
    }

    progressValue = 1;
    progressRange.setValues([[progressValue, JSON.stringify(customStyles), JSON.stringify(ssIds)]]);
  }

  styleClientSheets(satSsIds, actSsIds, customStyles, ssIds);
  progressRange.clearContent();

  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://drive.google.com/drive/u/0/folders/' + clientFolderId + '" target="_blank" onclick="google.script.host.close()">Client folder</a>')
    .setWidth(250)
    .setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Styling complete');
}

function styleClientSheets(satSsIds, actSsIds, customStyles, ssIds) {
  for (let ids of [satSsIds, actSsIds]) {
    for (let role of ['admin', 'student']) {
      const ssCompleteKey = role + 'SsComplete';
      const sheetsCompleteKey = role + 'SheetsComplete';

      if (!ids[ssCompleteKey]) {
        const ss = SpreadsheetApp.openById(ids[role]);
        const ssName = ss.getName();
        const satTestSheets = getSatTestCodes();
        const satDataSheets = ['question bank data', 'practice test data', 'rev sheet backend'];
        const actDataSheets = ['data', 'scoring'];
        const actTestCodes = getActTestCodes();

        const primaryColor = customStyles.primaryColor;
        const primaryContrastColor = customStyles.primaryContrastColor;
        const secondaryColor = customStyles.secondaryColor;
        const secondaryContrastColor = customStyles.secondaryContrastColor;
        const fontColor = customStyles.fontColor;
        const imgUrl = customStyles.img;

        if (ssName.includes('ACT')) {
          for (let j in ss.getSheets()) {
            const sh = ss.getSheets()[j];
            const shName = sh.getName().toLowerCase();

            if (!ids[sheetsCompleteKey].includes(shName)) {
              Logger.log(`Starting ${shName}`);

              const shRange = sh.getDataRange();
              shRange.setBackground('white');
              shRange.setFontColor(fontColor);

              if (actTestCodes.includes(shName)) {
                sh.getRange('A1:P4').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);

                sh.getRangeList(['B3', 'F3', 'J3', 'N3']).setBorder(true, true, true, true, true, true, '#93c47d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

                sh.getRange('F1').setBackground('#93c47d');
              } //
              else if (shName === 'test analysis' || shName === 'opportunity areas') {
                sh.getRange(1, 1, 8, sh.getMaxColumns())
                  .setBackground(primaryColor)
                  .setFontColor(primaryContrastColor)
                  .setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID)
                  .setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

                if (shName === 'test analysis') {
                  var correctRange = 'F6:J6';
                } //
                else {
                  var correctRange = 'E6:I6';
                }
                sh.getRange(correctRange).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

                if (imgUrl && shName === 'test analysis') {
                  const businessNameCell = sh.getRange('B3');
                  businessNameCell.setValue('=image("' + customStyles.img + '")');
                }

                applyConditionalFormatting(sh, customStyles);
              } //
              else if (actDataSheets.includes(shName)) {
                sh.getRange(1, 1, 1, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor);
              } //
              else if (shName === 'student responses') {
                sh.getRange(1, 1, 3, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
              }
          
              ids[sheetsCompleteKey].push(shName);
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tech').getRange('K1').setValue(JSON.stringify(ssIds));
            }
          }
        } //
        else if (ssName.includes('SAT')) {
          for (let j in ss.getSheets()) {
            const sh = ss.getSheets()[j];
            const shRange = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
            const shName = sh.getName();
            const shNameLower = shName.toLowerCase();

            Logger.log(`Starting ${shName}`);

            if (!shNameLower.includes('rev sheet')) {
              shRange.setFontColor(fontColor);
            }

            // practice SAT answer sheets
            if (satTestSheets.includes(shName)) {
              shRange.setBackground('white');
              if (customStyles.sameHeaderColor) {
                sh.getRangeList(['B2:L4', 'B33:L35']).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
              } //
              else {
                sh.getRangeList(['B2:L4', 'B33:L35']).setBackground(secondaryColor).setFontColor(secondaryContrastColor).setBorder(true, true, true, true, true, true, secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
              }
              sh.getRangeList(['A1:A', 'E5:E', 'I5:I']).setFontColor('white');
            } //
            else if (shNameLower.includes('analysis') || shNameLower.includes('opportunity')) {
              if (shNameLower === 'rev analysis' || shNameLower === 'opportunity areas') {
                sh.getRange('A1:K7').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
              } //
              else if (shNameLower === 'time series analysis') {
                sh.getRange('A1:K6').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
                sh.getRange('D5:E6').setFontColor(fontColor);
              } //
              else {
                sh.getRange('A1:K6').setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, false, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID).setBorder(null, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
                sh.getRange('A7:A8').setFontColor('white');
                applyConditionalFormatting(sh, customStyles);
              }

              if (imgUrl && shNameLower.includes('opportunity')) {
                const imgCell = sh.getRange('B2');
                imgCell.setValue('=image("' + customStyles.img + '")');
              }
            } //
            else if (shNameLower === 'reading & writing') {
              styleSatWorksheets(sh, 6, 11, customStyles);
            } //
            else if (shNameLower === 'math') {
              styleSatWorksheets(sh, 9, 11, customStyles);
            } //
            else if (shNameLower === 'slt uniques') {
              styleSatWorksheets(sh, 1, 7, customStyles);
            } //
            else if (satDataSheets.includes(shNameLower)) {
              sh.getRange(1, 1, 1, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor);
            } //
            else if (shNameLower === 'student responses') {
              sh.getRange(1, 1, 3, sh.getMaxColumns()).setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
            } //
            else if (shNameLower === 'rev sheets') {
              let revSheetHeaderRange;
              if (ssName.toLowerCase().includes('admin answer analysis')) {
                revSheetHeaderRange = sh.getRangeList(['B2:E4', 'G2:J4']);
              } //
              else {
                revSheetHeaderRange = sh.getRangeList(['B2:D4', 'F2:I4']);
              }

              if (customStyles.sameHeaderColor) {
                revSheetHeaderRange.setBackground(primaryColor).setFontColor(primaryContrastColor).setBorder(true, true, true, true, true, true, primaryColor, SpreadsheetApp.BorderStyle.SOLID);
              } //
              else {
                revSheetHeaderRange.setBackground(secondaryColor).setFontColor(secondaryContrastColor).setBorder(true, true, true, true, true, true, secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
              }
            }
          }
        }

        ids[ssCompleteKey] = true;
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tech').getRange('K1').setValue(JSON.stringify(ssIds));
        Logger.log(ssCompleteKey);
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
  }
) {
  const cats = TestPrepAnalysis.cats;
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
    } //
    else {
      highlightRange.setBackground(customStyles.secondaryColor).setFontColor(customStyles.secondaryContrastColor).setBorder(true, true, true, true, true, true, customStyles.secondaryColor, SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}

function applyConditionalFormatting(sheet = SpreadsheetApp.openById('1nwG8o6Rd0ArGQMzrfUjRP16FkSw9urEIK-V7UD2ayJM'), customStyles = null) {
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
      newRule = rules[i].copy();
      newRules.push(newRule);
    }
  }

  if (sheet.getName().toLowerCase().includes('opportunity')) {
    var subtotalStart = 'B';
    var domainStart = 'C';
  } //
  else {
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
  } //
  else if (isDark(fontColor)) {
    primaryContrastColor = fontColor;
  } //
  else {
    primaryContrastColor = 'black';
  }

  if (isDark(secondaryColor)) {
    var secondaryContrastColor = 'white';
  } //
  else if (isDark(fontColor)) {
    secondaryContrastColor = fontColor;
  } //
  else {
    secondaryContrastColor = 'black';
  }

  if (isDark(tertiaryColor)) {
    var tertiaryContrastColor = 'white';
  } //
  else if (isDark(fontColor)) {
    tertiaryContrastColor = fontColor;
  } //
  else {
    tertiaryContrastColor = 'black';
  }

  if (sameHeaderColor === ui.Button.YES) {
    sameHeaderColor = true;
  } //
  else {
    sameHeaderColor = false;
  }

  let imgUrl;
  if (imgFilename.toLowerCase().includes('.com/') || imgFilename === '') {
    imgUrl = imgFilename;
  } //
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
