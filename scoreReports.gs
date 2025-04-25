function findNewCompletedTests(fileList) {
  const testCodes = getTestCodes();
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
    const scores = [];

    // Loop through test sheets within analysis spreadsheet
    Logger.log('Starting new test check for ' + studentName)
    for (testCode of testCodes) {
      const sheet = ss.getSheetByName(testCode);
      if (sheet) {
        const sheetData = sheet.getRange('A1:M57').getValues();
        const rwScore = parseInt(sheetData[0][6]) || 0;
        const mScore = parseInt(sheetData[0][8]) || 0;
        const totalScore = rwScore + mScore;
        const completionCheck = sheetData[0][12];
        let isTestNewAndComplete = false;

        let dateSubmitted = sheet.getRange('D2').getValue();
        if ((!dateSubmitted)) {
          dateSubmitted = Utilities.formatDate(new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() - 1), 'UTC', 'MM/dd/yyyy');
          sheet.getRange('D2').setValue(dateSubmitted);
        }

        // Will not run if report previously generated or scores not entered
        if (completionCheck !== '✔') {
          // Last answer for each module
          const mod1RWEnd = sheetData[30][2];
          const mod2RWEnd = sheetData[30][6];
          const mod3RWEnd = sheetData[30][10];
          const mod1MathEnd = sheetData[56][2];
          const mod2MathEnd = sheetData[56][6];
          const mod3MathEnd = sheetData[56][10];

          const values = [mod1RWEnd, mod2RWEnd, mod3RWEnd, mod1MathEnd, mod2MathEnd, mod3MathEnd];

          // Filter out blank or null values and count the remaining ones
          const completedModCount = values.filter(function(value) {
            return value !== "" && value !== null;
          }).length;

          if (completedModCount === 4) {
            isTestNewAndComplete = true;
          }

          // if (sheetName.slice(0, 3) === 'apt') {
          //   testCode = testCode.replace('apt', 'psat').toUpperCase();
          // } else if (testCode.slice(0, 2) === 'at') {
          //   testCode = testCode.replace('at', 'sat').toUpperCase();
          // }

          // If test is newly completed, add to scores array
          if (isTestNewAndComplete) {
            if (rwScore && mScore) {
              Logger.log(ssName + ' ' + testCode + ' score report started');

              const rowData = [[studentName, 'Practice', testCode, dateSubmitted, totalScore, rwScore, mScore]];
              scoreSheet.getRange(nextOpenRow, 1, 1, 7).setValues(rowData);
              nextOpenRow += 1;

              scores.push({
                test: testCode,
                rw: rwScore,
                m: mScore,
                total: totalScore,
                date: dateSubmitted,
                isNew: isTestNewAndComplete
              });

              scores = scores.sort((a, b) => new Date(a.date) - new Date(b.date))

              // scores array will include reported scores of all completed tests
              createSatScoreReport(ssId, scores);
              sheet.getRange('M1').setValue('✔');
              Logger.log(ssName + ' ' + testCode + ' score report complete');
            }
            else if (completionCheck !== '?') {
              Logger.log(`Add scores for ${studentName} on ${testCode}`);
              const email = getOPTPermissionsList(ssId);
              if (email) {
                MailApp.sendEmail({
                  to: email,
                  subject: `Enter scores for ${studentName}`,
                  htmlBody: `It appears that ${testCode} was completed for ${studentName}, but scores are missing. Please add them asap to generate a score analysis. \n` +
                  `<a href="https://docs.google.com/spreadsheets/d/${ssId}/edit?gid=${sheet.getSheetId()}">${studentName}'s admin spreadsheet</a>`,
                });
                sheet.getRange('M1').setValue('?');
              }
            }
          }
        }
        else {
          scores.push({
            test: testCode,
            rw: rwScore,
            m: mScore,
            total: totalScore,
            date: dateSubmitted,
            isNew: isTestNewAndComplete
          });
        }
      }
    }
  }
}

function createSatScoreReport(spreadsheetId, scores) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();

    var sheetsToPrint = [scores.test, scores.test + ' Analysis'];
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
          if (sheetsToPrint.includes(sh.getName())) {
            sh.showSheet();
            if (sh.getName().includes('Analysis')) {
              analysisIndex = sh.getIndex();
              spreadsheet.setActiveSheet(sh);
              // Move analysis sheet to first position so that it displays first in PDF
              spreadsheet.moveActiveSheet(1);
              // Hide column H if student did not omit any answers
              // if (sh.getRange('H7').getValue() === '-') {
              //   sh.hideColumns(8);
              // } else if (sh.getRange('H7').getValue() === 'BLANK') {
              //   sh.showColumns(8);
              // }
            }
          } else {
            sh.hideSheet();
          }
        } catch (error) {
          Logger.log(error);
        }
      });

    SpreadsheetApp.flush();
    sendPdfScoreReport(spreadsheetId, studentName, scores);
    Logger.log(testCode + ' Score report created for ' + studentName);
  }
  catch (err) {
    Logger.log(err.message + '\n\n' + err.stack);
  }

  showAllExcept(spreadsheetId);
  // Move analysis sheet back to original position
  spreadsheet.moveActiveSheet(analysisIndex);
}

function sendPdfScoreReport(spreadsheetId, studentName, scores = []) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();
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

    for (score of scores) {
      if (score.isNew) {
        // Create PDF
        var currentScore = scores.slice(-1)[0];
        var pdfName = 'SAT answer analysis - ' + studentName + ' - ' + currentScore.test;
        var studentFirstName = studentName.split(' ')[0];
        const studentData = getStudentHours(studentName);
        var response = UrlFetchApp.fetch(url_base + url_ext, options);
        var blob = response.getBlob().setName(pdfName + '.pdf');
        var scoreReportFolder = DriveApp.getFolderById(scoreReportFolderId);
        scoreReportFolder.createFile(blob);
        const prevScores = [];

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

            for (i = 0; i < prevScores.length; i++) {
              message += '<li>' + prevScores[i].test + ': ' + prevScores[i].total + ' (' + prevScores[i].rw + ' RW, ' + prevScores[i].m + ' M)</li>';
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

        var email = getOPTPermissionsList(spreadsheetId);
        if (email) {
          MailApp.sendEmail({
            to: email,
            subject: currentScore.test + ' score report for ' + studentFirstName,
            htmlBody: message,
            attachments: [blob.getAs(MimeType.PDF)],
          });
        }

        break;
      }
      else {
        prevScores.push(score)
      }
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