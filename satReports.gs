function findNewSatScoreReports(students, folderName) {
  if (!students || students.triggerUid) {
    // if students is null, get OPT data row
    const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
    const clientSheet = clientDataSs.getSheetByName('Clients');
    const myDataRange = clientSheet.getRange(2, 1, 1, 17).getValues();
    const myStudentDataValue = myDataRange[0][16];
    folderName = myDataRange[0][1];
    students = JSON.parse(myStudentDataValue);
  }

  const fileList = [];

  for (student of students) {
    if (student.satAdminSsId && !student.name.toLowerCase().includes('template')) {
      const satAdminFile = DriveApp.getFileById(student.satAdminSsId);
      const satStudentFile = DriveApp.getFileById(student.satStudentSsId);
      const lastUpdated = Math.max(satAdminFile.getLastUpdated(), satStudentFile.getLastUpdated());
      const now = new Date();
      const msInTimeLimit = 5 * 24 * 60 * 60 * 1000;

      if (now - lastUpdated <= msInTimeLimit) {
        fileList.push(satAdminFile);
      } else {
        Logger.log(`${student.name} unchanged`);
      }
    }
  }

  // Sort by most recently updated first
  fileList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  Logger.log(`${folderName}: ${fileList}`);

  findNewCompletedSats(fileList);
}

function findTeamSatScoreReports() {
  const studentDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const teamDataSheet = studentDataSs.getSheetByName('Team OPT');
  const teamDataValues = teamDataSheet.getRange(2, 1, getLastFilledRow(teamDataSheet, 1) - 1, 4).getValues();

  for (let i = 0; i < teamDataValues.length; i++) {
    const studentsStr = teamDataValues[i][3];
    const folderName = teamDataValues[i][1];
    const students = JSON.parse(studentsStr);
    findNewSatScoreReports(students, folderName);
  }
}

function findNewCompletedSats(fileList) {
  const testCodes = getSatTestCodes();
  const scoreSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('SAT scores');
  let nextOpenRow = getLastFilledRow(scoreSheet, 1) + 1;

  // Loop through analysis spreadsheets
  for (var i = 0; i < fileList.length; i++) {
    const ssId = fileList[i].getId();
    const ss = SpreadsheetApp.openById(ssId);
    const ssName = ss.getName();
    const studentName = ssName.slice(ssName.indexOf('-') + 2);
    const practiceTestData = ss.getSheetByName('Practice test data').getDataRange().getValues();
    let scores = [];

    // Loop through test sheets within analysis spreadsheet
    Logger.log('Starting new test check for ' + studentName);

    for (testCode of testCodes) {
      const completedRwTestRows = practiceTestData.filter((row) => row[0] === testCode && row[1] === 'Reading & Writing' && row[10] !== '');
      const completedMathTestRows = practiceTestData.filter((row) => row[0] === testCode && row[1] === 'Math' && row[10] !== '');
      const completedRwQuestionCount = completedRwTestRows.length;
      const completedMathQuestionCount = completedMathTestRows.length;

      if (completedRwQuestionCount > 10 && completedMathQuestionCount > 10) {
        let testSheet = ss.getSheetByName(testCode);

        if (testSheet) {
          const testHeaderValues = testSheet.getRange('A1:M2').getValues();
          const rwScore = parseInt(testHeaderValues[0][6]) || 0;
          const mScore = parseInt(testHeaderValues[0][8]) || 0;
          const totalScore = rwScore + mScore;
          const dateSubmitted = testHeaderValues[1][3];
          const completionCheck = testHeaderValues[0][12];
          const sheetIndex = testSheet.getIndex();
          const isTestNew = completionCheck !== '✔';

          if (rwScore && mScore) {
            scores.push({
              test: testCode,
              rw: rwScore,
              m: mScore,
              total: totalScore,
              date: dateSubmitted,
              sheetIndex: sheetIndex,
              isNew: isTestNew,
            });
          } else if (completionCheck !== '?') {
            Logger.log(`Add scores for ${studentName} on ${testCode}`);
            const email = getOPTPermissionsList(ssId);
            if (email) {
              MailApp.sendEmail({
                to: email,
                subject: `Enter scores for ${studentName}`,
                htmlBody:
                  `It appears that ${testCode} was completed for ${studentName}, but scores are missing. Please add them asap to generate a score analysis. \n` + `<a href="https://docs.google.com/spreadsheets/d/${ssId}/edit?gid=${testSheet.getSheetId()}">${studentName}'s admin spreadsheet</a>`,
              });
              const completionCheckRange = testSheet.getRange('M1');
              completionCheckRange.setValue('?');
              completionCheckRange.setVerticalAlignment('middle');
            }
          }
        } else {
          createStudentFolders.addSatTestSheets(ssId);
        }
      }
    }

    scores = scores.sort((a, b) => new Date(a.date) - new Date(b.date));

    // scores array will include reported scores of all completed tests
    createSatScoreReports(ssId, scores);
  }
}

function createSatScoreReports(spreadsheetId, allTestData) {
  spreadsheetId = spreadsheetId ? spreadsheetId : SpreadsheetApp.getActiveSpreadsheet().getId();
  const pastTestData = [];

  try {
    for (testData of allTestData) {
      if (testData.isNew) {
        sendSatScoreReportPdf(spreadsheetId, testData, pastTestData);
      }
      pastTestData.push(testData);
    }
  } catch (err) {
    Logger.log(err.message + '\n\n' + err.stack);
  }
}

async function sendSatScoreReportPdf(spreadsheetId, currentTestData, pastTestData = []) {
  try {
    const spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();
    const ssName = spreadsheet.getName();
    const studentName = ssName.slice(ssName.indexOf('-') + 2);
    const practiceDataSheet = spreadsheet.getSheetByName('Practice test data');
    let scoreReportFolderId;

    if (practiceDataSheet.getRange('V1').getValue() === 'Score report folder ID:' && practiceDataSheet.getRange('W1').getValue() !== '') {
      scoreReportFolderId = practiceDataSheet.getRange('W1').getValue();
    } else {
      var parentId = DriveApp.getFileById(spreadsheetId).getParents().next().getId();
      const subfolderIds = getSubFolderIdsByFolderId(parentId);

      for (let i in subfolderIds) {
        let subfolderId = subfolderIds[i];
        let subfolder = DriveApp.getFolderById(subfolderId);
        let subfolderName = subfolder.getName();
        if (subfolderName.toLowerCase().includes('score report')) {
          scoreReportFolderId = subfolder.getId();
        }
      }
    }

    if (!scoreReportFolderId) {
      scoreReportFolderId = DriveApp.getFolderById(parentId).createFolder('Score reports').getId();
      practiceDataSheet.getRange('V1:W1').setValues([['Score report folder ID:', scoreReportFolderId]]);
    }

    const pdfName = currentTestData.test + ' answer analysis - ' + studentName + '.pdf';
    const answerSheetId = spreadsheet.getSheetByName(currentTestData.test).getSheetId();
    const analysisSheetId = spreadsheet.getSheetByName(currentTestData.test + ' analysis').getSheetId();

    Logger.log(`Starting ${currentTestData.test} score report for ${studentName}`);

    const answerFileId = savePdfSheet(spreadsheetId, answerSheetId, studentName);
    const analysisFileId = savePdfSheet(spreadsheetId, analysisSheetId, studentName);

    const fileIdsToMerge = [analysisFileId, answerFileId];

    const mergedFile = await mergePDFs(fileIdsToMerge, scoreReportFolderId, pdfName);
    const mergedBlob = mergedFile.getBlob();

    const studentFirstName = studentName.split(' ')[0];
    const studentData = getStudentsSpreadsheetData(studentName);

    if (studentData.hours) {
      var message =
        'Hi PARENTNAME, please find the score report from ' +
        studentFirstName +
        "'s recent practice test attached. " +
        currentTestData.total +
        ' overall (' +
        currentTestData.rw +
        ' Reading & Writing, ' +
        currentTestData.m +
        ' Math)<br><br>As of the session on ' +
        studentData.recentSessionDate +
        ', we have ' +
        studentData.hours +
        ' hours remaining on the current package. Let me know if you have any questions. Thanks!<br><br>';

      if (pastTestData.length > 1) {
        message += 'All scores - most recent last:<br><ul>';

        for (i = 0; i < pastTestData.length; i++) {
          const data = pastTestData[i];
          message += '<li>' + data.test + ': ' + data.total + ' (' + data.rw + ' RW, ' + data.m + ' M)</li>';
        }
        message += '</ul><br>';
      }
    } else {
      var message = 'Hi PARENTNAME, please find the score report from ' + studentFirstName + "'s recent practice test attached. " + currentTestData.total + ' overall (' + currentTestData.rw + ' Reading & Writing, ' + currentTestData.m + ' Math)<br><br>';
    }

    const email = getOPTPermissionsList(spreadsheetId);
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: currentTestData.test + ' score report for ' + studentFirstName,
        htmlBody: message,
        attachments: [mergedBlob],
      });
    }

    const testSheet = spreadsheet.getSheetByName(currentTestData.test);
    const completionCheckCell = testSheet.getRange('M1');
    completionCheckCell.setValue('✔');
    completionCheckCell.setVerticalAlignment('middle');
    Logger.log(studentName + ' ' + currentTestData.test + ' score report complete');
  } catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack);
  }
}
