function findNewActScoreReports(students, folderName) {
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
    if (student.actAdminSsId && !student.name.toLowerCase().includes('template')) {
      const actAdminFile = DriveApp.getFileById(student.actAdminSsId);
      const actStudentFile = DriveApp.getFileById(student.actStudentSsId);
      const lastUpdated = Math.max(actAdminFile.getLastUpdated(), actStudentFile.getLastUpdated());
      const now = new Date();
      const msInTimeLimit = 5 * 24 * 60 * 60 * 1000;

      if (now - lastUpdated <= msInTimeLimit) {
        fileList.push(actAdminFile);
      } else {
        Logger.log(`${student.name} unchanged`);
      }
    }
  }

  // Sort by most recently updated first
  fileList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  Logger.log(`${folderName}: ${fileList}`);

  findNewCompletedActs(fileList);
}

function findTeamActScoreReports() {
  const studentDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const teamDataSheet = studentDataSs.getSheetByName('Team OPT');
  const teamDataValues = teamDataSheet.getRange(2, 1, getLastFilledRow(teamDataSheet, 1) - 1, 4).getValues();

  for (let i = 0; i < teamDataValues.length; i++) {
    const studentsStr = teamDataValues[i][3];
    const folderName = teamDataValues[i][1];
    const students = studentsStr ? JSON.parse(studentsStr) : [];
    findNewActScoreReports(students, folderName);
  }
}

function findNewCompletedActs(fileList) {
  const testCodes = getActTestCodes();
  const scoreSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('ACT scores');
  let nextOpenRow = getLastFilledRow(scoreSheet, 1) + 1;

  // Loop through analysis spreadsheets
  for (var i = 0; i < fileList.length; i++) {
    const ssId = fileList[i].getId();
    const ss = SpreadsheetApp.openById(ssId);
    const ssName = ss.getName();
    const studentName = ssName.slice(ssName.indexOf('-') + 2);
    const allData = ss.getSheetByName('Data').getDataRange().getValues();
    let scores = [];

    // Loop through test sheets within analysis spreadsheet
    Logger.log('Starting new test check for ' + studentName);

    for (testCode of testCodes) {
      const completedEnglishCount = allData.filter((row) => row[0] === testCode && row[1] === 'English' && row[7] !== '').length;
      const completedMathCount = allData.filter((row) => row[0] === testCode && row[1] === 'Math' && row[7] !== '').length;
      const completedReadingCount = allData.filter((row) => row[0] === testCode && row[1] === 'Reading' && row[7] !== '').length;
      const completedScienceCount = allData.filter((row) => row[0] === testCode && row[1] === 'Science' && row[7] !== '').length;

      if (completedEnglishCount > 37 && completedMathCount > 30 && completedReadingCount > 20 && completedScienceCount > 20) {
        let testSheet = ss.getSheetByName(testCode);

        if (testSheet) {
          const testHeaderValues = testSheet.getRange('A1:N3').getValues();
          const eScore = parseInt(testHeaderValues[2][1]) || 0;
          const mScore = parseInt(testHeaderValues[2][5]) || 0;
          const rScore = parseInt(testHeaderValues[2][9]) || 0;
          const sScore = parseInt(testHeaderValues[2][13]) || 0;
          const totalScore = parseInt(testHeaderValues[0][5]) || '';
          const dateSubmitted = testHeaderValues[0][9];
          const isTestNew = testHeaderValues[0][6] !== 'Submitted on:';

          if (totalScore) {
            scores.push({
              studentName: studentName,
              test: testCode,
              eScore: eScore,
              mScore: mScore,
              rScore: rScore,
              sScore: sScore,
              total: totalScore,
              date: dateSubmitted,
              isNew: isTestNew,
            });
          }
        } else {
          createStudentFolders.addSatTestSheets(ssId);
        }
      }
    }

    scores = scores.sort((a, b) => new Date(a.date) - new Date(b.date));

    // scores array will include reported scores of all completed tests
    createActScoreReports(ssId, scores);
  }
}

function createActScoreReports(spreadsheetId, allTestData) {
  spreadsheetId = spreadsheetId ? spreadsheetId : SpreadsheetApp.getActiveSpreadsheet().getId();
  const pastTestData = [];

  try {
    for (testData of allTestData) {
      if (testData.isNew) {
        sendActScoreReportPdf(spreadsheetId, testData, pastTestData);
      }
      pastTestData.push(testData);
    }
  } catch (err) {
    Logger.log(err.message + '\n\n' + err.stack);
  }
}

async function sendActScoreReportPdf(spreadsheetId, currentTestData, pastTestData = []) {
  try {
    const spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();
    const ssName = spreadsheet.getName();
    const studentName = ssName.slice(ssName.indexOf('-') + 2);
    const dataSheet = spreadsheet.getSheetByName('Data');
    let scoreReportFolderId, studentFolderId;

    if (dataSheet.getRange('V1').getValue() === 'Score report folder ID:' && dataSheet.getRange('W1').getValue() !== '') {
      scoreReportFolderId = dataSheet.getRange('W1').getValue();
    } else {
      var parentFolderId = DriveApp.getFileById(spreadsheetId).getParents().next().getId();
      const subfolderIds = getSubFolderIdsByFolderId(parentFolderId);

      for (let i in subfolderIds) {
        let subfolderId = subfolderIds[i];
        let subfolder = DriveApp.getFolderById(subfolderId);
        let subfolderName = subfolder.getName();
        if (subfolderName.toLowerCase().includes('score report')) {
          scoreReportFolderId = subfolderId;
          break;
        } else if (subfolderName.includes(studentName)) {
          studentFolderId = subfolderId;
        }
      }

      if (studentFolderId && !scoreReportFolderId) {
        const subSubfolderIds = getSubFolderIdsByFolderId(studentFolderId);

        for (let id in subSubfolderIds) {
          let subSubfolderId = subSubfolderIds[id];
          let subSubfolder = DriveApp.getFolderById(subSubfolderId);
          let subSubfolderName = subSubfolder.getName();

          if (subSubfolderName.toLowerCase().includes('score report')) {
            scoreReportFolderId = subfolderId;
            break;
          }
        }

        if (!scoreReportFolderId) {
          scoreReportFolderId = DriveApp.getFolderById(studentFolderId).createFolder('Score reports').getId();
        }
      }
    }

    if (!scoreReportFolderId) {
      scoreReportFolderId = DriveApp.getFolderById(parentId).createFolder('Score reports').getId();
    }

    if (dataSheet.getRange('W1').getValue() !== scoreReportFolderId) {
      dataSheet.getRange('V1:W1').setValues([['Score report folder ID:', scoreReportFolderId]]);
    }

    const pdfName = `ACT ${currentTestData.test} answer analysis - ${studentName}.pdf`;
    const answerSheetId = spreadsheet.getSheetByName(currentTestData.test).getSheetId();
    const analysisSheetName = currentTestData.test + ' analysis';
    let analysisSheet = spreadsheet.getSheetByName(analysisSheetName);

    if (!analysisSheet) {
      const testAnalysisSheet = spreadsheet.getSheetByName('Test analysis');
      analysisSheet = testAnalysisSheet.copyTo(spreadsheet).setName(analysisSheetName);
      const analysisPivot = analysisSheet.getPivotTables()[0];

      // Clear the "Test code" filter, then set it to currentTestData.test
      if (analysisPivot) {
        // Remove any existing "Test code" filter and set it to currentTestData.test
        const filters = analysisPivot.getFilters();
        // Remove any filter for "Test code"
        const newFilters = filters.filter((f) => f.getSourceDataColumn().getName() !== 'Test code');
        // Add the new filter for the current test code
        const testCodeColumn = analysisPivot.getSourceDataColumnByName ? analysisPivot.getSourceDataColumnByName('Test code') : null;
        if (testCodeColumn) {
          newFilters.push(SpreadsheetApp.newPivotFilter().setSourceDataColumn(testCodeColumn).setValues([currentTestData.test]));
          analysisPivot.setFilters(newFilters);
        }
      }
    }

    const analysisSheetId = analysisSheet.getSheetId();

    Logger.log(`Starting ${currentTestData.test} score report for ${studentName}`);

    const answerSheetMargins = { top: '0.3', bottom: '0.25', left: '0.35', right: '0.35' };
    const answerFileId = savePdfSheet(spreadsheetId, answerSheetId, studentName, answerSheetMargins);

    const analysisSheetWidth = 1296;
    const analysisSheetMargin = { top: '0.25', bottom: null, left: '0.25', right: '0.25' };
    const pageScaleFactor = 576 / analysisSheetWidth; // fitw=true scales page to 576px given 0.25in left/right margins
    const headerHeightPixels = 24 * 8;
    const mathTotalRow = findActMathTotalRow(analysisSheet, 3);
    const bodyHeightPixels = (mathTotalRow - 8) * 21;
    const pageHeightPixels = headerHeightPixels + bodyHeightPixels;
    const pageHeightScaled = pageHeightPixels / pageScaleFactor;
    const marginTopPixels = parseInt(analysisSheetMargin.top);
    const pageBreakHeight = pageHeightScaled + marginTopPixels;
    analysisSheetMargin.bottom = String((792 - pageBreakHeight) / 72); // PDF is 792px tall and 72px/in when fitw=true
    // const exportRangeEquals = '&range=A1:L' + grandTotalRow;
    const analysisFileId = savePdfSheet(spreadsheetId, analysisSheetId, studentName, analysisSheetMargin);

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
        currentTestData.eScore +
        ' English, ' +
        currentTestData.mScore +
        ' Math, ' +
        currentTestData.rScore +
        ' Reading, ' +
        currentTestData.sScore +
        ' Science, )<br><br>As of the session on ' +
        studentData.recentSessionDate +
        ', we have ' +
        studentData.hours +
        ' hours remaining on the current package. Let me know if you have any questions. Thanks!<br><br>';

      if (pastTestData.length > 1) {
        message += 'All scores - most recent last:<br><ul>';

        for (i = 0; i < pastTestData.length; i++) {
          const data = pastTestData[i];
          message += '<li>' + data.test + ': ' + data.total + ' (' + data.eScore + 'E, ' + data.mScore + 'M, ' + data.rScore + 'R, ' + data.sScore + 'S)</li>';
        }
        message += '</ul><br>';
      }
    } else {
      var message =
        'Hi PARENTNAME, please find the score report from ' + studentFirstName + "'s recent practice test attached. " + currentTestData.total + ' overall (' + currentTestData.eScore + 'E, ' + currentTestData.mScore + 'M, ' + currentTestData.rScore + 'R, ' + currentTestData.sScore + 'S)<br><br>';
    }

    const email = getOPTPermissionsList(spreadsheetId);
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: 'ACT ' + currentTestData.test + ' score report for ' + studentFirstName,
        htmlBody: message,
        attachments: [mergedBlob],
      });
    }

    const testSheet = spreadsheet.getSheetByName(currentTestData.test);
    const completionCheckCell = testSheet.getRange('G1:I1').merge();
    completionCheckCell.setValue('Submitted on:');
    completionCheckCell.setHorizontalAlignment('right').setFontWeight('normal');
    testSheet.getRange('J1').setValue(currentTestData.date).setHorizontalAlignment('center').setFontWeight('normal');
    Logger.log(studentName + ' ' + currentTestData.test + ' score report complete');
  } catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack);
  }
}
