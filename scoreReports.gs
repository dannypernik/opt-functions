function findNewScoreReports(students, folderName) {
  if (!students || students.triggerUid) {
    // if students is null, get OPT data row
    const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
    const clientSheet = clientDataSs.getSheetByName('Clients');
    const myDataRange = clientSheet.getRange(2,1,1,17).getValues();
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
      const msInThreeDays = 5 * 24 * 60 * 60 * 1000;

      if ((now - lastUpdated) <= msInThreeDays) {
        fileList.push(satAdminFile);
      }
      else {
        Logger.log(`${student.name} unchanged`)
      }
    }
  }

  // Sort by most recently updated first
  fileList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  Logger.log(`${folderName}: ${fileList}`);

  findNewCompletedTests(fileList);
}

function findTeamScoreReports() {
  const studentDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const teamDataSheet = studentDataSs.getSheetByName('Team OPT');
  const teamDataValues = teamDataSheet.getRange(2,1,getLastFilledRow(teamDataSheet, 1) - 1, 4).getValues();

  for (let i = 0; i < teamDataValues.length; i ++) {
    const studentsStr = teamDataValues[i][3];
    const folderName = teamDataValues[i][1];
    const students = JSON.parse(studentsStr);
    findNewScoreReports(students, folderName);
  }
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
      'index': tutorIndex,
      'name': tutorFolderName,
      'studentsFolderId': tutorFolderId,
      'studentsDataJSON': tutorStudents
    }

    tutorStudents = createStudentFolders.findStudentFileIds(tutorData)

    teamDataSheet.getRange(tutorIndex + 2,1,1,4).setValues([[tutorIndex, tutorFolderName, tutorFolderId, JSON.stringify(tutorStudents)]]);
    tutorIndex ++;
  }
  
  const clientSheet = clientDataSs.getSheetByName('Clients')
  const myStudentsStr = clientSheet.getRange(2, 17).getValue();
  let myStudents = JSON.parse(myStudentsStr);

  const myStudentFolderData = {
    'index': 0,
    'name': 'Open Path Tutoring',
    'studentsFolderId': clientSheet.getRange(2, 15).getValue(),
    'studentsDataJSON': myStudents
  }

  myStudents = createStudentFolders.findStudentFileIds(myStudentFolderData);
  clientSheet.getRange(2, 17).setValue(JSON.stringify(myStudents));
}

function findNewCompletedTests(fileList) {
  const testCodes = getTestCodes();
  const scoreSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('Scores');
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
      const completedRwTestRows = practiceTestData.filter(row => row[0] === testCode && row[1] === 'Reading & Writing' && row[10] !== '');
      const completedMathTestRows = practiceTestData.filter(row => row[0] === testCode && row[1] === 'Math' && row[10] !== '');
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
          const isTestNew = completionCheck !== '✔';

          if (rwScore && mScore) {
            scores.push({
              'test': testCode,
              'rw': rwScore,
              'm': mScore,
              'total': totalScore,
              'date': dateSubmitted,
              'isNew': isTestNew
            })
          }
          else if (completionCheck !== '?') {
            Logger.log(`Add scores for ${studentName} on ${testCode}`);
            const email = getOPTPermissionsList(ssId);
            if (email) {
              MailApp.sendEmail({
                to: email,
                subject: `Enter scores for ${studentName}`,
                htmlBody: `It appears that ${testCode} was completed for ${studentName}, but scores are missing. Please add them asap to generate a score analysis. \n` +
                `<a href="https://docs.google.com/spreadsheets/d/${ssId}/edit?gid=${testSheet.getSheetId()}">${studentName}'s admin spreadsheet</a>`,
              });
              const completionCheckRange = testSheet.getRange('M1');
              completionCheckRange.setValue('?');
              completionCheckRange.setVerticalAlignment('middle');
            }
          }
        }
        else {
          createStudentFolders.addTestSheets(ssId);
        }

        
      }
    }

    scores = scores.sort((a, b) => new Date(a.date) - new Date(b.date));

    // scores array will include reported scores of all completed tests
    createSatScoreReports(ssId, scores);
  }
}

function createSatScoreReports(spreadsheetId, scores) {
  spreadsheetId = spreadsheetId ? spreadsheetId : SpreadsheetApp.getActiveSpreadsheet().getId();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const upToPresentScores = [];
  
  try {
    for (score of scores) {
      if (score.isNew) {
        const filename = spreadsheet.getName();
        const studentName = filename.slice(filename.indexOf('-') + 2);
        upToPresentScores.push(score);

        sendPdfScoreReport(spreadsheetId, studentName, upToPresentScores);
      }
      else {
        upToPresentScores.push(score)
      }
    }
  }
  catch (err) {
    Logger.log(err.message + '\n\n' + err.stack);
  }
}

async function sendPdfScoreReport(spreadsheetId, studentName, scoresUpToCurrent = []) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();
    var practiceDataSheet = spreadsheet.getSheetByName('Practice test data');
    let scoreReportFolderId;

    if (practiceDataSheet.getRange('V1').getValue() === 'Score report folder ID:' && practiceDataSheet.getRange('W1').getValue() !== '') {
      scoreReportFolderId = practiceDataSheet.getRange('W1').getValue();
    }
    else {
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

    const currentScore = scoresUpToCurrent.slice(-1)[0];
    const pdfName = currentScore.test + ' answer analysis - ' + studentName + '.pdf'
    const answerSheetId = spreadsheet.getSheetByName(currentScore.test).getSheetId();
    const analysisSheetId = spreadsheet.getSheetByName(currentScore.test + ' analysis').getSheetId();

    Logger.log(`Starting ${currentScore.test} score report for ${studentName}`);

    const answerFileId = savePdfSheet(spreadsheetId, answerSheetId, studentName);
    const analysisFileId = savePdfSheet(spreadsheetId, analysisSheetId, studentName);

    const fileIdsToMerge= [analysisFileId, answerFileId];

    const mergedFile = await mergePDFs(fileIdsToMerge, scoreReportFolderId, pdfName);
    const mergedBlob = mergedFile.getBlob();

    const studentFirstName = studentName.split(' ')[0];
    const studentData = getStudentsSpreadsheetData(studentName);

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
        studentData.hours +
        ' hours remaining on the current package. Let me know if you have any questions. Thanks!<br><br>';

      if (scoresUpToCurrent.length > 1) {
        message += 'All scores - most recent last:<br><ul>';

        for (i = 0; i < scoresUpToCurrent.length; i++) {
          const score = scoresUpToCurrent[i];
          message += '<li>' + score.test + ': ' + score.total + ' (' + score.rw + ' RW, ' + score.m + ' M)</li>';
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

    const email = getOPTPermissionsList(spreadsheetId);
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: currentScore.test + ' score report for ' + studentFirstName,
        htmlBody: message,
        attachments: [mergedBlob],
      });
    }

    const testSheet = spreadsheet.getSheetByName(currentScore.test);
    const completionCheckCell = testSheet.getRange('M1');
    completionCheckCell.setValue('✔');
    completionCheckCell.setVerticalAlignment('middle');
    Logger.log(studentName + ' ' + currentScore.test + ' score report complete');
  }

  catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack)
  }
}

async function mergePDFs(fileIds, destinationFolderId, name="merged.pdf") {
  // Retrieve PDF data as byte arrays
  const data = fileIds.map(id => new Uint8Array(DriveApp.getFileById(id).getBlob().getBytes()));

  // Load pdf-lib from CDN
  const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
  eval(UrlFetchApp.fetch(cdnjs).getContentText().replace(/setTimeout\(.*?,.*?(\d*?)\)/g, "Utilities.sleep($1);return t();"));

  // Merge PDFs
  const pdfDoc = await PDFLib.PDFDocument.create();
  for (let i = 0; i < data.length; i++) {
    const pdfData = await PDFLib.PDFDocument.load(data[i]);
    const pages = await pdfDoc.copyPages(pdfData, pdfData.getPageIndices());
    pages.forEach(page => pdfDoc.addPage(page));
  }

  // Save merged PDF to Drive
  const bytes = await pdfDoc.save();
  const mergedBlob = Utilities.newBlob([...new Int8Array(bytes)], MimeType.PDF, "merged.pdf");
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const mergedFile = destinationFolder.createFile(mergedBlob).setName(name);

  return mergedFile;
}


function savePdfSheet(spreadsheetId, sheetId, studentName) {
  try {
    var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId();

    var url_base = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export';
    var url_ext =
      '?format=pdf' + //export as pdf
      // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
      (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId)) +
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
    const pdfName = spreadsheet.getSheetById(sheetId).getName() + ' sheet for ' + studentName;
    const response = UrlFetchApp.fetch(url_base + url_ext, options);
    const blob = response.getBlob().setName(pdfName + '.pdf');
    const rootFolder = DriveApp.getRootFolder();
    const pdfSheet = rootFolder.createFile(blob);

    return pdfSheet.getId();
  }

  catch (err) {
    Logger.log(err.stack);
    throw new Error(err.message + '\n\n' + err.stack)
  }
}

function getStudentsSpreadsheetData(studentName) {
  const summarySheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('optSheetId')).getSheetByName('Summary');
  const lastFilledRow = getLastFilledRow(summarySheet, 1);
  const summaryData = summarySheet.getRange(1, 1, lastFilledRow, 26).getValues();
  const studentData = {
    'name': null,
    'hours': null,
    'recentSessionDate': null
  };

  for (let r = 0; r < lastFilledRow; r++) {
    if (summaryData[r][0] === studentName) {
      studentData.name = summaryData[r][0],
      studentData.hours = summaryData[r][3],
      studentData.recentSessionDate = Utilities.formatDate(new Date(summaryData[r][16]), 'GMT', 'EEE M/d');
      break;
    }
  }
  return studentData;
}