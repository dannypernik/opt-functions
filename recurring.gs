function findNewScoreReports(students = null) {
  if (students.triggerUid) {
    const clientDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
    const clientSheet = clientDataSs.getSheetByName('Clients');
    const myStudentDataValue = clientSheet.getRange(2, 17).getValue();
    students = JSON.parse(myStudentDataValue);
  }
  
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

  findNewCompletedTests(fileList);
}

function findTeamScoreReports() {
  const studentDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const teamDataSheet = studentDataSs.getSheetByName('Team OPT');
  const teamDataValues = teamDataSheet.getRange(2,1,getLastFilledRow(teamDataSheet, 1) - 1, 17).getValues();

  for (let i = 0; i < teamDataValues.length; i ++) {
    const studentsStr = teamDataValues[i][16];
    const students = JSON.parse(studentsStr);
    findNewScoreReports(students);
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

  return isRunning;
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

    tutorData = {
      'index': tutorIndex,
      'name': tutorFolderName,
      'studentsFolderId': tutorFolderId
    }

    const students = updateStudentFolderData(tutorData, teamDataSheet);

    teamDataSheet.getRange(tutorIndex + 2,1,1,4).setValues([[tutorIndex, tutorFolderName, tutorFolderId, JSON.stringify(students)]]);
    tutorIndex ++;
  }
  
  const clientSheet = clientDataSs.getSheetByName('Clients')
  const myStudentFolderData = {
    'index': 0,
    'name': 'Open Path Tutoring',
    'studentsFolderId': clientSheet.getRange(2, 15).getValue()
  }

  const myStudents = updateStudentFolderData(myStudentFolderData, clientSheet);
  clientSheet.getRange(2, 17).setValue(JSON.stringify(myStudents));
  
}