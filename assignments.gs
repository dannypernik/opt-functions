function checkAllNewAssignments() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const studentsStr = ss.getSheetByName('Clients').getRange(2, 17).getValue();
  const students = studentsStr ? JSON.parse(studentsStr) : [];

  for (s of students) {
    if (s.homeworkSsId) {
      Homework.checkNewAssignments(s);
    }
  }
}

function checkAllDueAssignments() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const studentsStr = ss.getSheetByName('Clients').getRange(2, 17).getValue();
  const students = studentsStr ? JSON.parse(studentsStr) : [];

  for (s of students) {
    if (s.homeworkSsId) {
      Homework.checkDueTodayAssignments(s);
      Homework.checkPastDueAssignments(s);
    }
  }
}

function checkTeamNewAssignments() {
  const studentDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const teamDataSheet = studentDataSs.getSheetByName('Team OPT');
  const teamDataValues = teamDataSheet.getRange(2, 1, getLastFilledRow(teamDataSheet, 1) - 1, 4).getValues();

  for (let i = 0; i < teamDataValues.length; i++) {
    const studentsStr = teamDataValues[i][3];
    const students = studentsStr ? JSON.parse(studentsStr) : [];

    for (s of students) {
      if (s.homeworkSsId) {
        Homework.checkNewAssignments(s);
      }
    }
  }
}

function checkTeamDueAssignments() {
  const studentDataSs = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const teamDataSheet = studentDataSs.getSheetByName('Team OPT');
  const teamDataValues = teamDataSheet.getRange(2, 1, getLastFilledRow(teamDataSheet, 1) - 1, 4).getValues();

  for (let i = 0; i < teamDataValues.length; i++) {
    const studentsStr = teamDataValues[i][3];
    const students = studentsStr ? JSON.parse(studentsStr) : [];

    for (s of students) {
      if (s.homeworkSsId) {
        Homework.checkDueTodayAssignments(s);
        Homework.checkPastDueAssignments(s);
      }
    }
  }
}

function addHomeworkSs(
  studentData = {
    name: null,
    folderId: null,
    satAdminSsId: null,
    actAdminSsId: null,
    satStudentSsId: null,
    actStudentSsId: null,
  }
) {
  const adminFolder = DriveApp.getFolderById(studentData.folderId);
  const subfolders = adminFolder.getFolders();
  let studentFolder, homeworkSsId;

  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();

    if (subfolder.getName().includes(studentData.name)) {
      studentFolder = subfolder;
      break;
    }
  }

  const studentFiles = studentFolder.getFiles();
  while (studentFiles.hasNext()) {
    const studentFile = studentFiles.next();

    if (studentFile.getName() === `Homework - ${studentData.name}`) {
      homeworkSsId = studentFile.getId();
      break;
    }
  }

  const studentSubfolders = studentFolder.getFolders();
  while (studentSubfolders.hasNext() && !homeworkSsId) {
    const studentSubfolder = studentSubfolders.next();
    const studentSubfiles = studentSubfolder.getFiles();

    while (studentSubfiles.hasNext()) {
      const studentSubfile = studentSubfiles.next();

      if (studentSubfile.getName() === `Homework - ${studentData.name}`) {
        homeworkSsId = studentSubfile.getId();
        break;
      }
    }
  }

  if (homeworkSsId) {
    const satAdminSs = SpreadsheetApp.openById(studentData.satAdminSsId);
    const actAdminSs = SpreadsheetApp.openById(studentData.actAdminSsId);
    const satStudentSs = SpreadsheetApp.openById(studentData.satStudentSsId);
    const actStudentSs = SpreadsheetApp.openById(studentData.actStudentSsId);
    const homeworkSs = SpreadsheetApp.openById(homeworkSsId);
    const homeworkInfoSheet = homeworkSs.getSheetByName('Info');

    satAdminSs.getSheetByName('Rev sheet backend').getRange('U8').setValue(homeworkSsId);
    satStudentSs.getSheetByName('Question bank data').getRange('U8').setValue(homeworkSsId);

    actAdminSs.getSheetByName('Data').getRange('U2').setValue(homeworkSsId);
    actStudentSs.getSheetByName('Data').getRange('U2').setValue(homeworkSsId);

    
    homeworkInfoSheet.getRange('C16:C17').setValues([
      [studentData.satStudentSsId],
      [studentData.actStudentSsId]
    ]);

    const studentNameSplit = studentData.name.split(' ', 2)
    const studentFirstName = studentNameSplit[0];
    const studentLastName = studentNameSplit[1] || '';
    homeworkInfoSheet.getRange('C4:D4').setValues([[studentFirstName,studentLastName]])
  } //
  else {
    Logger.log('Homework file not found.');
    return;
  }
}
