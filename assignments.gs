function checkAllNewAssignments() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('clientDataSsId'));
  const studentsStr = ss.getSheetByName('Clients').getRange(2, 17).getValue();
  const students = studentsStr ? JSON.parse(studentsStr) : [];

  for (s of students) {
    if (s.homeworkSsId) {
      s.homeworkSs = SpreadsheetApp.openById(s.homeworkSsId);
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
      s.homeworkSs = SpreadsheetApp.openById(s.homeworkSsId);
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
        s.homeworkSs = SpreadsheetApp.openById(s.homeworkSsId);
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
        s.homeworkSs = SpreadsheetApp.openById(s.homeworkSsId);
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
  },
  allStudentsDataStr='[]'
) {
  
  let ui, studentName, homeworkSsId; 

  if (!studentData.name) {
    ui = SpreadsheetApp.getUi();
    const prompt = ui.prompt('Student name', ui.ButtonSet.OK_CANCEL);

    if (prompt.getSelectedButton() === ui.Button.CANCEL) {
      return;
    }
    studentName = prompt.getResponseText();
  }

  let allStudentsData = JSON.parse(allStudentsDataStr);

  if (allStudentsData.length === 0) {
    const clientSheet = SpreadsheetApp.openById(CLIENT_DATA_SS_ID).getSheetByName('Clients');
    allStudentsDataStr = clientSheet.getRange(getRowByColSearch(clientSheet, 2, 'Open Path Tutoring'), 17).getValue();
    allStudentsData = JSON.parse(allStudentsDataStr);
  }

  studentData = allStudentsData.find((student) => String(student.name).toLowerCase() === studentName.toLowerCase());
  
  const adminFolder = DriveApp.getFolderById(studentData.folderId);
  let studentFolder = adminFolder;

  const adminSubfolders = adminFolder.getFolders();
  while (adminSubfolders.hasNext()) {
    const adminSubfolder = adminSubfolders.next();

    if (adminSubfolder.getName().includes(studentData.name)) {
      studentFolder = adminSubfolder;
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

  if (!homeworkSsId) {
    const homeworkTemplate = DriveApp.getFileById(HOMEWORK_TEMPLATE_SS_ID);
    const newHomeworkFile = homeworkTemplate.makeCopy().setName(`Homework - ${studentData.name}`);

    newHomeworkFile.moveTo(studentFolder);
    homeworkSsId = newHomeworkFile.getId();
  }

  studentData.homeworkSsId = homeworkSsId;
  const studentIndex = allStudentsData.findIndex((student) => student.name === studentData.name);
  if (studentIndex !== -1) {
    allStudentsData[studentIndex] = studentData;
  } //
  else {
    allStudentsData.push(studentData);
  }

  if (homeworkSsId) {
    const homeworkSs = SpreadsheetApp.openById(homeworkSsId);
    const homeworkInfoSheet = homeworkSs.getSheetByName('Info');

    if (studentData.satAdminSsId) {
      const satAdminSs = SpreadsheetApp.openById(studentData.satAdminSsId);
      const satStudentSs = SpreadsheetApp.openById(studentData.satStudentSsId);
      satAdminSs.getSheetByName('Rev sheet backend').getRange('U8').setValue(homeworkSsId);
      satStudentSs.getSheetByName('Question bank data').getRange('U8').setValue(homeworkSsId);
      homeworkInfoSheet.getRange('C19').setValue(studentData.satStudentSsId);
    }

    if (studentData.actAdminSsId) {
      const actAdminSs = SpreadsheetApp.openById(studentData.actAdminSsId);
      const actStudentSs = SpreadsheetApp.openById(studentData.actStudentSsId);
      actAdminSs.getSheetByName('Data').getRange('U2').setValue(homeworkSsId);
      actStudentSs.getSheetByName('Data').getRange('U2').setValue(homeworkSsId);
      homeworkInfoSheet.getRange('C20').setValue(studentData.actStudentSsId);
    }

    const studentNameSplit = studentData.name.split(' ', 2);
    const studentFirstName = studentNameSplit[0];
    const studentLastName = studentNameSplit[1] || '';
    homeworkInfoSheet.getRange('C4:D4').setValues([[studentFirstName, studentLastName]]);
  } //
  else {
    Logger.log('Homework file not found.');
  }

  allStudentsDataStr = JSON.stringify(allStudentsData);
  return allStudentsDataStr;
}
