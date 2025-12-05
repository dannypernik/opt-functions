CLIENT_DATA_SS_ID = PropertiesService.getScriptProperties().getProperty('clientDataSsId');
const HOMEWORK_TEMPLATE_SS_ID = PropertiesService.getScriptProperties().getProperty('homeworkTemplateSsId');

let satSsIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
  rev: null,
  adminSsComplete: null,
  studentSsComplete: null,
  adminSheetsComplete: null,
  studentSheetsComplete: null
};

let actSsIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
  adminSsComplete: null,
  studentSsComplete: null,
  adminSheetsComplete: null,
  studentSheetsComplete: null
};

const dataLatestDate = TestPrepAnalysis.dataLatestDate;