CLIENT_DATA_SS_ID = PropertiesService.getScriptProperties().getProperty('clientDataSsId');
const HOMEWORK_TEMPLATE_SS_ID = PropertiesService.getScriptProperties().getProperty('homeworkTemplateSsId');

const satSsIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
  rev: null,
};

const actSsIds = {
  admin: null,
  student: null,
  studentData: null,
  adminData: null,
};

const dataLatestDate = TestPrepAnalysis.dataLatestDate;