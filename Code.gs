const SHEET_ID = '1IIVGMkbANWYbWAsRXGBKUJuvkXxPUsECatdzRPauvtw';
const SHEET_NAME = 'Projects';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getProjects') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: getProjects()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if (action === 'saveProject') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: saveProject(data.project)
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'deleteProject') {
     return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: deleteProject(data.id)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function initializeSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME);
  const headers = ['id', 'name', 'contractType', 'contractor', 'supervisor', 'budget', 'po', 'disbursed', 'tor_date', 'announce_date', 'consider_status', 'appeal_status', 'wait_sign_status', 'signed_status', 'duration_days', 'notes_json'];
  sheet.appendRow(headers);
  sheet.getRange('1:1').setFontWeight('bold').setBackground('#1bb295').setFontColor('white');
  return sheet;
}

function initialSetup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) initializeSheet(ss);
}

function getProjects() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = initializeSheet(ss);
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); 
  if(data.length === 0) return [];
  
  return data.map(row => {
    let notes = {};
    try {
        notes = row[15] ? JSON.parse(String(row[15])) : {};
    } catch(e) {}
    
    return {
      id: Number(row[0]),
      name: String(row[1]),
      contractType: String(row[2]),
      contractor: String(row[3]),
      supervisor: String(row[4]),
      budget: Number(row[5]),
      po: Number(row[6]),
      disbursed: Number(row[7]),
      duration: Number(row[14]) || 0,
      dates: {
        tor: String(row[8]),
        announce: String(row[9]),
        consideration: String(row[10]),
        appeal: String(row[11]),
        waitsign: String(row[12]),
        signed: String(row[13])
      },
      notes: notes
    };
  });
}

function saveProject(project) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = initializeSheet(ss);
  
  const data = sheet.getDataRange().getValues();
  const isEdit = project.id && project.id !== '';
  let targetRowIndex = -1;
  let newId = isEdit ? Number(project.id) : 1;
  
  if (isEdit) {
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] == project.id) {
            targetRowIndex = i + 1;
            break;
        }
    }
  } else {
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
        if (Number(data[i][0]) > maxId) maxId = Number(data[i][0]);
    }
    newId = maxId + 1;
  }
  
  const rowData = [
    newId,
    project.name || '',
    project.contractType || '',
    project.contractor || '',
    project.supervisor || '',
    Number(project.budget) || 0,
    Number(project.po) || 0,
    Number(project.disbursed) || 0,
    project.dates.tor || '',
    project.dates.announce || '',
    project.dates.consideration || '',
    project.dates.appeal || '',
    project.dates.waitsign || '',
    project.dates.signed || '',
    Number(project.duration) || 0,
    JSON.stringify(project.notes || {})
  ];
  
  if (targetRowIndex > -1) {
    sheet.getRange(targetRowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { id: newId };
}

function deleteProject(id) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
     if (Number(data[i][0]) === Number(id)) {
        sheet.deleteRow(i + 1);
        return true;
     }
  }
  return false;
}
