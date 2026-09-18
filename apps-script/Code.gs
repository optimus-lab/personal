const SHEET_NAME = 'MOTO_SOS_NFC';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const headers = [
    'id','updatedAt','name','birth','phone','photoUrl',
    'blood','allergies','medications','conditions',
    'bike','bikeColor','plates','insurance',
    'contactName','contactPhone','instructions','publicMedical'
  ];
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return ContentService.createTextOutput(JSON.stringify({
    success:true,
    message:'Hoja preparada',
    sheet:SHEET_NAME
  })).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  if (action === 'get') return getProfile(e.parameter.id);
  return json({success:true, service:'MOTO SOS NFC', version:'3.0'});
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = (body.action || 'save').toLowerCase();
    if (action === 'save') return saveProfile(body.profile || {});
    return json({success:false,error:'Acción no válida'});
  } catch (err) {
    return json({success:false,error:String(err)});
  }
}

function saveProfile(p) {
  if (!p.id || !p.name || !p.blood || !p.contactName || !p.contactPhone) {
    return json({success:false,error:'Faltan campos obligatorios'});
  }

  const sh = getSheet();
  const rowValues = [
    p.id, new Date(), clean(p.name), clean(p.birth), clean(p.phone), clean(p.photoUrl),
    clean(p.blood), clean(p.allergies), clean(p.medications), clean(p.conditions),
    clean(p.bike), clean(p.bikeColor), clean(p.plates), clean(p.insurance),
    clean(p.contactName), clean(p.contactPhone), clean(p.instructions),
    p.publicMedical !== false
  ];

  const values = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.id)) { row = i + 1; break; }
  }

  if (row === -1) sh.appendRow(rowValues);
  else sh.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);

  return json({success:true,id:p.id,updatedAt:new Date().toISOString()});
}

function getProfile(id) {
  if (!id) return json({success:false,error:'ID requerido'});
  const sh = getSheet();
  const values = sh.getDataRange().getValues();
  const headers = values[0] || [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const p = {};
      headers.forEach((h,j) => p[h] = values[i][j]);
      // Never expose NSS: V3 deliberately does not store or publish it.
      p.updatedAt = p.updatedAt instanceof Date ? p.updatedAt.toISOString() : p.updatedAt;
      return json({success:true,profile:p});
    }
  }
  return json({success:false,error:'Perfil no encontrado'});
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow([
      'id','updatedAt','name','birth','phone','photoUrl','blood',
      'allergies','medications','conditions','bike','bikeColor',
      'plates','insurance','contactName','contactPhone','instructions','publicMedical'
    ]);
  }
  return sh;
}

function clean(v) {
  return v == null ? '' : String(v).trim();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}