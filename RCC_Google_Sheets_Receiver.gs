const SHEET_NAME = 'RCC Applications';

function doPost(e) {
  try {
    const payload = JSON.parse(e.parameter.payload || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    const headers = [
      'Timestamp','Full Name','Preferred Name','Date of Birth','Phone','Email','City/Area','Occupation',
      'Owns Vehicle','Brand','Model','Modification Type','Modification Details','Other Vehicles',
      'Automotive Interests','Join Reasons','Skills','Contribution','Participated in Events','Other Club',
      'Emergency Contact Name','Relationship','Emergency Phone','Facebook','Instagram','Other Social',
      'Consent'
    ];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }
    const row = [
      new Date(), payload.fullName || '', payload.preferredName || '', payload.dob || '',
      payload.phone || '', payload.email || '', payload.area || '', payload.occupation || '',
      payload.ownsVehicle || '', payload.brand || '', payload.model || '', payload.modType || '',
      payload.modDetails || '', payload.otherVehicles || '',
      Array.isArray(payload.interests) ? payload.interests.join(', ') : (payload.interests || ''),
      Array.isArray(payload.joinReason) ? payload.joinReason.join(', ') : (payload.joinReason || ''),
      payload.skills || '', payload.contribution || '', payload.events || '', payload.otherClub || '',
      payload.emergencyName || '', payload.relationship || '', payload.emergencyPhone || '',
      payload.facebook || '', payload.instagram || '', payload.social || '', payload.consent || ''
    ];
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({success:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('RCC Membership receiver is running.');
}