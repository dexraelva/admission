const SPREADSHEET_ID = '16lZUCQlsZHJ5m-_CqpPoFtGDojLzmGBpIB4nfZouufE';
const SHEET_GID = 0;
const SCHOOL_NAME = 'AIMTECH CHRISTIAN COLLEGE';

const HEADERS = [
  'Timestamp',
  'Email Status',
  'Full Name',
  'Date of Birth',
  'Place of Birth',
  'Age',
  'Gender',
  'Height (cm)',
  'Weight (kg)',
  'Civil Status',
  'Spouse Name',
  'Citizenship',
  'TIN',
  'Cedula',
  'Provincial Address',
  'Present Address',
  'Mobile Number',
  'Email Address',
  'Primary School',
  'Primary Year',
  'Secondary School',
  'Secondary Year',
  'Tertiary School',
  'Tertiary Course',
  'Tertiary Year',
  'Vocational Program',
  'Vocational Year',
  'Other Education',
  'Honors / Awards',
  'Local Church Name',
  'Pastor Name',
  'Pastor Contact',
  'Years of Membership',
  'Church Address',
  'Program Applied For',
  'How Applicant Came to Know Jesus Christ',
  'Call to Ministry',
  'Date Accepted Jesus Christ',
  'Date Baptized',
  'Baptized By',
  'Church Affiliation',
  'Baptism Venue',
  'Denomination',
  'Family Members',
  'Work Experiences',
  'Lectures / Seminars',
  'Character References',
  'Primary ID',
  'Documentary Requirements',
  'Emergency Contact Name',
  'Emergency Relationship',
  'Emergency Contact Number',
  'Applicant Signature',
  'Application Date'
];

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const sheet = getTargetSheet_();
    ensureHeaders_(sheet);

    const p = e.parameter || {};
    const ps = e.parameters || {};

    let emailStatus = 'Not sent: no email address';
    if (p.email) {
      emailStatus = sendConfirmationEmail_(p);
    }

    sheet.appendRow([
      new Date(),
      emailStatus,
      value_(p.fullname),
      value_(p.dob),
      value_(p.placeOfBirth),
      value_(p.age),
      value_(p.gender),
      value_(p.height),
      value_(p.weight),
      value_(p.civilStatus),
      value_(p.spouseName),
      value_(p.citizenship),
      value_(p.tin),
      value_(p.cedula),
      value_(p.provincialAddress),
      value_(p.presentAddress),
      value_(p.mobile),
      value_(p.email),
      value_(p.primary_school),
      value_(p.primary_year),
      value_(p.secondary_school),
      value_(p.secondary_year),
      value_(p.tertiary_school),
      value_(p.tertiary_course),
      value_(p.tertiary_year),
      value_(p.vocational_program),
      value_(p.vocational_year),
      value_(p.education_others),
      value_(p.honors_awards),
      value_(p.churchName),
      value_(p.pastorName),
      value_(p.pastorContact),
      value_(p.membershipYears),
      value_(p.churchAddress),
      value_(p.program),
      value_(p.testimonyJesus),
      value_(p.callingDescription),
      value_(p.dateAccepted),
      value_(p.dateBaptized),
      value_(p.baptizedBy),
      value_(p.churchAffiliation),
      value_(p.baptismVenue),
      value_(p.denomination),
      combineRows_(ps, ['family_name[]', 'family_relation[]', 'family_age[]', 'family_occupation[]'], ['Name', 'Relationship', 'Age', 'Occupation']),
      combineRows_(ps, ['work_company[]', 'work_description[]', 'work_year[]'], ['Company', 'Description', 'Year']),
      combineRows_(ps, ['lecture_title[]', 'lecture_year[]'], ['Title', 'Year']),
      combineRows_(ps, ['ref_name[]', 'ref_contact[]', 'ref_relation[]'], ['Name', 'Contact', 'Relationship']),
      value_(p.primary_id),
      multi_(ps, 'requirements').join(', '),
      value_(p.emergencyName),
      value_(p.emergencyRelationship),
      value_(p.emergencyContact),
      value_(p.applicantSignature),
      value_(p.applicationDate)
    ]);

    return textResponse_('Success');
  } catch (error) {
    return textResponse_('Error: ' + error.message);
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      // Lock may not exist if waitLock failed.
    }
  }
}

function doGet() {
  return textResponse_('AIMTECH admission web app is running.');
}

function setupSheet() {
  const sheet = getTargetSheet_();
  ensureHeaders_(sheet);
}

function authorizeEmail() {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'AIMTECH admission email authorization test',
    body: 'If you received this message, MailApp is authorized for the admission form script.',
    name: SCHOOL_NAME
  });
}

function getTargetSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = spreadsheet.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) {
      return sheets[i];
    }
  }

  return sheets[0];
}

function ensureHeaders_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
}

function sendConfirmationEmail_(p) {
  try {
    const fullName = value_(p.fullname) || 'Applicant';
    const program = value_(p.program) || 'your selected program';
    const subject = 'Admission Application Received - ' + SCHOOL_NAME;
    const body = [
      'Dear ' + fullName + ',',
      '',
      'Thank you for submitting your admission application to ' + SCHOOL_NAME + '.',
      '',
      'Program applied for: ' + program,
      '',
      'We have received your application and our admissions team will review your information.',
      '',
      'Blessings,',
      SCHOOL_NAME
    ].join('\n');

    MailApp.sendEmail({
      to: p.email,
      subject: subject,
      body: body,
      name: SCHOOL_NAME
    });

    return 'Sent';
  } catch (error) {
    return 'Failed: ' + error.message;
  }
}

function value_(value) {
  return value || '';
}

function multi_(parameters, name) {
  const value = parameters[name];

  if (!value) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

function combineRows_(parameters, names, labels) {
  const columns = names.map(function(name) {
    return multi_(parameters, name);
  });
  const rowCount = Math.max.apply(null, columns.map(function(column) {
    return column.length;
  }).concat([0]));
  const rows = [];

  for (let i = 0; i < rowCount; i++) {
    const parts = [];

    for (let j = 0; j < names.length; j++) {
      if (columns[j][i]) {
        parts.push(labels[j] + ': ' + columns[j][i]);
      }
    }

    if (parts.length) {
      rows.push(parts.join(' | '));
    }
  }

  return rows.join('\n');
}

function textResponse_(message) {
  return ContentService.createTextOutput(message);
}
