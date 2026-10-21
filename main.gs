/**
 * CAPE SOLAR BOOKINGS â€” PROJECT LIGHTHOUSE
 * Consolidated Google Apps Script. Single-file edition.
 * Provider: SSESA only | Region: Western Cape | Currency: ZAR
 *
 * SAFETY: Email sending is deliberately disabled by two independent locks.
 * This script never sends prospect email while either lock remains enabled.
 * Setup adds missing sheets/headers only; it does not delete tabs or rows.
 */
var CSB_EMAIL_SAFETY_LOCK = true;
var CSB_CONFIG = {
  BUSINESS_NAME: 'Cape Solar Bookings',
  COMPANY_EMAIL: 'capesolarbookings@gmail.com',
  PROVIDER_NAME: 'SSESA',
  REGION: 'Western Cape',
  CURRENCY: 'ZAR',
  APPOINTMENT_VALUE: 250,
  COMMISSION_RATE: 0.035,
  DAILY_EMAIL_LIMIT: 30,
  FOLLOW_UP_DAYS: [3, 7, 14],
  MAX_FOLLOW_UPS: 3,
  TIMEZONE: 'Africa/Johannesburg',
  MIN_LEAD_SCORE: 40,
  APPROVAL_REQUIRED: true,
  AUTO_SEND_EMAILS: false,
  REPLY_LOOKBACK_DAYS: 30,
  BATCH_SIZE: 50,
  BACKUP_FOLDER_NAME: 'CSB Project Lighthouse Backups'
};

var CSB_SHEETS = {
  DASHBOARD: 'Dashboard', PROSPECTS: 'Prospects', LEAD_INTAKE: 'Lead Intake',
  COMPANIES: 'Companies', CONTACTS: 'Contacts', LEADS: 'Leads',
  APPOINTMENTS: 'Appointments', FOLLOW_UPS: 'Follow-Ups', SSESA: 'SSESA',
  COMMISSIONS: 'Commissions', TASKS: 'Tasks', ACTIVITY_LOG: 'Activity Log',
  EMAIL_HISTORY: 'Email History', SETTINGS: 'Settings', ERRORS: 'Automation Errors'
};

var CSB_HEADERS = {
  'Dashboard': ['Metric','Value','Updated'],
  'Prospects': ['Prospect ID','Company','Contact Name','Job Title','Email','Phone','Website','Industry','Region','Status','Lead Score','Opt Out','Approved','Source','Source URL','First Contact','Last Contact','Next Follow-Up','Follow-Up Count','Last Reply','Last Reply Message ID','Email Status','Notes','Created','Updated'],
  'Lead Intake': ['Intake ID','Company','Contact Name','Job Title','Email','Phone','Website','Industry','Region','Source','Source URL','Status','Processed','Import Result','Created','Updated'],
  'Companies': ['Company ID','Company','Industry','Region','Website','Address','Electricity Usage','Electricity Cost','Supply Type','SSEG Status','Phase','Roof Condition','Source','Notes','Created','Updated'],
  'Contacts': ['Contact ID','Company ID','Company','Contact Name','Job Title','Email','Phone','Preferred Contact Method','Opt Out','Notes','Created','Updated'],
  'Leads': ['Lead ID','Prospect ID','Company','Contact Name','Job Title','Email','Phone','Industry','Region','Lead Score','Status','Electricity Usage','Electricity Cost','Supply Type','SSEG Status','Phase','Roof Condition','Notes','Created','Updated'],
  'Appointments': ['Appointment ID','Lead ID','Company','Contact Name','Email','Phone','Appointment Date','Appointment Time','Appointment Status','Qualified','Value','Provider','Notes','Created','Updated'],
  'Follow-Ups': ['Follow-Up ID','Prospect ID','Company','Email','Due Date','Follow-Up Number','Status','Last Attempt','Result','Notes','Created','Updated'],
  'SSESA': ['SSESA Record ID','Lead ID','Company','Project Value','Project Status','Contract Signed Date','Paid Percentage','Amount Paid','Provider Commission Rate','Expected Commission','Commission Status','Notes','Created','Updated'],
  'Commissions': ['Commission ID','SSESA Record ID','Lead ID','Company','Project Value','Commission Rate','Commission Amount','Status','Invoice Date','Payment Due Date','Paid Date','Notes','Created','Updated'],
  'Tasks': ['Task ID','Task','Related Type','Related ID','Due Date','Priority','Status','Assigned To','Notes','Created','Updated'],
  'Activity Log': ['Timestamp','Level','Function','Message','Details'],
  'Email History': ['Email Log ID','Timestamp','Prospect ID','Company','Recipient','Subject','Email Type','Status','Gmail Message ID','Error','Notes'],
  'Settings': ['Setting','Value','Description','Updated'],
  'Automation Errors': ['Error ID','Timestamp','Function','Message','Stack','Resolved','Resolution Notes']
};

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('CSB Project Lighthouse')
      .addItem('1. Set Up / Repair System', 'CSB_setupSystemManual')
      .addItem('2. Run Health Check', 'CSB_systemHealthCheckManual')
      .addItem('3. Run Safe Tests', 'CSB_runSafeTestsManual')
      .addSeparator()
      .addItem('Process Lead Intake', 'CSB_processLeadIntakeManual')
      .addItem('Update Lead Scores', 'CSB_updateLeadScoresManual')
      .addItem('Approve Selected Prospect', 'CSB_approveSelectedProspect')
      .addItem('Create Lead from Selected Prospect', 'CSB_createLeadFromSelectedProspect')
      .addItem('Create Appointment from Selected Lead', 'CSB_createAppointmentFromSelectedLead')
      .addItem('Add SSESA Record from Selected Lead', 'CSB_addSSESARecordForSelectedLead')
      .addSeparator()
      .addItem('Refresh Dashboard', 'CSB_updateDashboardManual')
      .addItem('Create Backup', 'CSB_createBackupManual')
      .addItem('Create New Prospect', 'CSB_createNewProspectPrompt')
      .addSeparator()
      .addItem('Run Safe Automation (Email OFF)', 'CSB_runAutomationManual')
      .addToUi();
  } catch (e) { CSB_logError_('onOpen', e); }
}

function CSB_setupSystemManual() {
  var r = CSB_setupSystem();
  SpreadsheetApp.getUi().alert('CSB setup finished.\nSheets checked: ' + r.sheets + '\nHeaders added: ' + r.headersAdded + '\nExisting data was not deleted.\nAutomatic emails remain OFF.');
}
function CSB_setupSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Open this script from the CSB spreadsheet.');
  var made = 0, added = 0;
  Object.keys(CSB_HEADERS).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); made++; }
    added += CSB_ensureHeaders_(sh, CSB_HEADERS[name]);
    if (name === CSB_SHEETS.SETTINGS) CSB_setupSettings_();
  });
  CSB_log_('INFO', 'CSB_setupSystem', 'System setup/repair complete.', 'Created sheets: ' + made + '; headers added: ' + added);
  CSB_updateDashboard();
  return {sheets: Object.keys(CSB_HEADERS).length, created: made, headersAdded: added};
}

function CSB_setupSettings() { return CSB_setupSettings_(); }
function CSB_setupSettings_() {
  var sh = CSB_sheet_(CSB_SHEETS.SETTINGS);
  CSB_ensureHeaders_(sh, CSB_HEADERS.Settings);
  var existing = CSB_rows_(sh), seen = {};
  existing.forEach(function(r) { if (r.Setting) seen[String(r.Setting)] = true; });
  var defaults = [
    ['BUSINESS_NAME', CSB_CONFIG.BUSINESS_NAME, 'Business display name'],
    ['PROVIDER_NAME', CSB_CONFIG.PROVIDER_NAME, 'Only solar provider for CSB'],
    ['REGION', CSB_CONFIG.REGION, 'Target region'],
    ['APPOINTMENT_VALUE', CSB_CONFIG.APPOINTMENT_VALUE, 'ZAR value per qualified appointment'],
    ['COMMISSION_RATE', CSB_CONFIG.COMMISSION_RATE, 'Commission rate as decimal'],
    ['DAILY_EMAIL_LIMIT', CSB_CONFIG.DAILY_EMAIL_LIMIT, 'Maximum emails/day if sending is later authorised'],
    ['MIN_LEAD_SCORE', CSB_CONFIG.MIN_LEAD_SCORE, 'Minimum lead score'],
    ['APPROVAL_REQUIRED', CSB_CONFIG.APPROVAL_REQUIRED, 'Manual approval required before outreach'],
    ['AUTO_SEND_EMAILS', false, 'Safety default; do not enable before complete testing and explicit approval'],
    ['EMAIL_SAFETY_LOCK', true, 'Independent hard safety lock; keep true until authorised'],
    ['FOLLOW_UP_DAYS', CSB_CONFIG.FOLLOW_UP_DAYS.join(','), 'Days after initial contact'],
    ['MAX_FOLLOW_UPS', CSB_CONFIG.MAX_FOLLOW_UPS, 'Maximum follow-ups after initial contact'],
    ['TIMEZONE', CSB_CONFIG.TIMEZONE, 'Spreadsheet/business timezone'],
    ['BATCH_SIZE', CSB_CONFIG.BATCH_SIZE, 'Maximum rows processed per batch']
  ];
  var now = new Date();
  defaults.forEach(function(x) { if (!seen[x[0]]) sh.appendRow([x[0], x[1], x[2], now]); });
  return true;
}

function CSB_getSheetDefinitions() { return CSB_HEADERS; }
function CSB_emailSendingAllowed_() {
  return CSB_EMAIL_SAFETY_LOCK !== true && CSB_CONFIG.AUTO_SEND_EMAILS === true;
}

function runAutomation() {
  var results = [];
  var steps = [
    ['CSB_processLeadIntake', CSB_processLeadIntake],
    ['CSB_updateLeadScores', CSB_updateLeadScores],
    ['CSB_processFollowUps', CSB_processFollowUps],
    ['CSB_processEmailReplies', CSB_processEmailReplies],
    ['CSB_updateSSESAAndCommissions', CSB_updateSSESAAndCommissions],
    ['CSB_updateDashboard', CSB_updateDashboard]
  ];
  steps.forEach(function(s) {
    try { var out = s[1](); results.push(s[0] + ': OK' + (out && out.message ? ' â€” ' + out.message : '')); }
    catch (e) { CSB_logError_(s[0], e); results.push(s[0] + ': ERROR â€” ' + e.message); }
  });
  CSB_log_('INFO', 'runAutomation', 'Safe automation cycle completed; outreach sending skipped by safety lock.', results.join('\n'));
  return {message: results.join(' | '), results: results};
}
function CSB_runAutomationManual() {
  var r = runAutomation();
  SpreadsheetApp.getUi().alert('Safe automation finished.\n\n' + r.results.join('\n') + '\n\nProspect email sending is LOCKED OFF.');
}

function CSB_processLeadIntake() {
  var intake = CSB_sheet_(CSB_SHEETS.LEAD_INTAKE), prospects = CSB_sheet_(CSB_SHEETS.PROSPECTS);
  CSB_ensureHeaders_(intake, CSB_HEADERS['Lead Intake']); CSB_ensureHeaders_(prospects, CSB_HEADERS.Prospects);
  var rows = CSB_rows_(intake), existing = CSB_rows_(prospects), index = CSB_buildProspectIndex_(existing);
  var processed = 0, imported = 0, skipped = 0, errors = 0;
  rows.some(function(r) {
    if (processed >= CSB_CONFIG.BATCH_SIZE) return true;
    if (CSB_truthy_(r.Processed)) return false;
    if (!r.Company && !r.Email && !r['Contact Name']) return false;
    processed++;
    var result = '', status = 'Imported';
    try {
      var company = CSB_clean_(r.Company), email = CSB_clean_(r.Email).toLowerCase();
      if (!company) throw new Error('Company is required.');
      if (email && !CSB_validEmail_(email)) throw new Error('Email address is invalid.');
      var key = email ? 'email:' + email : 'company:' + CSB_norm_(company) + '|contact:' + CSB_norm_(r['Contact Name']);
      if (index[key]) { result = 'Skipped: duplicate of ' + index[key]; status = 'Duplicate'; skipped++; }
      else {
        var p = {
          'Prospect ID': CSB_id_('PRO'), 'Company': company, 'Contact Name': CSB_clean_(r['Contact Name']),
          'Job Title': CSB_clean_(r['Job Title']), 'Email': email, 'Phone': CSB_clean_(r.Phone),
          'Website': CSB_clean_(r.Website), 'Industry': CSB_clean_(r.Industry),
          'Region': CSB_clean_(r.Region) || CSB_CONFIG.REGION, 'Status': 'New', 'Lead Score': 0,
          'Opt Out': false, 'Approved': false, 'Source': CSB_clean_(r.Source) || 'Lead Intake',
          'Source URL': CSB_clean_(r['Source URL']), 'Email Status': 'Not sent - safety lock',
          'Notes': '', 'Created': new Date(), 'Updated': new Date()
        };
        CSB_append_(CSB_SHEETS.PROSPECTS, p); index[key] = p['Prospect ID']; imported++;
        result = 'Imported as ' + p['Prospect ID'];
      }
    } catch (e) { status = 'Error'; result = 'Error: ' + e.message; errors++; CSB_logError_('CSB_processLeadIntake', e); }
    CSB_setRow_(intake, r._rowIndex, {'Processed': true, 'Import Result': result, 'Status': status, 'Updated': new Date()});
    return false;
  });
  CSB_log_('INFO', 'CSB_processLeadIntake', 'Lead intake batch finished.', 'Checked: ' + processed + '; imported: ' + imported + '; duplicates: ' + skipped + '; errors: ' + errors);
  return {processed: processed, imported: imported, skipped: skipped, errors: errors, message: 'imported ' + imported + ', duplicates ' + skipped + ', errors ' + errors};
}
function CSB_processLeadIntakeManual() {
  var r = CSB_processLeadIntake();
  SpreadsheetApp.getUi().alert('Lead Intake result\nChecked: ' + r.processed + '\nImported: ' + r.imported + '\nDuplicates: ' + r.skipped + '\nErrors: ' + r.errors + '\n\nOpen Lead Intake â†’ Import Result for row-by-row details.');
}

function CSB_updateLeadScores() {
  var sh = CSB_sheet_(CSB_SHEETS.PROSPECTS); CSB_ensureHeaders_(sh, CSB_HEADERS.Prospects);
  var rows = CSB_rows_(sh), count = 0;
  rows.forEach(function(r) {
    if (!r['Prospect ID']) return;
    var score = 0, email = CSB_clean_(r.Email), phone = CSB_clean_(r.Phone), website = CSB_clean_(r.Website);
    if (email && CSB_validEmail_(email)) score += 20;
    if (phone) score += 10;
    if (website) score += 5;
    if (r['Contact Name']) score += 10;
    if (r['Job Title']) score += 10;
    if (r.Industry) score += 10;
    if (CSB_norm_(r.Region) === CSB_norm_(CSB_CONFIG.REGION)) score += 10;
    var industry = CSB_norm_(r.Industry);
    ['warehouse','logistics','cold storage','hotel','school','mall','shopping centre','infrastructure','manufacturing','commercial property'].forEach(function(term) {
      if (industry.indexOf(term) !== -1) score += 10;
    });
    if (CSB_truthy_(r['Opt Out'])) score = 0;
    score = Math.min(score, 100);
    CSB_setRow_(sh, r._rowIndex, {'Lead Score': score, 'Updated': new Date()}); count++;
  });
  CSB_log_('INFO', 'CSB_updateLeadScores', 'Lead scores updated.', 'Prospects scored: ' + count);
  return {count: count, message: count + ' prospects scored'};
}
function CSB_updateLeadScoresManual() { var r = CSB_updateLeadScores(); SpreadsheetApp.getUi().alert('Updated lead scores for ' + r.count + ' prospects.'); }

function CSB_processApprovedOutreach() {
  // Deliberately non-sending until both safety gates are explicitly changed after testing.
  if (!CSB_emailSendingAllowed_()) {
    CSB_logEmail_('', '', '', '', 'Initial outreach', 'BLOCKED', '', 'Safety lock active', 'No email was sent.');
    return {sent: 0, message: 'Blocked by CSB_EMAIL_SAFETY_LOCK / AUTO_SEND_EMAILS.'};
  }
  // This branch is intentionally fail-closed. Enabling outreach requires a separately reviewed release.
  return {sent: 0, message: 'Outreach sending is not implemented in this safe release.'};
}
function CSB_processApprovedOutreachManual() { var r = CSB_processApprovedOutreach(); SpreadsheetApp.getUi().alert(r.message); }

function CSB_processFollowUps() {
  // Create/refresh due-follow-up queue only. Sending is never attempted by this function.
  var prospects = CSB_rows_(CSB_sheet_(CSB_SHEETS.PROSPECTS));
  var fsh = CSB_sheet_(CSB_SHEETS.FOLLOW_UPS); CSB_ensureHeaders_(fsh, CSB_HEADERS['Follow-Ups']);
  var followups = CSB_rows_(fsh), queued = {}, added = 0, now = new Date();
  followups.forEach(function(f) { if (f['Prospect ID'] && f['Follow-Up Number']) queued[String(f['Prospect ID']) + ':' + String(f['Follow-Up Number'])] = true; });
  prospects.forEach(function(p) {
    if (!p['Prospect ID'] || CSB_truthy_(p['Opt Out']) || !CSB_validEmail_(CSB_clean_(p.Email))) return;
    var status = CSB_norm_(p.Status), count = Number(p['Follow-Up Count'] || 0);
    if (['replied','won','lost','opt out','do not contact','unsubscribed'].indexOf(status) !== -1) return;
    if (!p['First Contact']) return;
    if (count >= CSB_CONFIG.MAX_FOLLOW_UPS) return;
    var day = CSB_CONFIG.FOLLOW_UP_DAYS[count]; if (day === undefined) return;
    var first = new Date(p['First Contact']); if (isNaN(first.getTime())) return;
    var due = new Date(first.getTime()); due.setDate(due.getDate() + Number(day));
    if (due > now) return;
    var key = String(p['Prospect ID']) + ':' + String(count + 1);
    if (queued[key]) return;
    CSB_append_(CSB_SHEETS.FOLLOW_UPS, {
      'Follow-Up ID': CSB_id_('FU'), 'Prospect ID': p['Prospect ID'], 'Company': p.Company,
      'Email': p.Email, 'Due Date': due, 'Follow-Up Number': count + 1,
      'Status': 'Due - manual review', 'Last Attempt': '', 'Result': 'Not sent - safety lock',
      'Notes': 'Review manually; no automatic email will be sent.', 'Created': now, 'Updated': now
    });
    queued[key] = true; added++;
  });
  CSB_log_('INFO', 'CSB_processFollowUps', 'Follow-up queue refreshed; no messages sent.', 'New due follow-ups: ' + added);
  return {added: added, message: added + ' follow-ups queued for manual review'};
}
function CSB_processFollowUpsManual() { var r = CSB_processFollowUps(); SpreadsheetApp.getUi().alert(r.message + '. No emails were sent.'); }

function CSB_processEmailReplies() {
  // Read-only inbox scan. Does not mark threads/messages read or change prospect statuses automatically.
  var lookback = new Date(); lookback.setDate(lookback.getDate() - CSB_CONFIG.REPLY_LOOKBACK_DAYS);
  var query = 'newer_than:' + CSB_CONFIG.REPLY_LOOKBACK_DAYS + 'd -from:me';
  var threads = GmailApp.search(query, 0, 50), found = 0;
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (msg.getDate() < lookback) return;
      var from = String(msg.getFrom() || '').toLowerCase();
      if (from.indexOf(CSB_CONFIG.COMPANY_EMAIL.toLowerCase()) !== -1) return;
      var email = CSB_extractEmail_(from); if (!email) return;
      var match = CSB_findProspectByEmail_(email); if (!match) return;
      var id = String(msg.getId());
      if (String(match['Last Reply Message ID'] || '') === id) return;
      CSB_setRow_(CSB_sheet_(CSB_SHEETS.PROSPECTS), match._rowIndex, {
        'Last Reply': msg.getDate(), 'Last Reply Message ID': id, 'Status': 'Replied',
        'Email Status': 'Reply detected - review manually', 'Updated': new Date()
      });
      CSB_log_('INFO', 'CSB_processEmailReplies', 'Reply detected for prospect.', match['Prospect ID'] + ' / ' + email);
      found++;
    });
  });
  return {found: found, message: found + ' new replies detected'};
}
function CSB_processEmailRepliesManual() { var r = CSB_processEmailReplies(); SpreadsheetApp.getUi().alert('Reply scan complete. ' + r.message + '. Messages were not marked read.'); }

function CSB_updateSSESAAndCommissions() {
  var ssesa = CSB_rows_(CSB_sheet_(CSB_SHEETS.SSESA));
  var commissionsSh = CSB_sheet_(CSB_SHEETS.COMMISSIONS); CSB_ensureHeaders_(commissionsSh, CSB_HEADERS.Commissions);
  var commissions = CSB_rows_(commissionsSh), byRecord = {};
  commissions.forEach(function(c) { if (c['SSESA Record ID']) byRecord[String(c['SSESA Record ID'])] = c; });
  var made = 0, updated = 0;
  ssesa.forEach(function(s) {
    if (!s['SSESA Record ID']) return;
    var projectValue = Number(s['Project Value'] || 0), rate = s['Provider Commission Rate'] === '' || s['Provider Commission Rate'] == null ? CSB_CONFIG.COMMISSION_RATE : Number(s['Provider Commission Rate']);
    if (!isFinite(projectValue) || projectValue <= 0 || !isFinite(rate) || rate < 0) return;
    var expected = Math.round(projectValue * rate * 100) / 100;
    CSB_setRow_(CSB_sheet_(CSB_SHEETS.SSESA), s._rowIndex, {'Expected Commission': expected, 'Updated': new Date()});
    var c = byRecord[String(s['SSESA Record ID'])];
    if (c) {
      CSB_setRow_(commissionsSh, c._rowIndex, {'Project Value': projectValue, 'Commission Rate': rate, 'Commission Amount': expected, 'Updated': new Date()}); updated++;
    } else {
      CSB_append_(CSB_SHEETS.COMMISSIONS, {
        'Commission ID': CSB_id_('COM'), 'SSESA Record ID': s['SSESA Record ID'], 'Lead ID': s['Lead ID'],
        'Company': s.Company, 'Project Value': projectValue, 'Commission Rate': rate,
        'Commission Amount': expected, 'Status': 'Pending - verify contract terms', 'Invoice Date': '',
        'Payment Due Date': '', 'Paid Date': '', 'Notes': 'Calculated from SSESA record; verify signed agreement and payment milestone.',
        'Created': new Date(), 'Updated': new Date()
      }); made++;
    }
  });
  return {created: made, updated: updated, message: 'commissions created ' + made + ', updated ' + updated};
}
function CSB_updateSSESAAndCommissionsManual() { var r = CSB_updateSSESAAndCommissions(); SpreadsheetApp.getUi().alert(r.message); }

function CSB_updateDashboard() {
  var sh = CSB_sheet_(CSB_SHEETS.DASHBOARD);
  var metrics = [
    ['CSB DASHBOARD METRICS', 'Value', 'Updated'],
    ['Prospects', CSB_count_(CSB_SHEETS.PROSPECTS, 'Prospect ID'), new Date()],
    ['Approved Prospects', CSB_countWhere_(CSB_SHEETS.PROSPECTS, 'Approved', true), new Date()],
    ['Lead Intake - Unprocessed', CSB_countWhere_(CSB_SHEETS.LEAD_INTAKE, 'Processed', false, true), new Date()],
    ['Leads', CSB_count_(CSB_SHEETS.LEADS, 'Lead ID'), new Date()],
    ['Appointments', CSB_count_(CSB_SHEETS.APPOINTMENTS, 'Appointment ID'), new Date()],
    ['Due Follow-Ups', CSB_countStatus_(CSB_SHEETS.FOLLOW_UPS, 'Status', 'Due'), new Date()],
    ['SSESA Records', CSB_count_(CSB_SHEETS.SSESA, 'SSESA Record ID'), new Date()],
    ['Commissions', CSB_count_(CSB_SHEETS.COMMISSIONS, 'Commission ID'), new Date()],
    ['Automatic Email Sending', 'LOCKED OFF', new Date()],
    ['Solar Provider', CSB_CONFIG.PROVIDER_NAME, new Date()],
    ['Region', CSB_CONFIG.REGION, new Date()]
  ];
  // Safe write: update matching labels in column A, otherwise append below existing content.
  // Never clear a range or overwrite unrelated cells.
  var last = Math.max(sh.getLastRow(), 1), vals = sh.getRange(1, 1, last, 1).getDisplayValues(), rowByLabel = {};
  vals.forEach(function(v, i) { if (v[0]) rowByLabel[String(v[0])] = i + 1; });
  var next = last + 1;
  metrics.forEach(function(m, i) {
    var target = rowByLabel[m[0]];
    if (i === 0 && !target && last === 1 && !sh.getRange(1,1).getValue()) target = 1;
    if (!target) target = next++;
    sh.getRange(target, 1, 1, 3).setValues([m]);
  });
  return {message: 'Dashboard metrics refreshed without clearing existing cells.'};
}
function CSB_updateDashboardManual() { var r = CSB_updateDashboard(); SpreadsheetApp.getUi().alert(r.message); }

function CSB_approveSelectedProspect() {
  var sh = SpreadsheetApp.getActiveSheet(); if (sh.getName() !== CSB_SHEETS.PROSPECTS) throw new Error('Select a row on the Prospects sheet first.');
  var row = sh.getActiveRange().getRow(); if (row < 2) throw new Error('Select a prospect data row, not the header.');
  var p = CSB_rowObject_(sh, row); if (!p['Prospect ID']) throw new Error('Selected row has no Prospect ID.');
  if (CSB_truthy_(p['Opt Out'])) throw new Error('Cannot approve a prospect marked Opt Out.');
  CSB_setRow_(sh, row, {'Approved': true, 'Status': 'Approved - manual review', 'Updated': new Date()});
  SpreadsheetApp.getUi().alert('Prospect approved for workflow review. No email was sent; the safety lock remains active.');
}

function CSB_createLeadFromSelectedProspect() {
  var sh = SpreadsheetApp.getActiveSheet(); if (sh.getName() !== CSB_SHEETS.PROSPECTS) throw new Error('Select a row on Prospects first.');
  var row = sh.getActiveRange().getRow(); if (row < 2) throw new Error('Select a prospect data row.');
  var p = CSB_rowObject_(sh, row); if (!p['Prospect ID']) throw new Error('Selected row has no Prospect ID.');
  var existing = CSB_rows_(CSB_sheet_(CSB_SHEETS.LEADS));
  for (var i=0;i<existing.length;i++) if (String(existing[i]['Prospect ID']) === String(p['Prospect ID'])) { SpreadsheetApp.getUi().alert('A lead already exists for this prospect: ' + existing[i]['Lead ID']); return; }
  var lead = {'Lead ID': CSB_id_('LEAD'), 'Prospect ID': p['Prospect ID'], 'Company': p.Company, 'Contact Name': p['Contact Name'], 'Job Title': p['Job Title'], 'Email': p.Email, 'Phone': p.Phone, 'Industry': p.Industry, 'Region': p.Region, 'Lead Score': p['Lead Score'] || 0, 'Status': 'New - qualify details', 'Created': new Date(), 'Updated': new Date()};
  CSB_append_(CSB_SHEETS.LEADS, lead);
  CSB_setRow_(sh, row, {'Status': 'Lead Created', 'Updated': new Date()});
  SpreadsheetApp.getUi().alert('Lead created: ' + lead['Lead ID']);
}

function CSB_createAppointmentFromSelectedLead() {
  var sh = SpreadsheetApp.getActiveSheet(); if (sh.getName() !== CSB_SHEETS.LEADS) throw new Error('Select a row on Leads first.');
  var row = sh.getActiveRange().getRow(); if (row < 2) throw new Error('Select a lead data row.');
  var l = CSB_rowObject_(sh, row); if (!l['Lead ID']) throw new Error('Selected row has no Lead ID.');
  var existing = CSB_rows_(CSB_sheet_(CSB_SHEETS.APPOINTMENTS));
  for (var i=0;i<existing.length;i++) if (String(existing[i]['Lead ID']) === String(l['Lead ID']) && ['cancelled','canceled'].indexOf(CSB_norm_(existing[i]['Appointment Status'])) === -1) { SpreadsheetApp.getUi().alert('An appointment record already exists for this lead: ' + existing[i]['Appointment ID']); return; }
  var a = {'Appointment ID': CSB_id_('APT'), 'Lead ID': l['Lead ID'], 'Company': l.Company, 'Contact Name': l['Contact Name'], 'Email': l.Email, 'Phone': l.Phone, 'Appointment Status': 'Proposed - schedule manually', 'Qualified': false, 'Value': CSB_CONFIG.APPOINTMENT_VALUE, 'Provider': CSB_CONFIG.PROVIDER_NAME, 'Notes': 'Confirm qualification and appointment with prospect and SSESA manually.', 'Created': new Date(), 'Updated': new Date()};
  CSB_append_(CSB_SHEETS.APPOINTMENTS, a);
  SpreadsheetApp.getUi().alert('Appointment record created: ' + a['Appointment ID'] + '. Schedule and qualify it manually.');
}

function CSB_addSSESARecordForSelectedLead() {
  var sh = SpreadsheetApp.getActiveSheet(); if (sh.getName() !== CSB_SHEETS.LEADS) throw new Error('Select a row on Leads first.');
  var row = sh.getActiveRange().getRow(); if (row < 2) throw new Error('Select a lead data row.');
  var l = CSB_rowObject_(sh, row); if (!l['Lead ID']) throw new Error('Selected row has no Lead ID.');
  var existing = CSB_rows_(CSB_sheet_(CSB_SHEETS.SSESA));
  for (var i=0;i<existing.length;i++) if (String(existing[i]['Lead ID']) === String(l['Lead ID'])) { SpreadsheetApp.getUi().alert('SSESA record already exists: ' + existing[i]['SSESA Record ID']); return; }
  var s = {'SSESA Record ID': CSB_id_('SSESA'), 'Lead ID': l['Lead ID'], 'Company': l.Company, 'Project Value': '', 'Project Status': 'Referred - pending SSESA update', 'Paid Percentage': 0, 'Amount Paid': 0, 'Provider Commission Rate': CSB_CONFIG.COMMISSION_RATE, 'Expected Commission': 0, 'Commission Status': 'Pending project value and agreement verification', 'Notes': 'Update only from verified SSESA/project documentation.', 'Created': new Date(), 'Updated': new Date()};
  CSB_append_(CSB_SHEETS.SSESA, s);
  SpreadsheetApp.getUi().alert('SSESA record created: ' + s['SSESA Record ID']);
}

function CSB_createNewProspectPrompt() {
  var ui = SpreadsheetApp.getUi();
  var company = ui.prompt('New Prospect', 'Company name (required):', ui.ButtonSet.OK_CANCEL); if (company.getSelectedButton() !== ui.Button.OK) return;
  var name = ui.prompt('New Prospect', 'Contact name (optional):', ui.ButtonSet.OK_CANCEL); if (name.getSelectedButton() !== ui.Button.OK) return;
  var email = ui.prompt('New Prospect', 'Email (optional; must be valid if supplied):', ui.ButtonSet.OK_CANCEL); if (email.getSelectedButton() !== ui.Button.OK) return;
  var c = CSB_clean_(company.getResponseText()), n = CSB_clean_(name.getResponseText()), e = CSB_clean_(email.getResponseText()).toLowerCase();
  if (!c) throw new Error('Company is required.'); if (e && !CSB_validEmail_(e)) throw new Error('Email address is invalid.');
  var existing = CSB_buildProspectIndex_(CSB_rows_(CSB_sheet_(CSB_SHEETS.PROSPECTS)));
  var key = e ? 'email:' + e : 'company:' + CSB_norm_(c) + '|contact:' + CSB_norm_(n);
  if (existing[key]) { ui.alert('Duplicate detected; existing Prospect ID: ' + existing[key]); return; }
  var p = {'Prospect ID': CSB_id_('PRO'), 'Company': c, 'Contact Name': n, 'Email': e, 'Region': CSB_CONFIG.REGION, 'Status': 'New', 'Lead Score': 0, 'Opt Out': false, 'Approved': false, 'Source': 'Manual entry', 'Email Status': 'Not sent - safety lock', 'Created': new Date(), 'Updated': new Date()};
  CSB_append_(CSB_SHEETS.PROSPECTS, p); ui.alert('Prospect created: ' + p['Prospect ID'] + '. No email was sent.');
}

function CSB_createBackupManual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), folderName = CSB_CONFIG.BACKUP_FOLDER_NAME;
  var folders = DriveApp.getFoldersByName(folderName), folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  var stamp = Utilities.formatDate(new Date(), CSB_CONFIG.TIMEZONE, 'yyyy-MM-dd_HHmmss');
  var copy = DriveApp.getFileById(ss.getId()).makeCopy('CSB_Backup_' + stamp, folder);
  CSB_log_('INFO', 'CSB_createBackupManual', 'Spreadsheet backup created.', copy.getUrl());
  SpreadsheetApp.getUi().alert('Backup created: ' + copy.getName() + '\n' + copy.getUrl());
}

function CSB_systemHealthCheckManual() {
  var report = CSB_systemHealthCheck();
  SpreadsheetApp.getUi().alert('CSB Health Check\nStatus: ' + report.status + '\n\n' + report.checks.join('\n') + '\n\nAutomatic emails: LOCKED OFF');
}
function CSB_systemHealthCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), checks = [], fail = 0, warn = 0;
  Object.keys(CSB_HEADERS).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { checks.push('FAIL: Missing sheet ' + name); fail++; return; }
    var headers = sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0] : [];
    var missing = CSB_HEADERS[name].filter(function(h) { return headers.indexOf(h) === -1; });
    if (missing.length) { checks.push('FAIL: ' + name + ' missing headers: ' + missing.join(', ')); fail++; }
    else checks.push('PASS: ' + name + ' and required headers');
  });
  if (CSB_EMAIL_SAFETY_LOCK === true && CSB_CONFIG.AUTO_SEND_EMAILS === false) checks.push('PASS: Email safety lock enabled and automatic sending disabled.');
  else { checks.push('FAIL: Email safety configuration is not locked off.'); fail++; }
  if (CSB_CONFIG.PROVIDER_NAME !== 'SSESA') { checks.push('FAIL: Provider config differs from SSESA-only requirement.'); fail++; }
  else checks.push('PASS: SSESA is the configured provider.');
  if (CSB_CONFIG.REGION !== 'Western Cape') { checks.push('WARN: Region differs from expected Western Cape.'); warn++; }
  if (CSB_CONFIG.BATCH_SIZE < 1 || CSB_CONFIG.BATCH_SIZE > 500) { checks.push('FAIL: BATCH_SIZE outside safe range.'); fail++; }
  var report = {status: fail ? 'FAIL' : (warn ? 'WARN' : 'PASS'), checks: checks, failures: fail, warnings: warn};
  CSB_log_(fail ? 'ERROR' : 'INFO', 'CSB_systemHealthCheck', 'Health check result: ' + report.status, checks.join('\n'));
  return report;
}

// ---------- Shared utilities ----------
function CSB_sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); if (!ss) throw new Error('No active spreadsheet.');
  var sh = ss.getSheetByName(name); if (!sh) throw new Error('Missing sheet: ' + name + '. Run CSB_setupSystem first.');
  return sh;
}
function CSB_ensureHeaders_(sh, required) {
  if (sh.getLastRow() === 0) { sh.getRange(1,1,1,required.length).setValues([required]); sh.setFrozenRows(1); return required.length; }
  var lastCol = Math.max(sh.getLastColumn(), 1), current = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  var added = 0;
  required.forEach(function(h) { if (current.indexOf(h) === -1) { current.push(h); sh.getRange(1,current.length).setValue(h); added++; } });
  sh.setFrozenRows(1); return added;
}
function CSB_rows_(sh) {
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn(); if (lastRow < 2 || lastCol < 1) return [];
  var values = sh.getRange(1,1,lastRow,lastCol).getValues(), headers = values[0], out = [];
  for (var r=1;r<values.length;r++) {
    var obj = {_rowIndex:r+1};
    for (var c=0;c<headers.length;c++) if (headers[c]) obj[String(headers[c])] = values[r][c];
    out.push(obj);
  }
  return out;
}
function CSB_rowObject_(sh, row) {
  var lastCol = sh.getLastColumn(), headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0], values = sh.getRange(row,1,1,lastCol).getValues()[0], out = {_rowIndex:row};
  headers.forEach(function(h,i) { if (h) out[String(h)] = values[i]; }); return out;
}
function CSB_append_(sheetName, obj) {
  var sh = CSB_sheet_(sheetName), headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  if (!headers.length) throw new Error('Sheet ' + sheetName + ' has no headers. Run setup first.');
  var row = headers.map(function(h) { return Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''; });
  sh.appendRow(row); return sh.getLastRow();
}
function CSB_setRow_(sh, row, fields) {
  var lastCol = sh.getLastColumn(), headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  Object.keys(fields).forEach(function(k) { var i = headers.indexOf(k); if (i >= 0) sh.getRange(row,i+1).setValue(fields[k]); });
}
function CSB_buildProspectIndex_(rows) {
  var idx = {};
  rows.forEach(function(p) {
    if (p.Email && CSB_validEmail_(CSB_clean_(p.Email))) idx['email:' + CSB_clean_(p.Email).toLowerCase()] = p['Prospect ID'] || 'existing row ' + p._rowIndex;
    var company = CSB_norm_(p.Company), contact = CSB_norm_(p['Contact Name']);
    if (company) idx['company:' + company + '|contact:' + contact] = p['Prospect ID'] || 'existing row ' + p._rowIndex;
  }); return idx;
}
function CSB_findProspectByEmail_(email) {
  var rows = CSB_rows_(CSB_sheet_(CSB_SHEETS.PROSPECTS)), target = CSB_clean_(email).toLowerCase();
  for (var i=0;i<rows.length;i++) if (CSB_clean_(rows[i].Email).toLowerCase() === target) return rows[i];
  return null;
}
function CSB_count_(sheetName, key) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); if (!sh || sh.getLastRow() < 2) return 0;
  return CSB_rows_(sh).filter(function(r) { return r[key] !== '' && r[key] !== null && r[key] !== undefined; }).length;
}
function CSB_countWhere_(sheetName, key, expected, blankAsFalse) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); if (!sh || sh.getLastRow() < 2) return 0;
  return CSB_rows_(sh).filter(function(r) {
    if (blankAsFalse && (r[key] === '' || r[key] === null || r[key] === undefined)) return true;
    return String(r[key]).toLowerCase() === String(expected).toLowerCase() || (expected === true && CSB_truthy_(r[key]));
  }).length;
}
function CSB_countStatus_(sheetName, key, term) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); if (!sh || sh.getLastRow() < 2) return 0;
  return CSB_rows_(sh).filter(function(r) { return CSB_norm_(r[key]).indexOf(CSB_norm_(term)) === 0; }).length;
}
function CSB_validEmail_(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim()); }
function CSB_extractEmail_(from) { var m = String(from || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i); return m ? m[0].toLowerCase() : ''; }
function CSB_clean_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function CSB_norm_(v) { return CSB_clean_(v).toLowerCase().replace(/\s+/g,' ').trim(); }
function CSB_truthy_(v) { return v === true || String(v).toLowerCase() === 'true' || String(v).toLowerCase() === 'yes' || String(v) === '1'; }
function CSB_id_(prefix) { return prefix + '-' + Utilities.getUuid().replace(/-/g,'').substring(0,12).toUpperCase(); }
function CSB_log_(level, fn, message, details) {
  try { var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CSB_SHEETS.ACTIVITY_LOG); if (!sh) return; CSB_append_(CSB_SHEETS.ACTIVITY_LOG, {'Timestamp':new Date(),'Level':level,'Function':fn,'Message':message,'Details':details || ''}); } catch (e) { console.error('CSB log failure: ' + e.message); }
}
function CSB_logEmail_(prospectId, company, recipient, subject, type, status, messageId, error, notes) {
  try { var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CSB_SHEETS.EMAIL_HISTORY); if (!sh) return; CSB_append_(CSB_SHEETS.EMAIL_HISTORY, {'Email Log ID':CSB_id_('MAIL'),'Timestamp':new Date(),'Prospect ID':prospectId,'Company':company,'Recipient':recipient,'Subject':subject,'Email Type':type,'Status':status,'Gmail Message ID':messageId,'Error':error,'Notes':notes}); } catch (e) { console.error('CSB email log failure: ' + e.message); }
}
function CSB_logError_(fn, err) {
  var msg = err && err.message ? err.message : String(err), stack = err && err.stack ? err.stack : '';
  try { var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CSB_SHEETS.ERRORS); if (sh) CSB_append_(CSB_SHEETS.ERRORS, {'Error ID':CSB_id_('ERR'),'Timestamp':new Date(),'Function':fn,'Message':msg,'Stack':stack,'Resolved':false,'Resolution Notes':''}); } catch (e) { console.error('CSB error logging failure: ' + e.message); }
  console.error(fn + ': ' + msg + '\n' + stack);
}

// ---------- Compatibility helpers retained for older menu items/triggers ----------
function CSB_sheet(name) { return CSB_sheet_(name); }
function CSB_ensureHeaders(sh, headers) { return CSB_ensureHeaders_(sh, headers); }
function CSB_rows(sh) { return CSB_rows_(sh); }
function CSB_append(sheetName, obj) { return CSB_append_(sheetName, obj); }
function CSB_set(sheetName, row, fields) { return CSB_setRow_(CSB_sheet_(sheetName), row, fields); }
function CSB_log(level, fn, message, details) { return CSB_log_(level, fn, message, details); }
function CSB_logEmail(prospectId, company, recipient, subject, type, status, messageId, error, notes) { return CSB_logEmail_(prospectId, company, recipient, subject, type, status, messageId, error, notes); }
function CSB_logError(fn, err) { return CSB_logError_(fn, err); }
function CSB_mergeTemplate(template, prospect) {
  // Only substitutes supplied values; it does not invent claims, savings, prior work, or incentives.
  var p = prospect || {}, out = String(template || '');
  var fields = {
    '{{BUSINESS_NAME}}': CSB_CONFIG.BUSINESS_NAME,
    '{{PROVIDER_NAME}}': CSB_CONFIG.PROVIDER_NAME,
    '{{COMPANY}}': CSB_clean_(p.Company),
    '{{CONTACT_NAME}}': CSB_clean_(p['Contact Name']) || 'there',
    '{{REGION}}': CSB_CONFIG.REGION,
    '{{COMPANY_EMAIL}}': CSB_CONFIG.COMPANY_EMAIL
  };
  Object.keys(fields).forEach(function(k) { out = out.split(k).join(fields[k]); });
  return out;
}

function CSB_runSafeTestsManual() {
  var report = CSB_runSafeTests();
  SpreadsheetApp.getUi().alert('CSB Safe Tests\nStatus: ' + report.status + '\n\n' + report.checks.join('\n') + '\n\nThese checks do not send email or create test records.');
}
function CSB_runSafeTests() {
  var checks = [], failures = 0;
  function check(name, passed, detail) {
    checks.push((passed ? 'PASS: ' : 'FAIL: ') + name + (detail ? ' â€” ' + detail : ''));
    if (!passed) failures++;
  }
  check('Email safety lock remains enabled', CSB_EMAIL_SAFETY_LOCK === true);
  check('Automatic emails remain disabled', CSB_CONFIG.AUTO_SEND_EMAILS === false);
  check('SSESA-only provider configuration', CSB_CONFIG.PROVIDER_NAME === 'SSESA', CSB_CONFIG.PROVIDER_NAME);
  check('Western Cape region configuration', CSB_CONFIG.REGION === 'Western Cape', CSB_CONFIG.REGION);
  check('Email validator accepts a valid sample', CSB_validEmail_('test@example.com'));
  check('Email validator rejects malformed sample', !CSB_validEmail_('test@@example.com'));
  check('Duplicate index catches normalized email', (function() { var x = CSB_buildProspectIndex_([{Email:'TEST@example.com','Prospect ID':'P1',Company:'Example','Contact Name':'A'}]); return x['email:test@example.com'] === 'P1'; })());
  check('Outreach gate blocks sending', CSB_emailSendingAllowed_() === false);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    check('Active spreadsheet available', !!ss);
    if (ss) {
      Object.keys(CSB_HEADERS).forEach(function(name) {
        var sh = ss.getSheetByName(name);
        check('Sheet exists: ' + name, !!sh);
        if (sh) {
          var h = sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0] : [];
          var missing = CSB_HEADERS[name].filter(function(k) { return h.indexOf(k) < 0; });
          check('Headers valid: ' + name, missing.length === 0, missing.length ? 'missing ' + missing.join(', ') : '');
        }
      });
    }
  } catch (e) { check('Spreadsheet inspection completed', false, e.message); }
  return {status: failures ? 'FAIL' : 'PASS', failures: failures, checks: checks};
}
