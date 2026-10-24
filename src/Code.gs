/**
 * Construction RFI & Submittal Tracker — Google Apps Script
 * Demo by Mario Herrera. All sample data is fictional.
 *
 * Menu: Tracker > Set up sheet | Send overdue reminders now | Enable daily reminders
 */
const CONFIG = {
  SHEET: 'Log',
  DASH: 'Dashboard',
  STATUSES: ['Open', 'Submitted', 'Under Review', 'Answered', 'Revise & Resubmit', 'Approved', 'Closed'],
  CLOSED: ['Answered', 'Approved', 'Closed'],
  TYPES: ['RFI', 'Submittal'],
  CC: '',                // optional: PM email copied on every reminder
  REMINDER_HOUR: 7,      // daily trigger hour (script time zone)
  HEADERS: ['ID', 'Type', 'Title', 'Spec Section', 'Owner', 'Owner Email',
            'Date Submitted', 'Due Date', 'Status', 'Notes', 'Days Overdue']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Tracker')
    .addItem('Set up sheet', 'setupTracker')
    .addItem('Send overdue reminders now', 'sendOverdueReminders')
    .addItem('Enable daily reminders', 'installDailyTrigger')
    .addToUi();
}

/** Builds headers, dropdowns, conditional formatting, and the dashboard. Safe to re-run. */
function setupTracker() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEET) || ss.insertSheet(CONFIG.SHEET);
  const n = CONFIG.HEADERS.length;
  sh.getRange(1, 1, 1, n).setValues([CONFIG.HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f4e78');
  sh.setFrozenRows(1);

  const rows = sh.getMaxRows() - 1;
  const list = v => SpreadsheetApp.newDataValidation().requireValueInList(v, true).setAllowInvalid(false).build();
  sh.getRange(2, 2, rows).setDataValidation(list(CONFIG.TYPES));
  sh.getRange(2, 9, rows).setDataValidation(list(CONFIG.STATUSES));
  sh.getRange(2, 7, rows, 2).setNumberFormat('yyyy-mm-dd')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  // Days Overdue formula
  sh.getRange('K1').setValue('Days Overdue');
  sh.getRange('K2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="",,IF(REGEXMATCH(I2:I,"^(Answered|Approved|Closed)$"),0,IF(TODAY()>H2:H,TODAY()-H2:H,0))))');

  const data = sh.getRange(2, 1, rows, n);
  const open = 'NOT(REGEXMATCH($I2,"^(Answered|Approved|Closed)$"))';
  const rule = (f, bg, fc) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(f).setBackground(bg).setFontColor(fc || null).setRanges([data]).build();
  sh.setConditionalFormatRules([
    rule('=AND($A2<>"",$H2<TODAY(),' + open + ')', '#f8d7da', '#842029'),                 // overdue
    rule('=AND($A2<>"",$H2>=TODAY(),$H2-TODAY()<=3,' + open + ')', '#fff3cd', '#664d03'), // due in 3 days
    rule('=AND($A2<>"",NOT(' + open + '))', '#e2efda', '#3b5d2a')                          // closed out
  ]);
  sh.autoResizeColumns(1, n);
  buildDashboard_();
}

function buildDashboard_() {
  const ss = SpreadsheetApp.getActive();
  const d = ss.getSheetByName(CONFIG.DASH) || ss.insertSheet(CONFIG.DASH, 0);
  d.clear();
  d.getCharts().forEach(c => d.removeChart(c));
  d.getRange('A1').setValue('RFI & Submittal Dashboard').setFontSize(16).setFontWeight('bold');
  d.getRange('A3:B3').setValues([['Status', 'Count']]).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#fff');
  CONFIG.STATUSES.forEach((s, i) => {
    d.getRange(4 + i, 1).setValue(s);
    d.getRange(4 + i, 2).setFormula(`=COUNTIF(${CONFIG.SHEET}!I:I,A${4 + i})`);
  });
  const r = 4 + CONFIG.STATUSES.length + 1;
  d.getRange(r, 1, 4, 2).setValues([
    ['Total items', `=COUNTA(${CONFIG.SHEET}!A2:A)`],
    ['Open items', `=B${r}-SUM(COUNTIF(${CONFIG.SHEET}!I:I,{"Answered","Approved","Closed"}))`],
    ['Overdue items', `=COUNTIF(${CONFIG.SHEET}!K2:K,">0")`],
    ['Avg days overdue', `=IFERROR(AVERAGEIF(${CONFIG.SHEET}!K2:K,">0"),0)`]
  ]).setFontWeight('bold');
  d.getRange(r + 2, 2).setFontColor('#c62828');
  d.insertChart(d.newChart().asColumnChart()
    .addRange(d.getRange(3, 1, CONFIG.STATUSES.length + 1, 2))
    .setPosition(3, 4, 0, 0).setOption('title', 'Items by Status').setOption('legend', { position: 'none' })
    .build());
  d.autoResizeColumns(1, 2);
}

/** Emails each owner a list of their overdue, still-open items. */
function sendOverdueReminders() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET);
  const rows = sh.getDataRange().getValues().slice(1);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const byOwner = {};
  rows.forEach(r => {
    const [id, type, title, , owner, email, , due, status] = r;
    if (!id || !email || !(due instanceof Date) || CONFIG.CLOSED.includes(status)) return;
    if (due >= today) return;
    const days = Math.round((today - due) / 864e5);
    (byOwner[email] = byOwner[email] || { owner, items: [] }).items.push(`• ${id} (${type}) — ${title} — due ${Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd')}, ${days} day(s) overdue [${status}]`);
  });
  let sent = 0;
  Object.keys(byOwner).forEach(email => {
    const { owner, items } = byOwner[email];
    const opts = CONFIG.CC ? { cc: CONFIG.CC } : {};
    MailApp.sendEmail(email, `[Tracker] ${items.length} overdue item(s)`,
      `Hi ${owner},\n\nThe following items assigned to you are past due:\n\n${items.join('\n')}\n\nSheet: ${SpreadsheetApp.getActive().getUrl()}\n`, opts);
    sent++;
  });
  return sent;
}

/** Installs one daily time-driven trigger for sendOverdueReminders. */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendOverdueReminders')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendOverdueReminders').timeBased().everyDays(1).atHour(CONFIG.REMINDER_HOUR).create();
}
