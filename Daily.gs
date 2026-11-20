/**
 * Afrobeats FC daily jobs (every day at 8am UK time):
 *  1. Saves a dated copy of this Sheet in the Drive folder "ABFC Sheet backups" (keeps 30 days).
 *  2. Emails a morning club summary to the address in Settings!B6.
 *  3. When awards voting closes (Settings B8), emails the private winners and cup list to Settings B9.
 * Run installDaily() once to switch them on. Run it again any time; it never creates duplicates.
 */
const BACKUP_FOLDER = 'ABFC Sheet backups';
const BACKUP_KEEP_DAYS = 30;
const APP_URL = 'https://dinadolapo-ui.github.io/afrobeats-fc/';

function installDaily() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyJobs').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJobs').timeBased().everyDays(1).atHour(8).inTimezone('Europe/London').create();
  installVotingClose();
  const st = sheet_(T.settings);
  if (st.getRange('A6').getValue() !== 'Daily summary to') {
    st.getRange('A6:C6').setValues([['Daily summary to', Session.getEffectiveUser().getEmail(),
      'Who gets the 8am club summary email. Separate several addresses with commas. Leave empty to stop the email.']]);
  }
}

function dailyJobs() {
  try { backupSheet_(); } catch (e) { console.error('Backup failed: ' + e); }
  try { sendSummary_(); } catch (e) { console.error('Summary failed: ' + e); }
  try { installVotingClose(); sendResultsEmail_(); } catch (e) { console.error('Results email failed: ' + e); }
}

/* ---------- awards: private results email when voting closes ---------- */
// Sets a one-off trigger 2 minutes after the closing time in Settings B8. Safe to run again (e.g. after changing the date).
function installVotingClose() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'votingClosedJob').forEach(t => ScriptApp.deleteTrigger(t));
  const w = votingWindow_();
  if (w.closes && w.closes > Date.now()) ScriptApp.newTrigger('votingClosedJob').timeBased().at(new Date(w.closes + 2 * 60000)).create();
}
function votingClosedJob() { sendResultsEmail_(); }
// Sends once per closing time. Pass true to send a copy again (e.g. from the editor).
function sendResultsEmail_(force) {
  const st = sheet_(T.settings);
  const to = st.getRange('A9').getValue() === 'Results email to' ? String(st.getRange('B9').getDisplayValue()).trim() : '';
  const w = votingWindow_();
  if (!to || votingPhase_() !== 'closed') return false;
  const props = PropertiesService.getScriptProperties(), flag = 'resultsSent:' + (w.closes || 'manual');
  if (!force && props.getProperty(flag)) return false;
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const ballots = ballotCount_(), members = members_().length, awards = winners_();
  const cell = 'padding:9px 12px;border-top:1px solid #24453a;vertical-align:top;color:#eef4ef;font-size:14px';
  const engrave = awards.map(a => {
    const first = a.top.filter(t => t.place === 1);
    const who = !first.length ? '<i>No votes</i>' : first.length > 1
      ? '<b style="color:#ff9b8f">TIE:</b> ' + first.map(t => esc(t.name)).join(' &amp; ') + ' <span style="color:#93a99d">(' + first[0].votes + ' each, decide before ordering)</span>'
      : '<b>' + esc(first[0].name) + '</b> <span style="color:#93a99d">(' + first[0].votes + ' votes)</span>';
    return '<tr><td style="' + cell + '">' + esc(a.name) + '</td><td style="' + cell + '">' + who + '</td></tr>';
  }).join('');
  const podium = awards.map(a => '<tr><td style="' + cell + '"><b>' + esc(a.name) + '</b><br>' +
    (a.top.length ? a.top.map(t => ['', '🥇', '🥈', '🥉'][t.place] + ' ' + esc(t.name) + ' · ' + t.votes + (t.votes === 1 ? ' vote' : ' votes')).join('<br>') : '<i>No votes</i>') + '</td></tr>').join('');
  const box = (title, body) => '<tr><td style="padding:16px 18px 4px"><div style="font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#f2c14e;font-weight:700">' + title + '</div></td></tr><tr><td style="padding:0 6px 8px"><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%">' + body + '</table></td></tr>';
  const html = '<div style="background:#07110c;padding:24px 12px;font-family:Arial,Helvetica,sans-serif">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;margin:0 auto;background:#12281e;border-radius:14px;overflow:hidden">' +
    '<tr><td style="height:5px;background:linear-gradient(90deg,#1f9d55 0 30%,#f2c14e 30% 55%,#d94b3f 55% 80%,#000 80%)"></td></tr>' +
    '<tr><td style="padding:20px 18px 6px"><div style="font-size:13px;color:#ff9b8f;font-weight:700">PRIVATE · organisers only until awards night</div>' +
    '<div style="font-size:22px;font-weight:800;color:#fff;margin-top:6px">ABFC Awards 2026: voting has closed</div>' +
    '<div style="font-size:15px;color:#cfdcd4;margin-top:6px">' + ballots + ' of ' + members + ' members voted' + (members ? ' (' + Math.round(ballots / members * 100) + '%)' : '') + '.</div></td></tr>' +
    box('Cup engraving list', engrave) + box('Full podium', podium) +
    '<tr><td style="padding:16px 18px 20px;font-size:12px;color:#93a99d;line-height:1.5">Members can\'t see these results yet. On awards night, tick "Show winners" in the Settings tab to reveal them in the app. ' +
    'Write-in names are counted exactly as typed, so check the Votes tab for near-duplicate spellings before ordering.</td></tr>' +
    '</table></div>';
  MailApp.sendEmail({ to: to, subject: 'ABFC Awards 2026 · Winners for the cups (private)', htmlBody: html, name: 'Afrobeats FC app' });
  props.setProperty(flag, String(Date.now()));
  return true;
}

function backupSheet_() {
  const found = DriveApp.getFoldersByName(BACKUP_FOLDER);
  const folder = found.hasNext() ? found.next() : DriveApp.createFolder(BACKUP_FOLDER);
  const stamp = Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd');
  DriveApp.getFileById(SHEET_ID).makeCopy('ABFC club backup ' + stamp, folder);
  const cutoff = Date.now() - BACKUP_KEEP_DAYS * 864e5;
  const files = folder.getFiles();
  while (files.hasNext()) { const f = files.next(); if (f.getDateCreated().getTime() < cutoff) f.setTrashed(true); }
}

function sendSummary_() {
  const to = String(sheet_(T.settings).getRange('B6').getDisplayValue()).trim();
  if (!to) return;
  const today = Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd');
  const nice = d => Utilities.formatDate(new Date(d + 'T12:00:00Z'), 'Europe/London', 'EEEE d MMMM');
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  const joins = read_(T.joins).filter(r => r.Status === 'pending');
  const bdays = birthdaysToday_().names;
  const club = clubMap_(), game = club.nextDate || '';
  const cap = Number(club.capacity) || 36;
  const signups = game ? read_(T.signups).filter(r => r.Game === game && r.Status !== 'dropped') : [];
  const playing = signups.filter(r => r.Status === 'in'), waiting = signups.filter(r => r.Status === 'wait');
  const unpaid = playing.filter(r => r.Paid !== 'yes').length;
  const regDone = read_(T.register).length, members = members_().length;

  const rows = [];
  const add = (title, body) => rows.push('<tr><td style="padding:14px 18px;border-top:1px solid #24453a"><div style="font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#f2c14e;font-weight:700">' + title + '</div><div style="margin-top:6px;font-size:15px;line-height:1.5;color:#eef4ef">' + body + '</div></td></tr>');

  add('Join requests', joins.length
    ? '<b>' + joins.length + ' waiting for approval:</b> ' + joins.map(r => esc(r.Name)).join(', ') + '. Approve them in the app\'s Players tab.'
    : 'None waiting.');
  if (bdays.length) add('Birthdays today', 'Happy birthday to <b>' + bdays.map(esc).join('</b> and <b>') + '</b>! Give them a shout in the group.');
  const phase = votingPhase_(), vw = votingWindow_(), ballots = ballotCount_(), mCount = members_().length;
  const when = ms => Utilities.formatDate(new Date(ms), 'Europe/London', "EEE d MMM 'at' h:mma").replace('AM', 'am').replace('PM', 'pm');
  add('Awards 2026', phase === 'before'
    ? 'Voting opens <b>' + (vw.opens ? when(vw.opens) : 'soon') + '</b>.'
    : phase === 'open'
      ? '<b>' + ballots + ' of ' + mCount + ' members have voted</b>, ' + Math.max(0, mCount - ballots) + ' still to go.' + (vw.closes ? ' Voting closes ' + when(vw.closes) + '.' : '')
      : ballots + ' ballots in. Voting is <b>closed</b>' + (showWinners_() ? ' and the winners are showing in the app.' : '. Winners are hidden from members until you tick "Show winners".'));
  if (game && game >= today) {
    add('Next game', '<b>' + esc(nice(game)) + '</b>' + (club.nextTime ? ', ' + esc(club.nextTime) : '') + (club.nextVenue ? ' at ' + esc(club.nextVenue) : '') + '<br>' +
      playing.length + ' of ' + cap + ' spots taken' + (waiting.length ? ', ' + waiting.length + ' on the waiting list' : '') + (playing.length ? '. ' + unpaid + ' not marked as paid.' : '.'));
  } else {
    add('Next game', 'No upcoming game is set up.' + (clubOpen_() ? ' Add one in the app\'s Matchday tab.' : ''));
  }
  add('Members', members + ' members. ' + regDone + ' have filled in My details.' + (clubOpen_() ? '' : ' Club features are hidden from members (Settings: Show club features).'));

  const html = '<div style="background:#07110c;padding:24px 12px;font-family:Arial,Helvetica,sans-serif">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;margin:0 auto;background:#12281e;border-radius:14px;overflow:hidden">' +
    '<tr><td style="height:5px;background:linear-gradient(90deg,#1f9d55 0 30%,#f2c14e 30% 55%,#d94b3f 55% 80%,#000 80%)"></td></tr>' +
    '<tr><td style="padding:20px 18px 10px"><div style="font-size:13px;color:#93a99d">' + esc(nice(today)) + '</div>' +
    '<div style="font-size:22px;font-weight:800;color:#fff;margin-top:4px">Good morning from Afrobeats FC</div></td></tr>' +
    rows.join('') +
    '<tr><td style="padding:18px"><a href="' + APP_URL + '" style="display:inline-block;background:#f2c14e;color:#1b1404;text-decoration:none;font-weight:800;padding:12px 18px;border-radius:10px">Open the club app</a>' +
    '<div style="margin-top:14px;font-size:12px;color:#93a99d">Backup saved to "' + BACKUP_FOLDER + '" in your Google Drive. To change who gets this email, edit "Daily summary to" in the Settings tab.</div></td></tr>' +
    '</table></div>';
  const subject = 'Afrobeats FC · ' + (joins.length ? joins.length + ' join request' + (joins.length === 1 ? '' : 's') + ' · ' : '') + Utilities.formatDate(new Date(), 'Europe/London', 'EEE d MMM');
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: 'Afrobeats FC app' });
}
