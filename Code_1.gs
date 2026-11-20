/**
 * Afrobeats FC club app — Google Sheets backend
 * One web app for the whole club: phone sign-in, matchday sign-ups and teams,
 * season results and stats, Man of the Match, club notices and fixtures,
 * the private player register, and the end-of-season awards.
 *
 * After pasting a new version: save, run setUp() once, then
 * Deploy > Manage deployments > pencil > Version: New version > Deploy.
 */

// The Google Sheet this code saves into. Change it only if you move to a new Sheet.
const SHEET_ID = '1RNhvqW7520k0USLq72JLJuyE0zzcX8YCldeDlCY-g3c';
function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); }

/* ---------- awards ---------- */
const CATEGORY_IDS = ['goalkeeper', 'defender', 'midfielder', 'winger', 'striker', 'improved', 'teamlead', 'omg'];
const CATEGORY_NAMES = ['Best Goalkeeper', 'Best Defender', 'Best Midfielder', 'Best Winger', 'Best Striker',
  'Most Improved Player', 'Best Captain / Coach', 'ABFC OMG Moment 2026'];
const FIRST_PICK_COL = 4; // Votes: Submitted, Voter, Phone, then one column per award

/* ---------- tabs and their columns ---------- */
const T = {
  members:  { name: 'Members',   head: ['Name', 'Phone', 'Organiser', 'Ballot name'] },
  register: { name: 'Register',  head: ['Key', 'Member name', 'Nickname', 'Full name', 'Date of birth', 'Position', 'Contact name', 'Contact relation', 'Contact phone', 'Updated'] },
  signups:  { name: 'Signups',   head: ['Id', 'Game', 'Name', 'Key', 'Paid', 'Status', 'Added', 'Guest'] },
  teams:    { name: 'Teams',     head: ['Game', 'Name', 'Team', 'Signup'] },
  results:  { name: 'Games',     head: ['Id', 'Game', 'Orange', 'Blue', 'Recorded'] },
  stats:    { name: 'Stats',     head: ['Game', 'Name', 'Goals', 'Assists'] },
  motm:     { name: 'MOTM',      head: ['Game', 'Voter', 'Pick', 'At'] },
  club:     { name: 'Club',      head: ['Setting', 'Value'] },
  fixtures: { name: 'Fixtures',  head: ['Id', 'Date', 'Time', 'Title', 'Venue'] },
  apart:    { name: 'KeepApart', head: ['Player A', 'Player B'] },
  votes:    { name: 'Votes',     head: ['Submitted', 'Voter', 'Phone'].concat(CATEGORY_NAMES) },
  settings: { name: 'Settings',  head: ['Setting', 'Value', 'What it does'] },
  joins:    { name: 'JoinRequests', head: ['Id', 'Name', 'Phone', 'Requested', 'Status'] },
};
// football-data.org free competitions
const FOOTBALL_CODES = ['PL', 'CL', 'PD', 'SA', 'BL1', 'FL1', 'ELC', 'DED', 'PPL', 'BSA', 'EC', 'WC'];
const POSITIONS = ['GK', 'DEF', 'MID', 'ATT'];
const TEAM_NAMES = ['Orange', 'Blue'];
const SESSION_DAYS = 180;

/* ====================================================================
   Entry points
   ==================================================================== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.winners) return reply(showWinners_() ? { ok: true, published: true, ballots: ballotCount_(), awards: winners_() } : { ok: true, published: false });
  // Legacy awards page check (kept so old links keep working)
  const phone = phone_(p.check);
  return reply({ ok: true, voting: votingPhase_() === 'open' ? 'open' : 'closed', member: phone ? !!member_(phone) : null, voted: phone ? hasVoted_(phone) : false });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return reply({ ok: false, error: 'Bad request.' }); }
  if (!data.action) return reply(legacyVote_(data));
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    return reply(handle_(data));
  } catch (err) {
    return reply({ ok: false, error: err && err.userMessage ? err.userMessage : 'Something went wrong. Try again.' });
  } finally {
    lock.releaseLock();
  }
}

function fail_(msg) { const e = new Error(msg); e.userMessage = msg; throw e; }

function handle_(d) {
  if (d.action === 'signin') return signIn_(d.phone);
  if (d.action === 'requestJoin') return requestJoin_(d.name, d.phone);
  const me = session_(d.token);
  if (!me) return { ok: false, error: 'Please sign in again.', signin: true };
  const org = () => { if (!me.organiser) fail_('Only organisers can do that.'); };
  const open = () => { if (!me.organiser && !clubOpen_()) fail_("That part of the app isn't open yet."); };
  if (d.action === 'football') { open(); return Object.assign({ ok: true }, football_(d)); }
  if (['join', 'leave', 'motmVote', 'saveRegister'].indexOf(d.action) >= 0) open();
  switch (d.action) {
    case 'state': break;
    case 'join': join_(me); break;
    case 'leave': leave_(me); break;
    case 'saveRegister': saveRegister_(me, d.data || {}); break;
    case 'motmVote': motmVote_(me, d.pick); break;
    case 'awardsVote': awardsVote_(me, d.picks || {}); break;
    case 'addGuest': org(); addGuest_(d.name); break;
    case 'setPaid': org(); setPaid_(d.id, !!d.paid); break;
    case 'removeSignup': org(); removeSignup_(d.id); break;
    case 'pickTeams': org(); pickTeams_(); break;
    case 'clearTeams': org(); clearTeams_(); break;
    case 'newGame': org(); newGame_(d.game || {}); break;
    case 'recordResult': org(); recordResult_(d.result || {}); break;
    case 'setNotice': org(); setClub_('notice', clean_(d.text, 600)); break;
    case 'addFixture': org(); addFixture_(d.fixture || {}); break;
    case 'removeFixture': org(); removeFixture_(d.id); break;
    case 'approveJoin': org(); approveJoin_(d.id); break;
    case 'declineJoin': org(); declineJoin_(d.id); break;
    default: fail_('Unknown request.');
  }
  return { ok: true, state: state_(me) };
}

/* ====================================================================
   Sign-in: phone number checked against Members, then a signed token
   ==================================================================== */
function signIn_(raw) {
  const phone = phone_(raw);
  if (!phone) return { ok: false, error: 'Enter your full phone number, e.g. 07123 456789.' };
  const m = member_(phone);
  if (!m) return { ok: false, error: "This number isn't on the ABFC members list. Message the committee to be added." };
  const exp = Date.now() + SESSION_DAYS * 864e5;
  const body = m.key + '.' + exp;
  return { ok: true, token: body + '.' + sign_(body) };
}

function session_(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 3) return null;
  const body = parts[0] + '.' + parts[1];
  if (sign_(body) !== parts[2] || Number(parts[1]) < Date.now()) return null;
  const m = memberByKey_(parts[0]);
  if (!m) return null;
  const reg = registerFor_(m.key);
  m.display = (reg && reg.Nickname) || m.name;
  return m;
}

function sign_(body) {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('SESSION_SECRET');
  if (!secret) { secret = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SESSION_SECRET', secret); }
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body, secret)).replace(/=+$/, '');
}

let membersCache_ = null;
function members_() {
  if (membersCache_) return membersCache_;
  const list = read_(T.members).filter(r => key10_(r.Phone)).map(r => ({
    name: clean_(r.Name, 80) || 'Member', key: key10_(r.Phone), organiser: r.Organiser === 'TRUE',
    ballot: clean_(r['Ballot name'], 200),
  }));
  membersCache_ = list;
  return list;
}
// While the Members tab is empty, anyone can vote in the awards (handy for testing)
function member_(phone) {
  const k = key10_(phone), list = members_();
  if (!list.length) return { name: 'Guest', key: k, organiser: false };
  return list.find(m => m.key === k) || null;
}
function memberByKey_(k) { const list = members_(); return list.find(m => m.key === k) || (!list.length ? { name: 'Guest', key: k, organiser: false } : null); }

/* ====================================================================
   Snapshot for the page (no phone numbers leave the Sheet)
   ==================================================================== */
function state_(me) {
  const club = clubMap_();
  const game = club.nextDate || '';
  const cap = Number(club.capacity) || 36;
  const signups = read_(T.signups).filter(r => r.Game === game && r.Status !== 'dropped');
  const inList = signups.filter(r => r.Status === 'in');
  const wait = signups.filter(r => r.Status === 'wait');
  const pub = r => ({ id: r.Id, name: r.Name, paid: r.Paid === 'yes', guest: r.Guest === 'yes', mine: r.Key === me.key });
  const teamsRows = read_(T.teams).filter(r => r.Game === game);
  const positions = {};
  read_(T.register).forEach(r => { if (r.Position) positions[r.Key] = r.Position; });
  const sig = {}; signups.forEach(r => sig[r.Id] = r);

  const results = read_(T.results).sort((a, b) => String(b.Game).localeCompare(String(a.Game)) || Number(b.Recorded) - Number(a.Recorded));
  const statsRows = read_(T.stats);
  const latest = results[0] || null;

  const bdays = birthdaysToday_();
  const s = {
    me: { name: me.name, display: me.display, organiser: me.organiser, birthday: bdays.keys.indexOf(me.key) >= 0 },
    flags: { clubOpen: clubOpen_(), football: !!footballKey_() },
    birthdays: bdays.names,
    club: {
      notice: club.notice || '',
      game: game ? { date: game, time: club.nextTime || '', venue: club.nextVenue || '', fee: club.nextFee || '', capacity: cap } : null,
    },
    fixtures: read_(T.fixtures).map(r => ({ id: r.Id, date: r.Date, time: r.Time, title: r.Title, venue: r.Venue }))
      .sort((a, b) => String(a.date).localeCompare(String(b.date))),
    matchday: {
      players: inList.map(pub), waiting: wait.map(pub),
      myStatus: (signups.find(r => r.Key === me.key) || {}).Status || '',
    },
    teams: {
      picked: teamsRows.length > 0,
      list: TEAM_NAMES.map(t => ({ name: t, players: teamsRows.filter(r => r.Team === t).map(r => ({
        name: r.Name, position: (sig[r.Signup] && positions[sig[r.Signup].Key]) || '' })) })),
    },
    season: {
      results: results.map(r => ({
        id: r.Id, game: r.Game, orange: Number(r.Orange) || 0, blue: Number(r.Blue) || 0,
        scorers: statsRows.filter(x => x.Game === r.Game && Number(x.Goals) > 0).map(x => ({ name: x.Name, goals: Number(x.Goals) })),
      })),
      leaders: leaders_(results, statsRows),
      motm: latest ? motmState_(me, latest.Game) : null,
    },
    awards: awardsState_(me),
    register: { mine: registerPublic_(registerFor_(me.key), me) },
  };
  if (me.organiser) {
    s.register.all = members_().map(m => {
      const r = registerFor_(m.key);
      return { member: m.name, organiser: m.organiser, nickname: r ? r.Nickname : '', fullName: r ? r['Full name'] : '',
        dob: r ? r['Date of birth'] : '', position: r ? r.Position : '', contactName: r ? r['Contact name'] : '',
        contactRelation: r ? r['Contact relation'] : '', contactPhone: r ? r['Contact phone'] : '', done: !!r };
    }).sort((a, b) => a.member.localeCompare(b.member));
    s.joinRequests = read_(T.joins).filter(r => r.Status === 'pending').map(r => ({ id: r.Id, name: r.Name, phone: r.Phone, requested: Number(r.Requested) || 0 }));
  }
  return s;
}

/* ====================================================================
   Matchday
   ==================================================================== */
function currentGame_() {
  const g = clubMap_().nextDate;
  if (!g) fail_('No game is set up yet. An organiser needs to add the next game.');
  return g;
}
function join_(me) {
  const game = currentGame_();
  const rows = read_(T.signups).filter(r => r.Game === game);
  const mine = rows.find(r => r.Key === me.key && r.Status !== 'dropped');
  if (mine) return;
  const cap = Number(clubMap_().capacity) || 36;
  const taken = rows.filter(r => r.Status === 'in').length;
  append_(T.signups, { Id: id_(), Game: game, Name: me.display, Key: me.key, Paid: '', Status: taken < cap ? 'in' : 'wait', Added: String(Date.now()), Guest: '' });
}
function addGuest_(name) {
  name = clean_(name, 40);
  if (!name) fail_('Type the guest\'s name.');
  const game = currentGame_();
  const rows = read_(T.signups).filter(r => r.Game === game);
  const cap = Number(clubMap_().capacity) || 36;
  const taken = rows.filter(r => r.Status === 'in').length;
  append_(T.signups, { Id: id_(), Game: game, Name: name, Key: '', Paid: '', Status: taken < cap ? 'in' : 'wait', Added: String(Date.now()), Guest: 'yes' });
}
function leave_(me) {
  const game = currentGame_();
  const r = read_(T.signups).find(x => x.Game === game && x.Key === me.key && x.Status !== 'dropped');
  if (r) drop_(r);
}
function removeSignup_(id) {
  const r = read_(T.signups).find(x => x.Id === String(id) && x.Status !== 'dropped');
  if (r) drop_(r);
}
// Dropping a player frees the spot for the first person on the waiting list
function drop_(r) {
  const wasIn = r.Status === 'in';
  update_(T.signups, r._row, { Status: 'dropped' });
  removeFromTeams_(r);
  if (wasIn) {
    const next = read_(T.signups).filter(x => x.Game === r.Game && x.Status === 'wait').sort((a, b) => Number(a.Added) - Number(b.Added))[0];
    if (next) update_(T.signups, next._row, { Status: 'in' });
  }
}
function removeFromTeams_(r) {
  const t = read_(T.teams).filter(x => x.Game === r.Game && x.Signup === r.Id);
  t.sort((a, b) => b._row - a._row).forEach(x => deleteRow_(T.teams, x._row));
}
function setPaid_(id, paid) {
  const r = read_(T.signups).find(x => x.Id === String(id));
  if (r) update_(T.signups, r._row, { Paid: paid ? 'yes' : '' });
}

/* Random teams, balanced by position, respecting keep-apart pairs */
function pickTeams_() {
  const game = currentGame_();
  const players = read_(T.signups).filter(r => r.Game === game && r.Status === 'in');
  if (players.length < 2) fail_('At least 2 players need to be signed up first.');
  const pos = {};
  read_(T.register).forEach(r => { if (r.Position) pos[r.Key] = r.Position; });
  const groups = {};
  players.forEach(p => { const g = (p.Key && pos[p.Key]) || 'ANY'; (groups[g] = groups[g] || []).push(p); });
  const teams = [[], []];
  ['GK', 'DEF', 'MID', 'ATT', 'ANY'].forEach(g => {
    shuffle_(groups[g] || []).forEach(p => {
      const t = teams[0].length < teams[1].length ? 0 : teams[1].length < teams[0].length ? 1 : Math.floor(Math.random() * 2);
      teams[t].push(Object.assign({ pos: g }, p));
    });
  });
  // Keep-apart: if a pair ends up together, swap one of them with a same-position player from the other team
  const norm = s => clean_(s, 80).toLowerCase();
  read_(T.apart).forEach(rule => {
    const a = norm(rule['Player A']), b = norm(rule['Player B']);
    if (!a || !b) return;
    for (let t = 0; t < 2; t++) {
      const ia = teams[t].findIndex(p => norm(p.Name) === a), ib = teams[t].findIndex(p => norm(p.Name) === b);
      if (ia < 0 || ib < 0) continue;
      const other = 1 - t, mover = teams[t][ib];
      let j = teams[other].findIndex(p => p.pos === mover.pos && norm(p.Name) !== a);
      if (j < 0) j = teams[other].findIndex(p => norm(p.Name) !== a);
      if (j < 0) continue;
      const swap = teams[other][j];
      teams[other][j] = mover; teams[t][ib] = swap;
    }
  });
  clearTeams_();
  const rows = [];
  teams.forEach((list, t) => list.forEach(p => rows.push({ Game: game, Name: p.Name, Team: TEAM_NAMES[t], Signup: p.Id })));
  appendMany_(T.teams, rows);
}
function clearTeams_() {
  const game = clubMap_().nextDate;
  read_(T.teams).filter(r => r.Game === game).sort((a, b) => b._row - a._row).forEach(r => deleteRow_(T.teams, r._row));
}
function newGame_(g) {
  const date = String(g.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) fail_('Pick a date for the next game.');
  setClub_('nextDate', date);
  setClub_('nextTime', clean_(g.time, 20));
  setClub_('nextVenue', clean_(g.venue, 80));
  setClub_('nextFee', clean_(g.fee, 20));
  const cap = Math.max(2, Math.min(60, Number(g.capacity) || 36));
  setClub_('capacity', String(cap));
}

/* ====================================================================
   Season: results, stats, leaderboards, Man of the Match
   ==================================================================== */
function recordResult_(r) {
  const game = String(r.game || clubMap_().nextDate || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(game)) fail_('Pick the date of the game.');
  const orange = Math.max(0, Math.round(Number(r.orange) || 0)), blue = Math.max(0, Math.round(Number(r.blue) || 0));
  const existing = read_(T.results).find(x => x.Game === game);
  if (existing) update_(T.results, existing._row, { Orange: orange, Blue: blue, Recorded: String(Date.now()) });
  else append_(T.results, { Id: id_(), Game: game, Orange: orange, Blue: blue, Recorded: String(Date.now()) });
  read_(T.stats).filter(x => x.Game === game).sort((a, b) => b._row - a._row).forEach(x => deleteRow_(T.stats, x._row));
  const merged = {};
  (r.scorers || []).forEach(s => {
    const name = clean_(s.name, 60); if (!name) return;
    const k = name.toLowerCase();
    merged[k] = merged[k] || { Game: game, Name: name, Goals: 0, Assists: 0 };
    merged[k].Goals += Math.max(0, Math.round(Number(s.goals) || 0));
    merged[k].Assists += Math.max(0, Math.round(Number(s.assists) || 0));
  });
  appendMany_(T.stats, Object.values(merged).filter(x => x.Goals || x.Assists));
}
function leaders_(results, statsRows) {
  const played = {}; results.forEach(r => played[r.Game] = true);
  const add = (map, name, n) => { const k = name.toLowerCase(); map[k] = map[k] || { name: name, n: 0 }; map[k].n += n; };
  const goals = {}, assists = {}, apps = {};
  statsRows.filter(x => played[x.Game]).forEach(x => { add(goals, x.Name, Number(x.Goals) || 0); add(assists, x.Name, Number(x.Assists) || 0); });
  read_(T.teams).filter(x => played[x.Game]).forEach(x => add(apps, x.Name, 1));
  const top = map => Object.values(map).filter(x => x.n > 0).sort((a, b) => b.n - a.n || a.name.localeCompare(b.name)).slice(0, 10);
  return { goals: top(goals), assists: top(assists), apps: top(apps) };
}
function motmState_(me, game) {
  const eligible = read_(T.teams).filter(r => r.Game === game).map(r => r.Name).sort((a, b) => a.localeCompare(b));
  const votes = read_(T.motm).filter(r => r.Game === game);
  const mine = votes.find(v => v.Voter === me.key);
  const tally = {};
  votes.forEach(v => tally[v.Pick] = (tally[v.Pick] || 0) + 1);
  return {
    game: game, eligible: eligible, myVote: mine ? mine.Pick : '', votes: votes.length,
    tally: mine || me.organiser ? Object.keys(tally).map(k => ({ name: k, votes: tally[k] })).sort((a, b) => b.votes - a.votes || a.name.localeCompare(b.name)) : [],
  };
}
function motmVote_(me, pick) {
  const results = read_(T.results).sort((a, b) => String(b.Game).localeCompare(String(a.Game)));
  if (!results.length) fail_('There is no finished game to vote on yet.');
  const game = results[0].Game;
  pick = clean_(pick, 60);
  const eligible = read_(T.teams).filter(r => r.Game === game).map(r => r.Name);
  if (eligible.length && eligible.indexOf(pick) < 0) fail_('Pick a player who played in that game.');
  if (!pick) fail_('Pick a player.');
  if (read_(T.motm).some(v => v.Game === game && v.Voter === me.key)) fail_('You have already voted for this game.');
  append_(T.motm, { Game: game, Voter: me.key, Pick: pick, At: String(Date.now()) });
}

/* ====================================================================
   Club notices and fixtures
   ==================================================================== */
function clubMap_() {
  const m = {};
  read_(T.club).forEach(r => { if (r.Setting) m[r.Setting] = r.Value; });
  return m;
}
function setClub_(key, value) {
  const r = read_(T.club).find(x => x.Setting === key);
  if (r) update_(T.club, r._row, { Value: value });
  else append_(T.club, { Setting: key, Value: value });
}
function addFixture_(f) {
  const date = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) fail_('Pick a date for the fixture.');
  const title = clean_(f.title, 80);
  if (!title) fail_('Give the fixture a title, e.g. "vs Lagos Lions".');
  append_(T.fixtures, { Id: id_(), Date: date, Time: clean_(f.time, 20), Title: title, Venue: clean_(f.venue, 80) });
}
function removeFixture_(id) {
  const r = read_(T.fixtures).find(x => x.Id === String(id));
  if (r) deleteRow_(T.fixtures, r._row);
}

/* ====================================================================
   Player register (each member edits their own; organisers see all)
   ==================================================================== */
function registerFor_(key) { return read_(T.register).find(r => r.Key === key) || null; }
function registerPublic_(r, me) {
  return {
    nickname: r ? r.Nickname : '', fullName: r ? r['Full name'] : '', dob: r ? r['Date of birth'] : '', position: r ? r.Position : '',
    contactName: r ? r['Contact name'] : '', contactRelation: r ? r['Contact relation'] : '', contactPhone: r ? r['Contact phone'] : '',
    memberName: me.name,
  };
}
function saveRegister_(me, d) {
  const dob = String(d.dob || '').trim();
  if (dob && !/^\d{2}\/\d{2}\/\d{4}$/.test(dob)) fail_('Write the date of birth as DD/MM/YYYY.');
  const position = POSITIONS.indexOf(d.position) >= 0 ? d.position : '';
  const obj = {
    Key: me.key, 'Member name': me.name, Nickname: clean_(d.nickname, 30), 'Full name': clean_(d.fullName, 80), 'Date of birth': dob,
    Position: position, 'Contact name': clean_(d.contactName, 80), 'Contact relation': clean_(d.contactRelation, 40),
    'Contact phone': clean_(d.contactPhone, 30), Updated: String(Date.now()),
  };
  const r = registerFor_(me.key);
  if (r) update_(T.register, r._row, obj); else append_(T.register, obj);
  // Show the new name on this week's list and team sheet straight away
  const display = obj.Nickname || me.name, game = clubMap_().nextDate;
  read_(T.signups).filter(x => x.Game === game && x.Key === me.key && x.Name !== display).forEach(x => {
    update_(T.signups, x._row, { Name: display });
    read_(T.teams).filter(t => t.Signup === x.Id).forEach(t => update_(T.teams, t._row, { Name: display }));
  });
}

/* ====================================================================
   Joining by QR code: requests wait for an organiser to approve them
   ==================================================================== */
function requestJoin_(name, raw) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    name = clean_(name, 60);
    const phone = phone_(raw);
    if (name.length < 2) return { ok: false, error: 'Enter your name.' };
    if (!phone) return { ok: false, error: 'Enter your full phone number, e.g. 07123 456789.' };
    const k = key10_(phone);
    if (members_().some(m => m.key === k)) return { ok: false, error: "You're already a member. Sign in with this number instead.", member: true };
    const pending = read_(T.joins).filter(r => r.Status === 'pending');
    if (pending.some(r => key10_(r.Phone) === k)) return { ok: true, already: true };
    if (pending.length >= 300) return { ok: false, error: 'Too many requests are waiting. Please message the committee.' };
    append_(T.joins, { Id: id_(), Name: name, Phone: String(raw).trim().slice(0, 30), Requested: String(Date.now()), Status: 'pending' });
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function approveJoin_(id) {
  const r = read_(T.joins).find(x => x.Id === String(id) && x.Status === 'pending');
  if (!r) fail_('That request has already been handled.');
  if (!members_().some(m => m.key === key10_(r.Phone))) {
    const sh = sheet_(T.members);
    sh.appendRow(["'" + r.Name, "'" + r.Phone, false]);
    sh.getRange(sh.getLastRow(), 3).insertCheckboxes();
    delete readCache_[T.members.name]; membersCache_ = null;
  }
  update_(T.joins, r._row, { Status: 'approved' });
}
function declineJoin_(id) {
  const r = read_(T.joins).find(x => x.Id === String(id) && x.Status === 'pending');
  if (r) update_(T.joins, r._row, { Status: 'declined' });
}

/* ====================================================================
   Birthdays (names only; dates of birth never leave the Sheet)
   ==================================================================== */
function birthdaysToday_() {
  const today = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM');
  const names = [], keys = [];
  read_(T.register).forEach(r => {
    if (String(r['Date of birth']).slice(0, 5) !== today) return;
    const m = memberByKey_(r.Key); if (!m) return;
    names.push(r.Nickname || m.name); keys.push(r.Key);
  });
  return { names: names, keys: keys };
}

/* ====================================================================
   Football scores from football-data.org (free key in Settings!B5)
   ==================================================================== */
function footballKey_() { const sh = ss_().getSheetByName(T.settings.name); return sh ? String(sh.getRange('B5').getDisplayValue()).trim() : ''; }
function football_(d) {
  const key = footballKey_();
  if (!key) return { configured: false };
  const cache = CacheService.getScriptCache();
  if (d.kind === 'table') {
    const code = FOOTBALL_CODES.indexOf(d.code) >= 0 ? d.code : 'PL';
    const ck = 'tbl:' + code;
    let body = cache.get(ck);
    if (!body) {
      const res = UrlFetchApp.fetch('https://api.football-data.org/v4/competitions/' + code + '/standings', { headers: { 'X-Auth-Token': key }, muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) return { configured: true, error: footballError_(res.getResponseCode()) };
      const j = JSON.parse(res.getContentText());
      const total = (j.standings || []).find(t => t.type === 'TOTAL') || (j.standings || [])[0] || { table: [] };
      body = JSON.stringify({ code: code, name: (j.competition || {}).name || code, table: (total.table || []).map(t => ({
        pos: t.position, team: (t.team || {}).shortName || (t.team || {}).name, crest: (t.team || {}).crest, played: t.playedGames,
        won: t.won, draw: t.draw, lost: t.lost, gd: t.goalDifference, pts: t.points })) });
      cache.put(ck, body, 1800);
    }
    return { configured: true, table: JSON.parse(body) };
  }
  const re = /^\d{4}-\d{2}-\d{2}$/;
  const from = re.test(d.from) ? d.from : Utilities.formatDate(new Date(), 'Europe/London', 'yyyy-MM-dd');
  const to = re.test(d.to) ? d.to : from;
  const ck = 'fx:' + from + ':' + to;
  let body = cache.get(ck);
  if (!body) {
    const url = 'https://api.football-data.org/v4/matches?competitions=' + FOOTBALL_CODES.join(',') + '&dateFrom=' + from + '&dateTo=' + to;
    const res = UrlFetchApp.fetch(url, { headers: { 'X-Auth-Token': key }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      const stale = cache.get('stale:' + ck);
      if (stale) return { configured: true, stale: true, matches: JSON.parse(stale) };
      return { configured: true, error: footballError_(res.getResponseCode()) };
    }
    const j = JSON.parse(res.getContentText());
    const list = (j.matches || []).map(m => ({
      id: m.id, utc: m.utcDate, status: m.status, minute: m.minute || null,
      comp: (m.competition || {}).code, compName: (m.competition || {}).name, compEmblem: (m.competition || {}).emblem,
      home: (m.homeTeam || {}).shortName || (m.homeTeam || {}).name || 'TBC', homeCrest: (m.homeTeam || {}).crest,
      away: (m.awayTeam || {}).shortName || (m.awayTeam || {}).name || 'TBC', awayCrest: (m.awayTeam || {}).crest,
      hs: ((m.score || {}).fullTime || {}).home, as: ((m.score || {}).fullTime || {}).away,
    }));
    body = JSON.stringify(list);
    const live = list.some(m => m.status === 'IN_PLAY' || m.status === 'PAUSED');
    cache.put(ck, body, live ? 60 : 300);
    cache.put('stale:' + ck, body, 21600);
  }
  return { configured: true, matches: JSON.parse(body) };
}
function footballError_(code) {
  if (code === 400 || code === 403) return 'The football data key in the Settings tab isn\'t working. Check it was copied correctly.';
  if (code === 429) return 'Scores are busy right now. Try again in a minute.';
  return 'Football scores are unavailable right now.';
}
function clubOpen_() { const sh = ss_().getSheetByName(T.settings.name); return !!sh && sh.getRange('B4').getValue() === true; }

/* ====================================================================
   Awards
   ==================================================================== */
function awardsState_(me) {
  const published = showWinners_(), phase = votingPhase_(), w = votingWindow_();
  const s = { open: phase === 'open', phase: phase, opensAt: w.opens || null, closesAt: w.closes || null,
    voted: hasVoted_(me.key), published: published, ballots: published ? ballotCount_() : null, winners: published ? winners_() : null,
    selfNames: selfNames_(me) };
  if (me.organiser) {
    s.turnout = { ballots: ballotCount_(), members: members_().length };
    // Organisers see the results privately once voting has closed (for ordering the cups)
    if (!published && phase === 'closed') s.private = { ballots: ballotCount_(), awards: winners_() };
  }
  return s;
}
function votingClosedMessage_(phase) {
  if (phase === 'before') {
    const w = votingWindow_();
    return 'Voting opens ' + (w.opens ? Utilities.formatDate(new Date(w.opens), 'Europe/London', "EEE d MMM 'at' h:mma").replace('AM', 'am').replace('PM', 'pm') : 'soon') + '.';
  }
  return 'Voting has closed. Thanks for taking part!';
}
function awardsVote_(me, picks) {
  const phase = votingPhase_();
  if (phase !== 'open') fail_(votingClosedMessage_(phase));
  if (hasVoted_(me.key)) fail_('This phone number has already voted.');
  const self = selfNames_(memberByKey_(me.key) || me);
  const row = [new Date(), me.name, "'" + me.key];
  CATEGORY_IDS.forEach((id, i) => {
    const name = clean_(picks[id], 200);
    if (!name) fail_('Every award needs a pick.');
    if (id !== 'omg' && self.indexOf(norm_(name)) >= 0) fail_("You can't vote for yourself for " + CATEGORY_NAMES[i] + '. Pick someone else.');
    row.push(/^[=+\-@]/.test(name) ? "'" + name : name);
  });
  sheet_(T.votes).appendRow(row);
}
/* ---------- no voting for yourself ---------- */
function norm_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
// Every spelling that means "this member": their name, the parts in and out of brackets, and Members!D (Ballot name)
function selfNames_(m) {
  const out = [];
  const add = x => { const n = norm_(x); if (n.length >= 2 && out.indexOf(n) < 0) out.push(n); };
  const name = String((m && m.name) || '');
  if (name && name !== 'Guest') { add(name); name.split(/[()]/).forEach(add); }
  String((m && m.ballot) || '').split(/[,;\/]/).forEach(add);
  return out;
}
/* ---------- voting window (Settings B7 and B8, UK time) ---------- */
function parseUk_(v) {
  const t = String(v || '').trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T]+(\d{1,2}):(\d{2})/), y, mo, d, h, mi;
  if (m) { y = m[1]; mo = m[2]; d = m[3]; h = m[4]; mi = m[5]; }
  else { m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ ,]+(\d{1,2}):(\d{2})/); if (!m) return null; d = m[1]; mo = m[2]; y = m[3]; h = m[4]; mi = m[5]; }
  const two = x => ('0' + x).slice(-2);
  try { return Utilities.parseDate(y + '-' + two(mo) + '-' + two(d) + ' ' + two(h) + ':' + mi, 'Europe/London', 'yyyy-MM-dd HH:mm').getTime(); }
  catch (e) { return null; }
}
function votingWindow_() {
  const sh = ss_().getSheetByName(T.settings.name);
  if (!sh || sh.getRange('A7').getValue() !== 'Voting opens') return {};
  const v = sh.getRange('B7:B8').getDisplayValues();
  return { opens: parseUk_(v[0][0]), closes: parseUk_(v[1][0]) };
}
// 'before' | 'open' | 'closed'. The Voting open tickbox is a master switch on top of the dates.
function votingPhase_() {
  if (isClosed_()) return 'closed';
  const w = votingWindow_(), now = Date.now();
  if (w.opens && now < w.opens) return 'before';
  if (w.closes && now > w.closes) return 'closed';
  return 'open';
}
// Old awards page (no sign-in): name + phone in the request
function legacyVote_(data) {
  const phase = votingPhase_();
  if (phase !== 'open') return { ok: false, error: votingClosedMessage_(phase) };
  const phone = phone_(data.phone), m = phone && member_(phone);
  if (!phone) return { ok: false, error: 'Enter a full phone number to vote.' };
  if (!m) return { ok: false, error: "This number isn't on the ABFC members list. Message the committee to be added." };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { awardsVote_({ key: m.key, name: clean_(data.voter, 80) || m.name }, data.picks || {}); return { ok: true }; }
  catch (err) { return { ok: false, error: err.userMessage || 'Something went wrong saving your vote. Try again.' }; }
  finally { lock.releaseLock(); }
}
function hasVoted_(phoneOrKey) {
  const k = key10_(phoneOrKey), sh = sheet_(T.votes), last = sh.getLastRow();
  if (last < 2) return false;
  return sh.getRange(2, 3, last - 1, 1).getDisplayValues().some(r => key10_(r[0]) === k);
}
function isClosed_() { const sh = ss_().getSheetByName(T.settings.name); return !!sh && sh.getRange('B2').getValue() === false; }
function showWinners_() { const sh = ss_().getSheetByName(T.settings.name); return !!sh && sh.getRange('B3').getValue() === true; }
function ballotCount_() { return Math.max(0, sheet_(T.votes).getLastRow() - 1); }
function winners_() {
  const sh = sheet_(T.votes), last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, FIRST_PICK_COL, last - 1, CATEGORY_IDS.length).getDisplayValues() : [];
  return CATEGORY_IDS.map((id, c) => {
    const tally = {};
    rows.forEach(r => { const n = clean_(r[c], 200); if (!n) return; const k = n.toLowerCase(); (tally[k] = tally[k] || { name: n, votes: 0 }).votes++; });
    const ranked = Object.values(tally).sort((a, b) => b.votes - a.votes || a.name.localeCompare(b.name));
    const top = [];
    ranked.forEach(t => { const place = 1 + ranked.filter(x => x.votes > t.votes).length; if (place <= 3) top.push({ name: t.name, votes: t.votes, place: place }); });
    return { id: id, name: CATEGORY_NAMES[c], top: top };
  });
}

/* ====================================================================
   Run once from the editor after each update: creates any missing tabs
   ==================================================================== */
function setUp() {
  const ss = ss_();
  Object.keys(T).forEach(k => sheet_(T[k]));
  // Members: Organiser tickbox column
  const m = sheet_(T.members);
  m.getRange('B:B').setNumberFormat('@');
  if (m.getRange('C1').getValue() !== 'Organiser') m.getRange('C1').setValue('Organiser').setFontWeight('bold');
  if (m.getRange('D1').getValue() !== 'Ballot name') {
    m.getRange('D1').setValue('Ballot name').setFontWeight('bold')
      .setNote('How this member appears on the awards ballot, if different from their name (e.g. Dolapo). Separate several with commas. Members can\'t vote for themselves.');
  }
  if (m.getLastRow() > 1) {
    const c = m.getRange(2, 3, m.getLastRow() - 1, 1);
    const vals = c.getValues();
    c.insertCheckboxes();
    c.setValues(vals.map(v => [v[0] === true]));
  }
  // Settings rows for the awards
  const st = sheet_(T.settings);
  if (st.getRange('A2').getValue() !== 'Voting open') {
    st.getRange('A2:C2').setValues([['Voting open', true, 'Untick to stop new award ballots. Tick to reopen.']]);
    st.getRange('B2').insertCheckboxes().setValue(true);
  }
  if (st.getRange('A3').getValue() !== 'Show winners') {
    st.getRange('A3:C3').setValues([['Show winners', false, 'Tick to show the top 3 for each award in the app.']]);
    st.getRange('B3').insertCheckboxes().setValue(false);
  }
  if (st.getRange('A4').getValue() !== 'Show club features') {
    st.getRange('A4:C4').setValues([['Show club features', false, 'Tick to show Matchday, Teams, Season, Football, Club and My details to every member. Unticked: members only see Awards; organisers see everything.']]);
    st.getRange('B4').insertCheckboxes().setValue(false);
  }
  st.getRange('C2').setValue('Master switch. Untick to stop voting at any time. Even when ticked, voting only runs between the two dates below.');
  if (st.getRange('A7').getValue() !== 'Voting opens') {
    st.getRange('B7:B8').setNumberFormat('@');
    st.getRange('A7:C8').setValues([
      ['Voting opens', '2026-11-21 12:00', 'UK time, written as yyyy-mm-dd hh:mm (24-hour). Leave empty to open as soon as the tickbox is ticked.'],
      ['Voting closes', '2026-11-25 23:59', 'UK time, written as yyyy-mm-dd hh:mm (24-hour). After this, organisers see the results privately and the cup list is emailed.'],
    ]);
  }
  if (st.getRange('A9').getValue() !== 'Results email to') {
    st.getRange('A9:C9').setValues([['Results email to', Session.getEffectiveUser().getEmail(), 'Who gets the private winners and cup engraving list when voting closes. Separate several addresses with commas.']]);
  }
  if (st.getRange('A5').getValue() !== 'Football data key') {
    st.getRange('A5:C5').setValues([['Football data key', '', 'Free key from football-data.org (Register, then check your email). Paste it in the Value cell.']]);
  }
  // Club defaults
  if (!clubMap_().capacity) setClub_('capacity', '36');
  // Plain-text columns so dates and numbers stay as typed
  ['register', 'signups', 'teams', 'results', 'stats', 'motm', 'club', 'fixtures', 'joins'].forEach(k => {
    const sh = sheet_(T[k]); sh.getRange(1, 1, sh.getMaxRows(), T[k].head.length).setNumberFormat('@');
  });
}

// Run once: fills Members!D (Ballot name) where the ballot spelling differs from the member's name.
// Only fills empty cells, so it never overwrites your edits.
function fillBallotNames() {
  const map = {
    'Oluwaleke Oyefolu': 'Leke', 'Ajudua Gaskiya Paul, Esq': 'Gaskiya', 'Bishop7': 'Bishop', 'Bolaji Bernard': 'Bolaji',
    'Deena': 'Dolapo', 'Efe Obels': 'Efe', 'Ezekamba (Fred Bosah)': 'Fred', 'Fabulous Chimaobi': 'Chimaobi',
    'Jesse Ugwu': 'Jesse', 'Kun Ibrahim': 'Ibrahim', 'Mamuzo': 'Mamuzio', 'Vitalis Idachaba': 'Vitalis',
    'Ifeanyi Tupac Shakur': 'Shakur', 'Alemabari Aadum': 'Alema',
  };
  const sh = sheet_(T.members), last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 4).getValues();
  let n = 0;
  vals.forEach((r, i) => {
    const want = map[String(r[0]).trim()];
    if (want && !String(r[3]).trim()) { sh.getRange(i + 2, 4).setValue(want); n++; }
  });
  return n;
}

/* ====================================================================
   Sheet helpers
   ==================================================================== */
const sheetCache_ = {};
function sheet_(t) {
  if (sheetCache_[t.name]) return sheetCache_[t.name];
  const ss = ss_();
  let sh = ss.getSheetByName(t.name);
  if (!sh) sh = ss.insertSheet(t.name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sheetCache_[t.name] = sh;
  return sh;
}
const readCache_ = {};
function read_(t) {
  if (readCache_[t.name]) return readCache_[t.name];
  const sh = sheet_(t), last = sh.getLastRow();
  if (last < 2) return (readCache_[t.name] = []);
  const width = Math.max(t.head.length, sh.getLastColumn());
  const head = sh.getRange(1, 1, 1, width).getDisplayValues()[0];
  const vals = sh.getRange(2, 1, last - 1, width).getDisplayValues();
  const out = vals.map((v, i) => { const o = { _row: i + 2 }; head.forEach((h, j) => { if (h) o[h] = v[j]; }); return o; })
    .filter(o => head.some(h => h && String(o[h]).trim() !== ''));
  return (readCache_[t.name] = out);
}
function toCell_(v) { return typeof v === 'string' ? "'" + v : v; }
function rowFor_(t, obj) {
  const sh = sheet_(t);
  const head = sh.getRange(1, 1, 1, Math.max(t.head.length, sh.getLastColumn())).getDisplayValues()[0];
  return head.map(h => (h && obj[h] !== undefined) ? toCell_(obj[h]) : '');
}
function append_(t, obj) { sheet_(t).appendRow(rowFor_(t, obj)); delete readCache_[t.name]; }
function appendMany_(t, list) {
  if (!list.length) return;
  const sh = sheet_(t), rows = list.map(o => rowFor_(t, o));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  delete readCache_[t.name];
}
function update_(t, row, obj) {
  const sh = sheet_(t);
  const head = sh.getRange(1, 1, 1, Math.max(t.head.length, sh.getLastColumn())).getDisplayValues()[0];
  Object.keys(obj).forEach(k => { const c = head.indexOf(k); if (c >= 0) sh.getRange(row, c + 1).setValue(toCell_(obj[k])); });
  delete readCache_[t.name];
}
function deleteRow_(t, row) { sheet_(t).deleteRow(row); delete readCache_[t.name]; }

/* ---------- small helpers ---------- */
function id_() { return Utilities.getUuid().slice(0, 8); }
function shuffle_(a) { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; }
function clean_(s, max) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, max || 200); }
function phone_(raw) {
  let d = String(raw == null ? '' : raw).replace(/[^\d+]/g, '');
  if (d.charAt(0) === '+') d = d.slice(1);
  else if (d.slice(0, 2) === '00') d = d.slice(2);
  else if (d.charAt(0) === '0') d = '44' + d.slice(1);
  return /^\d{10,15}$/.test(d) ? d : '';
}
// Numbers match on their last 10 digits: +44 7… = 07…, +234 803… = 0803…
function key10_(raw) { const d = String(raw == null ? '' : raw).replace(/\D/g, ''); return d.length >= 10 ? d.slice(-10) : ''; }
function reply(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
