/* Node test suite for gantt-core.js — proves the shipped logic is correct.
 * Run: node gantt-core.test.js
 * No test framework needed; a tiny assert harness keeps it dependency-free. */
const C = require('./gantt-core.js');

let passed = 0;
let failed = 0;
function ok(name, cond) {
  if (cond) {
    passed++;
    console.log('  \u2713 ' + name);
  } else {
    failed++;
    console.log('  \u2717 ' + name + '   <-- FAILED');
  }
}
function eq(name, a, b) {
  ok(name + '  (got ' + JSON.stringify(a) + ')', JSON.stringify(a) === JSON.stringify(b));
}

console.log('\nCSV parsing');
{
  // Quoted commas, doubled-quote escapes, and a quoted newline inside a field.
  const csv = 'UID,Name,Note\r\nT1,"Auth, SSO","line1\nline2"\r\nT2,"He said ""hi""",ok\r\n';
  const t = C.parseTable(csv);
  eq('detects 3 headers', t.headers, ['UID', 'Name', 'Note']);
  eq('field with comma preserved', t.records[0].Name, 'Auth, SSO');
  eq('embedded newline preserved', t.records[0].Note, 'line1\nline2');
  eq('doubled quotes unescaped', t.records[1].Name, 'He said "hi"');
}
{
  // Semicolon-delimited export with a UTF-8 BOM (a common European locale case).
  const csv = '\uFEFFUID;Name;Priority\r\nT1;Login;High\r\n';
  const t = C.parseTable(csv);
  eq('auto-detects semicolon delimiter', t.delimiter, ';');
  eq('BOM stripped from first header', t.headers[0], 'UID');
  eq('semicolon record parsed', t.records[0].Name, 'Login');
}

console.log('\nCSV round-trip (parse -> serialize -> parse is stable)');
{
  const headers = ['UID', 'Name', 'Note'];
  const records = [
    { UID: 'T1', Name: 'Auth, SSO', Note: 'has "quotes" and, commas' },
    { UID: 'T2', Name: 'Line\nbreak', Note: '' },
  ];
  const out = C.toCSV(headers, records, ',');
  const back = C.parseTable(out);
  eq('round-trip preserves records', back.records, records);
}

console.log('\nWeek math');
{
  const start = new Date(2025, 0, 6); // Monday 6 Jan 2025
  eq('project start is week 0', C.weekIndex(new Date(2025, 0, 6), start), 0);
  eq('one week later is week 1', C.weekIndex(new Date(2025, 0, 13), start), 1);
  eq('mid-week snaps to same week', C.weekIndex(new Date(2025, 0, 9), start), 0);
  eq('weekToDate inverts weekIndex', C.weekToDate(3, start).getTime(), new Date(2025, 0, 27).getTime());
}

console.log('\nPriority coercion');
{
  eq('numeric priority', C.priorityToRank('2'), 2);
  eq('"High" ranks above "Low"', C.priorityToRank('High') < C.priorityToRank('Low'), true);
  eq('blank priority sinks to bottom', C.priorityToRank(''), 999);
}

console.log('\nTopological order respects precedence, breaks ties by rank');
{
  const tasks = [
    { uid: 'A', name: 'A', duration: 2, rank: 3, deps: [], type: 'task' },
    { uid: 'B', name: 'B', duration: 2, rank: 1, deps: ['A'], type: 'task' },
    { uid: 'C', name: 'C', duration: 2, rank: 2, deps: [], type: 'task' },
  ];
  const order = C.topoOrder(tasks).order;
  ok('B never precedes its prerequisite A', order.indexOf('A') < order.indexOf('B'));
  ok('among ready tasks, lower rank wins (C before A despite A rank 3)', order.indexOf('C') < order.indexOf('A'));
}
{
  const cyc = [
    { uid: 'X', name: 'X', duration: 1, rank: 1, deps: ['Y'], type: 'task' },
    { uid: 'Y', name: 'Y', duration: 1, rank: 2, deps: ['X'], type: 'task' },
  ];
  const r = C.topoOrder(cyc);
  ok('cycle is detected', r.hadCycle === true);
  eq('no task dropped on cycle', r.order.length, 2);
}

console.log('\nForward scheduler: precedence + capacity');
{
  // A(2) -> B(3); C(4) independent. With 1 lane everything serializes.
  const tasks = [
    { uid: 'A', name: 'A', duration: 2, rank: 1, deps: [], type: 'task' },
    { uid: 'B', name: 'B', duration: 3, rank: 2, deps: ['A'], type: 'task' },
    { uid: 'C', name: 'C', duration: 4, rank: 3, deps: [], type: 'task' },
  ];
  const s1 = C.scheduleForward(tasks, 1);
  ok('B starts only after A finishes', s1.B.startWeek >= s1.A.finishWeek);
  ok('with 1 lane, no two tasks overlap', noOverlap([s1.A, s1.B, s1.C]));

  const s2 = C.scheduleForward(tasks, 2);
  ok('B still waits for A even with 2 lanes', s2.B.startWeek >= s2.A.finishWeek);
  ok('with 2 lanes, A and C run in parallel (both start week 0)', s2.A.startWeek === 0 && s2.C.startWeek === 0);
}
{
  // Anchor (pinned start) is honoured as an earliest-start floor.
  const tasks = [
    { uid: 'A', name: 'A', duration: 2, rank: 1, deps: [], type: 'task', anchorStart: 5 },
  ];
  const s = C.scheduleForward(tasks, 2);
  eq('anchored task cannot start before its pin', s.A.startWeek, 5);
}
{
  // Milestone is zero-width and sits at the finish of its prerequisite.
  const tasks = [
    { uid: 'A', name: 'A', duration: 3, rank: 1, deps: [], type: 'task' },
    { uid: 'M', name: 'Release', duration: 0, rank: 2, deps: ['A'], type: 'milestone' },
  ];
  const s = C.scheduleForward(tasks, 2);
  eq('milestone has zero width', s.M.finishWeek - s.M.startWeek, 0);
  eq('milestone sits at prerequisite finish', s.M.startWeek, s.A.finishWeek);
}

console.log('\nFixed-date scheduler');
{
  const tasks = [{ uid: 'A', name: 'A', duration: 3, rank: 1, deps: [], type: 'task', completionWeek: 10 }];
  const s = C.scheduleFixed(tasks, 2);
  eq('bar ends on its completion week', s.A.finishWeek, 10);
  eq('bar starts duration-before completion', s.A.startWeek, 7);
}

console.log('\nValidation & auto-fix');
{
  const tasks = [
    { uid: 'A', name: 'Design', duration: 2, rank: 5, deps: [], type: 'task' },
    { uid: 'B', name: 'Build', duration: 2, rank: 1, deps: ['A'], type: 'task' },
    { uid: 'C', name: 'Ghost', duration: 1, rank: 2, deps: ['ZZ'], type: 'task' },
  ];
  const w = C.validate(tasks);
  ok('flags unknown dependency reference', w.some((x) => /unknown id/.test(x.message)));
  ok('flags rank-above-prerequisite conflict', w.some((x) => /ranked above its prerequisite/.test(x.message)));
  const fixed = C.autoFixRanks(tasks);
  ok('auto-fixed rank puts prerequisite A before dependent B', fixed['A'] < fixed['B']);
}

console.log('\nColor contrast & coercion helpers');
{
  ok('white text on a dark fill', C.readableText('#1E3A5F') === '#FFFFFF');
  ok('dark text on a light fill', C.readableText('#F4D35E') === '#16202B');
  ok('luminance of black < white', C.relLuminance('#000000') < C.relLuminance('#FFFFFF'));
  ok('hexToRgb parses shorthand', JSON.stringify(C.hexToRgb('#fff')) === JSON.stringify([255, 255, 255]));
  ok('parseBool TRUE', C.parseBool('TRUE') === true);
  ok('parseBool FALSE', C.parseBool('FALSE') === false);
  eq('usDate formats M/D/YYYY', C.usDate(new Date(2026, 8, 18)), '9/18/2026');
}

console.log('\nUploaded schema round-trip (#, deliveryDate M/D/YYYY, status, lane, dan-label)');
{
  const csv = '#,Name,priority,deliveryDate,FTE-weeks,status,lane,dan-label\r\n' +
    '7,name7,3,9/18/2026,3,Not Scoped,Requirements Engineering,TRUE\r\n';
  const t = C.parseTable(csv);
  eq('parses the hash id column header', t.headers[0], '#');
  eq('delivery date field intact', t.records[0].deliveryDate, '9/18/2026');
  eq('dan-label value intact', t.records[0]['dan-label'], 'TRUE');
  const d = C.parseDate(t.records[0].deliveryDate);
  ok('delivery date parses to a real date', d && d.getFullYear() === 2026 && d.getMonth() === 8);
}

/** True when no two [startWeek, finishWeek) intervals overlap. */
function noOverlap(bars) {
  const s = bars.slice().sort((a, b) => a.startWeek - b.startWeek);
  for (let i = 1; i < s.length; i++) if (s[i].startWeek < s[i - 1].finishWeek) return false;
  return true;
}

console.log('\n----------------------------------------');
console.log('  ' + passed + ' passed, ' + failed + ' failed');
console.log('----------------------------------------\n');
process.exit(failed ? 1 : 0);
