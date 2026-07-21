/* Headless smoke test (jsdom) for the Use Case Roadmap. */
const fs = require('fs');
const { JSDOM } = require('jsdom');
const html = fs.readFileSync('/mnt/user-data/outputs/capability-roadmap.html', 'utf8');
let capturedCSV = null;
const dom = new JSDOM(html, {
  runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost:8000/',
  beforeParse(win) {
    win.URL.createObjectURL = () => 'blob:stub'; win.URL.revokeObjectURL = () => {};
    const RealBlob = win.Blob;
    win.Blob = function (parts, opts) { capturedCSV = parts.join(''); return new RealBlob(parts, opts); };
    win.HTMLAnchorElement.prototype.click = function () {};
    win.fetch = () => Promise.reject(new Error('no server'));
    win.confirm = () => true; win.prompt = () => 'New Lane';
  },
});
const win = dom.window, doc = win.document;
let pass = 0, fail = 0;
function ok(n, c) { c ? (pass++, console.log('  \u2713 ' + n)) : (fail++, console.log('  \u2717 ' + n + '  <-- FAILED')); }
function run() {
  console.log('\nBoot & sample');
  ok('GanttCore present', typeof win.GanttCore === 'object');
  doc.getElementById('emptySample').click();
  ok('13 sample rows render', doc.querySelectorAll('#mount .row').length === 13);

  console.log('\nSwimlanes');
  const laneHdrs = doc.querySelectorAll('#mount .lanehdr');
  ok('three lanes render as swimlanes', laneHdrs.length === 3);
  const laneNames = Array.from(laneHdrs).map(h => h.dataset.lane).sort();
  ok('lane names come from data', laneNames.join(',') === 'Model Analysis,Model Making,Requirements Engineering');
  // render() rebuilds the DOM on each toggle, so re-query before each click.
  const clickLane0 = () => doc.querySelectorAll('#mount .lanehdr')[0].querySelector('.laneinfo').dispatchEvent(new win.Event('click', { bubbles: true }));
  clickLane0();
  ok('collapsing a lane hides its rows', doc.querySelectorAll('#mount .row').length < 13);
  clickLane0();
  ok('re-expanding restores all rows', doc.querySelectorAll('#mount .row').length === 13);

  console.log('\nStatus color + contrast + dan');
  const bar = doc.querySelector('#mount .bar');
  ok('bars are colored', /background/.test(bar.getAttribute('style')));
  ok('bars set explicit contrast text color', /color:/.test(bar.getAttribute('style')));
  ok('dan-labeled task gets gold ring', doc.querySelectorAll('#mount .bar.dan').length === 1);
  ok('dan star shows in row', doc.querySelectorAll('#mount .danstar').length >= 1);
  ok('legend renders status chips', doc.querySelectorAll('#legend .chip[data-status]').length === 3);

  console.log('\nFilters');
  const danBtn = doc.getElementById('danOnly');
  danBtn.dispatchEvent(new win.Event('click', { bubbles: true }));
  ok('dan-only filter -> 1 row', doc.querySelectorAll('#mount .row').length === 1);
  danBtn.dispatchEvent(new win.Event('click', { bubbles: true }));
  const chip = doc.querySelector('#legend .chip[data-status="Not Scoped"]');
  chip.dispatchEvent(new win.Event('click', { bubbles: true }));
  ok('hiding a status removes rows', doc.querySelectorAll('#mount .row').length < 13);
  doc.querySelector('#legend .chip[data-status="Not Scoped"]').dispatchEvent(new win.Event('click', { bubbles: true }));
  const s = doc.getElementById('search'); s.value = 'name7'; s.dispatchEvent(new win.Event('input'));
  ok('search narrows results', doc.querySelectorAll('#mount .row').length === 1);
  s.value = ''; s.dispatchEvent(new win.Event('input'));

  console.log('\nInline edit');
  const dur = doc.querySelector('#mount .durcell input'); dur.value = '12'; dur.dispatchEvent(new win.Event('change'));
  ok('editing FTE-weeks keeps board consistent', doc.querySelectorAll('#mount .row').length === 13);

  console.log('\nDrawer edit');
  doc.querySelector('#mount .more').dispatchEvent(new win.Event('click', { bubbles: true }));
  ok('drawer opens', doc.getElementById('drawer').classList.contains('show'));
  ok('status select present', !!doc.getElementById('dkStatus'));
  ok('lane select present', !!doc.getElementById('dkLane'));
  doc.getElementById('dkDan').click();
  ok('dan toggle adds a 2nd gold bar', doc.querySelectorAll('#mount .bar.dan').length === 2);
  doc.getElementById('dkClose').click();

  console.log('\nCanonical CSV round-trip');
  doc.getElementById('btnDownloadCanonical').click();
  ok('canonical CSV downloaded', typeof capturedCSV === 'string' && capturedCSV.length > 0);
  const back = win.GanttCore.parseTable(capturedCSV.replace(/^\ufeff/, ''));
  ok('exactly the 8 input columns', back.headers.join(',') === '#,Name,priority,deliveryDate,FTE-weeks,status,lane,dan-label');
  ok('13 rows', back.records.length === 13);
  ok('deliveryDate stays M/D/YYYY', /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(back.records[0].deliveryDate));
  ok('dan-label stays TRUE/FALSE', /^(TRUE|FALSE)$/.test(back.records[0]['dan-label']));

  console.log('\nAnalysis CSV');
  doc.getElementById('btnDownloadAnalysis').click();
  ok('analysis CSV adds ScheduledFinishDate', /ScheduledFinishDate/.test(capturedCSV));

  console.log('\nImport uploaded schema');
  const csv = '#,Name,priority,deliveryDate,FTE-weeks,status,lane,dan-label\n' +
    '1,alpha,2,9/18/2026,5,Not Scoped,Requirements Engineering,FALSE\n' +
    '2,beta,1,10/19/2026,3,Currently Scoped,Model Analysis,TRUE\n';
  doc.getElementById('btnImport').click();
  doc.querySelector('.tabsmini button[data-it="paste"]').click();
  doc.getElementById('pasteArea').value = csv;
  doc.getElementById('parsePaste').click();
  ok('import enabled after auto-map', doc.getElementById('importApply').disabled === false);
  doc.getElementById('importApply').click();
  ok('imported 2 use cases', doc.querySelectorAll('#mount .row').length === 2);
  ok('priority sort -> beta first', /beta/.test(doc.querySelector('#mount .nm').value));

  console.log('\nPresent view');
  doc.querySelector('.tab[data-view="present"]').click();
  ok('present board renders', doc.querySelectorAll('#pBoard .row').length === 2);
  ok('present legend renders', doc.querySelectorAll('#pLegend .chip').length >= 2);

  console.log('\n----------------------------------------');
  console.log('  ' + pass + ' passed, ' + fail + ' failed');
  console.log('----------------------------------------\n');
  process.exit(fail ? 1 : 0);
}
let tries = 0;
(function wait() {
  if (doc.getElementById('emptySample')) return run();
  if (++tries > 100) { console.error('boot timeout'); process.exit(1); }
  setTimeout(wait, 20);
})();
