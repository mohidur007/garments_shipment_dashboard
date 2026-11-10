// ---- Settings ----
var FILE = 'data.xlsx';      // Excel file name (same folder as index.html)
var SHEET = '';              // sheet name; leave empty to use the first sheet that has the schedule table
var REFRESH_SECONDS = 30;    // how often the page re-reads the Excel file

// ---- State ----
var KEYS = ['floor','brand','style','insp','ccd','sec'];
var DIMS = {floor:'Floor', brand:'Brand', insp:'Inspection', ccd:'CCD'};
var F = {floor:'', brand:'', style:'', insp:'', ccd:'', sec:''};   // current filters
var data = null, dim = 'floor', chart = null, lastSig = '', shownRows = [];
function $(id){ return document.getElementById(id); }
function fmt(n){ return Number(n).toLocaleString('en-US'); }
function esc(s){ return String(s).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function nat(a,b){ return a.localeCompare(b, undefined, {numeric:true}); }
function uniq(arr){ return Array.from(new Set(arr)).sort(nat); }
function sumQty(rows){ return rows.reduce(function(s,r){ return s + r.qty; }, 0); }

// ---- Status colors ----
function color(s){
  s = String(s).toLowerCase();
  if (/fail/.test(s)) return '#c62f2f';
  if (/done/.test(s)) return '#2e7d4f';
  if (/running/.test(s)) return '#d98e04';
  if (/offer/.test(s)) return '#7a4fb5';
  if (/due/.test(s)) return '#2f6fd1';
  return '#8a949e';
}

// ---- Read the Excel file, then repeat ----
function load(){
  var x = new XMLHttpRequest();
  x.open('GET', FILE + '?t=' + Date.now(), true);
  x.responseType = 'arraybuffer';
  x.onload = function(){
    if ((x.status === 0 || x.status === 200) && x.response && x.response.byteLength) parse(x.response, x.getResponseHeader('Last-Modified'));
    else fail();
  };
  x.onerror = fail;
  x.send();
}
function fail(){
  $('status').textContent = 'Cannot read "' + FILE + '". Check that the file is in this folder with that exact name. ' +
    'If it is, open this page with the special shortcut described in README.txt.';
}
function parse(buf, lm){
  try {
    var wb = XLSX.read(buf, {type:'array'}), m = null;
    var names = SHEET ? [SHEET] : wb.SheetNames;
    for (var i = 0; i < names.length && !m; i++) if (wb.Sheets[names[i]]) m = readSheet(wb.Sheets[names[i]]);
    if (!m) { $('status').textContent = 'Could not find the schedule table (a header row with BRAND and QTY) in "' + FILE + '".'; return; }
    var ft = savedTime(wb, lm);
    if (ft) m.updated = ft;
    var sig = JSON.stringify(m);
    if (sig !== lastSig) { lastSig = sig; data = m; build(); }
    $('today').textContent = new Date().toLocaleDateString('en-GB', {weekday:'long', day:'numeric', month:'long', year:'numeric'});
    $('status').textContent = 'Checked ' + new Date().toLocaleTimeString() + '. The page updates by itself every ' + REFRESH_SECONDS + ' seconds.' + (logoNote ? ' ' + logoNote : '');
  } catch (err) {
    $('status').textContent = 'Found "' + FILE + '" but could not read it. Check that it is a valid .xlsx file.';
  }
}

// The time the Excel file was last saved (Excel stores it inside the file each time you press Save)
function savedTime(wb, lm){
  var d = lm ? new Date(lm) : null;
  if (!d || isNaN(d)) { var p = wb.Props && wb.Props.ModifiedDate; d = p ? new Date(p) : null; }
  if (!d || isNaN(d)) return '';
  var t = d.toLocaleTimeString('en-US', {hour: 'numeric', minute: '2-digit'});
  return d.toDateString() === new Date().toDateString() ? t : d.toLocaleDateString('en-GB', {day: 'numeric', month: 'short'}) + ', ' + t;
}

// ---- Understand the sheet layout ----
function readSheet(ws){
  var a = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''}), h = -1, i;
  for (i = 0; i < a.length && i < 30; i++) {
    var t = a[i].map(function(c){ return String(c).trim().toUpperCase(); });
    if (t.indexOf('BRAND') > -1 && t.some(function(x){ return x.indexOf('QTY') === 0; })) { h = i; break; }
  }
  if (h < 0) return null;
  var heads = a[h].map(function(c){ return String(c).replace(/\s+/g, ' ').trim(); });
  function ix(re){ for (var k = 0; k < heads.length; k++) if (re.test(heads[k])) return k; return -1; }
  var c = {floor:ix(/^floor/i), brand:ix(/^brand/i), style:ix(/^style name/i), insp:ix(/^inspection/i),
           ccd:ix(/^ccd/i), qty:ix(/^qty/i), po:ix(/^po/i)};
  var shown = [];
  heads.forEach(function(x, k){ if (x) shown.push(k); });
  function g(r, k){ return k < 0 || r[k] == null ? '' : String(r[k]).trim(); }

  var rows = [], sec = 'Today', curCamp = '';
  for (i = h + 1; i < a.length; i++) {
    var r = a[i];
    var hasData = g(r, c.brand) || g(r, c.style) || g(r, c.po);
    if (r.some(function(x){ return /^balance$/i.test(String(x).trim()); })) { sec = 'Balance'; continue; }   // the BALANCE heading row
    if (!hasData && r.some(function(x){ return /balance/i.test(String(x)); })) { sec = 'Balance'; continue; }   // heading like "BALANCE SHIPMENT"
    // a new "CAMPUS-xx" heading in the middle of the sheet starts a separate campus block (it is NOT balance)
    if (!hasData && rows.length) {
      var ch = r.map(function(x){ return String(x).trim(); }).filter(function(x){ return /^campus/i.test(x); })[0];
      if (ch) { curCamp = ch; sec = 'Today'; continue; }
    }
    if (!hasData) continue;   // skips blank lines and the total row
    rows.push({
      sec: sec, camp: curCamp, floor: g(r, c.floor) || '(blank)', brand: g(r, c.brand) || '(blank)', style: g(r, c.style) || '(blank)',
      insp: g(r, c.insp) || 'No status', ccd: g(r, c.ccd) || '(blank)',
      qty: Number(g(r, c.qty).replace(/,/g, '')) || 0,
      cells: shown.map(function(k){ return g(r, k); })
    });
  }
  var company = '', campus = '', updated = '';
  for (i = 0; i < h; i++) a[i].forEach(function(x){
    x = String(x).trim();
    if (!x) return;
    if (/^campus/i.test(x)) campus = campus || x;
    else if (!company && !/shipment|schedule/i.test(x) && !/^\d{1,2}-[a-z]{3}/i.test(x)) company = x;
  });
  a.forEach(function(r){ r.forEach(function(x){
    x = String(x);
    if (/last\s*updated/i.test(x)) updated = x.replace(/^.*?last\s*updated\s*(on)?\s*/i, '').trim();
  }); });
  rows.forEach(function(x){ if (!x.camp) x.camp = campus; });
  return {company:company, campus:campus, updated:updated,
          heads: shown.map(function(k){ return heads[k]; }),
          qtyPos: shown.indexOf(c.qty), inspPos: shown.indexOf(c.insp), rows: rows};
}

// ---- Search: PO number / Style No (letters optional, partial text works) ----
var Q = '';
function sKey(s){ return String(s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function searchCols(){
  var h = data.heads;
  return {st: h.findIndex(function(x){ return /^style\s*no/i.test(x); }), po: h.findIndex(function(x){ return /^po\b/i.test(x); })};
}
function prepSearch(){
  var c = searchCols();
  data.rows.forEach(function(r){
    r.sk = sKey(c.po > -1 ? r.cells[c.po] : '') + ' ' + sKey(c.st > -1 ? r.cells[c.st] : '');
  });
}
function matchQ(r){
  var terms = Q.split(/[,;\n]/).map(sKey).filter(Boolean);
  if (!terms.length) return true;
  return terms.some(function(t){ return (r.sk || '').indexOf(t) > -1; });
}
function lev(a, b){
  var m = a.length, n = b.length, d = [], i, j;
  for (i = 0; i <= m; i++) d[i] = [i];
  for (j = 1; j <= n; j++) d[0][j] = j;
  for (i = 1; i <= m; i++) for (j = 1; j <= n; j++)
    d[i][j] = Math.min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + (a[i-1] === b[j-1] ? 0 : 1));
  return d[m][n];
}
function near(qk, vk){
  if (qk.length < 4) return false;
  for (var len = qk.length - 1; len <= qk.length + 1; len++)
    for (var st = 0; st + len <= vk.length; st++) if (lev(qk, vk.substr(st, len)) <= 1) return true;
  return false;
}
// suggestions come ONLY from values that really exist in the data table
function suggestions(term){
  var qk = sKey(term);
  if (!qk) return [];
  var c = searchCols(), seen = {}, list = [];
  data.rows.forEach(function(r){
    [[c.po, 'PO'], [c.st, 'Style No']].forEach(function(p){
      if (p[0] < 0) return;
      var v = String(r.cells[p[0]] || '').trim();
      if (!v) return;
      var id = p[1] + '|' + v;
      if (!seen[id]) { seen[id] = {v: v, t: p[1], n: 0, k: sKey(v)}; list.push(seen[id]); }
      seen[id].n++;
    });
  });
  var hit = list.filter(function(s){ return s.k.indexOf(qk) > -1; });
  hit.forEach(function(s){ s.sc = s.k.indexOf(qk) === 0 ? 0 : (s.k.slice(-qk.length) === qk ? 1 : 2); });
  if (!hit.length) { hit = list.filter(function(s){ return near(qk, s.k); }); hit.forEach(function(s){ s.sc = 3; }); }
  hit.sort(function(a, b){ return a.sc - b.sc || nat(a.v, b.v); });
  return hit.slice(0, 8);
}
// ---- end search core ----

// ---- Build the page when data changes ----
function build(){
  $('app').hidden = false;
  $('company').textContent = data.company || 'Shipment Schedule';
  $('campus').textContent = data.campus;
  $('updated').textContent = data.updated ? 'Last updated ' + data.updated : '';
  $('updated').hidden = !data.updated;
  prepSearch();
  KEYS.forEach(function(k){
    var vals = uniq(data.rows.map(function(r){ return r[k]; }));
    if (F[k] && vals.indexOf(F[k]) < 0) F[k] = '';
    $('f-' + k).innerHTML = '<option value="">All</option>' + vals.map(function(v){ return '<option>' + esc(v) + '</option>'; }).join('');
  });
  render();
}

function pass(r, skip){ skip = [].concat(skip || []); return matchQ(r) && KEYS.every(function(k){ return skip.indexOf(k) > -1 || !F[k] || r[k] === F[k]; }); }

function render(){
  KEYS.forEach(function(k){ $('f-' + k).value = F[k]; });
  var rows = data.rows.filter(function(r){ return pass(r); });
  var fin = rows.filter(isFinal);
  $('kpi-rows').textContent = fmt(rows.length);
  $('kpi-final').textContent = fmt(fin.length);
  function split(list, f){
    var t = list.filter(function(r){ return r.sec === 'Today'; }), b = list.filter(function(r){ return r.sec === 'Balance'; });
    return 'Today ' + f(t) + '  |  Balance ' + f(b);
  }
  $('kpi-rows-s').textContent = split(rows, function(a){ return fmt(a.length); });
  $('kpi-final-q').textContent = split(fin, function(a){ return fmt(a.length); });
  $('kpi-qty-s').textContent = split(rows, function(a){ return fmt(sumQty(a)); });
  $('kpi-qty').textContent = fmt(sumQty(rows));
  $('k-final').className = 'kpi k2 go' + (F.insp && isFinal({insp: F.insp}) ? ' on' : '');
  chips('g-floor', 'floor'); chips('g-insp', 'insp'); chips('g-brand', 'brand');
  drawChart(rows);
  drawTable(rows);
  showInfo(rows.length);
}

// Clickable cards/buttons. Counts follow all other filters.
function counts(m){
  var t = m.filter(function(r){ return r.sec === 'Today'; }).length;
  function box(l, n){ return '<u><small>' + l + '</small><b>' + n + '</b></u>'; }
  return '<span class="nums">' + box('Today', t) + box('Balance', m.length - t) + box('Total', m.length) + '</span>';
}
function chips(id, key){
  var isI = key === 'insp';
  var base = data.rows.filter(function(r){ return pass(r, isI ? ['insp', 'sec'] : key); });
  var out = isI ? '<div class="sumcard"><b>All inspections</b>' + counts(base) + '</div>' : '';
  $(id).innerHTML = out + uniq(data.rows.map(function(r){ return r[key]; })).map(function(v){
    var m = base.filter(function(r){ return r[key] === v; });
    var dot = isI ? '<i style="background:' + color(v) + '"></i>' : '';
    var name = key === 'floor' ? 'Floor ' + v : v;
    var head = isI ? '<span class="hd">' + dot + '<b>' + esc(name) + '</b></span>' : '<b>' + esc(name) + '</b>';
    var info = isI ? counts(m) : '<span>' + fmt(sumQty(m)) + '</span>' + (key === 'floor' ? '<em>' + m.length + ' styles</em>' : '');
    return '<button type="button" class="chip' + (F[key] === v ? ' on' : '') + (m.length ? '' : ' off') +
      '" data-k="' + key + '" data-v="' + esc(v) + '">' + head + info + '</button>';
  }).join('');
}

function drawChart(rows){
  var type = $('f-chart').value, map = {}, labels = [];
  rows.forEach(function(r){
    if (!(r[dim] in map)) { map[r[dim]] = 0; labels.push(r[dim]); }
    map[r[dim]] += 1;
  });
  if (dim === 'brand' || dim === 'insp') labels.sort(function(a, b){ return map[b] - map[a]; });
  else labels.sort(nat);
  var pal = ['#243b6b','#2f6fd1','#d98e04','#2e7d4f','#7a4fb5','#c62f2f','#0f8b8d','#8a949e'];
  var cols = labels.map(function(l, i){ return dim === 'insp' ? color(l) : pal[i % pal.length]; });
  var round = type === 'pie' || type === 'doughnut';
  if (chart) chart.destroy();
  chart = new Chart($('chart'), {
    type: type,
    data: {labels: labels, datasets: [{label: 'Styles', data: labels.map(function(l){ return map[l]; }),
      backgroundColor: type === 'line' ? '#243b6b' : cols, borderColor: round ? '#fff' : '#243b6b', borderWidth: 2, tension: .25}]},
    options: {responsive: true, maintainAspectRatio: false,
      scales: round ? {} : {y: {beginAtZero: true, ticks: {precision: 0}, title: {display: true, text: 'Number of styles'}}},
      plugins: {legend: {display: round, position: 'right'}, title: {display: true, text: 'Styles by ' + DIMS[dim], font: {size: 16}}}}
  });
}

function drawTable(rows){
  shownRows = rows;
  var n = data.heads.length, q = data.qtyPos, prev = null;
  var h = '<thead><tr>' + data.heads.map(function(x){ return '<th>' + esc(x) + '</th>'; }).join('') + '</tr></thead><tbody>';
  var prevSec = null;
  rows.forEach(function(r, ri){
    if (r.sec !== prevSec) {
      prevSec = r.sec; prev = null;
      h += '<tr class="sec"><td colspan="' + n + '">' + r.sec + ': ' + rows.filter(function(x){ return x.sec === r.sec; }).length + ' styles</td></tr>';
    }
    var fl = r.floor, cls = prev !== null && fl !== prev ? ' class="grp"' : '';
    prev = fl;
    h += '<tr data-i="' + ri + '"' + cls + '>' + r.cells.map(function(c, i){
      if (i === data.inspPos && c) return '<td><span class="tag" style="--c:' + color(c) + '">' + esc(c) + '</span></td>';
      return '<td' + (i === q ? ' class="num"' : '') + '>' + esc(c) + '</td>';
    }).join('') + '</tr>';
  });
  if (!rows.length) h += '<tr class="nodata"><td colspan="' + n + '">No data found' + (sKey(Q) ? ' for &ldquo;' + esc(Q.trim()) + '&rdquo;' : '') + '</td></tr>';
  h += '</tbody><tfoot><tr>';
  if (q > 0) h += '<td colspan="' + q + '">Total: ' + rows.length + ' styles</td><td class="num">' + fmt(sumQty(rows)) + '</td>' + (n - q - 1 > 0 ? '<td colspan="' + (n - q - 1) + '"></td>' : '');
  else h += '<td colspan="' + n + '">Total: ' + rows.length + ' styles</td>';
  $('table').innerHTML = h + '</tr></tfoot>';
}

// ---- Clicks ----
KEYS.forEach(function(k){ $('f-' + k).addEventListener('change', function(){ F[k] = this.value; render(); }); });
$('f-chart').addEventListener('change', function(){ render(); });
$('reset').addEventListener('click', function(){ KEYS.forEach(function(k){ F[k] = ''; }); Q = ''; $('q').value = ''; closeSugg(); render(); });
$('by').addEventListener('click', function(e){
  var b = e.target.closest('button');
  if (!b) return;
  dim = b.getAttribute('data-d');
  Array.prototype.forEach.call($('by').children, function(x){ x.className = x === b ? 'on' : ''; });
  render();
});
document.addEventListener('click', function(e){
  var b = e.target.closest('.chip');
  if (!b) return;
  var k = b.getAttribute('data-k'), v = b.getAttribute('data-v');
  F[k] = F[k] === v ? '' : v;
  render();
});

// ---- Search box (suggestions) ----
var qTimer = null, sgItems = [], sgPos = -1;
function showInfo(n){ $('sbar').classList.toggle('bad', !!sKey(Q) && n === 0); }   // red ring when nothing matches
function closeSugg(){ $('sugg').hidden = true; sgPos = -1; }
function showSugg(){
  sgItems = suggestions(Q.split(/[,;\n]/).pop());
  if (!sgItems.length) { closeSugg(); return; }
  $('sugg').innerHTML = sgItems.map(function(s, i){
    return '<div class="sg" data-i="' + i + '"><b>' + esc(s.v) + '</b><small>' + s.t + ' &middot; ' + s.n + (s.n > 1 ? ' rows' : ' row') + '</small></div>';
  }).join('');
  $('sugg').hidden = false; sgPos = -1;
}
function markSugg(){
  Array.prototype.forEach.call($('sugg').children, function(el, i){
    el.className = 'sg' + (i === sgPos ? ' on' : '');
    if (i === sgPos) el.scrollIntoView({block: 'nearest'});
  });
}
function pickSugg(i){
  var s = sgItems[i];
  if (!s) return;
  var last = Q.split(/[,;\n]/).pop();
  Q = Q.slice(0, Q.length - last.length) + s.v;
  $('q').value = Q;
  closeSugg();
  render();
}
$('q').addEventListener('input', function(){
  Q = this.value;
  showSugg();
  clearTimeout(qTimer);
  qTimer = setTimeout(render, 150);
});
$('q').addEventListener('focus', function(){ if (sKey(Q)) showSugg(); });
$('qgo').addEventListener('click', function(){ $('q').focus(); if (sKey(Q)) showSugg(); render(); });
$('q').addEventListener('keydown', function(e){
  var open = !$('sugg').hidden, n = sgItems.length;
  if (e.key === 'ArrowDown' && open) { e.preventDefault(); sgPos = (sgPos + 1) % n; markSugg(); }
  else if (e.key === 'ArrowUp' && open) { e.preventDefault(); sgPos = (sgPos - 1 + n) % n; markSugg(); }
  else if (e.key === 'Enter') {
    if (open && sgPos > -1) { e.preventDefault(); pickSugg(sgPos); }
    else closeSugg();
  }
  else if (e.key === 'Escape') closeSugg();
});
$('sugg').addEventListener('mousedown', function(e){
  e.preventDefault();
  var d = e.target.closest('.sg');
  if (d) pickSugg(+d.getAttribute('data-i'));
});
document.addEventListener('click', function(e){ if (!e.target.closest('.sbox')) closeSugg(); });

// ---- Final Completed card ----
function isFinal(r){ return /final\s*(done|complete|pass)/i.test(r.insp); }
function toggleFinal(){
  var lab = uniq(data.rows.filter(isFinal).map(function(r){ return r.insp; }))[0];
  if (!lab) return;
  F.insp = F.insp === lab ? '' : lab;
  render();
}
$('k-final').addEventListener('click', toggleFinal);
$('k-final').addEventListener('keydown', function(e){ if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleFinal(); } });

// ---- Row popup ----
function openRow(r){
  $('mtitle').textContent = r.style;
  $('msub').textContent = r.brand + '  |  Floor ' + r.floor + '  |  ' + r.sec;
  $('mcards').innerHTML = data.heads.map(function(hd, i){
    var v = r.cells[i], body = v === '' ? '-' : esc(v);
    if (i === data.inspPos && v) body = '<span class="tag" style="--c:' + color(v) + '">' + esc(v) + '</span>';
    return '<div class="mc"><small>' + esc(hd) + '</small><b>' + body + '</b></div>';
  }).join('');
  $('modal').hidden = false;
}
function closeRow(){ $('modal').hidden = true; }
$('table').addEventListener('click', function(e){
  var tr = e.target.closest('tbody tr');
  if (tr && tr.getAttribute('data-i') !== null) openRow(shownRows[Number(tr.getAttribute('data-i'))]);
});
$('mclose').addEventListener('click', closeRow);
$('modal').addEventListener('click', function(e){ if (e.target === $('modal')) closeRow(); });
document.addEventListener('keydown', function(e){ if (e.key === 'Escape') closeRow(); });

// ---- Print and download ----
var FL = {floor:'Floor', brand:'Brand', style:'Style', insp:'Inspection', ccd:'CCD', sec:'Schedule'};
function stamp(){ return new Date().toLocaleString('en-GB', {day:'numeric', month:'long', year:'numeric', hour:'numeric', minute:'2-digit', hour12:true}); }
function fdate(d){
  var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return d.getDate() + '-' + m[d.getMonth()] + '-' + String(d.getFullYear()).slice(2);
}
// Excel print/download colours: Final Pass = green cell, Final Fail = red cell, Final Done / Final Fail = gray row
function xlStat(r){
  var t = String(r.insp || '').replace(/[_\-\s]+/g, ' ').toLowerCase();
  if (/final fail/.test(t)) return 'fail';
  if (/final pass/.test(t)) return 'pass';
  if (/final (done|complete)/.test(t)) return 'done';
  return '';
}
function styleCol(){ return data.heads.findIndex(function(x){ return /^style name/i.test(x); }); }

// The rows of the Excel-style sheet (same order and look as the data file)
function sheetModel(){
  var L = [], camps = [];
  shownRows.forEach(function(r){ if (camps.indexOf(r.camp) < 0) camps.push(r.camp); });
  var multi = camps.length > 1;
  camps.forEach(function(cp, ci){
    var inCamp = shownRows.filter(function(r){ return r.camp === cp; });
    if (ci > 0 || cp !== data.campus) L.push({t:'camp', v:cp});
    ['Today', 'Balance'].forEach(function(sec){
      var list = inCamp.filter(function(r){ return r.sec === sec; }), prev = null;
      if (!list.length) return;
      if (sec === 'Balance') L.push({t:'bal'});
      list.forEach(function(r){
        if (prev !== null && r.floor !== prev) L.push({t:'gap'});
        prev = r.floor;
        var st = xlStat(r);
        L.push({t:'row', cells:r.cells, stat:st, gray:(st === 'fail' || st === 'done')});
      });
      L.push({t:'total', n:list.length, q:sumQty(list)});
      if (sec === 'Today' && data.updated && !multi) L.push({t:'upd', v:'Last Updated On ' + data.updated});
    });
  });
  if (multi && data.updated) L.push({t:'upd', v:'Last Updated On ' + data.updated});
  return L;
}

function xlHtml(){
  var n = data.heads.length, sp = styleCol();
  function td(v, i, x){
    var cl = [];
    if (i === sp) cl.push('l');
    if (i === data.inspPos && x && x.stat === 'pass') cl.push('sp');
    if (i === data.inspPos && x && x.stat === 'fail') cl.push('sf');
    return '<td' + (cl.length ? ' class="' + cl.join(' ') + '"' : '') + '>' + esc(v) + '</td>';
  }
  var h = (logoUrl ? '<img class="xlogo" src="' + logoUrl + '" alt="">' : '') + '<div class="xt1">' + esc(data.company) + '</div><div class="xt2">TODAY SHIPMENT SCHEDULE</div>' +
    '<div class="xt3"><span>' + esc(data.campus) + '</span><span>' + fdate(new Date()) + '</span></div>' +
    '<table class="xl"><thead><tr>' + data.heads.map(function(x){ return '<th>' + esc(x) + '</th>'; }).join('') + '</tr></thead><tbody>';
  sheetModel().forEach(function(x){
    if (x.t === 'gap') h += '<tr class="xg"><td colspan="' + n + '"></td></tr>';
    else if (x.t === 'bal') h += '<tr class="xb"><td colspan="' + n + '">BALANCE</td></tr>';
    else if (x.t === 'camp') h += '<tr class="xc"><td colspan="' + n + '">' + esc(x.v) + '</td></tr>';
    else if (x.t === 'upd') h += '<tr class="xu"><td colspan="' + n + '">' + esc(x.v) + '</td></tr>';
    else if (x.t === 'total') h += '<tr class="xs">' + data.heads.map(function(_, i){ return '<td>' + (i === 0 ? x.n : i === data.qtyPos ? fmt(x.q) : '') + '</td>'; }).join('') + '</tr>';
    else h += '<tr' + (x.gray ? ' class="gy"' : '') + '>' + x.cells.map(function(v, i){ return td(v, i, x); }).join('') + '</tr>';
  });
  return h + '</tbody></table>';
}

// Excel-style print: up to 40 data rows -> everything shrinks onto ONE A4 landscape page.
// More than 40 rows -> normal pages (the table is only fitted to the page width, header repeats on each page).
var XL_ONE_PAGE_ROWS = 40;
function fitXprint(multi){
  var xp = $('xprint'), sc = $('xscale');
  xp.className = multi ? 'multi' : '';
  sc.style.cssText = '';
  xp.style.cssText = 'display:block;position:absolute;left:-9999px;top:0';
  var aw = xp.clientWidth, ah = xp.clientHeight - 4;
  if (multi) {
    var kw = Math.min(1, aw / sc.offsetWidth) * 0.99;
    sc.style.zoom = kw;
    sc.style.width = (aw / kw) + 'px';
  } else {
    var k = Math.min(1, aw / sc.offsetWidth, ah / sc.offsetHeight) * 0.99;
    sc.style.width = (aw / k) + 'px';
    sc.style.transform = 'scale(' + k + ')';
  }
  xp.style.cssText = '';
}
function printExcel(){
  $('xprint').innerHTML = '<div class="xscale" id="xscale">' + xlHtml() + '</div>';
  imagesReady($('xprint'), function(){
    fitXprint(sheetModel().filter(function(x){ return x.cells; }).length > XL_ONE_PAGE_ROWS);
    document.body.className = 'print-xl';
    setTimeout(function(){ window.print(); }, 250);
  });
}

// Executive summary tables for the print view (Today / Balance / Total)
function sumTable(title, head, key, withQty, nameFn){
  var vals = uniq(shownRows.map(function(r){ return r[key]; })), T = 0, B = 0, Q = 0;
  var body = vals.map(function(v){
    var m = shownRows.filter(function(r){ return r[key] === v; });
    var t = m.filter(function(r){ return r.sec === 'Today'; }).length, b = m.length - t, q = sumQty(m);
    T += t; B += b; Q += q;
    return '<tr><td>' + esc(nameFn ? nameFn(v) : v) + '</td><td>' + t + '</td><td>' + b + '</td><td>' + m.length + '</td>' + (withQty ? '<td>' + fmt(q) + '</td>' : '') + '</tr>';
  }).join('') || '<tr><td colspan="' + (withQty ? 5 : 4) + '">No data</td></tr>';
  return '<h4>' + title + '</h4><table class="pt"><thead><tr><th>' + head + '</th><th>Today</th><th>Balance</th><th>Total</th>' + (withQty ? '<th>QTY</th>' : '') +
    '</tr></thead><tbody>' + body + '</tbody><tfoot><tr><td>Total</td><td>' + T + '</td><td>' + B + '</td><td>' + (T + B) + '</td>' +
    (withQty ? '<td>' + fmt(Q) + '</td>' : '') + '</tr></tfoot></table>';
}

// One A4 landscape report page: header, KPIs, chart, summary tables, footer. Shrinks to fit if needed.
function buildPview(mode){
  var pv = $('pview');
  function kp(id){ return esc($(id).textContent); }
  function kbox(label, num, sub){ return '<div class="pkb"><small>' + label + '</small><b>' + num + '</b><span>' + sub + '</span></div>'; }
  var fl = KEYS.filter(function(k){ return F[k]; }).map(function(k){ return FL[k] + ': ' + F[k]; });
  if (sKey(Q)) fl.push('Search: ' + Q.trim());
  pv.innerHTML =
    '<div class="pscale" id="pscale">' +
    '<div class="ph"><div class="pl">' + (logoUrl ? '<img src="' + logoUrl + '" alt="">' : '') + '<div><h1>' + esc(data.company) + '</h1><p>' + esc(data.campus) + '</p></div></div>' +
    '<div class="phr"><h2>Today\'s Shipment Schedule</h2><p>' + esc($('today').textContent) + '</p>' +
    (data.updated ? '<p>Last updated ' + esc(data.updated) + '</p>' : '') + '</div></div>' +
    '<p class="pfil">Filters applied: ' + (fl.length ? esc(fl.join('   |   ')) : 'none (all data)') + '</p>' +
    '<div class="pk">' + kbox('Total Style', kp('kpi-rows'), kp('kpi-rows-s')) + kbox('Final Completed', kp('kpi-final'), kp('kpi-final-q')) +
    kbox('Total QTY', kp('kpi-qty'), kp('kpi-qty-s')) + '</div>' +
    '<div class="pg"><div class="pbx pchart"><img alt="Chart" src="' + chart.toBase64Image() + '"></div>' +
    '<div class="pbx">' + sumTable('Floor Summary (Styles &amp; QTY)', 'Floor', 'floor', true, function(v){ return 'Floor ' + v; }) + '</div>' +
    '<div class="pbx">' + sumTable('Brand Summary (Styles &amp; QTY)', 'Brand', 'brand', true) + '</div>' +
    '<div class="pbx">' + sumTable('Inspection Summary (Number of Inspections)', 'Inspection', 'insp', false) + '</div></div></div>' +
    '<div class="pfoot"><span>' + esc(data.company + '   |   ' + data.campus) + '</span><span>' +
    (mode === 'download' ? 'Downloaded' : 'Printed') + ' on ' + stamp() + '   |   Page 1 of 1</span></div>';
  return pv;
}
// Shrinks the report to fit one A4 page. Must run AFTER the logo and chart pictures are ready.
function fitPview(pv){
  pv.style.cssText = 'display:block;position:absolute;left:-9999px;top:0';
  var cs = getComputedStyle(pv), sc = $('pscale');
  var aw = pv.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
  var ah = pv.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) - 4;
  var k = Math.min(1, ah / sc.offsetHeight);
  if (k < 1) sc.style.width = (aw / k) + 'px';
  sc.style.transform = 'scale(' + k + ')';
  pv.style.cssText = '';
}
// Waits until every picture inside a box is fully loaded (fixes blank logo/chart when printing)
function imagesReady(root, cb){
  var imgs = Array.prototype.slice.call(root.querySelectorAll('img')), left = imgs.length, fired = false;
  function fin(){ if (fired) return; fired = true; cb(); }
  if (!left) { fin(); return; }
  function one(){ left--; if (left <= 0) fin(); }
  imgs.forEach(function(im){
    if (im.decode) im.decode().then(one, one);
    else if (im.complete) one();
    else { im.addEventListener('load', one); im.addEventListener('error', one); }
  });
  setTimeout(fin, 4000);
}
function printDash(){
  var pv = buildPview('print');
  imagesReady(pv, function(){
    fitPview(pv);
    document.body.className = 'print-dash';
    setTimeout(function(){ window.print(); }, 250);
  });
}
window.addEventListener('afterprint', function(){ document.body.className = ''; $('pview').innerHTML = ''; $('xprint').innerHTML = ''; });

function save(blob, name){
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

// Real .xlsx with the same layout and green "Final Done" rows (needs libs/exceljs.min.js)
function downloadExcel(){
  if (typeof ExcelJS === 'undefined') {
    alert('Excel download needs one more library file.\n\n1. Open: https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js\n2. Press Ctrl+S and save it in the "libs" folder as exceljs.min.js\n3. Reload this page.');
    return;
  }
  var n = data.heads.length, q = data.qtyPos, sp = styleCol();
  var wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Sheet1'), r = 0;
  var thin = {style:'thin', color:{argb:'FF000000'}}, med = {style:'medium', color:{argb:'FF000000'}};
  var box = {top:thin, left:thin, bottom:thin, right:thin};
  var mid = {horizontal:'center', vertical:'middle', wrapText:true};
  function fill(argb){ return {type:'pattern', pattern:'solid', fgColor:{argb:argb}}; }
  function line(text, size, align){
    r++; ws.mergeCells(r, 1, r, n);
    var c = ws.getCell(r, 1);
    c.value = text; c.font = {bold:true, size:size}; c.alignment = {horizontal:align, vertical:'middle'};
  }
  var ext = logoUrl ? logoUrl.slice(11, logoUrl.indexOf(';')) : '';
  if (ext === 'png' || ext === 'jpeg') {
    var lh = 52, lw = logoImg && logoImg.naturalHeight ? Math.round(lh * logoImg.naturalWidth / logoImg.naturalHeight) : 120;
    ws.addImage(wb.addImage({base64: logoUrl, extension: ext}), {tl: {col: 0.1, row: 0.1}, ext: {width: lw, height: lh}});
    ws.getRow(1).height = 32;
  }
  line(data.company, 16, 'center');
  line('TODAY SHIPMENT SCHEDULE', 12, 'center');
  r++; ws.mergeCells(r, 1, r, n - 3); ws.mergeCells(r, n - 2, r, n);
  ws.getCell(r, 1).value = data.campus; ws.getCell(r, 1).font = {bold:true, size:12};
  ws.getCell(r, n - 2).value = fdate(new Date()); ws.getCell(r, n - 2).font = {bold:true, size:12};
  ws.getCell(r, n - 2).alignment = {horizontal:'center'};
  r++;
  data.heads.forEach(function(x, i){
    var c = ws.getCell(r, i + 1);
    c.value = x; c.font = {bold:true}; c.alignment = mid; c.border = box; c.fill = fill('FFE2EFDA');
  });
  sheetModel().forEach(function(x){
    r++;
    if (x.t === 'gap') {
      ws.getRow(r).height = 5;
      for (var i = 1; i <= n; i++) ws.getCell(r, i).border = box;
    } else if (x.t === 'camp') {
      ws.mergeCells(r, 1, r, n);
      var cc = ws.getCell(r, 1);
      cc.value = x.v; cc.font = {bold:true, size:12}; cc.alignment = {horizontal:'left', vertical:'middle'};
      ws.getRow(r).height = 22;
    } else if (x.t === 'bal' || x.t === 'upd') {
      ws.mergeCells(r, 1, r, n);
      var c = ws.getCell(r, 1);
      c.value = x.t === 'bal' ? 'BALANCE' : x.v;
      c.font = {bold:true, size: x.t === 'bal' ? 12 : 10};
      c.alignment = {horizontal: x.t === 'bal' ? 'center' : 'right'};
    } else if (x.t === 'total') {
      for (var j = 0; j < n; j++) {
        var t = ws.getCell(r, j + 1);
        t.font = {bold:true}; t.alignment = mid; t.border = {top:med, bottom:med, left:thin, right:thin};
      }
      ws.getCell(r, 1).value = x.n;
      ws.getCell(r, q + 1).value = x.q;
    } else {
      x.cells.forEach(function(v, j){
        var c = ws.getCell(r, j + 1), num = Number(String(v).replace(/,/g, ''));
        c.value = (j === 0 || j === q) && v !== '' && !isNaN(num) ? num : v;
        c.alignment = {horizontal: j === sp ? 'left' : 'center', vertical:'middle', wrapText:true};
        c.border = box;
        if (x.gray) c.fill = fill('FFD9D9D9');
        if (j === data.inspPos && x.stat === 'pass') { c.fill = fill('FFC6EFCE'); c.font = {bold:true, color:{argb:'FF006100'}}; }
        if (j === data.inspPos && x.stat === 'fail') { c.fill = fill('FFFFC7CE'); c.font = {bold:true, color:{argb:'FF9C0006'}}; }
      });
    }
  });
  data.heads.forEach(function(hd, i){
    var w = hd.length;
    shownRows.forEach(function(row){ w = Math.max(w, String(row.cells[i]).length); });
    ws.getColumn(i + 1).width = Math.min(40, w + 3);
  });
  ws.pageSetup = {paperSize:9, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight:1};
  var d = new Date();
  function p(k){ return String(k).padStart(2, '0'); }
  var name = 'Shipment_Schedule_' + d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()) + '.xlsx';
  wb.xlsx.writeBuffer().then(function(buf){
    save(new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), name);
  }).catch(function(){ alert('Could not create the Excel file.'); });
}

// ---- Direct PDF download (needs libs/html2canvas.min.js) ----
function fileStamp(){
  var d = new Date();
  function p(k){ return String(k).padStart(2, '0'); }
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes());
}

// Wraps one JPEG picture into a one-page A4 landscape PDF (no library needed)
function jpegPdf(dataUrl, pxW, pxH){
  var bin = atob(dataUrl.split(',')[1]), img = new Uint8Array(bin.length), i;
  for (i = 0; i < bin.length; i++) img[i] = bin.charCodeAt(i);
  var PW = 841.89, PH = 595.28, m = 0, w = PW - 2 * m, h = w * pxH / pxW;
  if (h > PH - 2 * m) { h = PH - 2 * m; w = h * pxW / pxH; }
  var x = (PW - w) / 2, y = (PH - h) / 2;
  var parts = [], offs = [], len = 0;
  function add(v){ var b = typeof v === 'string' ? new TextEncoder().encode(v) : v; parts.push(b); len += b.length; }
  function obj(n, body){ offs[n] = len; add(n + ' 0 obj\n' + body + '\nendobj\n'); }
  add('%PDF-1.4\n');
  obj(1, '<< /Type /Catalog /Pages 2 0 R >>');
  obj(2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>');
  obj(3, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + PW + ' ' + PH + '] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>');
  var content = 'q ' + w.toFixed(2) + ' 0 0 ' + h.toFixed(2) + ' ' + x.toFixed(2) + ' ' + y.toFixed(2) + ' cm /Im0 Do Q';
  obj(4, '<< /Length ' + content.length + ' >>\nstream\n' + content + '\nendstream');
  offs[5] = len;
  add('5 0 obj\n<< /Type /XObject /Subtype /Image /Width ' + pxW + ' /Height ' + pxH + ' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ' + img.length + ' >>\nstream\n');
  add(img);
  add('\nendstream\nendobj\n');
  var xref = len, t = 'xref\n0 6\n0000000000 65535 f \n';
  for (i = 1; i <= 5; i++) t += String(offs[i]).padStart(10, '0') + ' 00000 n \n';
  add(t + 'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n' + xref + '\n%%EOF');
  return new Blob(parts, {type: 'application/pdf'});
}

function downloadPdf(){
  if (typeof html2canvas === 'undefined') {
    alert('PDF download needs one more library file.\n\n1. Open: https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js\n2. Press Ctrl+S and save it in the "libs" folder as html2canvas.min.js\n3. Reload this page.');
    return;
  }
  var y0 = window.scrollY, pv = buildPview('download');
  function done(){ pv.style.cssText = ''; pv.innerHTML = ''; window.scrollTo(0, y0); }
  imagesReady(pv, function(){
    fitPview(pv);
    window.scrollTo(0, 0);
    pv.style.cssText = 'display:block;position:absolute;left:0;top:0;z-index:-1';
    html2canvas(pv, {scale: 3, backgroundColor: '#ffffff', useCORS: true}).then(function(cv){
      var url = cv.toDataURL('image/jpeg', 0.92), w = cv.width, h = cv.height;
      done();
      save(jpegPdf(url, w, h), 'Dashboard_' + fileStamp() + '.pdf');
    }).catch(function(){
      done();
      alert('Could not create the PDF. Use the Print button and choose "Save as PDF" instead.');
    });
  });
}

document.querySelector('.outs').addEventListener('click', function(e){
  var b = e.target.closest('button');
  if (!b) return;
  var a = b.getAttribute('data-act');
  if (a === 'pd') printDash();
  else if (a === 'dd') downloadPdf();
  else if (a === 'px') printExcel();
  else if (a === 'dx') downloadExcel();
});

// ---- Company logo: put logo.png (or logo.jpg) in the same folder as index.html ----
var LOGOS = ['logo.png', 'logo.PNG', 'logo.jpg', 'logo.JPG', 'logo.jpeg', 'logo.svg', 'logo.png.png', 'logo.jpg.jpg'];
var logoUrl = '', logoImg = null, logoNote = '';
function logoNotice(msg){ logoNote = msg; $('status').textContent += ' ' + msg; }
function logoEl(){
  var el = $('logo');
  if (!el) {
    el = document.createElement('img');
    el.id = 'logo';
    el.alt = 'Company logo';
    var hero = document.querySelector('.hero');
    hero.insertBefore(el, hero.firstChild);
  }
  return el;
}
function loadLogo(i){
  if (i >= LOGOS.length) { logoNotice('Logo not found: save it as logo.png in the same folder as index.html.'); return; }
  var im = new Image();
  im.onload = function(){
    var el = logoEl();
    el.src = im.src;
    el.hidden = false;
    logoImg = im;
    logoData(LOGOS[i]);
  };
  im.onerror = function(){ loadLogo(i + 1); };
  im.src = LOGOS[i] + '?t=' + Date.now();
}
// a copy of the logo for the printed report, PDF and Excel file
function logoData(name){
  var x = new XMLHttpRequest();
  x.open('GET', name + '?t=' + Date.now(), true);
  x.responseType = 'arraybuffer';
  x.onload = function(){
    if (!((x.status === 0 || x.status === 200) && x.response && x.response.byteLength)) { logoNotice('Logo shows here but could not be added to the print, PDF and Excel.'); return; }
    var b = new Uint8Array(x.response), t = '', k;
    for (k = 0; k < b.length; k++) t += String.fromCharCode(b[k]);
    var ext = name.split('.').pop().toLowerCase();
    logoUrl = 'data:image/' + (ext === 'svg' ? 'svg+xml' : ext === 'jpg' ? 'jpeg' : ext) + ';base64,' + btoa(t);
  };
  x.onerror = function(){ logoNotice('Logo shows here but could not be added to the print, PDF and Excel.'); };
  x.send();
}

loadLogo(0);
load();
setInterval(load, REFRESH_SECONDS * 1000);
