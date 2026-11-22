(function () {
  'use strict';

  // Superstore Fun Facts, VizSight Analytics style: a guess-then-reveal tour of seven facts, every sentence computed from live queries.
  var root = document.getElementById('app');
  var DS_NAME = 'Sample - Superstore';
  var reduceMQ = window.matchMedia ? window.matchMedia('(prefers-reduced-motion: reduce)') : null;
  function still() { return !!window.FF_STILL || !!(reduceMQ && reduceMQ.matches); }

  var DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var DAY_ORDER = [1, 2, 3, 4, 5, 6, 0];
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var MODES = ['Same Day', 'First Class', 'Second Class', 'Standard Class'];
  var NS = 'http://www.w3.org/2000/svg';
  var C = {}; // colour tokens, read from CSS so a style swap is a token change

  // ------------------------------------------------------------------ small helpers
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) { e.className = cls; }
    if (text != null) { e.textContent = text; }
    return e;
  }
  function sv(tag, attrs, parent) {
    var e = document.createElementNS(NS, tag);
    for (var k in attrs) { if (attrs.hasOwnProperty(k)) { e.setAttribute(k, attrs[k]); } }
    if (parent) { parent.appendChild(e); }
    return e;
  }
  function clamp(v, a, b) { return Math.max(a, Math.min(b, v)); }
  function sum(a, f) { var s = 0; a.forEach(function (x) { s += f ? f(x) : x; }); return s; }
  function nf(n) { return Math.round(n).toLocaleString('en-US'); }
  function money(v) { return (v < 0 ? '−' : '') + '$' + nf(Math.abs(v)); }
  function moneyAbs(v) { return '$' + nf(Math.abs(v)); }
  function money2(v) { return (v < 0 ? '−' : '') + '$' + Math.abs(v).toFixed(2); }
  function moneyShort(v) {
    var a = Math.abs(v), s = v < 0 ? '−' : '';
    if (a >= 1e6) { return s + '$' + (a / 1e6).toFixed(1) + 'M'; }
    if (a >= 1e4) { return s + '$' + Math.round(a / 1e3) + 'K'; }
    return s + '$' + nf(a);
  }
  function pct(x) { return Math.round(x * 100) + '%'; }
  function times(a, b) { return (a / b).toFixed(1) + '×'; }
  function plural(n, one, many) { return n === 1 ? one : many; }
  function B(t) { return { b: String(t) }; }
  function SG(t, v) { return { b: String(t), sign: v }; }
  function rich(node, parts) {
    node.textContent = '';
    parts.forEach(function (p) {
      if (typeof p === 'string') { node.appendChild(document.createTextNode(p)); return; }
      var s = el('b');
      if (p.sign) {
        s.className = p.sign < 0 ? 'neg' : 'pos';
        s.textContent = (p.sign < 0 ? '▼' : '▲') + ' ' + p.b;
      } else { s.textContent = p.b; }
      node.appendChild(s);
    });
    return node;
  }
  function num(v) { var n = typeof v === 'number' ? v : parseFloat(v); return isFinite(n) ? n : 0; }
  function str(v) { return v == null ? '' : String(v); }
  function pick(r, keys) { for (var i = 0; i < keys.length; i++) { if (r[keys[i]] !== undefined && r[keys[i]] !== null) { return r[keys[i]]; } } return null; }
  function parseDay(v) {
    if (v == null) { return null; }
    if (typeof v === 'number') { var d0 = new Date(v); return { y: d0.getUTCFullYear(), m: d0.getUTCMonth() + 1, d: d0.getUTCDate() }; }
    var m = /(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(v));
    if (m) { return { y: +m[1], m: +m[2], d: +m[3] }; }
    var t = Date.parse(String(v));
    if (!isNaN(t)) { var d1 = new Date(t); return { y: d1.getFullYear(), m: d1.getMonth() + 1, d: d1.getDate() }; }
    return null;
  }
  function weekday(p) { return new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay(); }
  function dayNum(p) { return Date.UTC(p.y, p.m - 1, p.d) / 864e5; }
  function argmax(a, f) { var bi = -1, bv = -Infinity; a.forEach(function (x, i) { var v = f(x); if (v > bv) { bv = v; bi = i; } }); return a[bi]; }
  function argmin(a, f) { return argmax(a, function (x) { return -f(x); }); }

  // ------------------------------------------------------------------ data access
  function extractData(result) { return result && Array.isArray(result.data) ? result.data : []; }
  function getDataSources() {
    var wb = tableau.extensions.workbook;
    if (wb && typeof wb.getAllDataSourcesAsync === 'function') { return wb.getAllDataSourcesAsync(); }
    var wc = tableau.extensions.worksheetContent;
    if (wc && wc.worksheet && typeof wc.worksheet.getDataSourcesAsync === 'function') { return wc.worksheet.getDataSourcesAsync(); }
    return Promise.resolve([]);
  }
  function findDataSource(list) {
    var found = null;
    (list || []).forEach(function (ds) { if (ds && ds.name === DS_NAME) { found = ds; } });
    return found || ((list && list.length) ? list[0] : null);
  }
  function F(caption, fn, alias) { return { fieldCaption: caption, 'function': fn, fieldAlias: alias }; }
  function D(caption, alias) { return { fieldCaption: caption, fieldAlias: alias }; }
  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function runFirst(ds, variants) {
    function pass() {
      var i = 0;
      function attempt() {
        return ds.queryAsync({ fields: variants[i] }).then(extractData).catch(function (err) {
          i += 1;
          if (i < variants.length) { return attempt(); }
          throw err;
        });
      }
      return attempt();
    }
    return pass().catch(function () { return wait(900).then(pass); });
  }
  function errText(err) { return (err && (err.message || err.errorCode)) || String(err); }

  var QUERIES = [
    ['orders', 'orders and baskets', [
      [D('Order ID', 'oid'), D('Order Date', 'od'), D('Country/Region', 'ctry'), F('Row ID', 'COUNT', 'lines'), F('Sales', 'SUM', 'sales')],
      [D('Order ID', 'oid'), F('Order Date', 'TRUNC_DAY', 'od'), D('Country/Region', 'ctry'), F('Row ID', 'COUNT', 'lines'), F('Sales', 'SUM', 'sales')],
      [D('Order ID', 'oid'), D('Order Date', 'od'), D('Country/Region', 'ctry'), F('Product ID', 'COUNT', 'lines'), F('Sales', 'SUM', 'sales')]
    ]],
    ['ship', 'shipping times', [
      [D('Ship Mode', 'mode'), { fieldCaption: 'Days to ship', calculation: "DATEDIFF('day', [Order Date], [Ship Date])", fieldAlias: 'days' },
        D('Country/Region', 'ctry'), F('Row ID', 'COUNT', 'lines'), F('Order ID', 'COUNTD', 'orders')],
      [D('Ship Mode', 'mode'), D('Order ID', 'oid'), D('Order Date', 'od'), D('Ship Date', 'sd'), D('Country/Region', 'ctry'), F('Product ID', 'COUNT', 'lines')]
    ]],
    ['lines', 'order lines and discounts', [
      [D('Row ID', 'rid'), D('Product Name', 'pn'), D('Sub-Category', 'sub'), F('Discount', 'MIN', 'disc'),
        F('Sales', 'SUM', 'sales'), F('Profit', 'SUM', 'profit'), F('Quantity', 'SUM', 'qty')],
      [D('Row ID', 'rid'), D('Product Name', 'pn'), D('Sub-Category', 'sub'), D('Discount', 'disc'),
        F('Sales', 'SUM', 'sales'), F('Profit', 'SUM', 'profit'), F('Quantity', 'SUM', 'qty')]
    ]],
    ['cust', 'customers', [
      [D('Customer ID', 'cid'), D('Customer Name', 'nm'), F('Sales', 'SUM', 'sales'), F('Profit', 'SUM', 'profit'), F('Order ID', 'COUNTD', 'orders')]
    ]],
    ['totals', 'store totals', [
      [F('Product ID', 'COUNTD', 'products'), F('Customer ID', 'COUNTD', 'customers'), F('State/Province', 'COUNTD', 'states')]
    ]]
  ];

  function normalize(name, rows) {
    if (name === 'orders') {
      return rows.map(function (r) {
        var p = parseDay(pick(r, ['od', 'Order Date', 'DAY(Order Date)']));
        return { oid: str(pick(r, ['oid', 'Order ID'])), date: p, ctry: str(pick(r, ['ctry', 'Country/Region'])),
          lines: num(pick(r, ['lines', 'COUNT(Row ID)'])), sales: num(pick(r, ['sales', 'SUM(Sales)'])) };
      }).filter(function (o) { return o.date; });
    }
    if (name === 'ship') {
      if (rows.length && pick(rows[0], ['sd', 'Ship Date']) != null) {
        var agg = {};
        rows.forEach(function (r) {
          var a = parseDay(pick(r, ['od', 'Order Date'])), b = parseDay(pick(r, ['sd', 'Ship Date']));
          if (!a || !b) { return; }
          var days = Math.round(dayNum(b) - dayNum(a));
          var k = str(pick(r, ['mode', 'Ship Mode'])) + '|' + days + '|' + str(pick(r, ['ctry', 'Country/Region']));
          if (!agg[k]) { agg[k] = { mode: k.split('|')[0], days: days, ctry: k.split('|')[2], lines: 0, ids: {} }; }
          agg[k].lines += num(pick(r, ['lines'])); agg[k].ids[str(pick(r, ['oid', 'Order ID']))] = 1;
        });
        return Object.keys(agg).map(function (k) { var g = agg[k]; return { mode: g.mode, days: g.days, ctry: g.ctry, lines: g.lines, orders: Object.keys(g.ids).length }; });
      }
      return rows.map(function (r) {
        return { mode: str(pick(r, ['mode', 'Ship Mode'])), days: Math.round(num(pick(r, ['days', 'Days to ship']))),
          ctry: str(pick(r, ['ctry', 'Country/Region'])), lines: num(pick(r, ['lines'])), orders: num(pick(r, ['orders'])) };
      });
    }
    if (name === 'lines') {
      return rows.map(function (r) {
        var d = num(pick(r, ['disc', 'MIN(Discount)', 'Discount']));
        return { pn: str(pick(r, ['pn', 'Product Name'])), sub: str(pick(r, ['sub', 'Sub-Category'])),
          disc: Math.round(d <= 1 ? d * 100 : d), sales: num(pick(r, ['sales', 'SUM(Sales)'])),
          profit: num(pick(r, ['profit', 'SUM(Profit)'])), qty: num(pick(r, ['qty', 'SUM(Quantity)'])) };
      });
    }
    if (name === 'cust') {
      return rows.map(function (r) {
        return { cid: str(pick(r, ['cid', 'Customer ID'])), nm: str(pick(r, ['nm', 'Customer Name'])),
          sales: num(pick(r, ['sales', 'SUM(Sales)'])), profit: num(pick(r, ['profit', 'SUM(Profit)'])), orders: num(pick(r, ['orders'])) };
      });
    }
    if (name === 'totals') {
      var t = rows[0] || {};
      return { products: num(pick(t, ['products'])), customers: num(pick(t, ['customers'])), states: num(pick(t, ['states'])) };
    }
    return rows;
  }

  // ------------------------------------------------------------------ state
  var DATA = {};
  var S = { ch: 0, guess: {}, revealed: {}, year: 'All', mode: 'Same Day', discMeasure: 'total', subSort: 'qty', spot: 'spend' };
  var loadedAt = new Date();
  var CH = []; // chapter definitions (built after data loads)

  // ------------------------------------------------------------------ animation engine
  function ease(t) { return 1 - Math.pow(1 - t, 3); }
  // owner: the chart instance, so a new update cancels only that chart's running animation
  function run(owner, duration, stagger, n, frame) {
    var token = owner.__tok = (owner.__tok || 0) + 1;
    if (still() || duration <= 0) { frame(function () { return 1; }); return; }
    var total = duration + stagger * Math.max(0, n - 1), t0 = null;
    function step(now) {
      if (token !== owner.__tok) { return; }
      if (t0 === null) { t0 = now; }
      var el0 = now - t0;
      frame(function (i) { return ease(clamp((el0 - (i || 0) * stagger) / duration, 0, 1)); });
      if (el0 < total) { requestAnimationFrame(step); }
    }
    requestAnimationFrame(step);
  }
  function lerp(a, b, t) { return a + (b - a) * t; }
  function countUp(node, to, fmt, dur) {
    if (still()) { node.textContent = fmt(to); return; }
    var t0 = null;
    function step(now) {
      if (t0 === null) { t0 = now; }
      var p = ease(clamp((now - t0) / dur, 0, 1));
      node.textContent = fmt(to * p);
      if (p < 1 && node.isConnected) { requestAnimationFrame(step); }
    }
    requestAnimationFrame(step);
  }

  // ------------------------------------------------------------------ tooltip
  var tip = el('div', 'tip');
  function showTip(evt, title, rows) {
    tip.textContent = '';
    tip.appendChild(el('div', 'tt', title));
    rows.forEach(function (r) { var d = el('div', 'tr'); d.appendChild(el('span', null, r[0])); d.appendChild(el('span', null, r[1])); tip.appendChild(d); });
    tip.classList.add('on');
    var x = evt.clientX + 14, y = evt.clientY + 14, w = tip.offsetWidth, h = tip.offsetHeight;
    if (x + w > window.innerWidth - 8) { x = evt.clientX - w - 14; }
    if (y + h > window.innerHeight - 8) { y = evt.clientY - h - 14; }
    tip.style.left = x + 'px'; tip.style.top = y + 'px';
  }
  function hideTip() { tip.classList.remove('on'); }

  // ------------------------------------------------------------------ scales
  function niceTicks(lo, hi, count) {
    if (hi === lo) { hi = lo + 1; }
    var span = hi - lo, step = Math.pow(10, Math.floor(Math.log10(span / count))), err = span / count / step;
    if (err >= 7.5) { step *= 10; } else if (err >= 3.5) { step *= 5; } else if (err >= 1.5) { step *= 2; }
    var a = Math.floor(lo / step) * step, b = Math.ceil(hi / step) * step, out = [];
    for (var v = a; v <= b + step / 2; v += step) { out.push(Math.round(v / step) * step); }
    return out;
  }

  // ------------------------------------------------------------------ column / lollipop chart
  function ColumnChart(host) {
    var svg = sv('svg', {}, host);
    var gGrid = sv('g', {}, svg), gMarks = sv('g', {}, svg), gLab = sv('g', {}, svg), gHit = sv('g', {}, svg);
    var marks = {};
    // square ends (VizSight charts: never mix square and rounded in one app)
    function columnPath(cx, w, y0, y1) {
      var x0 = cx - w / 2;
      return 'M' + x0 + ',' + y0 + 'V' + y1 + 'H' + (x0 + w) + 'V' + y0 + 'Z';
    }
    this.update = function (spec, opts) {
      opts = opts || {};
      var W = spec.w, H = spec.h, m = { l: 8, r: 8, t: 24, b: 28 };
      var items = spec.items, n = items.length;
      // Axis rule: a value axis only earns its place when some marks are unlabelled. If every mark carries its value, the axis would repeat it.
      var allLabelled = items.every(function (d) { return d.showLabel; });
      var showAxis = !spec.hidden && !allLabelled;
      if (showAxis) { m.l = 52; }
      svg.setAttribute('width', W); svg.setAttribute('height', H); svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);
      var pw = W - m.l - m.r, ph = H - m.t - m.b;
      var vals = items.map(function (d) { return d.value; });
      var lo = Math.min(0, Math.min.apply(null, vals)), hi = Math.max(0, Math.max.apply(null, vals));
      var ticks = spec.hidden ? [0, 1] : niceTicks(lo, hi, 4);
      var dLo = ticks[0], dHi = ticks[ticks.length - 1];
      function y(v) { return m.t + ph - (v - dLo) / (dHi - dLo) * ph; }
      var slot = pw / n, colW = Math.min(32, slot * 0.55), y0 = y(0);
      gGrid.textContent = '';
      if (showAxis) {
        // partly labelled: light gridlines and tick labels let readers size the unlabelled marks; the zero line is the one axis line
        ticks.forEach(function (t) {
          sv('line', { x1: m.l, x2: W - m.r, y1: Math.round(y(t)) + 0.5, y2: Math.round(y(t)) + 0.5, 'class': t === 0 ? 'zeroline' : 'gridline' }, gGrid);
          var tx = sv('text', { x: m.l - 8, y: y(t) + 4, 'text-anchor': 'end' }, gGrid); tx.textContent = spec.fmt(t);
        });
      } else {
        // every mark labelled (or still hidden): the baseline only
        sv('line', { x1: m.l, x2: W - m.r, y1: Math.round(y0) + 0.5, y2: Math.round(y0) + 0.5, 'class': 'zeroline' }, gGrid);
      }
      items.forEach(function (d, i) {
        var tx = sv('text', { x: m.l + slot * (i + 0.5), y: H - 8, 'text-anchor': 'middle' }, gGrid); tx.textContent = d.label;
      });
      if (spec.hidden) { gGrid.classList.add('fade-in'); }
      // marks: tween from current state to target
      var keep = {};
      var plan = items.map(function (d, i) {
        keep[d.key] = 1;
        var mk = marks[d.key], tv = spec.hidden ? 0.5 + 0.12 * Math.sin(i * 1.7) : d.value;
        var target = { cx: m.l + slot * (i + 0.5), w: colW, y1: y(tv), y0: y0 };
        if (!mk) {
          mk = marks[d.key] = { g: sv('g', {}, gMarks) };
          mk.cur = { cx: target.cx, w: colW, y1: y0, y0: y0 };
        }
        if (opts.fromBase) { mk.cur = { cx: target.cx, w: colW, y1: y0, y0: y0 }; }
        mk.from = { cx: mk.cur.cx, w: mk.cur.w, y1: mk.cur.y1, y0: mk.cur.y0 };
        mk.to = target;
        mk.color = spec.hidden ? C.placeholder : d.color;
        mk.lolli = !!spec.lollipop;
        mk.g.textContent = '';
        if (mk.lolli) {
          mk.stem = sv('line', { stroke: mk.color, 'stroke-width': 1, 'class': 'mark' }, mk.g);
          mk.dot = sv('circle', { r: 6, fill: mk.color, 'class': 'mark' }, mk.g);
        } else {
          mk.path = sv('path', { fill: mk.color, 'class': 'mark' }, mk.g);
        }
        return mk;
      });
      Object.keys(marks).forEach(function (k) { if (!keep[k]) { marks[k].g.remove(); delete marks[k]; } });
      // value labels (selective) and hit targets
      gLab.textContent = ''; gHit.textContent = '';
      var labs = items.map(function (d, i) {
        if (!(spec.hidden ? i % 2 === 0 || n < 6 : d.showLabel)) { return null; }
        var t = sv('text', { 'text-anchor': 'middle', 'class': 'vlab' + (spec.hidden ? ' q' : '') }, gLab);
        t.textContent = spec.hidden ? '?' : (d.labelText || spec.fmtLabel(d.value));
        return t;
      });
      items.forEach(function (d, i) {
        var r = sv('rect', { x: m.l + slot * i, y: m.t - 10, width: slot, height: ph + 10 + m.b, 'class': 'hit' }, gHit);
        r.addEventListener('mousemove', function (e) {
          if (spec.hidden) { showTip(e, d.label, [['Value', 'Make a guess first']]); return; }
          showTip(e, d.tipTitle || d.label, d.tip);
        });
        r.addEventListener('mouseleave', hideTip);
      });
      function draw(mk, i, t) {
        var c = { cx: lerp(mk.from.cx, mk.to.cx, t), w: lerp(mk.from.w, mk.to.w, t), y1: lerp(mk.from.y1, mk.to.y1, t), y0: lerp(mk.from.y0, mk.to.y0, t) };
        mk.cur = c;
        if (mk.lolli) {
          mk.stem.setAttribute('x1', c.cx); mk.stem.setAttribute('x2', c.cx); mk.stem.setAttribute('y1', c.y0); mk.stem.setAttribute('y2', c.y1);
          mk.dot.setAttribute('cx', c.cx); mk.dot.setAttribute('cy', c.y1);
        } else {
          mk.path.setAttribute('d', columnPath(c.cx, c.w, c.y0, c.y1));
        }
        var lb = labs[i];
        if (lb) {
          var up = c.y1 <= c.y0, off = mk.lolli ? 6 + 6 : 6;
          lb.setAttribute('x', c.cx);
          lb.setAttribute('y', up ? c.y1 - off : c.y1 + off + 10);
        }
      }
      run(this, opts.duration == null ? 650 : opts.duration, opts.stagger == null ? 35 : opts.stagger, plan.length, function (p) {
        plan.forEach(function (mk, i) { draw(mk, i, p(i)); });
      });
    };
  }

  // ------------------------------------------------------------------ horizontal ranked bars (rows re-sort with motion)
  var ROW = 24;
  function HBarChart(host) {
    var svg = sv('svg', {}, host);
    var gGrid = sv('g', {}, svg), gRows = sv('g', {}, svg), gHit = sv('g', {}, svg);
    var rows = {};
    function barPath(x0, x1, yc, th) {
      var y = yc - th / 2;
      return 'M' + x0 + ',' + y + 'H' + x1 + 'V' + (y + th) + 'H' + x0 + 'Z';
    }
    this.update = function (spec, opts) {
      opts = opts || {};
      var W = spec.w, items = spec.items, n = items.length;
      var m = { l: spec.nameW || 12, r: 64, t: 8, b: 8 }, H = m.t + n * ROW + m.b;
      svg.setAttribute('width', W); svg.setAttribute('height', H); svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);
      var vals = items.map(function (d) { return d.value; });
      var lo = Math.min(0, Math.min.apply(null, vals)), hi = Math.max(0, Math.max.apply(null, vals));
      var ticks = spec.hidden ? [0, 1] : niceTicks(lo, hi, 4), dLo = ticks[0], dHi = ticks[ticks.length - 1];
      if (lo < 0) { m.l += 52; }
      var pw = W - m.l - m.r;
      function x(v) { return m.l + (v - dLo) / (dHi - dLo) * pw; }
      var x0 = x(0), th = Math.min(13, ROW * 0.55);
      gGrid.textContent = '';
      // Every bar carries its value label, so there is no x-axis or gridlines: just the zero line.
      sv('line', { x1: Math.round(x0) + 0.5, x2: Math.round(x0) + 0.5, y1: m.t, y2: m.t + n * ROW, 'class': 'zeroline' }, gGrid);
      var keep = {};
      var plan = items.map(function (d, i) {
        keep[d.key] = 1;
        var rw = rows[d.key], tv = spec.hidden ? (0.35 + 0.3 * ((i * 7) % 5) / 5) : d.value;
        var target = { yc: m.t + ROW * (i + 0.5), x1: x(tv), x0: x0 };
        if (!rw) {
          rw = rows[d.key] = { g: sv('g', {}, gRows) };
          rw.cur = { yc: target.yc, x1: x0, x0: x0 };
        }
        if (opts.fromBase) { rw.cur = { yc: target.yc, x1: x0, x0: x0 }; }
        rw.from = { yc: rw.cur.yc, x1: rw.cur.x1, x0: rw.cur.x0 };
        rw.to = target;
        rw.g.textContent = '';
        rw.bar = sv('path', { fill: spec.hidden ? C.placeholder : d.color, 'class': 'mark' }, rw.g);
        rw.name = null;
        if (spec.nameW) {
          rw.name = sv('text', { x: 0, 'text-anchor': 'start' }, rw.g);
          rw.name.textContent = d.label;
          rw.name.setAttribute('class', 'name');
        }
        rw.val = sv('text', { 'class': 'vlab' + (spec.hidden ? ' q' : '') }, rw.g);
        rw.val.textContent = spec.hidden ? '?' : spec.fmtLabel(d.value);
        rw.d = d;
        return rw;
      });
      Object.keys(rows).forEach(function (k) { if (!keep[k]) { rows[k].g.remove(); delete rows[k]; } });
      gHit.textContent = '';
      items.forEach(function (d, i) {
        var r = sv('rect', { x: 0, y: m.t + ROW * i, width: W, height: ROW, 'class': 'hit' }, gHit);
        r.addEventListener('mousemove', function (e) { if (spec.hidden) { showTip(e, d.label, [['Value', 'Make a guess first']]); } else { showTip(e, d.label, d.tip); } });
        r.addEventListener('mouseleave', hideTip);
      });
      function draw(rw, t) {
        var c = { yc: lerp(rw.from.yc, rw.to.yc, t), x1: lerp(rw.from.x1, rw.to.x1, t), x0: lerp(rw.from.x0, rw.to.x0, t) };
        rw.cur = c;
        rw.bar.setAttribute('d', barPath(c.x0, c.x1, c.yc, th));
        if (rw.name) { rw.name.setAttribute('y', c.yc + 4); }
        var neg = c.x1 < c.x0;
        rw.val.setAttribute('x', neg ? c.x1 - 6 : c.x1 + 6);
        rw.val.setAttribute('text-anchor', neg ? 'end' : 'start');
        rw.val.setAttribute('y', c.yc + 4);
      }
      run(this, opts.duration == null ? 700 : opts.duration, opts.stagger == null ? 22 : opts.stagger, plan.length, function (p) {
        plan.forEach(function (rw, i) { draw(rw, p(i)); });
      });
      return H;
    };
  }

  // ------------------------------------------------------------------ scatter (dots fly from a huddle to their place)
  function ScatterChart(host) {
    var svg = sv('svg', {}, host);
    var gGrid = sv('g', {}, svg), gDots = sv('g', {}, svg), gSpot = sv('g', {}, svg), gHit = sv('g', {}, svg);
    var dots = {}, spotState = null;
    this.update = function (spec, opts) {
      opts = opts || {};
      var W = spec.w, H = spec.h, m = { l: 60, r: 16, t: 14, b: 40 };
      svg.setAttribute('width', W); svg.setAttribute('height', H); svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);
      var pw = W - m.l - m.r, ph = H - m.t - m.b, pts = spec.points;
      var xt = niceTicks(0, Math.max.apply(null, pts.map(function (p) { return p.x; })), 5);
      var yt = niceTicks(Math.min(0, Math.min.apply(null, pts.map(function (p) { return p.y; }))), Math.max(0, Math.max.apply(null, pts.map(function (p) { return p.y; }))), 5);
      function X(v) { return m.l + (v - xt[0]) / (xt[xt.length - 1] - xt[0]) * pw; }
      function Y(v) { return m.t + ph - (v - yt[0]) / (yt[yt.length - 1] - yt[0]) * ph; }
      gGrid.textContent = '';
      if (!spec.hidden) {
        yt.forEach(function (t) {
          sv('line', { x1: m.l, x2: W - m.r, y1: Math.round(Y(t)) + 0.5, y2: Math.round(Y(t)) + 0.5, 'class': t === 0 ? 'zeroline' : 'gridline' }, gGrid);
          var tx = sv('text', { x: m.l - 8, y: Y(t) + 4, 'text-anchor': 'end' }, gGrid); tx.textContent = moneyShort(t);
        });
        xt.forEach(function (t) {
          var tx = sv('text', { x: X(t), y: m.t + ph + 18, 'text-anchor': 'middle' }, gGrid); tx.textContent = moneyShort(t);
        });
        var xl = sv('text', { x: m.l + pw, y: H - 4, 'text-anchor': 'end' }, gGrid); xl.textContent = 'Sales per customer';
        var yl = sv('text', { x: m.l + 4, y: m.t + 10, 'text-anchor': 'start' }, gGrid); yl.textContent = 'Profit per customer';
      }
      var cx0 = m.l + pw / 2, cy0 = m.t + ph / 2;
      // pre-guess huddle: the dots line up as a dollar sign (an S-shaped stroke with a vertical bar through it)
      var nPts = pts.length, sh = Math.min(ph * 0.92, pw * 1.2), sw = sh * 0.5;
      var nBar = Math.round(nPts * 0.16), nS = nPts - nBar;
      // S stroke as cubic bezier segments in a unit box (x: -1..1, y: -1..1, y down): top-right tip, over the top, down the spine, round the bottom, bottom-left tip
      var S_SEGS = [
        [[0.85, -0.62], [0.78, -0.95], [0.40, -1.00], [0.0, -1.00]],
        [[0.0, -1.00], [-0.55, -1.00], [-0.85, -0.78], [-0.85, -0.50]],
        [[-0.85, -0.50], [-0.85, -0.12], [-0.30, -0.06], [0.0, 0.0]],
        [[0.0, 0.0], [0.45, 0.08], [0.85, 0.14], [0.85, 0.50]],
        [[0.85, 0.50], [0.85, 0.78], [0.55, 1.00], [0.0, 1.00]],
        [[0.0, 1.00], [-0.40, 1.00], [-0.78, 0.95], [-0.85, 0.62]]
      ];
      function bez(s, t) {
        var u = 1 - t, a = u * u * u, b = 3 * u * u * t, c = 3 * u * t * t, d3 = t * t * t;
        return { x: a * s[0][0] + b * s[1][0] + c * s[2][0] + d3 * s[3][0], y: a * s[0][1] + b * s[1][1] + c * s[2][1] + d3 * s[3][1] };
      }
      // sample the stroke densely, then pick points at equal arc length so dots are evenly spaced
      var sPath = [], sLen = [0];
      S_SEGS.forEach(function (s) { for (var k = 0; k < 40; k++) { sPath.push(bez(s, k / 40)); } });
      sPath.push(bez(S_SEGS[S_SEGS.length - 1], 1));
      for (var q = 1; q < sPath.length; q++) { sLen.push(sLen[q - 1] + Math.hypot(sPath[q].x - sPath[q - 1].x, sPath[q].y - sPath[q - 1].y)); }
      function onS(f) {
        var target = f * sLen[sLen.length - 1], j = 1;
        while (j < sLen.length - 1 && sLen[j] < target) { j++; }
        var k = (target - sLen[j - 1]) / Math.max(1e-9, sLen[j] - sLen[j - 1]);
        return { x: sPath[j - 1].x + (sPath[j].x - sPath[j - 1].x) * k, y: sPath[j - 1].y + (sPath[j].y - sPath[j - 1].y) * k };
      }
      function dollarPos(i) {
        var u;
        if (i < nS) { u = onS(i / Math.max(1, nS - 1)); return { x: cx0 + u.x * sw / 2, y: cy0 + u.y * sh / 2 }; }
        var b = (i - nS) / Math.max(1, nBar - 1);
        return { x: cx0, y: cy0 - sh * 0.56 + b * sh * 1.12 };
      }
      var plan = pts.map(function (p, i) {
        var d = dots[p.key];
        var huddle = dollarPos(i);
        var target = spec.hidden ? huddle : { x: X(p.x), y: Y(p.y) };
        if (!d) { d = dots[p.key] = { c: sv('circle', { r: 4, 'class': 'mark' }, gDots), cur: huddle }; }
        if (opts.fromBase) { d.cur = huddle; }
        d.from = { x: d.cur.x, y: d.cur.y }; d.to = target; d.p = p;
        d.c.setAttribute('fill', spec.hidden ? C.placeholder : p.color);
        d.c.setAttribute('fill-opacity', spec.hidden ? 0.8 : 0.7);
        return d;
      });
      // spotlight ring + label
      gSpot.textContent = '';
      var sp = null, ring = null, slab = null;
      if (!spec.hidden && spec.spot) {
        sp = dots[spec.spot.key];
        ring = sv('circle', { r: 10, fill: 'none', stroke: '#000', 'stroke-width': 4 }, gSpot);
        slab = sv('text', { 'class': 'vlab name' }, gSpot); slab.textContent = spec.spot.label;
      }
      // nearest-dot hover
      gHit.textContent = '';
      var hit = sv('rect', { x: m.l, y: m.t, width: pw, height: ph, 'class': 'hit' }, gHit);
      hit.addEventListener('mousemove', function (e) {
        if (spec.hidden) { showTip(e, 'Customers', [['Value', 'Make a guess first']]); return; }
        var b = svg.getBoundingClientRect(), mx = e.clientX - b.left, my = e.clientY - b.top, best = null, bd = 400;
        plan.forEach(function (d) { var dx = d.to.x - mx, dy = d.to.y - my, dd = dx * dx + dy * dy; if (dd < bd) { bd = dd; best = d; } });
        if (best) { showTip(e, best.p.label, best.p.tip); } else { hideTip(); }
      });
      hit.addEventListener('mouseleave', hideTip);
      var spotFrom = spotState, spotTo = sp ? { x: sp.to.x, y: sp.to.y } : null;
      function placeSpot(t) {
        if (!ring || !spotTo) { return; }
        var s = spotFrom ? { x: lerp(spotFrom.x, spotTo.x, t), y: lerp(spotFrom.y, spotTo.y, t) } : spotTo;
        spotState = s;
        ring.setAttribute('cx', s.x); ring.setAttribute('cy', s.y);
        var right = s.x < m.l + pw * 0.7;
        slab.setAttribute('x', right ? s.x + 15 : s.x - 15); slab.setAttribute('text-anchor', right ? 'start' : 'end');
        slab.setAttribute('y', s.y + 4);
      }
      var n = plan.length;
      run(this, opts.duration == null ? 650 : opts.duration, opts.stagger == null ? 0.6 : opts.stagger, n, function (p) {
        plan.forEach(function (d, i) {
          var t = p(i), x = lerp(d.from.x, d.to.x, t), y = lerp(d.from.y, d.to.y, t);
          d.cur = { x: x, y: y }; d.c.setAttribute('cx', x); d.c.setAttribute('cy', y);
        });
        if (ring) {
          var t2 = p(n - 1);
          ring.setAttribute('opacity', spotFrom ? 1 : t2); slab.setAttribute('opacity', spotFrom ? 1 : t2);
          placeSpot(spotFrom ? p(0) : 1);
        }
      });
    };
  }

  // ------------------------------------------------------------------ facts (all computed from DATA)
  function ordersFor(year) {
    return year === 'All' ? DATA.orders : DATA.orders.filter(function (o) { return o.date.y === +year; });
  }
  function years() {
    var ys = {}; DATA.orders.forEach(function (o) { ys[o.date.y] = 1; });
    return Object.keys(ys).map(Number).sort();
  }
  function yearWords(year) { return year === 'All' ? '' : 'In ' + year + ', '; }
  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

  function factBasket() {
    var o = DATA.orders, N = o.length, L = sum(o, function (x) { return x.lines; });
    var dist = {}, maxL = 0;
    o.forEach(function (x) { dist[x.lines] = (dist[x.lines] || 0) + 1; maxL = Math.max(maxL, x.lines); });
    var keys = []; for (var k = 1; k <= maxL; k++) { keys.push(k); }
    var mode = argmax(keys, function (k) { return dist[k] || 0; });
    var share = (dist[mode] || 0) / N;
    var ys = years(), byYear = ys.map(function (y) { return { y: y, n: o.filter(function (x) { return x.date.y === y; }).length }; });
    var ctrys = {}; o.forEach(function (x) { ctrys[x.ctry] = 1; });
    return { N: N, L: L, dist: dist, keys: keys, maxL: maxL, mode: mode, share: share, byYear: byYear,
      avgLines: L / N, avgValue: sum(o, function (x) { return x.sales; }) / N, nCtry: Object.keys(ctrys).length, ys: ys };
  }
  function factDays(year) {
    var o = ordersFor(year), by = DAYS.map(function () { return { n: 0, s: 0 }; });
    o.forEach(function (x) { var w = weekday(x.date); by[w].n += 1; by[w].s += x.sales; });
    var idx = DAY_ORDER.slice();
    var minD = argmin(idx, function (i) { return by[i].n; }), maxD = argmax(idx, function (i) { return by[i].n; });
    var topAvg = argmax(idx, function (i) { return by[i].n ? by[i].s / by[i].n : 0; });
    return { by: by, minD: minD, maxD: maxD, topAvg: topAvg };
  }
  function factMonths(year) {
    var o = ordersFor(year), by = MONTHS.map(function () { return 0; });
    o.forEach(function (x) { by[x.date.m - 1] += x.sales; });
    var idx = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
    var maxM = argmax(idx, function (i) { return by[i]; }), minM = argmin(idx, function (i) { return by[i]; });
    var total = sum(by), late = by[8] + by[9] + by[10] + by[11];
    var shares = years().map(function (y) {
      var oy = ordersFor(String(y)), t = sum(oy, function (x) { return x.sales; });
      var l = sum(oy.filter(function (x) { return x.date.m >= 9; }), function (x) { return x.sales; });
      return { y: y, share: t ? l / t : 0 };
    });
    return { by: by, maxM: maxM, minM: minM, lateShare: total ? late / total : 0, shares: shares };
  }
  function factShip() {
    var rows = DATA.ship, maxDays = 0;
    rows.forEach(function (r) { maxDays = Math.max(maxDays, r.days); });
    var modes = MODES.filter(function (md) { return rows.some(function (r) { return r.mode === md; }); });
    rows.forEach(function (r) { if (modes.indexOf(r.mode) < 0) { modes.push(r.mode); } });
    var per = {};
    modes.forEach(function (md) {
      var rs = rows.filter(function (r) { return r.mode === md; }), lines = sum(rs, function (r) { return r.lines; });
      var dist = []; for (var d = 0; d <= maxDays; d++) { dist.push(0); }
      rs.forEach(function (r) { if (r.days >= 0) { dist[r.days] += r.lines; } });
      per[md] = { lines: lines, dist: dist, avg: lines ? sum(rs, function (r) { return r.days * r.lines; }) / lines : 0,
        same: lines ? dist[0] / lines : 0, late: lines - dist[0] };
    });
    var slowRows = rows.filter(function (r) { return r.days === maxDays; });
    var slowCtry = {}; slowRows.forEach(function (r) { slowCtry[r.ctry] = 1; });
    var slowOrders = sum(slowRows, function (r) { return r.orders; });
    var slowMode = {}; slowRows.forEach(function (r) { slowMode[r.mode] = 1; });
    return { per: per, modes: modes, maxDays: maxDays, slowCtry: Object.keys(slowCtry), slowOrders: slowOrders, slowMode: Object.keys(slowMode),
      sameMode: per['Same Day'] ? 'Same Day' : modes[0] };
  }
  function factDiscount() {
    var lv = {};
    DATA.lines.forEach(function (r) {
      if (!lv[r.disc]) { lv[r.disc] = { d: r.disc, profit: 0, sales: 0, n: 0, loss: 0 }; }
      var g = lv[r.disc]; g.profit += r.profit; g.sales += r.sales; g.n += 1; if (r.profit < 0) { g.loss += 1; }
    });
    var levels = Object.keys(lv).map(Number).sort(function (a, b) { return a - b; }).map(function (k) { return lv[k]; });
    var t = null;
    for (var i = levels.length - 1; i >= 0; i--) { if (levels[i].profit < 0 && levels[i].d > 0) { t = levels[i].d; } else { break; } }
    var above = levels.filter(function (g) { return t !== null && g.d >= t; });
    var below = levels.filter(function (g) { return t === null || g.d < t; });
    var worst = argmin(DATA.lines, function (r) { return r.profit; });
    return { levels: levels, t: t, prev: below.length ? below[below.length - 1].d : 0,
      lost: sum(above, function (g) { return g.profit; }), n: sum(above, function (g) { return g.n; }),
      keptProfit: sum(below, function (g) { return g.profit; }),
      lossShare: sum(above, function (g) { return g.n; }) ? sum(above, function (g) { return g.loss; }) / sum(above, function (g) { return g.n; }) : 0,
      worst: worst };
  }
  function factSubs() {
    var sb = {};
    DATA.lines.forEach(function (r) {
      if (!sb[r.sub]) { sb[r.sub] = { sub: r.sub, qty: 0, profit: 0, sales: 0 }; }
      sb[r.sub].qty += r.qty; sb[r.sub].profit += r.profit; sb[r.sub].sales += r.sales;
    });
    var subs = Object.keys(sb).map(function (k) { var g = sb[k]; g.ppu = g.qty ? g.profit / g.qty : 0; return g; });
    var top = argmax(subs, function (g) { return g.profit; });
    var fewest = argmin(subs, function (g) { return g.qty; });
    var most = argmax(subs, function (g) { return g.qty; });
    var minP = argmin(subs, function (g) { return g.ppu; });
    return { subs: subs, top: top, fewest: fewest, most: most, minP: minP };
  }
  function factCust() {
    var cs = DATA.cust, N = cs.length;
    var spend = argmax(cs, function (c) { return c.sales; }), loyal = argmax(cs, function (c) { return c.orders; }),
      loss = argmin(cs, function (c) { return c.profit; });
    var losers = cs.filter(function (c) { return c.profit < 0; }).length;
    var maxO = 0; cs.forEach(function (c) { maxO = Math.max(maxO, c.orders); });
    var dist = []; for (var k = 1; k <= maxO; k++) { dist.push(0); }
    cs.forEach(function (c) { if (c.orders >= 1) { dist[c.orders - 1] += 1; } });
    return { N: N, spend: spend, loyal: loyal, loss: loss, losers: losers, dist: dist };
  }

  // ------------------------------------------------------------------ chapters
  function chartSizes(W, kind) {
    // charts sit in cards (24px side padding, 1px border), so plot widths are the card width less 50px
    var gap = 28, pad = 50;
    if (W < 760) { return { stack: true, mainW: W - pad, compW: W - pad, h: Math.round(clamp(W / 1.8, 240, 420)), compH: Math.round(clamp(W / 1.2, 240, 420)) }; }
    var h = Math.round(clamp((W - gap) / 2.95, 280, 430)), compW = Math.round(h * 1.15);
    if (kind === 'wide') { compW = Math.round(h * 1.05); }
    var mainW = W - gap - compW - pad * 2;
    return { mainW: mainW, compW: W - gap - (mainW + pad) - pad, h: h - 70, compH: h - 70 };
  }

  function buildChapters() {
    var basket = factBasket(), ship = factShip(), disc = factDiscount(), subs = factSubs(), cust = factCust(), allDays = factDays('All'),
      allMonths = factMonths('All');
    var tot = DATA.totals;
    return [
      { // 1. basket size
        tile: 'Basket Size', eyebrow: 'Meet the store',
        question: 'How many product lines does a typical order hold?',
        choices: (function () {
          var c = [1, 2, 3, 5]; if (c.indexOf(basket.mode) < 0) { c[3] = basket.mode; }
          return c.sort(function (a, b) { return a - b; }).map(function (v) { return { key: v, label: v + ' ' + plural(v, 'line', 'lines') }; });
        })(),
        answer: basket.mode,
        tileValue: function () { return { text: pct(basket.share) }; },
        headline: function () {
          return basket.mode === 1
            ? [B(pct(basket.share)), ' of the store’s ', B(nf(basket.N)), ' orders hold just one product line.']
            : ['The most common order holds ', B(basket.mode), ' product lines: ', B(pct(basket.share)), ' of ', B(nf(basket.N)), ' orders.'];
        },
        teaser: function () { return ['Some shoppers fill a cart; most do not. Guess the typical basket, then watch the columns settle.']; },
        lede: function () {
          var two = (basket.dist[2] || 0) / basket.N, five = 0;
          basket.keys.forEach(function (k) { if (k >= 5) { five += basket.dist[k] || 0; } });
          return ['Two-line orders come next at ', B(pct(two)), '. Only ', B(pct(five / basket.N)), ' of orders hold five lines or more.'];
        },
        finding: function () {
          return { label: 'Orders with one line', value: basket.share, fmt: pct, sign: 0,
            what: ['The biggest order held ', B(basket.maxL), ' lines.'],
            why: ['An order averages ', B(basket.avgLines.toFixed(1)), ' lines and ', B(money(basket.avgValue)), ', so baskets are small.'] };
        },
        hint: function () { return [B('Try this: '), 'hover any dot or column for the exact count.']; },
        charts: function (sz, hidden) {
          var first = basket.byYear[0], last = basket.byYear[basket.byYear.length - 1], g = last.n / first.n - 1;
          return [
            { type: 'col', w: sz.mainW, h: sz.h, title: 'Orders by number of product lines', sub: 'Each dot is a basket size; most orders are small',
              spec: { lollipop: true, hidden: hidden, fmt: nf, fmtLabel: nf, items: basket.keys.map(function (k) {
                var v = basket.dist[k] || 0;
                return { key: 'l' + k, label: String(k), value: v, color: k === basket.mode ? C.focus : C.context,
                  showLabel: k === basket.mode || k === basket.maxL || k === 2, tipTitle: k + ' ' + plural(k, 'line', 'lines') + ' per order',
                  tip: [['Orders', nf(v)], ['Share', (v / basket.N * 100).toFixed(1) + '%']] };
              }) } },
            { type: 'col', w: sz.compW, h: sz.compH, title: 'Orders per year',
              sub: (g >= 0 ? 'Up ' : 'Down ') + pct(Math.abs(g)) + ' from ' + first.y + ' to ' + last.y,
              spec: { hidden: hidden, fmt: nf, fmtLabel: nf, items: basket.byYear.map(function (r, i) {
                return { key: 'y' + r.y, label: String(r.y), value: r.n, color: i === basket.byYear.length - 1 ? C.focus : C.context, showLabel: true,
                  tip: [['Orders', nf(r.n)]] };
              }) } }
          ];
        }
      },
      { // 2. quiet day
        tile: 'Quiet Day', eyebrow: 'Ordering habits',
        question: 'Which day of the week gets the fewest orders?',
        choices: DAY_ORDER.map(function (i) { return { key: i, label: DAYS[i].slice(0, 3), full: DAYS[i] }; }),
        answer: allDays.minD,
        controls: 'year',
        tileValue: function () { return { text: DAYS[allDays.minD] }; },
        headline: function () {
          var f = factDays(S.year), n = f.by[f.minD].n;
          return [yearWords(S.year) ? yearWords(S.year) + DAYS[f.minD] : DAYS[f.minD], ' gets the fewest orders, ', B(nf(n)),
            f.topAvg === f.minD ? ', yet the biggest average basket.' : '.'];
        },
        teaser: function () { return ['Orders arrive every day of the week, but not evenly. Pick the quietest day, then watch the columns settle.']; },
        lede: function () {
          var f = factDays(S.year), mx = f.by[f.maxD], mn = f.by[f.minD], ta = f.by[f.topAvg];
          var parts = [B(DAYS[f.maxD]), ' is the busiest day with ', B(nf(mx.n)), ' orders, ', B(times(mx.n, mn.n)), ' as many as ', DAYS[f.minD] + '. '];
          if (f.topAvg === f.minD) { parts.push(DAYS[f.minD] + '’s orders average ', B(money(mn.s / mn.n)), ', the highest of any day.'); }
          else { parts.push('The biggest average order lands on ', B(DAYS[f.topAvg]), ': ', B(money(ta.s / ta.n)), '.'); }
          return parts;
        },
        finding: function () {
          var f = factDays(S.year), mx = f.by[f.maxD], mn = f.by[f.minD];
          return { label: 'Orders on ' + DAYS[f.minD] + 's', value: mn.n, fmt: nf, sign: 0,
            what: ['Just ', B(pct(mn.n / Math.max(1, sum(f.by, function (b) { return b.n; })))), ' of orders land on a ' + DAYS[f.minD] + '.'],
            why: (function () {
              var lowA = argmin(DAY_ORDER.filter(function (k) { return f.by[k].n; }), function (k) { return f.by[k].s / f.by[k].n; }), lo = f.by[lowA];
              return f.topAvg === f.minD
                ? ['Fewer orders, bigger baskets: a ' + DAYS[f.minD] + ' order beats a ' + DAYS[lowA] + ' order by ', B(money(mn.s / mn.n - lo.s / lo.n)), '.']
                : ['The average ' + DAYS[f.minD] + ' order is ', B(money(mn.s / mn.n)), '.'];
            })() };
        },
        hint: function () { return [B('Try this: '), 'pick a year. Is ' + DAYS[allDays.minD] + ' quiet every year?']; },
        charts: function (sz, hidden) {
          var f = factDays(S.year), avgAll = sum(f.by, function (b) { return b.s; }) / Math.max(1, sum(f.by, function (b) { return b.n; }));
          return [
            { type: 'col', w: sz.stack ? sz.mainW : Math.round((sz.W - 28 - 100) / 2), h: sz.h, title: 'Orders by day of the week', sub: S.year === 'All' ? 'All years' : String(S.year),
              spec: { hidden: hidden, fmt: nf, fmtLabel: nf, items: DAY_ORDER.map(function (i) {
                var b = f.by[i];
                return { key: 'd' + i, label: DAYS[i].slice(0, 3), value: b.n, color: i === f.minD ? C.focus : C.context,
                  showLabel: i === f.minD || i === f.maxD, tipTitle: DAYS[i], tip: [['Orders', nf(b.n)], ['Sales', money(b.s)]] };
              }) } },
            { type: 'col', w: sz.stack ? sz.compW : Math.round((sz.W - 28 - 100) / 2), h: sz.compH, title: 'Average order value by day', sub: 'All days average ' + money(avgAll),
              spec: { hidden: hidden, fmt: moneyShort, fmtLabel: money, items: DAY_ORDER.map(function (i) {
                var b = f.by[i], v = b.n ? b.s / b.n : 0;
                return { key: 'a' + i, label: DAYS[i].slice(0, 1), value: v, color: i === f.topAvg ? C.focus : C.context,
                  showLabel: i === f.topAvg, tipTitle: DAYS[i], tip: [['Average order', money(v)], ['Orders', nf(b.n)]] };
              }) } }
          ];
        }
      },
      { // 3. rush month
        tile: 'Busy Month', eyebrow: 'Seasons',
        question: 'Which month brings in the most sales?',
        choices: (function () {
          var idx = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11].sort(function (a, b) { return allMonths.by[b] - allMonths.by[a]; }).slice(0, 4);
          return idx.sort(function (a, b) { return a - b; }).map(function (i) { return { key: i, label: MONTHS[i] }; });
        })(),
        answer: allMonths.maxM,
        controls: 'year',
        tileValue: function () { return { text: MONTHS[allMonths.maxM].slice(0, 3) }; },
        headline: function () {
          var f = factMonths(S.year);
          return [yearWords(S.year) + (S.year === 'All' ? MONTHS[f.maxM] : MONTHS[f.maxM]), ' is the busiest month: ', B(money(f.by[f.maxM])),
            ' in sales, ', B(times(f.by[f.maxM], f.by[f.minM])), ' ', MONTHS[f.minM] + '.'];
        },
        teaser: function () { return ['Twelve months, one clear winner. Pick a month, then watch the year fill in from January.']; },
        lede: function () {
          var f = factMonths(S.year);
          return ['September to December, a third of the calendar, bring ', B(pct(f.lateShare)), ' of ', S.year === 'All' ? 'all' : S.year, ' sales.'];
        },
        finding: function () {
          var f = factMonths(S.year);
          return { label: 'Sales in ' + MONTHS[f.maxM], value: f.by[f.maxM], fmt: money, sign: 0,
            what: [B(pct(f.by[f.maxM] / Math.max(1, sum(f.by)))), ' of the year’s sales in one month.'],
            why: f.by[8] > f.by[7] ? ['The rush starts early: September sales are ', B(times(f.by[8], f.by[7])), ' August’s.'] : [] };
        },
        hint: function () { return [B('Try this: '), 'pick a year and watch the months reshuffle.']; },
        charts: function (sz, hidden) {
          var f = factMonths(S.year);
          return [
            { type: 'col', w: sz.mainW, h: sz.h, title: 'Sales by month', sub: S.year === 'All' ? 'All years combined' : String(S.year),
              spec: { hidden: hidden, fmt: moneyShort, fmtLabel: moneyShort, items: MONTHS.map(function (mn, i) {
                return { key: 'm' + i, label: mn.slice(0, 3), value: f.by[i], color: i === f.maxM ? C.focus : C.context,
                  showLabel: i === f.maxM || i === f.minM, tipTitle: mn, tip: [['Sales', money(f.by[i])]] };
              }) } },
            { type: 'col', w: sz.compW, h: sz.compH, title: 'September to December share', sub: 'Of each year’s sales',
              spec: { hidden: hidden, fmt: pct, fmtLabel: pct, items: f.shares.map(function (r) {
                var on = S.year === 'All' ? r.y === argmax(f.shares, function (q) { return q.share; }).y : String(r.y) === String(S.year);
                return { key: 's' + r.y, label: String(r.y), value: r.share, color: on ? C.focus : C.context, showLabel: true,
                  tip: [['Sep to Dec share', (r.share * 100).toFixed(1) + '%']] };
              }) } }
          ];
        }
      },
      { // 4. shipping
        tile: 'Same Day', eyebrow: 'Shipping',
        question: 'How often does “Same Day” shipping really leave the same day?',
        choices: (function () {
          var a = Math.round(ship.per[ship.sameMode].same * 100), c = [50, 75, a, 100];
          var uniq = c.filter(function (v, i) { return c.indexOf(v) === i; });
          if (uniq.length < 4) { uniq.push(a === 90 ? 85 : 90); }
          return uniq.sort(function (x, y) { return x - y; }).map(function (v) { return { key: v, label: v + '%' }; });
        })(),
        answer: Math.round(ship.per[ship.sameMode].same * 100),
        controls: 'mode',
        tileValue: function () { return { text: pct(ship.per[ship.sameMode].same) }; },
        headline: function () {
          var p = ship.per[ship.sameMode];
          var out = ['“' + ship.sameMode + '” leaves the same day ', B(pct(p.same)), ' of the time'];
          return p.late > 0 ? out.concat(['; ', B(nf(p.late)), ' ' + plural(p.late, 'line', 'lines') + ' shipped later.']) : out.concat(['.']);
        },
        teaser: function () { return ['Four ship modes promise four speeds. Guess how well the fastest one keeps its word.']; },
        lede: function () {
          var slow = argmax(ship.modes, function (md) { return ship.per[md].avg; }), fast = argmin(ship.modes, function (md) { return ship.per[md].avg; });
          var fa = ship.per[fast].avg;
          return [slow + ' is the slowest mode at ', B(ship.per[slow].avg.toFixed(1)), ' days on average; ', fast, ' averages ',
            fa < 0.5 ? B('under half a day') : B(fa.toFixed(1) + ' days'), '.'];
        },
        finding: function () {
          var p = ship.per[ship.sameMode], std = ship.per['Standard Class'];
          return { label: 'Same Day, same day', value: p.same, fmt: pct, sign: 0,
            what: ship.modes.every(function (m2) { return m2 === ship.sameMode || ship.per[m2].same === 0; })
              ? ['No other ship mode ever leaves the same day.'] : ['Other modes leave the same day too, but rarely.'],
            why: ['The longest wait was ', B(ship.maxDays + ' days'), ', for ', ship.slowOrders === 1 ? 'one order' : ship.slowOrders + ' orders',
              ' shipped to ' + ship.slowCtry.join(' and ') + '.'] };
        },
        hint: function () { return [B('Try this: '), 'switch ship modes and watch the waiting time spread out.']; },
        charts: function (sz, hidden) {
          var md = S.mode, p = ship.per[md] || ship.per[ship.sameMode], keys = [];
          for (var d = 0; d <= ship.maxDays; d++) { keys.push(d); }
          var top = argmax(keys, function (k) { return p.dist[k]; });
          return [
            { type: 'col', w: sz.mainW, h: sz.h, title: 'Order lines by days to ship', sub: md + ', all years',
              spec: { lollipop: true, hidden: hidden, fmt: nf, fmtLabel: nf, items: keys.map(function (k) {
                return { key: 'k' + k, label: String(k), value: p.dist[k], color: k === top ? C.focus : C.context, showLabel: k === top || (k === ship.maxDays && p.dist[k] > 0),
                  tipTitle: k + ' ' + plural(k, 'day', 'days') + ' to ship', tip: [['Lines', nf(p.dist[k])], ['Share of ' + md, (p.lines ? p.dist[k] / p.lines * 100 : 0).toFixed(1) + '%']] };
              }) } },
            { type: 'col', w: sz.compW, h: sz.compH, title: 'Average days to ship', sub: 'By ship mode',
              spec: { hidden: hidden, fmt: function (v) { return v.toFixed(0); }, fmtLabel: function (v) { return v.toFixed(1); }, items: ship.modes.map(function (m2) {
                return { key: 'm' + m2, label: m2.split(' ')[0], value: ship.per[m2].avg, color: m2 === md ? C.focus : C.context, showLabel: true,
                  tipTitle: m2, tip: [['Average days', ship.per[m2].avg.toFixed(2)], ['Lines', nf(ship.per[m2].lines)]] };
              }) } }
          ];
        }
      },
      { // 5. discounts
        tile: 'Discounts', eyebrow: 'Discounts',
        question: 'From what discount does Superstore start losing money overall?',
        choices: (function () {
          var c = [10, 20, 30, 50]; if (disc.t !== null && c.indexOf(disc.t) < 0) { c[3] = disc.t; }
          return c.sort(function (a, b) { return a - b; }).map(function (v) { return { key: v, label: v + '% off' }; });
        })(),
        answer: disc.t,
        controls: 'disc',
        tileValue: function () { return { text: moneyShort(disc.lost), sign: -1 }; },
        headline: function () {
          if (disc.t === null) { return ['Every discount level still earns a profit overall.']; }
          return ['Every discount of ', B(disc.t + '%'), ' or more loses money: ', SG(moneyAbs(disc.lost), -1), ' across ', B(nf(disc.n)), ' lines.'];
        },
        teaser: function () { return ['Discounts win sales, but at some point they cost more than they earn. Guess where the line is.']; },
        lede: function () {
          return ['Up to ', B(disc.prev + '%'), ' off, discounts still earn a profit overall. From ', B(disc.t + '%'), ' up, ', B(pct(disc.lossShare)),
            ' of discounted lines lose money.'];
        },
        finding: function () {
          var w = disc.worst;
          return { label: 'Lost at ' + disc.t + '%+ off', value: disc.lost, fmt: money, sign: -1,
            what: ['That cancels ', B(pct(-disc.lost / Math.max(1, disc.keptProfit))), ' of the profit earned at smaller discounts.'],
            why: ['Worst single line: a ' + w.pn + ' sold at ' + w.disc + '% off for ', B(money(w.sales)), ' lost ', SG(moneyAbs(w.profit), -1), '.'] };
        },
        hint: function () { return [B('Try this: '), 'switch to profit per line to see how deep each discount cuts.']; },
        charts: function (sz, hidden) {
          var per = S.discMeasure === 'avg';
          var vals = disc.levels.map(function (g) { return per ? g.profit / g.n : g.profit; });
          var mx = Math.max.apply(null, vals), mn = Math.min.apply(null, vals);
          return [
            { type: 'col', w: sz.mainW, h: sz.h, title: per ? 'Average profit per line, by discount' : 'Total profit by discount level',
              sub: 'Discount off list price', legend: true,
              spec: { lollipop: true, hidden: hidden, fmt: per ? money : moneyShort, fmtLabel: per ? money : moneyShort, items: disc.levels.map(function (g, i) {
                var v = vals[i];
                return { key: 'g' + g.d, label: g.d + '%', value: v, color: v < 0 ? C.neg : C.pos, showLabel: v === mx || v === mn,
                  tipTitle: g.d + '% off', tip: [['Profit', money(g.profit)], ['Per line', money(g.profit / g.n)], ['Lines', nf(g.n)], ['Sales', money(g.sales)]] };
              }) } },
            { type: 'col', w: sz.compW, h: sz.compH, title: 'Share of lines that lose money', sub: 'By discount level',
              spec: { hidden: hidden, fmt: pct, fmtLabel: pct, items: disc.levels.map(function (g) {
                var s = g.n ? g.loss / g.n : 0;
                return { key: 'z' + g.d, label: String(g.d), value: s, color: disc.t !== null && g.d >= disc.t ? C.focus : C.context,
                  showLabel: g.d === disc.t || g.d === disc.prev, tipTitle: g.d + '% off', tip: [['Losing lines', nf(g.loss) + ' of ' + nf(g.n)], ['Share', (s * 100).toFixed(1) + '%']] };
              }) } }
          ];
        }
      },
      { // 6. copiers
        tile: 'Profit Per Unit', eyebrow: 'Products',
        question: 'Which sub-category earns Superstore the most profit?',
        choices: (function () {
          var c = subs.subs.slice().sort(function (a, b) { return b.qty - a.qty; }).filter(function (g) { return g.sub !== subs.top.sub; }).slice(0, 3);
          c.push(subs.top);
          return c.sort(function (a, b) { return a.sub < b.sub ? -1 : 1; }).map(function (g) { return { key: g.sub, label: g.sub }; });
        })(),
        answer: subs.top.sub,
        controls: 'sub',
        tileValue: function () { return { text: money(subs.top.ppu), sign: subs.top.ppu < 0 ? -1 : 1 }; },
        headline: function () {
          var t = subs.top;
          if (t.sub === subs.fewest.sub) {
            return [t.sub + ' sell the fewest units of any sub-category, ', B(nf(t.qty)), ', yet earn the most profit: ', SG(moneyAbs(t.profit), 1), '.'];
          }
          return [t.sub + ' earn the most profit: ', SG(moneyAbs(t.profit), 1), ' from ', B(nf(t.qty)), ' units.'];
        },
        teaser: function () { return ['Seventeen sub-categories, from paper clips to copiers. Guess which one makes the most money.']; },
        lede: function () {
          var t = subs.top, mo = subs.most;
          var allQ = sum(subs.subs, function (g) { return g.qty; }), allP = sum(subs.subs, function (g) { return g.profit; }), uq = t.qty / allQ;
          return [t.sub + ' bring ', B(pct(t.profit / allP)), ' of all profit from ', B(uq < 0.01 ? 'under 1%' : pct(uq)), ' of units sold. ',
            mo.sub + ' move the most units, ', B(nf(mo.qty)), ', but earn ', B(money2(mo.ppu)), ' each.'];
        },
        finding: function () {
          var t = subs.top, mn = subs.minP;
          return { label: 'Profit Per Unit', value: t.ppu, fmt: money, sign: t.ppu < 0 ? -1 : 1,
            what: ['for every unit of ' + t.sub + ' sold.'],
            why: mn.ppu < 0 ? [mn.sub + ' lose ', SG(money2(Math.abs(mn.ppu)), -1), ' on every unit sold.'] : [] };
        },
        hint: function () { return [B('Try this: '), 'sort by profit and watch ' + subs.fewest.sub + ' race up the list.']; },
        charts: function (sz, hidden) {
          var key = S.subSort, list = subs.subs.slice().sort(function (a, b) { return b[key] - a[key]; });
          var mainW = Math.round((sz.W - 28 - 100) * 0.58), compW = sz.W - 28 - 100 - mainW;
          if (sz.stack) { mainW = sz.W - 50; compW = sz.W - 50; }
          function tipOf(g) { return [['Units', nf(g.qty)], ['Profit', money(g.profit)], ['Profit Per Unit', money2(g.ppu)], ['Sales', money(g.sales)]]; }
          return [
            { type: 'hbar', w: mainW, title: key === 'qty' ? 'Units sold by sub-category' : 'Profit by sub-category', sub: 'Sorted by ' + (key === 'qty' ? 'units sold' : 'profit'),
              spec: { nameW: 112, hidden: hidden, fmt: key === 'qty' ? nf : moneyShort, fmtLabel: key === 'qty' ? nf : money, items: list.map(function (g) {
                var v = g[key];
                return { key: g.sub, label: g.sub, value: v, bold: g.sub === subs.top.sub,
                  color: key === 'qty' ? (g.sub === subs.top.sub ? C.focus : C.context) : (v < 0 ? C.neg : C.pos), tip: tipOf(g) };
              }) } },
            { type: 'hbar', w: compW, title: 'Profit Per Unit Sold', sub: 'Same rows, same order',
              spec: { hidden: hidden, fmt: money, fmtLabel: money2, items: list.map(function (g) {
                return { key: g.sub, label: g.sub, value: g.ppu, color: g.ppu < 0 ? C.neg : C.pos, tip: tipOf(g) };
              }) } }
          ];
        }
      },
      { // 7. customers
        tile: 'Top Spender', eyebrow: 'Customers',
        question: 'Superstore’s biggest spender: did they make the store money?',
        choices: [{ key: 'lots', label: 'Yes, a lot' }, { key: 'little', label: 'Yes, a little' }, { key: 'no', label: 'No, a loss' }],
        answer: cust.spend.profit < 0 ? 'no' : (cust.spend.profit < cust.spend.sales * 0.05 ? 'little' : 'lots'),
        controls: 'spot',
        tileValue: function () { return { text: money(cust.spend.profit), sign: cust.spend.profit < 0 ? -1 : 1 }; },
        headline: function () {
          var c = cust.spend;
          return c.profit < 0
            ? ['The top spender, ' + c.nm + ', bought ', B(money(c.sales)), ' but cost the store ', SG(moneyAbs(c.profit), -1), ' in profit.']
            : ['The top spender, ' + c.nm + ', bought ', B(money(c.sales)), ' and earned the store ', SG(moneyAbs(c.profit), 1), '.'];
        },
        teaser: function () { return ['Every dot will be one of ' + nf(cust.N) + ' customers. Guess first, then watch them spread out by sales and profit.']; },
        lede: function () {
          return [B(nf(cust.losers)), ' of ', B(nf(cust.N)), ' customers lost the store money overall. ', cust.loyal.nm + ' placed the most orders, ', B(cust.loyal.orders), '; ',
            cust.loss.nm + ' cost the store the most, ', SG(moneyAbs(cust.loss.profit), -1), '.'];
        },
        finding: function () {
          var c = cust.spend;
          return { label: 'Top Spender’s Profit', value: c.profit, fmt: money, sign: c.profit < 0 ? -1 : 1,
            what: ['A profit margin of ', SG((Math.abs(c.profit / c.sales) * 100).toFixed(1) + '%', c.profit < 0 ? -1 : 1), ' on everything they bought.'],
            why: (function () {
              var top10 = DATA.cust.slice().sort(function (a, b) { return b.sales - a.sales; }).slice(0, 10);
              var k = top10.filter(function (x) { return x.profit < 0; }).length;
              return [B(k === 1 ? 'Just 1' : String(k)), ' of the 10 biggest spenders ', k === 1 ? 'costs' : 'cost', ' the store money.'];
            })() };
        },
        hint: function () { return [B('Try this: '), 'switch the spotlight.']; },
        spotlight: function () { return S.spot === 'orders' ? cust.loyal : (S.spot === 'loss' ? cust.loss : cust.spend); },
        charts: function (sz, hidden) {
          var sp = this.spotlight();
          var dist = cust.dist;
          return [
            { type: 'scatter', w: sz.mainW, h: sz.h, title: 'Every customer: sales against profit', sub: 'One dot per customer', legend: true,
              spec: { hidden: hidden, spot: { key: sp.cid, label: sp.nm }, points: cust.N ? DATA.cust.map(function (c) {
                return { key: c.cid, x: c.sales, y: c.profit, color: c.profit < 0 ? C.neg : C.pos, label: c.nm,
                  tip: [['Sales', money(c.sales)], ['Profit', money(c.profit)], ['Orders', nf(c.orders)]] };
              }) : [] } },
            { type: 'col', w: sz.compW, h: sz.compH, title: 'Customers by number of orders', sub: 'Highlighted: ' + sp.nm + '’s ' + sp.orders,
              spec: { hidden: hidden, fmt: nf, fmtLabel: nf, items: dist.map(function (v, i) {
                return { key: 'o' + (i + 1), label: (i + 1) % 2 === 1 ? String(i + 1) : '', value: v, color: i + 1 === sp.orders ? C.focus : C.context,
                  showLabel: i + 1 === sp.orders, tipTitle: (i + 1) + ' ' + plural(i + 1, 'order', 'orders'), tip: [['Customers', nf(v)]] };
              }) } }
          ];
        }
      }
    ];
  }

  // ------------------------------------------------------------------ rendering (VizSight markup)
  var UI = {};
  function score() {
    var got = 0, done = 0;
    CH.forEach(function (c, i) { if (S.revealed[i]) { done += 1; if (S.guess[i] !== undefined && S.guess[i] === c.answer) { got += 1; } } });
    return { got: got, done: done };
  }
  // VizSight headings state the finding as a sentence; no decorative full-stop treatment.
  function headingDot(node) { return node; }
  function eyebrow(text) {
    var e = el('div', 'eyebrow'); e.appendChild(document.createTextNode(text)); return e;
  }
  function renderShell() {
    root.textContent = '';
    // hero: the one ink band. Brand bar, headline, dataset in one sentence, score, fact tiles.
    var hero = el('div', 'hero');
    var heroWrap = el('div', 'wrap'); hero.appendChild(heroWrap);
    var brand = el('header', 'brand');
    var img = el('img', 'brand-mark'); img.src = 'src/vizsight-logo-white.svg'; img.alt = 'VizSight'; brand.appendChild(img);
    brand.appendChild(el('span', 'brand-name', 'VizSight'));
    var app = el('span', 'brand-app'); app.appendChild(el('b', null, 'Superstore')); app.appendChild(document.createTextNode(' Fun Facts')); brand.appendChild(app);
    brand.appendChild(el('span', 'brand-space'));
    var hh = String(loadedAt.getHours()).padStart(2, '0'), mm = String(loadedAt.getMinutes()).padStart(2, '0');
    brand.appendChild(el('span', 'brand-source', DS_NAME + (window.FF_SNAPSHOT ? ' · snapshot' : ' · live at ' + hh + ':' + mm)));
    var tt = el('button', 'theme-toggle'); tt.type = 'button';
    function syncToggle() { var d = document.documentElement.getAttribute('data-theme') === 'dark'; tt.textContent = d ? 'Light mode' : 'Dark mode'; tt.setAttribute('aria-pressed', d ? 'true' : 'false'); }
    tt.addEventListener('click', function () { setTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'); syncToggle(); });
    syncToggle(); brand.appendChild(tt);
    heroWrap.appendChild(brand);
    var top = el('section', 'top');
    var row = el('div', 'top-row');
    var left = el('div');
    left.appendChild(eyebrow('A guessing game in seven facts'));
    left.appendChild(el('h1', null, 'Seven fun facts hiding in Superstore'));
    var basket = factBasket(), tot = DATA.totals;
    left.appendChild(rich(el('p', 'intro'), ['The file holds ', B(nf(basket.L)), ' order lines: ', B(nf(basket.N)), ' orders from ', B(nf(tot.customers)),
      ' customers buying ', B(nf(tot.products)), ' products in ', B(tot.states), ' states and provinces, ', B(basket.ys[0]), ' to ', B(basket.ys[basket.ys.length - 1]),
      '. Guess each fact, then watch the charts reveal the answer.']));
    row.appendChild(left);
    var sc = el('div', 'score'); sc.appendChild(el('span', 'label', 'Your score'));
    UI.score = el('div', 'score-row'); sc.appendChild(UI.score); row.appendChild(sc);
    top.appendChild(row);
    heroWrap.appendChild(top);
    root.appendChild(hero);
    // fact tiles sit outside the hero, on the white surface
    var tilesBand = el('div', 'tiles-band'), tilesWrap = el('div', 'wrap');
    UI.tiles = el('nav', 'tiles'); UI.tiles.setAttribute('aria-label', 'Facts');
    tilesWrap.appendChild(UI.tiles); tilesBand.appendChild(tilesWrap);
    root.appendChild(tilesBand);
    // finding band (surface) and ink footer with the source and method note
    var band = el('main', 'body');
    UI.stage = el('div', 'wrap'); band.appendChild(UI.stage);
    root.appendChild(band);
    var foot = el('footer', 'site-footer'), fw = el('div', 'wrap');
    var fi = el('img'); fi.src = 'src/vizsight-logo-white.svg'; fi.alt = ''; fw.appendChild(fi);
    fw.appendChild(el('span', null, 'Designed by Chimdi Nwosu. Source: ' + DS_NAME + ', queried live at ' + hh + ':' + mm +
      '. Every number is counted from the order lines; “busiest”, “quietest” and “top” compare the totals shown in each chart.'));
    foot.appendChild(fw); root.appendChild(foot);
    document.body.appendChild(tip);
  }
  function renderScore(popIt) {
    var s = score();
    UI.score.textContent = '';
    var dots = el('span', 'score-dots');
    CH.forEach(function (c, i) { dots.appendChild(el('i', S.revealed[i] ? (S.guess[i] === c.answer ? 'got' : 'miss') : '')); });
    UI.score.appendChild(dots);
    UI.score.appendChild(el('span', null, s.got + ' of ' + CH.length + ' right'));
    if (s.done > 0) {
      var again = el('button', 'skip', 'Start over'); again.type = 'button';
      again.addEventListener('click', startOver);
      UI.score.appendChild(again);
    }
    if (popIt) { UI.score.classList.remove('pop'); void UI.score.offsetWidth; UI.score.classList.add('pop'); }
  }
  function renderTiles(flipIdx) {
    UI.tiles.textContent = '';
    // markers pop in once, on first load only
    UI.tiles.classList.toggle('settled', !!UI.tilesDrawn); UI.tilesDrawn = true;
    CH.forEach(function (c, i) {
      var b = el('button', 'tile' + (S.revealed[i] ? ' done' : '')); b.type = 'button';
      if (i === S.ch) { b.setAttribute('aria-current', 'step'); }
      var hd = el('span', 'tile-head'); hd.appendChild(el('span', 'tile-n', S.revealed[i] ? '✓' : String(i + 1))); hd.appendChild(el('span', 'tile-name', c.tile));
      b.appendChild(hd);
      var v = el('span', 'tile-big');
      if (S.revealed[i]) {
        var tv = c.tileValue(); v.textContent = tv.text;
        if (tv.sign) { v.classList.add(tv.sign < 0 ? 'neg' : 'pos'); b.setAttribute('data-tone', tv.sign < 0 ? 'neg' : 'pos'); }
        if (i === flipIdx) { v.classList.add('flip'); }
      } else { v.textContent = '?'; v.classList.add('hidden'); }
      b.appendChild(v);
      b.setAttribute('aria-label', 'Fact ' + (i + 1) + ': ' + c.tile + (S.revealed[i] ? ', ' + c.tileValue().text : ', not revealed yet'));
      b.addEventListener('click', function () { go(i); });
      UI.tiles.appendChild(b);
    });
  }

  var charts = [];
  function go(i) {
    if (i < 0 || i >= CH.length || i === S.ch && UI.card) { return; }
    var dir = i >= S.ch ? 'next' : 'prev';
    S.ch = i;
    renderTiles();
    renderChapter(dir);
  }

  function renderChapter(dir) {
    var c = CH[S.ch], i = S.ch, revealed = !!S.revealed[i];
    UI.stage.textContent = '';
    var card = el('section', 'card enter-' + (dir || 'next'));
    UI.card = card;
    var head = el('div', 'head');
    var left = el('div');
    left.appendChild(eyebrow('Fact ' + (i + 1) + ' of ' + CH.length + ' · ' + c.eyebrow));
    UI.h2 = el('h2'); left.appendChild(UI.h2);
    UI.lede = el('p', 'lede'); left.appendChild(UI.lede);
    head.appendChild(left);
    UI.finding = el('aside', 'finding'); head.appendChild(UI.finding);
    card.appendChild(head);
    var quiz = el('div', 'quiz');
    UI.qrow = el('div', 'qrow'); quiz.appendChild(UI.qrow);
    UI.verdictCol = el('div'); quiz.appendChild(UI.verdictCol);
    card.appendChild(quiz);
    var ctl = el('div', 'controls');
    UI.ctlLeft = el('div', 'ctl'); ctl.appendChild(UI.ctlLeft);
    UI.hint = el('div', 'hint'); ctl.appendChild(UI.hint);
    card.appendChild(ctl);
    UI.charts = el('div', 'charts'); card.appendChild(UI.charts);
    var foot = el('div', 'footer');
    foot.appendChild(el('span'));
    var nb = el('div', 'navbtns');
    var prev = el('button', 'btn', 'Previous'); prev.type = 'button'; prev.disabled = i === 0; prev.addEventListener('click', function () { go(i - 1); });
    UI.next = el('button', 'btn' + (revealed ? ' primary' : ''), i === CH.length - 1 ? 'Back to fact 1' : 'Next fact'); UI.next.type = 'button';
    UI.next.addEventListener('click', function () { go(i === CH.length - 1 ? 0 : i + 1); });
    nb.appendChild(prev); nb.appendChild(UI.next); foot.appendChild(nb);
    card.appendChild(foot);
    UI.stage.appendChild(card);

    renderQuiz();
    renderText(false);
    renderControls();
    buildCharts();
    drawCharts({ fromBase: true });
  }

  function renderText(animateIt) {
    var c = CH[S.ch], revealed = !!S.revealed[S.ch];
    if (revealed) { headingDot(rich(UI.h2, c.headline())); } else { UI.h2.textContent = c.question; }
    rich(UI.lede, revealed ? c.lede() : c.teaser());
    if (animateIt) { [UI.h2, UI.lede].forEach(function (n) { n.classList.remove('swap'); void n.offsetWidth; n.classList.add('swap'); }); }
    UI.finding.textContent = '';
    var f = c.finding();
    UI.finding.appendChild(el('div', 'label', revealed ? f.label : 'The answer'));
    var big = el('div', 'v' + (revealed ? '' : ' hidden'));
    UI.finding.appendChild(big);
    if (revealed) {
      if (f.sign) { big.classList.add(f.sign < 0 ? 'neg' : 'pos'); }
      if (animateIt) { countUp(big, f.value, f.fmt, 700); } else { big.textContent = f.fmt(f.value); }
      UI.finding.appendChild(rich(el('p', 'c'), f.what));
      if (f.why && f.why.length) {
        var why = el('div', 'why'); why.appendChild(el('span', 'label', 'Why it matters'));
        var wp = el('p'); wp.style.margin = '0'; why.appendChild(rich(wp, f.why)); UI.finding.appendChild(why);
      }
    } else {
      big.textContent = '?';
      UI.finding.appendChild(el('p', 'c', 'Pick an answer below. The number, and the charts behind it, appear as soon as you do.'));
    }
    UI.hint.textContent = '';
    UI.hint.appendChild(el('span', 'label', revealed ? 'Try this' : 'Guess first'));
    var hp = el('p'); hp.style.margin = '0';
    rich(hp, revealed ? c.hint().slice(1).map(function (x, k) { return k === 0 && typeof x === 'string' ? cap(x) : x; })
      : ['The charts wait in grey until you pick.']);
    UI.hint.appendChild(hp);
  }

  function renderQuiz() {
    var c = CH[S.ch], i = S.ch, revealed = !!S.revealed[i];
    UI.qrow.textContent = '';
    UI.qrow.appendChild(el('span', 'label', 'Your guess'));
    c.choices.forEach(function (ch) {
      var b = el('button', 'choice', ch.label); b.type = 'button';
      if (revealed) {
        b.disabled = true;
        if (ch.key === c.answer) { b.classList.add('right'); b.appendChild(el('span', 'tick', '✓')); }
        else if (ch.key === S.guess[i]) { b.classList.add('wrong'); }
        else { b.classList.add('faded'); }
      } else {
        b.addEventListener('click', function () { reveal(ch.key); });
      }
      UI.qrow.appendChild(b);
    });
    if (!revealed) {
      var skip = el('button', 'skip', 'Just show me'); skip.type = 'button';
      skip.addEventListener('click', function () { reveal(undefined); });
      UI.qrow.appendChild(skip);
      UI.qrow.appendChild(el('span', 'kbd-hint', 'Tip: ← and → move between facts'));
    }
    UI.verdictCol.textContent = '';
    UI.verdict = null;
    if (revealed) {
      var right = S.guess[i] === c.answer, skipped = S.guess[i] === undefined;
      var ans0 = c.choices.filter(function (x) { return x.key === c.answer; })[0], label = ans0.full || ans0.label;
      var v = el('p', 'verdict');
      v.appendChild(el('span', 'badge ' + (right ? 'right' : 'miss'), right ? 'Right' : (skipped ? 'Skipped' : 'Not quite')));
      v.appendChild(document.createTextNode((right ? 'Spot on: ' : 'The answer: ') + label + '.'));
      UI.verdictCol.appendChild(v);
      UI.verdict = v;
    }
  }

  function reveal(key) {
    var i = S.ch;
    if (S.revealed[i]) { return; }
    S.guess[i] = key; S.revealed[i] = true;
    saveProgress();
    renderQuiz();
    if (UI.verdict) { UI.verdict.classList.add('in'); }
    renderText(true);
    renderControls();
    renderTiles(i);
    renderScore(true);
    UI.next.classList.add('primary');
    drawCharts({ stagger: 40, duration: 650 });
  }

  // Progress survives a reload; the game can be replayed with Start over.
  function saveProgress() {
    try { localStorage.setItem('ff-progress', JSON.stringify({ guess: S.guess, revealed: S.revealed, ch: S.ch })); } catch (e) { /* storage may be blocked in the extension frame */ }
  }
  function loadProgress() {
    try {
      var p = JSON.parse(localStorage.getItem('ff-progress') || 'null');
      if (p && p.revealed) { S.guess = p.guess || {}; S.revealed = p.revealed; S.ch = Math.min(CH.length - 1, Math.max(0, p.ch || 0)); }
    } catch (e) { /* ignore a corrupt save */ }
  }
  function startOver() {
    S.guess = {}; S.revealed = {}; S.ch = 0;
    try { localStorage.removeItem('ff-progress'); } catch (e) { /* storage may be blocked */ }
    renderScore(false); renderTiles(); renderChapter('prev');
  }

  // Choice controls: ghost buttons, the selected option in ink fill.
  function tabs(options, current, onPick) {
    var t = el('div', 'seg'); t.setAttribute('role', 'group');
    options.forEach(function (o) {
      var on = String(o.key) === String(current);
      var b = el('button', null, o.label); b.type = 'button'; b.setAttribute('aria-pressed', String(on));
      b.addEventListener('click', function () { onPick(o.key); });
      t.appendChild(b);
    });
    return t;
  }
  function renderControls() {
    var c = CH[S.ch], revealed = !!S.revealed[S.ch];
    UI.ctlLeft.textContent = '';
    if (!revealed || !c.controls) { return; }
    function refresh() { renderControls(); renderText(false); drawCharts({ stagger: 25, duration: 550 }); }
    if (c.controls === 'year') {
      UI.ctlLeft.appendChild(tabs([{ key: 'All', label: 'All years' }].concat(years().map(function (y) { return { key: String(y), label: String(y) }; })), S.year,
        function (k) { S.year = k; refresh(); }));
    } else if (c.controls === 'mode') {
      var ship = factShip();
      UI.ctlLeft.appendChild(tabs(ship.modes.map(function (m) { return { key: m, label: m }; }), S.mode, function (k) { S.mode = k; refresh(); }));
      var p = ship.per[S.mode];
      if (p) { UI.ctlLeft.appendChild(rich(el('span', 'ctl-note'), [B(p.avg < 0.5 ? 'Under half a day' : p.avg.toFixed(1) + ' days'), ' on average, ', B(pct(p.same)), ' same day.'])); }
    } else if (c.controls === 'disc') {
      UI.ctlLeft.appendChild(tabs([{ key: 'total', label: 'Total profit' }, { key: 'avg', label: 'Profit per line' }], S.discMeasure, function (k) { S.discMeasure = k; refresh(); }));
    } else if (c.controls === 'sub') {
      UI.ctlLeft.appendChild(tabs([{ key: 'qty', label: 'Sort by units sold' }, { key: 'profit', label: 'Sort by profit' }], S.subSort, function (k) { S.subSort = k; refresh(); }));
    } else if (c.controls === 'spot') {
      UI.ctlLeft.appendChild(tabs([{ key: 'spend', label: 'Top Spender' }, { key: 'orders', label: 'Most orders' }, { key: 'loss', label: 'Biggest loss' }], S.spot,
        function (k) { S.spot = k; refresh(); }));
      var sp = c.spotlight();
      UI.ctlLeft.appendChild(rich(el('span', 'ctl-note'), [sp.nm + ': ', B(sp.orders), ' orders, ', B(money(sp.sales)), ' sales, ',
        SG(moneyAbs(sp.profit), sp.profit < 0 ? -1 : 1), ' profit.']));
    }
  }

  function contentWidth() { return UI.charts.clientWidth; }
  function buildCharts() {
    UI.charts.textContent = '';
    charts = [];
    var specs = currentSpecs();
    specs.forEach(function (s) {
      var box = el('div', 'chart');
      box.appendChild(el('p', 'ctitle', s.title));
      box.appendChild(el('p', 'csub', s.sub || ' '));
      if (s.legend) {
        var lg = el('div', 'legend');
        [['Profit', C.pos], ['Loss', C.neg]].forEach(function (p) { var sp = el('span'); var i0 = el('i'); i0.style.background = p[1]; sp.appendChild(i0); sp.appendChild(document.createTextNode(p[0])); lg.appendChild(sp); });
        box.appendChild(lg);
      }
      var plot = el('div'); box.appendChild(plot);
      UI.charts.appendChild(box);
      var inst = s.type === 'hbar' ? new HBarChart(plot) : (s.type === 'scatter' ? new ScatterChart(plot) : new ColumnChart(plot));
      charts.push({ box: box, inst: inst, sub: box.children[1], title: box.children[0] });
    });
    // every chart in the row gets the same header height so plots line up
  }
  function currentSpecs() {
    var c = CH[S.ch], W = UI.charts.clientWidth || 1000, sz = chartSizes(W, c.controls === 'spot' ? 'wide' : 'pair');
    sz.W = W;
    return c.charts(sz, !S.revealed[S.ch]);
  }
  function drawCharts(opts) {
    var specs = currentSpecs();
    specs.forEach(function (s, i) {
      var ch = charts[i]; if (!ch) { return; }
      ch.title.textContent = s.title; ch.sub.textContent = s.sub || ' ';
      ch.box.style.width = (s.w + 50) + 'px';
      var o = {}; for (var k in opts) { if (opts.hasOwnProperty(k)) { o[k] = opts[k]; } }
      var spec = s.spec; spec.w = s.w; spec.h = s.h;
      ch.inst.update(spec, o);
    });
    // a legend row exists on some charts only: pad the others so plots start on one line
    var tops = charts.map(function (ch) { return ch.box.querySelector('svg').getBoundingClientRect().top; });
    var maxTop = Math.max.apply(null, tops);
    charts.forEach(function (ch, i) {
      var svg = ch.box.querySelector('svg');
      svg.style.marginTop = (parseFloat(svg.style.marginTop) || 0) + (maxTop - tops[i]) + 'px';
    });
  }

  // ------------------------------------------------------------------ boot
  function setTheme(mode) {
    if (mode === 'dark') { document.documentElement.setAttribute('data-theme', 'dark'); } else { document.documentElement.removeAttribute('data-theme'); }
    try { localStorage.setItem('ff-theme', mode); } catch (e) { /* storage may be blocked in the extension frame */ }
    readTokens();
    // chart colours are baked into the SVG at build time: rebuild the current chapter so they pick up the new tokens
    if (UI.charts && CH) { renderChapter('next'); }
  }
  function readTokens() {
    var cs = getComputedStyle(document.documentElement);
    var map = { focus: 'data', context: 'data-neutral', pos: 'teal', neg: 'negative', placeholder: 'placeholder',
      ink: 'ink', muted: 'muted' };
    Object.keys(map).forEach(function (k) { C[k] = cs.getPropertyValue('--' + map[k]).trim(); });
  }
  function renderLoading(step) {
    root.textContent = '';
    var l = el('div', 'loading');
    var bars = el('div', 'bars'); for (var i = 0; i < 4; i++) { bars.appendChild(el('i')); } l.appendChild(bars);
    l.appendChild(el('div', null, 'Counting the ' + step + '…'));
    root.appendChild(l);
  }
  function renderError(msg, step) {
    root.textContent = '';
    var e = el('div', 'error');
    e.appendChild(el('b', null, step ? 'Could not load ' + step + '.' : 'Live query unavailable.'));
    e.appendChild(document.createTextNode(msg));
    root.appendChild(e);
  }

  function loadAll(ds) {
    var chain = Promise.resolve();
    QUERIES.forEach(function (q) {
      chain = chain.then(function () {
        renderLoading(q[1]);
        return runFirst(ds, q[2]).then(function (rows) { DATA[q[0]] = normalize(q[0], rows); }).catch(function (err) {
          var e = new Error(errText(err)); e.step = q[1]; throw e;
        });
      });
    });
    return chain;
  }

  function start() {
    CH = buildChapters();
    loadProgress();
    renderShell();
    renderScore(false);
    renderTiles();
    renderChapter('next');
    var rt = null;
    window.addEventListener('resize', function () {
      clearTimeout(rt);
      rt = setTimeout(function () { if (UI.charts) { buildCharts(); drawCharts({ duration: 0 }); } }, 150);
    });
    document.addEventListener('keydown', function (e) {
      if (e.target && /input|textarea|select/i.test(e.target.tagName)) { return; }
      if (e.key === 'ArrowRight') { go(Math.min(CH.length - 1, S.ch + 1)); }
      if (e.key === 'ArrowLeft') { go(Math.max(0, S.ch - 1)); }
    });
    window.__FF = { S: S, CH: CH, go: go, reveal: reveal, DATA: DATA };
  }

  function boot() {
    if (!root) { return; }
    if (window.FF_STILL) { document.documentElement.classList.add('still'); }
    try { if (localStorage.getItem('ff-theme') === 'dark') { document.documentElement.setAttribute('data-theme', 'dark'); } } catch (e) { /* default to light */ }
    readTokens();
    if (!window.tableau || !tableau.extensions) { renderError('The Tableau Extensions API did not load.'); return; }
    renderLoading('orders');
    tableau.extensions.initializeAsync().then(function () { return getDataSources(); }).then(function (list) {
      var ds = findDataSource(list);
      if (!ds) { throw new Error('No data source found in the workbook.'); }
      return loadAll(ds);
    }).then(function () {
      if (!DATA.orders || !DATA.orders.length) { throw new Error('The orders query returned no rows.'); }
      start();
    }).catch(function (err) {
      renderError(errText(err), err && err.step);
    });
  }

  boot();
})();
