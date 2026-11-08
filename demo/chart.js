// Charts for results: picks a sensible chart from the result's columns and draws it as plain SVG (no libraries).
// pickChart() is pure (tested in Node). drawChart() builds the SVG in the page.
(function (root) {
  const NUM = /^-?\d+(\.\d+)?$/;
  const isNum = (v) => typeof v === 'number' || (typeof v === 'string' && NUM.test(v.trim()));
  const num = (v) => (typeof v === 'number' ? v : parseFloat(v));
  const DATEISH = /^\d{4}([-/]\d{1,2}([-/]\d{1,2})?)?([T ].*)?$/;
  const MAX_BARS = 30;

  function pickChart(columns, rows) {
    if (!Array.isArray(columns) || !Array.isArray(rows) || rows.length < 2 || columns.length < 2) return null;
    const colVals = (i) => rows.map((r) => r[i]);
    const numeric = columns.map((_, i) => { const v = colVals(i).filter((x) => x !== null && x !== undefined && x !== ''); return v.length > 0 && v.every(isNum); });
    const isId = (n) => /(^|_)(id|key|code|no|number)$/i.test(String(n)) || /^id$/i.test(String(n));
    const measures = columns.map((c, i) => i).filter((i) => numeric[i] && !isId(columns[i]));
    const labelIdx = columns.findIndex((c, i) => !numeric[i] || isId(c));
    if (!measures.length || labelIdx < 0 || measures.includes(labelIdx)) return null;
    const labels = colVals(labelIdx).map((v) => (v === null || v === undefined ? '' : String(v)));
    if (new Set(labels).size < labels.length) return null; // repeated labels: a chart would mislead
    const dateLike = labels.length >= 3 && labels.every((l) => DATEISH.test(l));
    const use = measures.slice(0, 3);
    const cap = dateLike ? 200 : MAX_BARS;
    const cut = rows.length > cap;
    const take = (arr) => arr.slice(0, cap);
    return {
      type: dateLike ? 'line' : 'bar',
      labelName: columns[labelIdx],
      labels: take(labels),
      series: use.map((i) => ({ name: String(columns[i]), values: take(colVals(i)).map((v) => (v === null || v === undefined || v === '' ? 0 : num(v))) })),
      note: cut ? `Showing the first ${cap} of ${rows.length} rows.` : '',
    };
  }

  const compact = (n) => {
    const a = Math.abs(n);
    if (a >= 1e9) return `${+(n / 1e9).toFixed(1)}B`;
    if (a >= 1e6) return `${+(n / 1e6).toFixed(1)}M`;
    if (a >= 1e4) return `${+(n / 1e3).toFixed(1)}K`;
    return String(+n.toFixed(2));
  };
  const niceMax = (v) => { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))); const f = v / p; return ([1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10].find((k) => f <= k + 1e-9) || 10) * p; };
  const short = (s, n) => (s.length > n ? `${s.slice(0, n - 1)}…` : s);

  function drawChart(host, spec) {
    const NS = 'http://www.w3.org/2000/svg';
    const el = (name, attrs, text) => { const e = document.createElementNS(NS, name); for (const k in attrs) e.setAttribute(k, attrs[k]); if (text !== undefined) e.textContent = text; return e; };
    host.textContent = '';
    const COLORS = ['var(--accent)', 'var(--ok)', 'var(--warn)'];
    const n = spec.labels.length;
    const horizontal = spec.type === 'bar' && (n > 8 || spec.labels.some((l) => l.length > 10)) && spec.series.length === 1;
    const W = 860;
    const max = niceMax(Math.max(0, ...spec.series.flatMap((s) => s.values)));
    const min = Math.min(0, ...spec.series.flatMap((s) => s.values));
    const span = max - min || 1;
    let H; let svg;
    if (horizontal) {
      const rowH = 30; const left = 170; const right = 60; const top = 8;
      H = top + n * rowH + 30; svg = el('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': `Bar chart of ${spec.series[0].name} by ${spec.labelName}` });
      const x = (v) => left + ((v - min) / span) * (W - left - right);
      for (let g = 0; g <= 4; g++) { const gv = min + (span * g) / 4; svg.append(el('line', { x1: x(gv), x2: x(gv), y1: top, y2: top + n * rowH, class: 'ch-grid' }), el('text', { x: x(gv), y: H - 8, class: 'ch-tick', 'text-anchor': 'middle' }, compact(gv))); }
      spec.labels.forEach((l, i) => {
        const y = top + i * rowH; const v = spec.series[0].values[i];
        const lab = el('text', { x: left - 10, y: y + rowH / 2 + 4, class: 'ch-lab', 'text-anchor': 'end' }, short(l, 24)); lab.append(el('title', {}, l));
        const bar = el('rect', { x: Math.min(x(0), x(v)), y: y + 5, width: Math.max(2, Math.abs(x(v) - x(0))), height: rowH - 10, rx: 4, fill: COLORS[0], class: 'ch-bar' }); bar.append(el('title', {}, `${l}: ${v}`));
        svg.append(lab, bar, el('text', { x: x(v) + 6, y: y + rowH / 2 + 4, class: 'ch-val' }, compact(v)));
      });
    } else {
      const left = 52; const right = 16; const top = 12; const bottom = spec.type === 'line' || n > 8 ? 64 : 40; H = 340;
      svg = el('svg', { viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': `${spec.type === 'line' ? 'Line' : 'Bar'} chart by ${spec.labelName}` });
      const plotW = W - left - right; const plotH = H - top - bottom;
      const y = (v) => top + plotH - ((v - min) / span) * plotH;
      for (let g = 0; g <= 4; g++) { const gv = min + (span * g) / 4; svg.append(el('line', { x1: left, x2: W - right, y1: y(gv), y2: y(gv), class: 'ch-grid' }), el('text', { x: left - 8, y: y(gv) + 4, class: 'ch-tick', 'text-anchor': 'end' }, compact(gv))); }
      const step = plotW / n; const every = Math.ceil(n / 14);
      const sx = (i) => left + step * i + step / 2;
      const tickLabel = (i) => { if (i % every) return; const t = el('text', { x: sx(i), y: H - bottom + 18, class: 'ch-tick', 'text-anchor': n > 8 || spec.type === 'line' ? 'end' : 'middle', transform: n > 8 || spec.type === 'line' ? `rotate(-35 ${sx(i)} ${H - bottom + 18})` : '' }, short(spec.labels[i], 14)); t.append(el('title', {}, spec.labels[i])); svg.append(t); };
      for (let i = 0; i < n; i++) tickLabel(i);
      if (spec.type === 'bar') {
        const k = spec.series.length; const bw = Math.min(46, (step * 0.7) / k);
        spec.series.forEach((s, si) => s.values.forEach((v, i) => {
          const x0 = sx(i) - (bw * k) / 2 + si * bw; const top0 = Math.min(y(v), y(0)); const h = Math.max(2, Math.abs(y(v) - y(0)));
          const r = el('rect', { x: x0, y: top0, width: bw - 2, height: h, rx: 3, fill: COLORS[si], class: 'ch-bar' }); r.append(el('title', {}, `${spec.labels[i]} · ${s.name}: ${v}`)); svg.append(r);
          if (n <= 12 && k === 1) svg.append(el('text', { x: x0 + (bw - 2) / 2, y: top0 - 6, class: 'ch-val', 'text-anchor': 'middle' }, compact(v)));
        }));
      } else {
        spec.series.forEach((s, si) => {
          const pts = s.values.map((v, i) => `${sx(i)},${y(v)}`);
          if (si === 0 && spec.series.length === 1) svg.append(el('polygon', { points: `${sx(0)},${y(Math.max(0, min))} ${pts.join(' ')} ${sx(n - 1)},${y(Math.max(0, min))}`, fill: COLORS[0], opacity: '.12' }));
          svg.append(el('polyline', { points: pts.join(' '), fill: 'none', stroke: COLORS[si], 'stroke-width': 2.5, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
          if (n <= 40) s.values.forEach((v, i) => { const c = el('circle', { cx: sx(i), cy: y(v), r: 3.5, fill: COLORS[si], class: 'ch-dot' }); c.append(el('title', {}, `${spec.labels[i]} · ${s.name}: ${v}`)); svg.append(c); });
        });
      }
    }
    host.append(svg);
    if (spec.series.length > 1) {
      const lg = document.createElement('div'); lg.className = 'ch-legend';
      spec.series.forEach((s, i) => { const it = document.createElement('span'); const dot = document.createElement('i'); dot.style.background = COLORS[i]; it.append(dot, s.name); lg.append(it); });
      host.append(lg);
    }
    if (spec.note) { const p = document.createElement('p'); p.className = 'ch-note'; p.textContent = spec.note; host.append(p); }
  }

  const api = { pickChart, drawChart };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.AskaChart = api;
})(typeof window !== 'undefined' ? window : globalThis);
