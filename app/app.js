(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const WEBLLM = 'https://cdn.jsdelivr.net/npm/@mlc-ai/web-llm@0.2.79/+esm';
  const WEB_MODEL = 'Qwen2.5-Coder-1.5B-Instruct-q4f16_1-MLC';
  const XLSX_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  const MAX_BYTES = 50 * 1024 * 1024;
  const MAX_ROWS = 300000;
  let SQL = null; let db = null; let tables = []; let engine = null; let last = null;
  const files = []; // { name, tables: [..] }

  const msg = (t, err) => { $('msg').textContent = t || ''; $('msg').className = err ? 'err' : ''; };
  const q = (n) => `"${String(n).replace(/"/g, '""')}"`;

  // ---------- reading files ----------
  function parseCsv(text) {
    text = text.replace(/^﻿/, '');
    const first = text.split(/\r?\n/, 1)[0] || '';
    const delim = [',', ';', '\t'].map((d) => [d, first.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
    const rows = []; let row = []; let f = ''; let inq = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inq) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else inq = false; } else f += c; }
      else if (c === '"') inq = true;
      else if (c === delim) { row.push(f); f = ''; }
      else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(f); f = ''; if (row.length > 1 || row[0] !== '') rows.push(row); row = []; if (rows.length > MAX_ROWS) break; }
      else f += c;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  const ident = (s, fallback, used) => {
    let b = String(s || '').replace(/\.[^.]+$/, '').normalize('NFKD').replace(/[^\w]+/g, '_').replace(/^_+|_+$/g, '') || fallback;
    if (/^\d/.test(b)) b = `t_${b}`;
    let n = b; let i = 2; while (used.has(n.toLowerCase())) n = `${b}_${i++}`;
    used.add(n.toLowerCase()); return n;
  };
  const usedTables = () => new Set(tables.map((t) => t.name.toLowerCase()));
  function addTable(name, rows) {
    if (!rows.length) return null;
    const used = new Set(); const header = rows[0].map((h, i) => ident(h == null || h === '' ? `col${i + 1}` : h, `col${i + 1}`, used));
    const body = rows.slice(1, MAX_ROWS + 1);
    const types = header.map((_, c) => { const vals = body.map((r) => r[c]).filter((v) => v !== null && v !== undefined && v !== ''); if (!vals.length) return 'TEXT'; if (vals.every((v) => typeof v === 'number' || (typeof v === 'string' && /^-?\d+(\.\d+)?$/.test(v.trim())))) return vals.every((v) => Number.isInteger(+v)) ? 'INTEGER' : 'REAL'; return 'TEXT'; });
    db.exec(`CREATE TABLE ${q(name)} (${header.map((h, i) => `${q(h)} ${types[i]}`).join(', ')})`);
    const st = db.prepare(`INSERT INTO ${q(name)} VALUES (${header.map(() => '?').join(',')})`);
    db.exec('BEGIN');
    for (const r of body) st.run(header.map((_, c) => { let v = r[c]; if (v instanceof Date) v = v.toISOString().slice(0, 10); if (v === undefined || v === '') return null; return types[c] === 'TEXT' ? String(v) : +v; }));
    db.exec('COMMIT'); st.free();
    return name;
  }
  function loadScript(src) { return new Promise((ok, no) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => no(new Error('Could not load the Excel reader. Check your connection.')); document.head.append(s); }); }

  async function addFile(file) {
    if (file.size > MAX_BYTES) throw new Error(`${file.name} is larger than 50 MB.`);
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    const made = [];
    if (['sqlite', 'sqlite3', 'db'].includes(ext)) {
      const other = new SQL.Database(new Uint8Array(await file.arrayBuffer()));
      const names = other.exec("SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'");
      for (const [n] of (names[0] ? names[0].values : [])) {
        const tn = ident(n, 'table', usedTables());
        const cols = other.exec(`PRAGMA table_info(${q(n)})`)[0].values;
        db.exec(`CREATE TABLE ${q(tn)} (${cols.map((c) => `${q(c[1])} ${c[2] || ''}`).join(', ')})`);
        const data = other.exec(`SELECT * FROM ${q(n)} LIMIT ${MAX_ROWS}`);
        if (data[0]) { const st = db.prepare(`INSERT INTO ${q(tn)} VALUES (${cols.map(() => '?').join(',')})`); db.exec('BEGIN'); data[0].values.forEach((r) => st.run(r)); db.exec('COMMIT'); st.free(); }
        made.push(tn); tables.push({ name: tn });
      }
      other.close();
    } else if (['csv', 'tsv'].includes(ext)) {
      const tn = ident(file.name, 'data', usedTables());
      addTable(tn, parseCsv(await file.text())); made.push(tn); tables.push({ name: tn });
    } else if (['xlsx', 'xls', 'xlsm'].includes(ext)) {
      if (!window.XLSX) await loadScript(XLSX_URL);
      const wb = window.XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array', cellDates: true });
      const sheets = wb.SheetNames.filter((s) => wb.Sheets[s] && wb.Sheets[s]['!ref']);
      for (const s of sheets) {
        const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[s], { header: 1, raw: true, defval: null, blankrows: false });
        const tn = ident(sheets.length > 1 ? `${file.name.replace(/\.[^.]+$/, '')}_${s}` : file.name, 'sheet', usedTables());
        if (addTable(tn, rows)) { made.push(tn); tables.push({ name: tn }); }
      }
    } else throw new Error(`Can't read ${file.name}. Use Excel, CSV or a SQLite file.`);
    files.push({ name: file.name, tables: made });
  }

  function refreshSchema() {
    tables = db.exec("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")[0]?.values.map((v) => ({ name: v[0] })) || [];
    for (const t of tables) { t.cols = db.exec(`PRAGMA table_info(${q(t.name)})`)[0].values.map((c) => ({ name: c[1], type: c[2] || 'TEXT' })); t.rows = db.exec(`SELECT COUNT(*) FROM ${q(t.name)}`)[0].values[0][0]; }
    const box = $('files'); box.textContent = '';
    for (const t of tables) { const s = document.createElement('span'); s.className = 'tag'; const b = document.createElement('b'); b.textContent = t.name; s.append(b, ` · ${t.rows.toLocaleString()} rows · ${t.cols.length} columns`); s.title = t.cols.map((c) => c.name).join(', '); s.onclick = () => run(`Preview of ${t.name}`, `SELECT * FROM ${q(t.name)} LIMIT 20`, 'preview'); box.append(s); }
    $('askBox').hidden = !tables.length;
    $('fhint').hidden = !!tables.length;
    const chips = $('chips'); chips.textContent = '';
    const add = (label, sql) => { const c = document.createElement('button'); c.className = 'chip'; c.type = 'button'; c.textContent = label; c.onclick = () => run(label, sql, 'ready-made, no AI'); chips.append(c); };
    add('How many rows are in each table?', tables.map((t) => `SELECT ${`'${t.name.replace(/'/g, "''")}'`} AS "table", COUNT(*) AS "rows" FROM ${q(t.name)}`).join(' UNION ALL ') + ' ORDER BY "rows" DESC');
    if (tables[0]) add(`Show 20 rows of ${tables[0].name}`, `SELECT * FROM ${q(tables[0].name)} LIMIT 20`);
    if (tables[0]) add(`Describe ${tables[0].name}`, `SELECT name AS "column", type AS "type" FROM pragma_table_info(${`'${tables[0].name.replace(/'/g, "''")}'`})`);
  }
  const schemaText = () => tables.map((t) => `${t.name}(${t.cols.map((c) => `${c.name} ${c.type}`).join(', ')})`).join('\n');

  // ---------- safety and running ----------
  function guard(sql) {
    const s = String(sql || '').trim().replace(/;+\s*$/, '');
    const m = s.replace(/--[^\n]*/g, ' ').replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/'(?:[^']|'')*'/g, "''").replace(/"(?:[^"]|"")*"/g, '""');
    if (!/^\s*(select|with)\b/i.test(m)) return { ok: false, why: 'Only SELECT queries are allowed.' };
    if (m.includes(';')) return { ok: false, why: 'Only one statement is allowed.' };
    const bad = m.match(/\b(insert|update|delete|drop|alter|create|replace|attach|detach|pragma|vacuum|reindex|truncate|exec|merge)\b/i);
    if (bad && !/^\s*select name AS "column", type AS "type" FROM pragma_table_info/i.test(s)) return { ok: false, why: `"${bad[1]}" is not allowed. This app only reads data.` };
    return { ok: true, sql: s };
  }
  function run(title, sql, who) {
    const g = guard(sql); if (!g.ok) { msg(g.why, true); return false; }
    try {
      const t0 = performance.now(); const r = db.exec(g.sql)[0] || { columns: [], values: [] }; const ms = Math.round(performance.now() - t0);
      last = { columns: r.columns, rows: r.values, sql: g.sql };
      $('title').textContent = title; $('pSql').textContent = g.sql;
      $('meta').textContent = `${r.values.length.toLocaleString()} row${r.values.length === 1 ? '' : 's'}${r.values.length > 200 ? ' (first 200 shown)' : ''} · ${ms} ms · ${who}`;
      const t = $('tbl'); t.textContent = ''; const h = t.createTHead().insertRow();
      r.columns.forEach((c) => { const th = document.createElement('th'); th.textContent = c; h.append(th); });
      const b = t.createTBody(); r.values.slice(0, 200).forEach((row) => { const tr = b.insertRow(); row.forEach((v) => { const td = tr.insertCell(); td.textContent = v == null ? '' : v; if (typeof v === 'number') td.className = 'n'; }); });
      const spec = window.AskaChart.pickChart(r.columns, r.values);
      $('tChart').hidden = !spec; if (spec) window.AskaChart.drawChart($('pChart'), spec);
      tab(spec ? 'chart' : 'table'); $('res').style.display = 'block'; msg(''); return true;
    } catch (e) { msg(String(e.message || e), true); return false; }
  }
  function tab(n) { $('pChart').hidden = n !== 'chart'; $('pTable').hidden = n !== 'table'; $('pSql').hidden = n !== 'sql'; $('tChart').classList.toggle('on', n === 'chart'); $('tTable').classList.toggle('on', n === 'table'); $('tSql').classList.toggle('on', n === 'sql'); }
  $('tChart').onclick = () => tab('chart'); $('tTable').onclick = () => tab('table'); $('tSql').onclick = () => tab('sql');
  $('csv').onclick = () => {
    if (!last) return; const esc = (v) => (v == null ? '' : /[",\n\r]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v));
    const text = '﻿' + [last.columns, ...last.rows].map((r) => r.map(esc).join(',')).join('\r\n');
    const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type: 'text/csv' })); a.download = 'askadb-result.csv'; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  };

  // ---------- AI ----------
  const SYSTEM = () => 'You write ONE read-only SQLite query for the schema below. Reply with only the SQL in a single code block, no explanation. Use only these tables and columns, quote names with double quotes when needed. Use LIMIT for "top N" questions. Give result columns clear aliases. Match the user\'s words to table and column names even with synonyms or small typos.\nSchema:\n' + schemaText();
  const pull = (t) => (/```(?:sql)?\s*([\s\S]*?)```/i.exec(t) || [null, t])[1].trim();
  async function callAI(messages) {
    if ($('mode').value === 'web') {
      if (!engine) throw new Error('Start the browser AI first (button below), or choose an AI service.');
      const r = await engine.chat.completions.create({ messages, temperature: 0, max_tokens: 500 }); return r.choices[0].message.content;
    }
    const base = $('base').value.trim().replace(/\/+$/, ''); const model = $('model').value.trim();
    if (!base || !model) throw new Error('Choose a service and a model, and paste your key (open "AI for your own questions").');
    const headers = { 'Content-Type': 'application/json' }; if ($('key').value.trim()) headers.Authorization = `Bearer ${$('key').value.trim()}`;
    let res; try { res = await fetch(`${base}/chat/completions`, { method: 'POST', headers, body: JSON.stringify({ model, temperature: 0, messages }) }); }
    catch (e) { throw new Error('Could not reach the AI service. Check the URL, your key, and that the service allows browser calls (CORS).'); }
    if (!res.ok) { let t = ''; try { t = (await res.json()).error?.message || ''; } catch (_) { /* ignore */ } throw new Error(`The AI service said ${res.status}${t ? `: ${String(t).slice(0, 160)}` : ''}`); }
    return (await res.json()).choices[0].message.content;
  }
  async function ask() {
    const text = $('q').value.trim(); if (!text || !db) return;
    $('go').disabled = true; msg('Thinking…');
    try {
      const t0 = performance.now();
      let sql = pull(await callAI([{ role: 'system', content: SYSTEM() }, { role: 'user', content: text }]));
      for (let i = 0; i < 2; i++) {
        const g = guard(sql); if (!g.ok) throw new Error(g.why);
        try { db.exec(`EXPLAIN ${g.sql}`); sql = g.sql; break; }
        catch (e) { if (i) throw e; sql = pull(await callAI([{ role: 'system', content: SYSTEM() }, { role: 'user', content: `${text}\n\nYour last query failed. Fix it. Only the error text follows.\nSQL: ${sql}\nError: ${e.message}` }])); }
      }
      run(text, sql, `AI · ${((performance.now() - t0) / 1000).toFixed(1)} s`);
    } catch (e) { msg(String(e.message || e), true); } finally { $('go').disabled = false; }
  }
  $('go').onclick = ask;
  $('q').addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); ask(); } });

  function syncAi() {
    const web = $('mode').value === 'web';
    ['presetRow', 'baseRow', 'modelRow', 'keyRow'].forEach((id) => { $(id).hidden = web; }); $('webRow').hidden = !web;
  }
  $('mode').onchange = syncAi;
  const KEYS = { 'generativelanguage': ['https://aistudio.google.com/apikey', 'Get a free Gemini key', 'sign in with Google, click Create API key'], 'groq.com': ['https://console.groq.com/keys', 'Get a free Groq key', 'sign up, click Create API Key'], 'openrouter': ['https://openrouter.ai/keys', 'Get an OpenRouter key', 'sign in, click Create Key'], 'openai.com': ['https://platform.openai.com/api-keys', 'Get an OpenAI key (paid)', 'sign in, click Create new secret key'] };
  $('preset').onchange = () => {
    const [b, m] = $('preset').value.split('|'); $('base').value = b; $('model').value = m;
    const k = Object.entries(KEYS).find(([h]) => b.includes(h));
    $('getKey').hidden = !k; if (k) { $('keyLink').href = k[1][0]; $('getKey').innerHTML = `New to this? <a id="keyLink" href="${k[1][0]}" target="_blank" rel="noopener" style="color:var(--accent)">${k[1][1]}</a> (${k[1][2]}, copy it, paste it here). It takes about one minute and the key is yours.`; }
    if (b.includes('localhost')) { $('getKey').hidden = false; $('getKey').textContent = 'Ollama must be installed and started, and allow this site: set OLLAMA_ORIGINS to this site address.'; }
  };
  $('webBtn').onclick = async () => {
    if (!navigator.gpu) { $('webState').textContent = 'This browser has no WebGPU. Use Chrome or Edge on a computer, or choose an AI service.'; return; }
    $('webBtn').disabled = true; $('bar').style.display = 'block'; $('webState').textContent = 'Downloading… keep this tab open.';
    try {
      const w = await import(WEBLLM);
      engine = await w.CreateMLCEngine(WEB_MODEL, { initProgressCallback: (p) => { $('barI').style.width = `${Math.round((p.progress || 0) * 100)}%`; if (p.text) $('webState').textContent = p.text.slice(0, 80); } });
      $('bar').style.display = 'none'; $('webState').textContent = 'Browser AI is ready.'; $('webBtn').hidden = true;
    } catch (e) { $('bar').style.display = 'none'; $('webBtn').disabled = false; $('webState').textContent = `Could not start: ${String(e.message || e).slice(0, 100)}`; }
  };

  // ---------- files in ----------
  async function addMany(list) {
    if (!SQL) { msg('Still loading the engine, one moment…'); return; }
    msg('Reading…');
    try { for (const f of list) await addFile(f); refreshSchema(); msg(''); } catch (e) { refreshSchema(); msg(String(e.message || e), true); }
  }
  const drop = $('drop');
  drop.onclick = () => $('file').click(); drop.onkeydown = (e) => { if (e.key === 'Enter' || e.key === ' ') $('file').click(); };
  $('file').onchange = (e) => { addMany([...e.target.files]); e.target.value = ''; };
  ['dragenter', 'dragover'].forEach((n) => drop.addEventListener(n, (e) => { e.preventDefault(); drop.classList.add('on'); }));
  ['dragleave', 'drop'].forEach((n) => drop.addEventListener(n, (e) => { e.preventDefault(); drop.classList.remove('on'); }));
  drop.addEventListener('drop', (e) => addMany([...e.dataTransfer.files]));
  $('sample').onclick = async (e) => { e.preventDefault(); try { const b = await (await fetch('../demo/chinook.sqlite')).blob(); await addMany([new File([b], 'chinook.sqlite')]); } catch (_) { msg('Could not load the sample.', true); } };

  (async () => {
    try { SQL = await initSqlJs({ locateFile: (f) => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/${f}` }); db = new SQL.Database(); }
    catch (_) { msg('Could not load the engine. Check your connection and refresh.', true); }
    $('preset').onchange(); syncAi();
  })();
})();
