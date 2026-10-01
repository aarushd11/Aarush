(function () {
  'use strict';

  const P = window.TTParser;
  const $ = (id) => document.getElementById(id);
  const DEFAULT_FILES = ['data/timetable.xlsx', 'data/timetable.xls', 'data/timetable.csv'];
  const PREF_KEY = 'tt.lastSelection';

  const state = {
    entries: [],
    index: {},
    source: null,        // {name, kind: 'upload'|'default'}
    hasDefault: false,
    view: 'days',
    current: null,       // {year, category, gKey, gName}
  };

  // ---------- storage (uploaded file kept in IndexedDB so it survives reloads) ----------

  function idb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('timetable', 1);
      req.onupgradeneeded = () => req.result.createObjectStore('files');
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function saveUpload(name, buf) {
    try {
      const db = await idb();
      db.transaction('files', 'readwrite').objectStore('files').put({ name, buf, at: Date.now() }, 'upload');
    } catch (e) { /* storage unavailable: still works for this visit */ }
  }
  async function loadUpload() {
    try {
      const db = await idb();
      return await new Promise((resolve) => {
        const req = db.transaction('files').objectStore('files').get('upload');
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      });
    } catch (e) { return null; }
  }
  async function clearUpload() {
    try {
      const db = await idb();
      db.transaction('files', 'readwrite').objectStore('files').delete('upload');
    } catch (e) { /* ignore */ }
  }
  const prefs = {
    get() { try { return JSON.parse(localStorage.getItem(PREF_KEY)) || null; } catch (e) { return null; } },
    set(v) { try { localStorage.setItem(PREF_KEY, JSON.stringify(v)); } catch (e) { /* ignore */ } },
  };

  // ---------- loading ----------

  async function fetchDefault() {
    for (const path of DEFAULT_FILES) {
      try {
        const res = await fetch(path, { cache: 'no-cache' });
        if (res.ok) return { name: path.split('/').pop(), buf: await res.arrayBuffer() };
      } catch (e) { /* not available (e.g. opened from file://) */ }
    }
    return null;
  }

  function loadWorkbook(buf, name, kind) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const { entries, report } = P.parseWorkbook(wb, XLSX);
    renderDiag(report, entries.length);
    if (!entries.length) {
      throw new Error('Could not find any classes in this file. Open "How the sheet was read" below for details.');
    }
    state.entries = entries;
    state.index = P.buildIndex(entries);
    state.source = { name, kind };
    const groups = new Set();
    for (const y in state.index) for (const c in state.index[y]) for (const k of state.index[y][c].keys()) groups.add(y + '|' + c + '|' + k);
    $('dataStatusText').textContent = `${name} · ${groups.size} groups · ${entries.length.toLocaleString()} classes`;
    $('dataStatus').hidden = false;
    setupFinder();
  }

  async function init() {
    wireEvents();
    const def = await fetchDefault();
    state.hasDefault = !!def;
    const up = await loadUpload();
    const src = up || def;
    if (!src) return showUpload();
    try {
      loadWorkbook(src.buf, src.name, up ? 'upload' : 'default');
      showFinder();
      restoreSelection();
    } catch (e) {
      showUpload(e.message);
    }
  }

  async function handleFile(file) {
    if (!file) return;
    const msg = $('uploadMsg');
    msg.hidden = false; msg.className = 'msg'; msg.textContent = `Reading ${file.name}…`;
    try {
      const buf = await file.arrayBuffer();
      loadWorkbook(buf, file.name, 'upload');
      await saveUpload(file.name, buf);
      msg.hidden = true;
      showFinder();
      restoreSelection();
    } catch (e) {
      msg.className = 'msg error';
      msg.textContent = e.message || 'Could not read that file.';
    }
  }

  // ---------- views ----------

  function showUpload(err) {
    $('uploadView').hidden = false;
    $('finderView').hidden = true;
    $('resultView').hidden = true;
    $('useDefaultBtn').hidden = !(state.hasDefault && state.source && state.source.kind === 'upload');
    const msg = $('uploadMsg');
    if (err) { msg.hidden = false; msg.className = 'msg error'; msg.textContent = err; } else msg.hidden = true;
  }
  function showFinder() {
    $('uploadView').hidden = true;
    $('finderView').hidden = false;
  }

  const yearLabel = (y) => (/^\d$/.test(y) ? `${y}${['th', 'st', 'nd', 'rd'][y] || 'th'} Year` : y || 'All');

  function setupFinder() {
    const years = Object.keys(state.index).sort(P.naturalCompare);
    const ys = $('yearSelect');
    ys.innerHTML = '';
    if (years.length > 1) ys.append(new Option('Select year', ''));
    for (const y of years) ys.append(new Option(yearLabel(y), y));
    $('yearField').hidden = years.length === 1 && years[0] === '';
    onYearChange();
  }

  function onYearChange() {
    const y = $('yearSelect').value;
    const cats = y in state.index ? Object.keys(state.index[y]).sort(P.naturalCompare) : [];
    const cs = $('categorySelect');
    cs.innerHTML = '';
    const hasCats = cats.length > 1 || (cats.length === 1 && cats[0] !== '');
    $('categoryField').hidden = !hasCats;
    if (hasCats) {
      cs.append(new Option('Any', '*'));
      for (const c of cats) cs.append(new Option(c || '(none)', c));
    }
    fillGroups();
  }

  function groupsFor(y, cat) {
    const out = [];
    const cats = state.index[y] || {};
    for (const c in cats) {
      if (cat !== '*' && cat != null && c !== cat) continue;
      for (const [k, name] of cats[c]) out.push({ key: k, name, category: c });
    }
    out.sort((a, b) => P.naturalCompare(a.name, b.name) || P.naturalCompare(a.category, b.category));
    return out;
  }

  function fillGroups() {
    const y = $('yearSelect').value;
    const cat = $('categoryField').hidden ? '*' : $('categorySelect').value;
    const dl = $('groupList');
    dl.innerHTML = '';
    const list = groupsFor(y, cat);
    for (const g of list) {
      const o = document.createElement('option');
      o.value = g.name;
      if (g.category) o.label = g.category;
      dl.append(o);
    }
    $('groupInput').placeholder = list.length ? `e.g. ${list[Math.min(1, list.length - 1)].name}` : 'No groups for this year';
  }

  /** Resolves what the user typed into one group, tolerating "12" for "G12" etc. */
  function resolveGroup(y, cat, typed) {
    const list = groupsFor(y, cat);
    const key = P.groupKey(typed);
    if (!key) return { error: 'Please enter your group number.' };
    let hits = list.filter((g) => g.key === key);
    if (!hits.length) {
      const digits = key.replace(/\D/g, '');
      if (digits && digits === key) hits = list.filter((g) => g.key.replace(/^[A-Z]+/, '') === digits);
    }
    if (!hits.length) {
      const near = list.filter((g) => g.key.includes(key)).slice(0, 8).map((g) => g.name);
      return { error: `No group "${typed}" in ${yearLabel(y)}.` + (near.length ? ` Did you mean: ${near.join(', ')}?` : '') };
    }
    const cats = [...new Set(hits.map((h) => h.category))];
    if (cats.length > 1) {
      return { error: `Group "${typed}" exists in several categories (${cats.join(', ')}). Pick one from the Category list.` };
    }
    if (new Set(hits.map((h) => h.key)).size > 1) {
      return { error: `"${typed}" matches several groups: ${hits.map((h) => h.name).join(', ')}. Please type the full name.` };
    }
    return { group: hits[0] };
  }

  function submit(e) {
    if (e) e.preventDefault();
    const y = $('yearSelect').value;
    const msg = $('finderMsg');
    msg.hidden = true;
    if (!(y in state.index)) { msg.hidden = false; msg.className = 'msg error'; msg.textContent = 'Please select your year.'; return; }
    const cat = $('categoryField').hidden ? '*' : $('categorySelect').value;
    const r = resolveGroup(y, cat, $('groupInput').value);
    if (r.error) { msg.hidden = false; msg.className = 'msg error'; msg.textContent = r.error; $('resultView').hidden = true; return; }
    const g = r.group;
    $('groupInput').value = g.name;
    state.current = { year: y, category: g.category, gKey: g.key, gName: g.name };
    prefs.set({ y, c: cat, g: g.name });
    history.replaceState(null, '', '#' + new URLSearchParams({ year: y, ...(g.category ? { cat: g.category } : {}), group: g.name }).toString());
    renderResult();
  }

  function restoreSelection() {
    const h = new URLSearchParams(location.hash.slice(1));
    const p = h.get('group') ? { y: h.get('year') || '', c: h.get('cat') || '*', g: h.get('group') } : prefs.get();
    if (!p || !(p.y in state.index)) return;
    $('yearSelect').value = p.y;
    onYearChange();
    if (!$('categoryField').hidden && [...$('categorySelect').options].some((o) => o.value === p.c)) $('categorySelect').value = p.c;
    fillGroups();
    $('groupInput').value = p.g;
    submit();
  }

  // ---------- rendering ----------

  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function hue(text) {
    let h = 0;
    for (const ch of String(text).toLowerCase()) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
    return h % 360;
  }

  function timeText(s) {
    if (s.start == null) return esc(s.label || 'Time TBA');
    return `${P.fmtTime(s.start)}${s.end != null ? `<span class="to">–</span>${P.fmtTime(s.end)}` : ''}`;
  }

  function duration(a, b) {
    if (a == null || b == null || b <= a) return '';
    const m = b - a, h = Math.floor(m / 60), r = m % 60;
    return (h ? `${h}h` : '') + (h && r ? ' ' : '') + (r ? `${r}m` : '');
  }

  function typeClass(t) {
    const l = (t || '').toLowerCase();
    if (l.startsWith('lec')) return 'lec';
    if (l.startsWith('tut')) return 'tut';
    if (l.startsWith('lab') || l.startsWith('prac')) return 'lab';
    return 'other';
  }

  function togetherHtml(c) {
    if (c.togetherAll) return '<span class="chip together">With all groups</span>';
    if (!c.together.length) return '';
    const MAX = 6;
    const shown = c.together.slice(0, MAX).map(esc).join(', ');
    const more = c.together.length - MAX;
    if (more <= 0) return `<span class="chip together" title="Shared class">Together with ${shown}</span>`;
    return `<details class="chip together more"><summary>Together with ${shown} <b>+${more} more</b></summary>
      <div class="all-groups">${c.together.map(esc).join(', ')}</div></details>`;
  }

  function classHtml(c) {
    const meta = [];
    if (c.faculty) meta.push(`<span class="meta"><svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 12a5 5 0 1 0 0-10 5 5 0 0 0 0 10Zm0 2c-4.4 0-8 2.2-8 5v3h16v-3c0-2.8-3.6-5-8-5Z"/></svg>${esc(c.faculty)}</span>`);
    if (c.room) meta.push(`<span class="meta"><svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M12 2a7 7 0 0 0-7 7c0 5.3 7 13 7 13s7-7.7 7-13a7 7 0 0 0-7-7Zm0 9.5a2.5 2.5 0 1 1 0-5 2.5 2.5 0 0 1 0 5Z"/></svg>${esc(c.room)}</span>`);
    return `<article class="class" style="--h:${hue(c.subject)}">
      <div class="class-top">
        <h4>${esc(c.subject)}</h4>
        ${c.type ? `<span class="pill ${typeClass(c.type)}">${esc(c.type)}</span>` : ''}
      </div>
      ${c.code ? `<div class="code">${esc(c.code)}</div>` : ''}
      ${meta.length ? `<div class="metas">${meta.join('')}</div>` : ''}
      ${togetherHtml(c)}
    </article>`;
  }

  function renderResult() {
    const cur = state.current;
    const days = P.timetableFor(state.entries, cur.year, cur.category, cur.gKey);
    const todayIdx = (new Date().getDay() + 6) % 7;

    $('resultEyebrow').textContent = [cur.year !== '' ? yearLabel(cur.year) : '', cur.category].filter(Boolean).join(' · ');
    $('resultTitle').textContent = `Group ${cur.gName}`;
    const total = days.reduce((n, d) => n + d.slots.reduce((m, s) => m + s.classes.length, 0), 0);
    const mins = days.reduce((n, d) => n + d.slots.reduce((m, s) => m + (s.end != null && s.start != null ? s.end - s.start : 0), 0), 0);
    const shared = days.reduce((n, d) => n + d.slots.reduce((m, s) => m + s.classes.filter((c) => c.together.length || c.togetherAll).length, 0), 0);
    $('resultStats').textContent = `${total} classes a week across ${days.length} days` +
      (mins ? ` · ${Math.round(mins / 6) / 10} hours` : '') + (shared ? ` · ${shared} shared with other groups` : '');

    // Day jump tabs
    $('dayTabs').innerHTML = days.map((d) =>
      `<a href="#day-${d.day}" class="${d.day === todayIdx ? 'today' : ''}" data-day="${d.day}">${d.name.slice(0, 3)}</a>`).join('');

    // Day-wise cards
    $('daysView').innerHTML = days.length ? days.map((d) => {
      const first = d.slots.find((s) => s.start != null), last = [...d.slots].reverse().find((s) => s.start != null);
      const span = first && last ? `${P.fmtTime(first.start)} – ${P.fmtTime(last.end ?? last.start)}` : '';
      let prevEnd = null;
      const rows = d.slots.map((s) => {
        let gap = '';
        if (prevEnd != null && s.start != null && s.start - prevEnd >= 10) {
          gap = `<div class="gap"><span>Free · ${duration(prevEnd, s.start)}</span></div>`;
        }
        if (s.end != null) prevEnd = Math.max(prevEnd ?? 0, s.end);
        return `${gap}<div class="slot${s.classes.length > 1 ? ' multi' : ''}">
          <div class="slot-time"><div class="t">${timeText(s)}</div>${duration(s.start, s.end) ? `<div class="dur">${duration(s.start, s.end)}</div>` : ''}</div>
          <div class="slot-body">
            ${s.classes.length > 1 ? `<div class="parallel">${s.classes.length} classes at this time</div>` : ''}
            ${s.classes.map(classHtml).join('')}
          </div>
        </div>`;
      }).join('');
      const count = d.slots.reduce((m, s) => m + s.classes.length, 0);
      return `<section class="day card${d.day === todayIdx ? ' is-today' : ''}" id="day-${d.day}">
        <header class="day-head">
          <h3>${d.name}${d.day === todayIdx ? ' <span class="today-badge">Today</span>' : ''}</h3>
          <span class="muted">${count} class${count === 1 ? '' : 'es'}${span ? ' · ' + span : ''}</span>
        </header>
        <div class="slots">${rows}</div>
      </section>`;
    }).join('') : '<div class="card empty">No classes found for this group.</div>';

    renderWeek(days, todayIdx);
    setView(state.view);
    $('resultView').hidden = false;
    $('resultView').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderWeek(days, todayIdx) {
    if (!days.length) { $('weekView').innerHTML = ''; return; }
    const slotKeys = new Map();
    for (const d of days) for (const s of d.slots) {
      const k = s.start == null ? 'L' + s.label : `${String(s.start).padStart(4, '0')}-${s.end ?? ''}`;
      if (!slotKeys.has(k)) slotKeys.set(k, s);
    }
    const keys = [...slotKeys.keys()].sort();
    const head = `<tr><th class="corner">Time</th>${days.map((d) => `<th class="${d.day === todayIdx ? 'today' : ''}">${d.name.slice(0, 3)}</th>`).join('')}</tr>`;
    const body = keys.map((k) => {
      const s = slotKeys.get(k);
      const cells = days.map((d) => {
        const hit = d.slots.find((x) => (x.start == null ? 'L' + x.label : `${String(x.start).padStart(4, '0')}-${x.end ?? ''}`) === k);
        if (!hit) return `<td class="${d.day === todayIdx ? 'today' : ''}"></td>`;
        return `<td class="${d.day === todayIdx ? 'today' : ''}">${hit.classes.map((c) => `<div class="mini" style="--h:${hue(c.subject)}">
          <b>${esc(c.subject)}</b>${c.type ? ` <span class="pill ${typeClass(c.type)}">${esc(c.type)}</span>` : ''}
          ${c.room ? `<span class="mini-meta">${esc(c.room)}</span>` : ''}
          ${c.together.length || c.togetherAll ? `<span class="mini-meta shared" title="${esc(c.togetherAll ? 'All groups' : c.together.join(', '))}">+${c.togetherAll ? 'all' : c.together.length} group${c.together.length === 1 ? '' : 's'}</span>` : ''}
        </div>`).join('')}</td>`;
      }).join('');
      return `<tr><th class="time-col">${timeText(s)}</th>${cells}</tr>`;
    }).join('');
    $('weekView').innerHTML = `<div class="table-scroll"><table class="week-table">${head}${body}</table></div>`;
  }

  function setView(v) {
    state.view = v;
    document.querySelectorAll('.seg button').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.view === v)));
    $('daysView').hidden = v !== 'days';
    $('dayTabs').hidden = v !== 'days';
    $('weekView').hidden = v !== 'week';
  }

  function renderDiag(report, total) {
    $('diag').hidden = false;
    const label = { year: 'Year', category: 'Category', group: 'Group', day: 'Day', time: 'Time', start: 'Start', end: 'End', subject: 'Subject', code: 'Code', type: 'Type', faculty: 'Faculty', room: 'Room' };
    $('diagBody').innerHTML = `<p>${total.toLocaleString()} class entries found.</p><ul>` + report.map((r) => {
      if (r.layout === 'skipped') return `<li><b>${esc(r.sheet)}</b> – skipped: ${esc(r.note)}</li>`;
      if (r.layout === 'grid') return `<li><b>${esc(r.sheet)}</b> – grid layout, ${r.groupColumns} group columns, ${r.rows} entries</li>`;
      const cols = Object.entries(r.columns).map(([f, h]) => `${label[f]} ← “${esc(h)}”`).join(', ');
      return `<li><b>${esc(r.sheet)}</b> – list layout, ${r.rows} entries. Columns: ${cols}</li>`;
    }).join('') + '</ul>';
  }

  // ---------- events ----------

  function wireEvents() {
    $('fileInput').addEventListener('change', (e) => handleFile(e.target.files[0]));
    const dz = $('dropzone');
    ['dragenter', 'dragover'].forEach((t) => dz.addEventListener(t, (e) => { e.preventDefault(); dz.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((t) => dz.addEventListener(t, (e) => { e.preventDefault(); dz.classList.remove('over'); }));
    dz.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]));
    $('changeFileBtn').addEventListener('click', () => { $('fileInput').value = ''; showUpload(); });
    $('useDefaultBtn').addEventListener('click', async () => {
      await clearUpload();
      const def = await fetchDefault();
      if (def) { loadWorkbook(def.buf, def.name, 'default'); showFinder(); restoreSelection(); }
    });
    $('yearSelect').addEventListener('change', () => { onYearChange(); $('groupInput').value = ''; });
    $('categorySelect').addEventListener('change', fillGroups);
    $('finderForm').addEventListener('submit', submit);
    document.querySelectorAll('.seg button').forEach((b) => b.addEventListener('click', () => setView(b.dataset.view)));
    $('printBtn').addEventListener('click', () => window.print());
    $('copyLinkBtn').addEventListener('click', async () => {
      try { await navigator.clipboard.writeText(location.href); $('copyLinkBtn').textContent = 'Link copied'; }
      catch (e) { $('copyLinkBtn').textContent = 'Copy failed'; }
      setTimeout(() => { $('copyLinkBtn').textContent = 'Copy link'; }, 1800);
    });
    $('dayTabs').addEventListener('click', (e) => {
      const a = e.target.closest('a');
      if (!a) return;
      e.preventDefault();
      document.getElementById(a.getAttribute('href').slice(1)).scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }

  init();
})();
