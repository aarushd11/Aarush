/*
 * Timetable parser: turns an Excel workbook (read by SheetJS) into a flat list
 * of class entries, whatever reasonable layout the sheet uses.
 *
 * Supported layouts (auto-detected per sheet):
 *   1. "List" layout - one row per class, with columns such as
 *        Year | Category | Group | Day | Start | End | Subject | Type | Faculty | Room
 *      Column names are matched loosely (e.g. "Batch", "Grp", "Timing", "Venue").
 *      A single row may list several groups ("G1, G2", "G1-G4", "All") - those
 *      groups attend that class together.
 *   2. "Grid" layout - one row per (day, time slot) and one column per group:
 *        Day | Time | G1 | G2 | G3 ...
 *      Merged cells spanning several group columns mean a shared class.
 *
 * Works both in the browser (window.TTParser) and in Node (module.exports).
 */
(function (root) {
  'use strict';

  const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

  // Header synonyms, compared after lower-casing and stripping non-letters/digits.
  const SYN = {
    year: ['year', 'yr', 'yearofstudy', 'studyyear', 'academicyear', 'classyear'],
    category: ['category', 'branch', 'department', 'dept', 'program', 'programme', 'stream', 'discipline', 'categoryname', 'branchname'],
    group: ['group', 'groups', 'grp', 'groupno', 'groupnumber', 'groupname', 'batch', 'batches', 'batchno', 'batchname', 'section', 'sec', 'subgroup', 'groupid', 'batchid', 'batchcode', 'groupcode'],
    day: ['day', 'days', 'weekday', 'dayname', 'dayofweek'],
    time: ['time', 'timing', 'timings', 'slot', 'timeslot', 'period', 'hours', 'hour', 'classtime', 'lecturetime'],
    start: ['start', 'starttime', 'from', 'begin', 'begins', 'timefrom', 'fromtime', 'starts', 'starttiming'],
    end: ['end', 'endtime', 'to', 'till', 'until', 'timeto', 'totime', 'ends', 'finish', 'endtiming'],
    subject: ['subject', 'subjects', 'course', 'coursename', 'subjectname', 'module', 'modulename', 'paper', 'coursetitle', 'title', 'class', 'lecture', 'activity', 'subjecttitle'],
    code: ['code', 'subjectcode', 'coursecode', 'modulecode', 'papercode', 'courseno', 'subcode'],
    type: ['type', 'classtype', 'lecturetype', 'mode', 'ltp', 'kind', 'component', 'sessiontype', 'nature', 'lectype'],
    faculty: ['faculty', 'teacher', 'teachers', 'instructor', 'instructors', 'professor', 'prof', 'lecturer', 'staff', 'facultyname', 'tutor', 'teachername', 'facultyincharge'],
    room: ['room', 'rooms', 'venue', 'location', 'hall', 'classroom', 'roomno', 'roomnumber', 'place', 'lab', 'lt', 'block', 'labroom'],
  };

  const norm = (v) => String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]/g, '');
  const clean = (v) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim();

  function matchField(header) {
    const h = norm(header);
    if (!h) return null;
    for (const f in SYN) if (SYN[f].includes(h)) return f;
    // Looser "contains" matches for longer headers like "Name of Faculty".
    const contains = [
      ['start', /start|from/], ['end', /\bend|till|upto/],
      ['code', /code/], ['faculty', /faculty|teacher|instructor|professor/],
      ['room', /room|venue|hall/], ['group', /group|batch|section/],
      ['year', /year/], ['day', /^day|weekday/], ['time', /time|slot|period/],
      ['type', /type/], ['category', /branch|category|department|programme|program/],
      ['subject', /subject|course|module/],
    ];
    const lower = String(header).toLowerCase();
    for (const [f, re] of contains) if (re.test(lower)) return f;
    return null;
  }

  // ---------- value parsers ----------

  const DAY_ALIASES = {
    mon: 0, monday: 0, mo: 0,
    tue: 1, tues: 1, tuesday: 1, tu: 1,
    wed: 2, weds: 2, wednesday: 2, we: 2,
    thu: 3, thur: 3, thurs: 3, thursday: 3, th: 3,
    fri: 4, friday: 4, fr: 4,
    sat: 5, saturday: 5, sa: 5,
    sun: 6, sunday: 6, su: 6,
  };

  /** Returns an array of day indexes (a cell may hold "Mon, Wed"). */
  function parseDays(v) {
    if (v == null || v === '') return [];
    if (typeof v === 'number' && v >= 1 && v <= 7) return [v - 1];
    const s = String(v).toLowerCase();
    const out = [];
    // Ranges like "Mon-Fri"
    const range = s.match(/^\s*([a-z]+)\.?\s*(?:-|–|to)\s*([a-z]+)\.?\s*$/);
    if (range && range[1] in DAY_ALIASES && range[2] in DAY_ALIASES) {
      for (let d = DAY_ALIASES[range[1]]; d <= DAY_ALIASES[range[2]]; d++) out.push(d);
      return out;
    }
    for (const tok of s.split(/[^a-z]+/)) {
      if (tok in DAY_ALIASES && !out.includes(DAY_ALIASES[tok])) out.push(DAY_ALIASES[tok]);
    }
    return out;
  }

  /** Parses a single clock time. Returns {m: minutes, ap: 'am'|'pm'|null} or null. */
  function parseClock(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return { m: v.getHours() * 60 + v.getMinutes(), ap: 'x' };
    if (typeof v === 'number') {
      if (v >= 0 && v < 1) return { m: Math.round(v * 24 * 60) % 1440, ap: 'x' }; // Excel time fraction
      if (v >= 1 && v < 24) { // 9 or 9.30
        const h = Math.floor(v), mm = Math.round((v - h) * 100);
        return { m: h * 60 + (mm < 60 ? mm : 0), ap: null };
      }
      if (v >= 100 && v <= 2359) return { m: Math.floor(v / 100) * 60 + (v % 100), ap: 'x' };
      return null;
    }
    // "8:50:AM", "9:40: AM" -> "8:50am"
    const s = String(v).trim().toLowerCase().replace(/\s+/g, '').replace(/[:.]+(?=[ap]\.?m?\.?$)/, '');
    const m = s.match(/^(\d{1,2})(?:[:.](\d{2}))?(?::\d{2})?(am|pm|a\.m\.|p\.m\.|a|p)?$/) ||
              s.match(/^(\d{2})(\d{2})(am|pm)?$/);
    if (!m) return null;
    let h = +m[1], mm = m[2] ? +m[2] : 0;
    if (h > 23 || mm > 59) return null;
    let ap = m[3] ? (m[3][0] === 'p' ? 'pm' : 'am') : null;
    if (ap === 'pm' && h < 12) h += 12;
    if (ap === 'am' && h === 12) h = 0;
    if (!ap && h >= 13) ap = 'x'; // already 24h
    return { m: h * 60 + mm, ap };
  }

  // Without AM/PM, college hours 1-7 are almost always afternoon.
  const fix12 = (c) => (c && c.ap === null && c.m < 8 * 60 ? { m: c.m + 720, ap: 'pm' } : c);

  /** Parses "9:00-9:50", "9 to 10 AM", "09:00 AM - 10:00 AM" ... into {start, end}. */
  function parseRange(v) {
    if (v == null || v === '') return null;
    if (typeof v !== 'string') {
      const c = fix12(parseClock(v));
      return c ? { start: c.m, end: null } : null;
    }
    const s = v.trim();
    const parts = s.split(/\s*(?:-|–|—|to|till|\/)\s*/i).filter(Boolean);
    if (parts.length >= 2) {
      let a = parseClock(parts[0]), b = parseClock(parts[1]);
      if (a && b) {
        // "9 - 10 AM": share the meridiem.
        if (!a.ap && (b.ap === 'am' || b.ap === 'pm')) {
          const pm = b.ap === 'pm';
          const cand = (a.m % 720) + (pm ? 720 : 0);
          a = { m: cand > b.m ? cand - 720 : cand, ap: b.ap };
        }
        a = fix12(a);
        if (!b.ap || b.ap === null) b = b.m < a.m ? { m: b.m + 720, ap: 'pm' } : b;
        if (b.m <= a.m && b.m + 720 > a.m && b.m + 720 < 1440) b = { m: b.m + 720 };
        return { start: a.m, end: b.m };
      }
    }
    const c = fix12(parseClock(s));
    return c ? { start: c.m, end: null } : null;
  }

  const ROMAN = { i: 1, ii: 2, iii: 3, iv: 4, v: 5, vi: 6 };
  const WORDS = { first: 1, second: 2, third: 3, fourth: 4, fifth: 5, sixth: 6, final: null };

  /** Canonical year label ("1", "2", ...) or the cleaned raw text. */
  function parseYear(v) {
    const s = clean(v);
    if (!s) return '';
    const l = s.toLowerCase();
    let m = l.match(/(?:^|[^a-z0-9])(\d)(?:st|nd|rd|th)?(?:[^0-9]|$)/) || l.match(/^(\d)$/);
    if (m && /year|yr|^y?\d|\d(st|nd|rd|th)|^\d$/.test(l)) return m[1];
    for (const w in WORDS) if (WORDS[w] && new RegExp('\\b' + w + '\\b').test(l)) return String(WORDS[w]);
    m = l.match(/\b(i{1,3}|iv|v|vi)\b/);
    if (m && /year|yr|^(i{1,3}|iv|v|vi)$/.test(l)) return String(ROMAN[m[1]]);
    m = l.match(/^y(\d)$/);
    if (m) return m[1];
    return s;
  }

  /** Key used to compare groups: "Group 12" == "group12" == "G 12"? (no: G12 stays distinct). */
  function groupKey(name) {
    return String(name).toUpperCase().replace(/^(GROUP|GRP|BATCH|SECTION|SEC)\s*[-:#.]?\s*/, '').replace(/[^A-Z0-9]/g, '');
  }

  const ALL_RE = /^(all|all\s*groups?|all\s*batches|common|everyone|entire\s*class|whole\s*class)$/i;

  /** Splits a group cell into individual group names, expanding ranges like G1-G4. */
  function parseGroups(v) {
    const s = clean(v);
    if (!s) return { all: false, groups: [] };
    if (ALL_RE.test(s)) return { all: true, groups: [] };
    const out = [];
    const tokens = s.split(/\s*(?:,|;|\/|&|\+|\band\b|\n)\s*/i).filter(Boolean);
    for (const t of tokens) {
      if (ALL_RE.test(t)) return { all: true, groups: [] };
      // Range: "G1-G4", "1-4", "COE1 to COE5", "Group 1 - 4"
      const m = t.match(/^(.*?)(\d+)\s*(?:-|–|to)\s*(.*?)(\d+)$/i);
      if (m) {
        const p1 = m[1], p2 = m[3] || m[1], a = +m[2], b = +m[4];
        if (groupKey(p1 || 'x') === groupKey(p2 || 'x') && b > a && b - a <= 300 && !/\d$/.test(p1)) {
          for (let i = a; i <= b; i++) out.push(clean(p1) + String(i).padStart(m[2].length, '0'));
          continue;
        }
      }
      out.push(t);
    }
    return { all: false, groups: out };
  }

  const TYPE_MAP = [
    [/^(l|lec|lect|lecture|theory|th)$/i, 'Lecture'],
    [/^(t|tut|tutorial)$/i, 'Tutorial'],
    [/^(p|pr|prac|practical|lab|laboratory)$/i, 'Lab'],
    [/^(s|sem|seminar)$/i, 'Seminar'],
  ];
  function parseType(v) {
    const s = clean(v);
    for (const [re, t] of TYPE_MAP) if (re.test(s)) return t;
    return s;
  }

  /** Pulls a type marker like "(L)", "[Lab]" or "- P" out of free text. */
  function extractType(text) {
    const m = text.match(/[\(\[]\s*(L|T|P|Lec|Lecture|Tut|Tutorial|Lab|Practical|Prac|Seminar)\s*[\)\]]/i) ||
              text.match(/\b(Lecture|Tutorial|Practical|Lab)\b/i);
    return m ? parseType(m[1]) : '';
  }

  // ---------- sheet helpers ----------

  function sheetRows(XLSX, ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
    // Copy merged-cell values into every cell of the merge (shared classes/days).
    for (const mg of ws['!merges'] || []) {
      const v = rows[mg.s.r] && rows[mg.s.r][mg.s.c];
      for (let r = mg.s.r; r <= mg.e.r; r++) {
        rows[r] = rows[r] || [];
        for (let c = mg.s.c; c <= mg.e.c; c++) {
          if (r === mg.s.r && c === mg.s.c) continue;
          rows[r][c] = v;
          (rows[r].__merged = rows[r].__merged || {})[c] = `${mg.s.r}:${mg.s.c}`;
        }
      }
      if (rows[mg.s.r]) (rows[mg.s.r].__merged = rows[mg.s.r].__merged || {})[mg.s.c] = `${mg.s.r}:${mg.s.c}`;
    }
    return rows;
  }

  function detectHeader(rows) {
    let best = null;
    for (let r = 0; r < Math.min(rows.length, 30); r++) {
      const row = rows[r] || [];
      const map = {};
      let hits = 0;
      row.forEach((cell, c) => {
        const f = matchField(cell);
        if (f && map[f] == null) { map[f] = c; hits++; }
      });
      const hasTime = map.time != null || map.start != null;
      if (map.day != null && hasTime && hits >= 2 && (!best || hits > best.hits)) best = { r, map, hits };
    }
    return best;
  }

  function looksLikeGroupHeader(h) {
    const s = clean(h);
    return s && !matchField(s) && s.length <= 30;
  }


  // ---------- "block grid" layout (university master timetable) ----------
  //
  //            | LECTURE   |  1A1 (merged over its groups)          | 1A2 ...
  //            | TUTORIAL  |  1A11  | 1A12  | ...                   |
  //  DAY SR NO | PRACTICAL | HOURS | 1A1A  | 1A1B  | ...           |
  //   M    1   |           | 8:00  | UCB009L (merged over all groups taught together)
  //   O        |           |       | LP101 (room)            ... | SBA (teacher, last column)
  //
  // Every group owns a span of columns (usually 2). A subject cell is merged
  // across exactly the groups that attend it; the row under it holds the room
  // and, in the span's last column, the teacher. A "LAB" cell continues the
  // practical started in the slot above (its teacher is written below it).

  // Subject codes like "UCB009L", "UMA022 L", "PCYXXXP", "PMC L", "PBT110P G1".
  const CODE_RE = /^([A-Za-z]{2,}[A-Za-z0-9&-]*?(?:\d|XXX)[A-Za-z0-9]*?|[A-Z]{3,}(?=\s))\s*([LTP])(?![A-Za-z0-9])\s*(.*)$/;
  // Looser whole-cell form: "BEST3.3T", "BEST-SENP", "BEST-EE/HP L" (needs a separator or digit).
  const LINE_RE = /^([A-Z]{3,}[A-Z0-9 .\/-]*?[-. \d\/][A-Z0-9 .\/-]*?)[\s-]*([LTP])$/;
  const PAREN_RE = /^(.*?)\s*\(\s*([LTP])\s*\)\s*(.*)$/; // "BES -EE(P)" 
  const LABELS = { lecture: /^lectures?$/i, tutorial: /^tutorials?$/i, practical: /^practicals?$/i, branch: /^(branch|department|programme|program)$/i };

  function mergeIndex(ws) {
    const at = new Map();   // "r:c" of every merged cell -> merge
    for (const m of ws['!merges'] || []) {
      for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) at.set(r + ':' + c, m);
    }
    return at;
  }

  function detectBlockGrid(rows) {
    for (let r = 0; r < Math.min(rows.length, 20); r++) {
      const row = rows[r] || [];
      let dayCol = -1, timeCol = -1;
      for (let c = 0; c < Math.min(row.length, 10); c++) {
        const v = clean(row[c]).toLowerCase();
        if (v === 'day' && dayCol < 0) dayCol = c;
        if (/^(hours?|time|timings?)$/.test(v) && timeCol < 0) timeCol = c;
      }
      if (dayCol < 0 || timeCol < 0) continue;
      // The group columns end where the sheet repeats its DAY column on the right.
      // (the marker may sit on any header row, so check them all).
      let endCol = row.length - 1;
      for (let rr = 0; rr <= r + 1 && rr < rows.length; rr++) {
        const hr = rows[rr] || [];
        for (let c = timeCol + 1; c < Math.min(hr.length, endCol + 1); c++) {
          if (/^(day|sr\.?\s*no\.?|hours?|lecture|tutorial|practical)$/i.test(clean(hr[c]))) { endCol = c - 1; break; }
        }
      }
      return { hr: r, dayCol, timeCol, firstCol: timeCol + 1, endCol };
    }
    return null;
  }

  function parseBlockGrid(rows, ws, g, sheetName, out, aliases) {
    const merges = mergeIndex(ws);
    const isTimeRow = (r) => rows[r] && parseClock(rows[r][g.timeCol]) != null && rows[r][g.timeCol] !== '';
    const areaCount = (r) => { let n = 0; for (let c = g.firstCol; c <= g.endCol; c++) if (clean((rows[r] || [])[c])) n++; return n; };

    // Header rows: look for LECTURE / TUTORIAL / PRACTICAL / BRANCH labels left of the grid.
    const label = {};
    for (let r = 0; r <= g.hr + 1 && r < rows.length; r++) {
      for (let c = 0; c < g.firstCol; c++) {
        const v = clean((rows[r] || [])[c]);
        for (const k in LABELS) if (LABELS[k].test(v) && label[k] == null) label[k] = r;
      }
    }
    let groupRow, batchRow, aliasRow = null, branchRow = null;
    if (label.tutorial != null) {
      groupRow = label.tutorial;
      batchRow = label.lecture != null ? label.lecture : null;
      aliasRow = label.practical != null && label.practical !== groupRow ? label.practical : null;
      if (label.branch != null) branchRow = label.branch;
      else if (batchRow != null && batchRow > 0 && areaCount(batchRow - 1) >= 2 &&
        !(rows[batchRow - 1] || []).some((v) => /year|time\s*table/i.test(String(v)))) branchRow = batchRow - 1;
    } else {
      // No labels (e.g. PG sheet): programme row on top, group row just below it.
      let r = g.hr;
      while (r + 1 < rows.length && !isTimeRow(r + 1) && areaCount(r + 1) >= 3) r++;
      groupRow = r;
      batchRow = r > g.hr ? r - 1 : null;
    }

    // Value of a header cell, taking merged headers into account.
    const headerAt = (r, c) => {
      if (r == null) return '';
      const m = merges.get(r + ':' + c);
      return clean(m ? (rows[m.s.r] || [])[m.s.c] : (rows[r] || [])[c]);
    };

    // Units = the groups (column spans) along the group row.
    const units = [];
    for (let c = g.firstCol; c <= g.endCol;) {
      const m = merges.get(groupRow + ':' + c);
      const c1 = m ? Math.min(m.e.c, g.endCol) : (clean(rows[groupRow][c]) ? c + 1 : c);
      const name = headerAt(groupRow, c) || headerAt(batchRow, c);
      if (name) {
        units.push({
          c0: c, c1, name,
          batch: headerAt(batchRow, c),
          branch: headerAt(branchRow, c),
          alias: aliasRow != null ? headerAt(aliasRow, c) : '',
        });
      }
      c = c1 + 1;
    }
    if (!units.length) return { units: 0 };

    // Disambiguate repeated names (e.g. "G1" under two programmes).
    const byName = new Map();
    for (const u of units) { const k = groupKey(u.name); (byName.get(k) || byName.set(k, []).get(k)).push(u); }
    for (const list of byName.values()) {
      if (new Set(list.map((u) => u.batch)).size > 1) for (const u of list) if (u.batch && u.batch !== u.name) u.name = `${u.batch} ${u.name}`;
    }

    // Batch info shared by this sheet's entries, used to say "whole batch" for lectures.
    const bm = { of: {}, size: {}, branch: {} };
    for (const u of units) {
      const k = groupKey(u.name);
      if (!(k in bm.branch)) bm.branch[k] = u.branch || '';
      if (u.batch && !bm.of[k]) { bm.of[k] = u.batch; bm.size[u.batch] = (bm.size[u.batch] || 0) + 1; }
    }

    const year = /\bpg\b|post\s*grad/i.test(sheetName) ? 'PG' : (/^\d$/.test(parseYear(sheetName)) ? parseYear(sheetName) : clean(sheetName));
    for (const u of units) {
      if (u.alias && groupKey(u.alias) !== groupKey(u.name)) {
        (aliases[year] = aliases[year] || {})[groupKey(u.alias)] = u.name;
      }
    }
    const unitsIn = (a, b) => units.filter((u) => u.c1 >= a && u.c0 <= b);

    // Time rows and day blocks.
    const times = [];
    for (let r = groupRow + 1; r < rows.length; r++) if (isTimeRow(r)) times.push(r);
    let dayIdx = -1, prev = Infinity;
    const blocks = [];
    for (const r of times) {
      const t = fix12(parseClock(rows[r][g.timeCol])).m;
      if (t <= prev) { dayIdx++; blocks.push({ start: r, rows: [] }); }
      blocks[blocks.length - 1].rows.push({ r, t });
      prev = t;
    }
    blocks.forEach((b, i) => {
      const endR = i + 1 < blocks.length ? blocks[i + 1].start - 1 : b.rows[b.rows.length - 1].r + 1;
      let letters = '';
      for (let r = b.start; r <= endR; r++) letters += clean((rows[r] || [])[g.dayCol]);
      const d = parseDays(letters.replace(/[^a-z]/gi, ''));
      b.day = d.length === 1 ? d[0] : i;
    });

    let count = 0, unknown = 0;
    for (const b of blocks) {
      const open = new Map(); // unit -> practical entry that may continue
      const last = new Map(); // unit -> latest entry of the day
      b.rows.forEach(({ r, t }, i) => {
        const next = b.rows[i + 1];
        const len = next && next.t - t > 0 && next.t - t <= 120 ? next.t - t : 50;
        const row = rows[r];
        const seen = new Set();
        const current = new Map(); // unit -> entry started in this slot
        for (let c = g.firstCol; c <= g.endCol; c++) {
          const raw = row[c];
          if (raw === '' || raw == null) continue;
          const m = merges.get(r + ':' + c);
          if (m && (m.s.c !== c || m.s.r !== r)) continue;
          const a = c, bEnd = m ? Math.min(m.e.c, g.endCol) : c;
          const info = m && m.e.r > r ? m.e.r + 1 : r + 1;
          const infoRow = info < rows.length && !isTimeRow(info) ? rows[info] : [];
          const covered = unitsIn(a, bEnd);
          if (!covered.length) continue;
          const lines = String(raw).split(/\n+/).map(clean).filter(Boolean);
          const text = lines.join(' ');
          const roomM = merges.get(info + ':' + a);
          const roomEnd = roomM ? roomM.e.c : a;
          let spanEnd = bEnd;
          while (spanEnd < covered[covered.length - 1].c1 && !clean(row[spanEnd + 1]) && !merges.has(r + ':' + (spanEnd + 1)) ) spanEnd++;
          if (spanEnd < covered[covered.length - 1].c1) {
            // Nothing else in this unit's code row -> the teacher column at the unit's end is ours.
            const rest = []; for (let x = spanEnd + 1; x <= covered[covered.length - 1].c1; x++) rest.push(x);
            if (rest.every((x) => !clean(row[x]) && !(merges.get(r + ':' + x) && clean((rows[merges.get(r + ':' + x).s.r] || [])[merges.get(r + ':' + x).s.c])))) spanEnd = covered[covered.length - 1].c1;
          }
          const after = () => { for (let x = roomEnd + 1; x <= spanEnd; x++) { const v = clean(infoRow[x]); if (v) return v; } return ''; };
          if (typeof raw === 'number') continue; // stray serial numbers / times
          const tokens = (lines[0] || '').split('/').map((p) => p.trim()).filter(Boolean);
          let code = tokens.map((p) => CODE_RE.exec(p) || PAREN_RE.exec(p)).find(Boolean);
          let lineCode = null;
          if (!code && (lineCode = LINE_RE.exec(lines[0] || ''))) code = [lineCode[0], lineCode[1], lineCode[2], ''];
          const openFor = (u) => (open.get(u) || []).filter((e) => e.end === t);
          const continuing = covered.filter((u) => openFor(u).length);

          if (/^lab\b/i.test(text) || (!code && continuing.length)) {
            // Continuation of a practical from the previous slot.
            const fac = clean(infoRow[a]) || after();
            for (const u of continuing.length ? continuing : covered) {
              const list = openFor(u);
              const e = list.find((x) => x._a <= bEnd && x._b >= a) || list[0];
              if (!e || seen.has(e)) continue;
              seen.add(e);
              e.end = t + len;
              if (fac && !e.faculty) e.faculty = fac;
              if (!/^lab$/i.test(text) && !e.room) e.room = text;
            }
            continue;
          }
          if (!code) {
            // A teacher's initials written beside the subject, or spilling into the next slot's row.
            const cands = new Set();
            for (const u of covered) {
              const cur = current.get(u), prevE = last.get(u);
              if (cur && !cur.faculty) cands.add(cur);
              else if (prevE && prevE.end === t && !prevE.faculty) cands.add(prevE);
            }
            if (cands.size === 1 && /^[A-Z][A-Za-z0-9 .\/-]{0,24}$/.test(text)) {
              [...cands][0].faculty = text;
              continue;
            }
          }
          if (!code && (/^[A-Z]{2,5}\d{3}[A-Z]?$/i.test(text) || /project|seminar|training|elective|library|mentor/i.test(text))) {
            const e = {
              year, category: covered[0].branch || '', groups: covered.map((u) => u.name), all: false,
              day: b.day, start: t, end: t + len, label: '', subject: text, code: '',
              type: /project/i.test(text) ? 'Project' : '', faculty: after(), room: clean(infoRow[a]), bm,
            };
            out.push(e);
            for (const u of covered) { current.set(u, e); last.set(u, e); }
            count++;
            continue;
          }
          if (!code) { unknown++; if (root.TT_DEBUG) console.log("?", sheetName, r + 1, c, JSON.stringify(raw)); continue; }

          const type = parseType(code[2]);
          const subject = lineCode ? clean(lineCode[1]) : tokens.map((p) => { const x = CODE_RE.exec(p) || PAREN_RE.exec(p); return x ? (x[1] + (x[3] ? ' ' + x[3] : '')).trim() : p; }).join(' / ');
          const room = clean(infoRow[a]) || lines.slice(1).join(' ');
          const e = {
            year, category: covered[0].branch || '',
            groups: covered.map((u) => u.name), all: false,
            day: b.day, start: t, end: t + len, label: '',
            subject, code: '', type, faculty: after(), room,
            bm,
          };
          out.push(e);
          for (const u of covered) { current.set(u, e); last.set(u, e); }
          count++;
          if (type === 'Lab') {
            e._a = a; e._b = bEnd;
            for (const u of covered) (open.get(u) || open.set(u, []).get(u)).push(e);
          }
        }
      });
    }
    return { units: units.length, entries: count, unknown, days: blocks.length };
  }

  // ---------- main ----------

  function parseWorkbook(wb, XLSX) {
    const entries = [];
    const report = [];
    const aliases = {};

    wb.SheetNames.forEach((name) => {
      const ws = wb.Sheets[name];
      if (!ws || !ws['!ref']) return;
      const block = detectBlockGrid(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true }));
      if (block) {
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
        const res = parseBlockGrid(raw, ws, block, name, entries, aliases);
        if (res.entries) {
          report.push({ sheet: name, layout: 'block', rows: res.entries, groupColumns: res.units, unknown: res.unknown });
          return;
        }
      }
      const rows = sheetRows(XLSX, ws);
      const hdr = detectHeader(rows);
      if (!hdr) {
        report.push({ sheet: name, layout: 'skipped', note: 'No header row with a Day and Time column found.' });
        return;
      }
      const sheetYear = parseYear(name);
      const sheetYearOk = /^\d$/.test(sheetYear) ? sheetYear : '';
      const headerRow = rows[hdr.r];
      const map = hdr.map;
      const isList = map.group != null && (map.subject != null || map.code != null);
      const before = entries.length;

      if (isList) {
        parseList(rows, hdr.r, map, sheetYearOk, entries);
        report.push({
          sheet: name, layout: 'list', rows: entries.length - before,
          columns: Object.fromEntries(Object.entries(map).map(([f, c]) => [f, clean(headerRow[c])])),
        });
      } else {
        const known = new Set(Object.values(map));
        const groupCols = [];
        headerRow.forEach((h, c) => { if (!known.has(c) && looksLikeGroupHeader(h)) groupCols.push(c); });
        if (!groupCols.length) {
          report.push({ sheet: name, layout: 'skipped', note: 'Found Day/Time columns but no Group column or group-named columns.' });
          return;
        }
        parseGrid(rows, hdr.r, map, groupCols, headerRow, sheetYearOk, entries);
        report.push({ sheet: name, layout: 'grid', rows: entries.length - before, groupColumns: groupCols.length });
      }
    });

    finalize(entries);
    return { entries, report, aliases };
  }

  function parseList(rows, hr, map, sheetYear, out) {
    const get = (row, f) => (map[f] != null ? row[map[f]] : '');
    let lastDay = '', lastTime = '', lastStart = '', lastEnd = '', lastYear = '', lastCat = '';
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.every((c) => c === '' || c == null)) continue;

      // Forward-fill columns commonly left blank under a heading.
      let dayV = get(row, 'day'); if (dayV === '') dayV = lastDay; else lastDay = dayV;
      let timeV = get(row, 'time'); if (timeV === '' && map.time != null) timeV = lastTime; else lastTime = timeV;
      let startV = get(row, 'start'); if (startV === '' && map.start != null) startV = lastStart; else lastStart = startV;
      let endV = get(row, 'end'); if (endV === '' && map.end != null) endV = lastEnd; else lastEnd = endV;
      let yearV = get(row, 'year'); if (yearV === '' && map.year != null) yearV = lastYear; else lastYear = yearV;
      let catV = get(row, 'category'); if (catV === '' && map.category != null) catV = lastCat; else lastCat = catV;

      const subject = clean(get(row, 'subject'));
      const code = clean(get(row, 'code'));
      if (!subject && !code) continue;
      const days = parseDays(dayV);
      if (!days.length) continue;
      const g = parseGroups(get(row, 'group'));
      if (!g.all && !g.groups.length) continue;

      let start = null, end = null, label = '';
      if (map.start != null) {
        const a = parseRange(typeof startV === 'string' ? startV : startV);
        const b = parseRange(endV);
        if (a) { start = a.start; end = a.end; }
        if (b && b.start != null) end = b.start;
        if (start != null && end != null && end <= start && end + 720 < 1440) end += 720;
      }
      if (start == null && map.time != null) {
        const t = parseRange(timeV);
        if (t) { start = t.start; end = t.end; }
      }
      if (start == null) label = clean(timeV || startV);

      const typeRaw = get(row, 'type');
      for (const d of days) {
        out.push({
          year: map.year != null ? parseYear(yearV) : sheetYear,
          category: clean(catV),
          groups: g.groups.map(clean),
          all: g.all,
          day: d, start, end, label,
          subject: subject || code,
          code: subject ? code : '',
          type: typeRaw !== '' ? parseType(typeRaw) : extractType(subject),
          faculty: clean(get(row, 'faculty')),
          room: clean(get(row, 'room')),
        });
      }
    }
  }

  function parseGrid(rows, hr, map, groupCols, headerRow, sheetYear, out) {
    let lastDay = '';
    for (let r = hr + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      let dayV = row[map.day]; if (dayV === '' || dayV == null) dayV = lastDay; else lastDay = dayV;
      const days = parseDays(dayV);
      if (!days.length) continue;
      let start = null, end = null, label = '';
      if (map.start != null) {
        const a = parseRange(row[map.start]), b = parseRange(map.end != null ? row[map.end] : '');
        if (a) { start = a.start; end = b ? b.start : a.end; }
      } else {
        const t = parseRange(row[map.time]);
        if (t) { start = t.start; end = t.end; } else label = clean(row[map.time]);
      }
      if (start == null && !label) continue;
      const yearV = map.year != null ? parseYear(row[map.year]) : sheetYear;
      const catV = map.category != null ? clean(row[map.category]) : '';

      for (const c of groupCols) {
        const text = String(row[c] == null ? '' : row[c]).trim();
        if (!text || /^(-+|free|break|lunch|nil|none|x)$/i.test(text)) continue;
        const lines = text.split(/\n+/).map(clean).filter(Boolean);
        const first = lines[0];
        let room = '', faculty = '';
        const roomM = text.match(/(?:room|venue|lab|lt|hall)\s*[:#-]?\s*([A-Z0-9-]+)/i);
        if (roomM) room = roomM[0].trim();
        if (lines.length > 1) faculty = lines.slice(1).filter((l) => !roomM || !l.includes(roomM[0])).join(', ');
        for (const d of days) {
          out.push({
            year: yearV, category: catV,
            groups: [clean(headerRow[c])], all: false,
            day: d, start, end, label,
            subject: first.replace(/[\(\[]\s*(L|T|P|Lec|Lecture|Tut|Tutorial|Lab|Practical|Prac)\s*[\)\]]/i, '').trim() || first,
            code: '', type: extractType(text), faculty, room,
          });
        }
      }
    }
  }

  /** Adds keys and computes, for each class, which other groups share it. */
  function finalize(entries) {
    const shared = new Map();
    for (const e of entries) {
      e.groupKeys = e.groups.map(groupKey);
      e.slotKey = [e.year, e.day, e.start, e.end, e.label.toLowerCase(),
        norm(e.subject), norm(e.room)].join('|');
      if (!shared.has(e.slotKey)) shared.set(e.slotKey, { keys: new Map(), all: false });
      const s = shared.get(e.slotKey);
      if (e.all) s.all = true;
      e.groups.forEach((g, i) => s.keys.set(e.groupKeys[i], g));
    }
    for (const e of entries) e.shared = shared.get(e.slotKey);
  }

  /** Builds {year -> {category -> Map(groupKey -> displayName)}}. */
  function buildIndex(entries) {
    const idx = {};
    for (const e of entries) {
      const y = (idx[e.year] = idx[e.year] || {});
      e.groups.forEach((g, i) => {
        const k = e.groupKeys[i];
        const cat = e.bm ? e.bm.branch[k] || '' : e.category;
        const c = (y[cat] = y[cat] || new Map());
        if (!c.has(k)) c.set(k, g);
      });
    }
    return idx;
  }

  /**
   * Returns the timetable for one group: [{day, slots: [{start, end, label, classes: [...]}]}].
   * Classes at overlapping times are put in the same slot so they show side by side.
   */
  function timetableFor(entries, year, category, gKey) {
    const seen = new Set();
    const mine = [];
    for (const e of entries) {
      if (e.year !== year) continue;
      if (category != null && e.category !== category) continue;
      if (!e.all && !e.groupKeys.includes(gKey)) continue;
      if (seen.has(e.slotKey)) continue;
      seen.add(e.slotKey);
      const together = [];
      for (const [k, name] of e.shared.keys) if (k !== gKey) together.push(name);
      together.sort(naturalCompare);
      mine.push({ ...e, together, togetherAll: e.shared.all, togetherParts: summarize(e, gKey) });
    }
    const byDay = new Map();
    for (const e of mine) {
      if (!byDay.has(e.day)) byDay.set(e.day, []);
      byDay.get(e.day).push(e);
    }
    const days = [];
    for (const [day, list] of [...byDay].sort((a, b) => a[0] - b[0])) {
      list.sort((a, b) => (a.start ?? 1e9) - (b.start ?? 1e9) || (a.end ?? 0) - (b.end ?? 0) || a.label.localeCompare(b.label, undefined, { numeric: true }));
      const slots = [];
      for (const e of list) {
        const last = slots[slots.length - 1];
        const overlaps = last && (e.start == null
          ? last.start == null && last.label === e.label
          : last.start != null && e.start < (last.end ?? last.start + 1));
        if (overlaps) {
          last.classes.push(e);
          if (e.end != null && (last.end == null || e.end > last.end)) last.end = e.end;
        } else {
          slots.push({ start: e.start, end: e.end, label: e.label, classes: [e] });
        }
      }
      days.push({ day, name: DAYS[day], slots });
    }
    return days;
  }

  /**
   * Short description of who shares a class, by batch where possible:
   * ["all of 1A1", "1A2", "1A3"] instead of listing 24 groups.
   */
  function summarize(e, gKey) {
    if (!e.bm) return null;
    const byBatch = new Map();
    const loose = [];
    for (const [k, name] of e.shared.keys) {
      const b = e.bm.of[k];
      if (b) (byBatch.get(b) || byBatch.set(b, []).get(b)).push({ k, name }); else if (k !== gKey) loose.push(name);
    }
    const own = e.bm.of[gKey];
    const parts = [];
    let wholeOwn = false;
    for (const [b, list] of [...byBatch].sort((x, y) => naturalCompare(x[0], y[0]))) {
      const full = list.length === e.bm.size[b] && list.length > 1;
      if (b === own) {
        if (full) wholeOwn = true;
        else parts.push(...list.filter((x) => x.k !== gKey).map((x) => x.name));
      } else if (full) parts.push(b);
      else parts.push(...list.map((x) => x.name));
    }
    parts.push(...loose);
    return { wholeOwn: wholeOwn ? own : '', parts };
  }

  function naturalCompare(a, b) {
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
  }

  function fmtTime(m) {
    if (m == null) return '';
    const h = Math.floor(m / 60) % 24, mm = m % 60;
    const h12 = ((h + 11) % 12) + 1;
    return `${h12}:${String(mm).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`;
  }

  const api = {
    DAYS, parseWorkbook, buildIndex, timetableFor, groupKey, naturalCompare, fmtTime,
    // exported for tests
    _: { parseDays, parseClock, parseRange, parseYear, parseGroups, matchField, parseType },
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TTParser = api;
})(typeof window !== 'undefined' ? window : globalThis);
