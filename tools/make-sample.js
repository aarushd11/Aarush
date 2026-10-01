// Generates samples/sample-timetable.xlsx: a made-up timetable with 4 years,
// 3 categories and 10 groups per category/year (120 groups), in "list" layout.
// Run: npm i xlsx && node tools/make-sample.js
const path = require('path');
const XLSX = require('xlsx');

const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
const slots = ['09:00', '10:00', '11:00', '12:00', '14:00', '15:00', '16:00'];
const plus = (t, h) => `${String(+t.slice(0, 2) + h).padStart(2, '0')}:00`;
const subjects = {
  CSE: ['Data Structures', 'Operating Systems', 'Discrete Maths', 'Computer Networks', 'DBMS'],
  ECE: ['Signals & Systems', 'Analog Circuits', 'Digital Electronics', 'Electromagnetics', 'Control Systems'],
  ME: ['Thermodynamics', 'Fluid Mechanics', 'Machine Design', 'Strength of Materials', 'Manufacturing'],
};
const faculty = ['Dr. Rao', 'Dr. Mehta', 'Prof. Iyer', 'Dr. Kaur', 'Prof. Sen', 'Dr. Thomas', 'Dr. Gupta', 'Prof. Nair'];

const rows = [['Year', 'Branch', 'Group', 'Day', 'Start Time', 'End Time', 'Subject Code', 'Subject', 'Type', 'Faculty', 'Room']];
let seed = 7;
const rnd = (n) => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed % n; };

for (let year = 1; year <= 4; year++) {
  for (const [ci, cat] of Object.keys(subjects).entries()) {
    const busy = {}; // "G|day|slot" -> true
    const take = (gs, d, s, len = 1) => {
      for (const g of gs) for (let k = 0; k < len; k++) if (busy[`${g}|${d}|${s + k}`]) return false;
      for (const g of gs) for (let k = 0; k < len; k++) busy[`${g}|${d}|${s + k}`] = true;
      return true;
    };
    const prefix = `${cat[0]}${year}`;
    const G = (i) => `${prefix}${String(i).padStart(2, '0')}`;
    subjects[cat].forEach((subj, si) => {
      const code = `${cat}${year}0${si + 1}`;
      // Lectures: groups 1-5 together and 6-10 together, 3 times a week.
      for (const [a, b] of [[1, 5], [6, 10]]) {
        let placed = 0, tries = 0;
        while (placed < 3 && tries++ < 200) {
          const d = rnd(5), s = rnd(slots.length);
          const gs = Array.from({ length: b - a + 1 }, (_, i) => G(a + i));
          if (!take(gs, `${d}`, s)) continue;
          rows.push([`${year}`, cat, `${G(a)}-${G(b)}`, days[d], slots[s], plus(slots[s], 1), code, subj, 'L',
            faculty[(si + ci + a) % faculty.length], `LT-${ci + 1}0${(si % 4) + 1}`]);
          placed++;
        }
      }
      // Tutorial: pairs of groups, listed as one row per group.
      for (let g = 1; g <= 10; g += 2) {
        let tries = 0;
        while (tries++ < 200) {
          const d = rnd(5), s = rnd(slots.length);
          if (!take([G(g), G(g + 1)], `${d}`, s)) continue;
          for (const gg of [g, g + 1]) rows.push([`${year}`, cat, G(gg), days[d], slots[s], plus(slots[s], 1), code, subj, 'T',
            faculty[(si + g) % faculty.length], `TR-${(g % 6) + 1}`]);
          break;
        }
      }
    });
    // Labs: 2-hour practical per group for two subjects.
    for (let g = 1; g <= 10; g++) {
      for (const si of [0, 2]) {
        let tries = 0;
        while (tries++ < 300) {
          const d = rnd(5), s = rnd(slots.length - 1);
          if (s === 3) continue; // don't span lunch
          if (!take([G(g)], `${d}`, s, 2)) continue;
          rows.push([`${year}`, cat, G(g), days[d], slots[s], plus(slots[s], 2), `${cat}${year}0${si + 1}`, subjects[cat][si], 'P',
            faculty[(si + g + 3) % faculty.length], `${cat} Lab ${(g % 3) + 1}`]);
          break;
        }
      }
    }
    // A common seminar for everyone in the year+branch.
    rows.push([`${year}`, cat, 'All', 'Friday', '16:00', '17:00', '', 'Department Seminar', 'Seminar', '', 'Auditorium']);
  }
}

const ws = XLSX.utils.aoa_to_sheet(rows);
ws['!cols'] = [6, 8, 12, 11, 10, 10, 12, 24, 8, 14, 12].map((w) => ({ wch: w }));
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Timetable');
const out = path.join(__dirname, '..', 'samples', 'sample-timetable.xlsx');
XLSX.writeFile(wb, out);
console.log(`Wrote ${rows.length - 1} rows to ${out}`);
