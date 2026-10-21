// Pronunciation guides, narration-only rule (author decision 2026-10-20).
// Each word gets its guide in brackets at its FIRST mention in narration, never inside dialogue.
// Words that are only ever spoken get no inline guide; the front-matter page covers every word.
// Usage:
//   node pronounce.js <Indic Respelling Checker.html> <chapter dir> [--write]
// Without --write it prints what it would change. It first strips every existing guide (including
// the old "outside the closing quote" placements), then places them again under the new rule, and
// writes 00-names-and-pronunciation.md (the front-matter page) into the chapter dir.
const fs = require('fs'), path = require('path');
const [htmlPath, dir] = process.argv.slice(2);
const WRITE = process.argv.includes('--write');

const html = fs.readFileSync(htmlPath, 'utf8');
const core = html.split('/*CORE-START*/')[1].split('/*CORE-END*/')[0];
const api = new Function(core + '\nreturn {analyzeWord};')();
const ENTRIES = require('./pronounce-entries.js').map(([phrase, native, rules, hand]) => {
  let guide;
  if (rules === 'hand') guide = hand;
  else { const p = native.split(' '), r = rules.split(' '); guide = p.map((w, i) => api.analyzeWord(w, r[i] || r[0]).resp).join(' '); }
  return { phrase, guide };
});
// Old placements that used the italic-tail form: *Trishanku.* -> *Trishanku* (guide).
const ITALIC_TAIL = new Set(['Trishanku', 'Ashwatthama', 'Rahu']);

const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const files = fs.readdirSync(dir).filter(f => /^\d\d-.*merged\.md$/.test(f)).sort();
const text = {}; files.forEach(f => { text[f] = fs.readFileSync(path.join(dir, f), 'utf8'); });
const report = [];

// 1. Strip every existing guide, restoring the original punctuation.
let stripped = 0;
for (const f of files) {
  let s = text[f];
  for (const { phrase, guide } of ENTRIES) {
    const g = esc(guide);
    s = s.replace(new RegExp('(["”]) \\(' + g + '\\)([,.?!…—]*)', 'g'), (m, q, p) => { stripped++; return p + q; });
    if (ITALIC_TAIL.has(phrase))
      s = s.replace(new RegExp(esc(phrase) + '\\* \\(' + g + '\\)([,.?!…]+)', 'g'), (m, p) => { stripped++; return phrase + p + '*'; });
    s = s.replace(new RegExp(' \\(' + g + '\\)', 'g'), () => { stripped++; return ''; });
  }
  text[f] = s;
}
report.push(`Stripped ${stripped} existing guides.`);

function quoteOpenBefore(line, idx) {
  let open = false;
  for (let i = 0; i < idx; i++) { const c = line[i]; if (c === '"') open = !open; else if (c === '“') open = true; else if (c === '”') open = false; }
  return open;
}

// 2. Place each guide at the first narration mention.
const spokenOnly = [];
for (const { phrase, guide } of ENTRIES) {
  let placed = false;
  const re = new RegExp('(?<![\\p{L}\\p{M}])' + esc(phrase) + '(?![\\p{L}\\p{M}])', 'gu');
  for (const f of files) {
    if (placed) break;
    const eol = text[f].includes('\r\n') ? '\r\n' : '\n';
    const lines = text[f].split(/\r?\n/);
    let inBody = false;
    for (let i = 0; i < lines.length && !placed; i++) {
      const line = lines[i];
      if (/^## /.test(line)) { inBody = true; continue; }
      if (!inBody || /^#/.test(line)) continue;
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(line))) {
        if (quoteOpenBefore(line, m.index)) continue;          // dialogue: skip
        let end = m.index + m[0].length;
        const poss = /^['’]s(?![\p{L}])/u.exec(line.slice(end)); if (poss) end += poss[0].length;
        if (line[end] === '*' && line[m.index - 1] === '*') end += 1;
        const bracket = ' (' + guide + ')';
        const itTail = line[m.index - 1] === '*' ? /^([,.?!…]+)\*/.exec(line.slice(end)) : null;
        const out = itTail
          ? line.slice(0, end) + '*' + bracket + itTail[1] + line.slice(end + itTail[0].length)
          : line.slice(0, end) + bracket + line.slice(end);
        lines[i] = out; text[f] = lines.join(eol); placed = true;
        const s = Math.max(0, m.index - 45);
        report.push(`${f.slice(0, 2)}:${i + 1}  ${phrase} (${guide})\n    …${out.slice(s, end + 40 + bracket.length)}…`);
        break;
      }
    }
  }
  if (!placed) spokenOnly.push(phrase);
}
report.push('\nOnly ever spoken (front-matter page only): ' + (spokenOnly.join(', ') || 'none'));

// 3. Front-matter page.
const GROUPS = [
  ['Houses and places', ['Ashoka', 'Drona', 'Kubera', 'Chanakya', 'Meru', 'Patal Loka', 'Kalinga', 'Sākṣī', 'gurukul', 'gana']],
  ['Titles', ['Dhammapala', 'Dhamma', 'Arthashastri', 'Sarvajña', 'Draṣṭā']],
  ['People', ['Ravi', 'Rohan', 'Anandi', 'Vayun', 'Nayan', 'Baju', 'Kaan', 'Jeebh', 'Dev', 'Meera', 'Arjun', 'Shu Chan']],
  ['Stories, figures and old words', ['Vishnu', 'Ramayana', 'Sanjeevani', 'Vedas', 'karma', 'graha', 'Trishanku', 'Ashwatthama', 'Rahu', 'Rani Lakshmibai', 'Kusunoki Masashige']],
  ['Everyday words', ['dal', 'Wadō-ryū', 'dojo', 'dan', 'sensei', 'kimono']],
];
const byPhrase = Object.fromEntries(ENTRIES.map(e => [e.phrase, e.guide]));
const grouped = new Set(GROUPS.flatMap(g => g[1]));
const missing = ENTRIES.map(e => e.phrase).filter(p => !grouped.has(p));
if (missing.length) throw new Error('Not in any front-matter group: ' + missing.join(', '));
let page = '# Names and Pronunciation\n\n'
  + 'Many names in this book come from Sanskrit, Pali, Hindi, Japanese and Chinese. Stressed syllables are in capitals. '
  + 'Each guide also appears once in the story, in brackets, the first time the word is used outside of speech.\n';
for (const [title, words] of GROUPS) {
  page += `\n## ${title}\n`;
  for (const w of words) page += `\n**${w}** · ${byPhrase[w]}\n`;
}

console.log(report.join('\n'));
if (WRITE) {
  files.forEach(f => fs.writeFileSync(path.join(dir, f), text[f], 'utf8'));
  fs.writeFileSync(path.join(dir, '00-names-and-pronunciation.md'), page, 'utf8');
  console.log('\nWrote ' + files.length + ' chapter files and 00-names-and-pronunciation.md.');
} else console.log('\n(dry run; add --write to apply)');
