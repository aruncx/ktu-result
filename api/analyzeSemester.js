const pdf = require('pdf-parse');
const XLSX = require('xlsx-js-style');
const nodeFetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));

/* ── CONSTANTS ──────────────────────────────────────────────────── */
const GP = { S: 10, 'A+': 9, A: 8.5, 'B+': 8, B: 7.5, 'C+': 7, C: 6.5, D: 6, P: 5.5, F: 0, FE: 0, I: 0, Ab: 0, Absent: 0, Withheld: 0 };
const OK = new Set(['S', 'A+', 'A', 'B+', 'B', 'C+', 'C', 'D', 'P']);
const DEPT_MAP = {
  CSE: 'Computer Science & Engineering', ECE: 'Electronics & Communication',
  EEE: 'Electrical & Electronics', ME: 'Mechanical Engineering',
  CE: 'Civil Engineering', IT: 'Information Technology',
  AIM: 'Artificial Intelligence and Machine Learning'
};

/* ── ANALYSIS ENGINE ────────────────────────────────────────────── */
function analyze(semesters) {
  const allStusGlobal = semesters.flatMap(sem => sem.departments.flatMap(d => d.students));
  const yearsGlobal = allStusGlobal.map(s => parseInt(s.admissionYear)).filter(y => !isNaN(y));
  const globalMaxYear = yearsGlobal.length > 0 ? Math.max(...yearsGlobal) : null;

  return semesters.map(sem => {
    const depts = sem.departments.map(dept => {
      const targetStudents = globalMaxYear ? dept.students.filter(s => parseInt(s.admissionYear) === globalMaxYear) : dept.students;
      const tot = targetStudents.length;
      const pass = targetStudents.filter(s => s.allPassed).length;
      const avgSGPA = tot ? targetStudents.reduce((a, s) => a + parseFloat(s.sgpa), 0) / tot : 0;
      const sm = {};
      targetStudents.forEach(stu => stu.subjects.forEach(sub => {
        if (!sm[sub.code]) sm[sub.code] = { code: sub.code, name: sub.name, total: 0, passed: 0 };
        sm[sub.code].total++;
        if (sub.passed) sm[sub.code].passed++;
      }));
      const subjects = Object.values(sm).map(s => ({ ...s, failed: s.total - s.passed, pp: s.total ? ((s.passed / s.total) * 100).toFixed(1) : '0' }));
      return { ...dept, students: targetStudents, tot, pass, fail: tot - pass, pp: tot ? ((pass / tot) * 100).toFixed(1) : '0', avgSGPA: avgSGPA.toFixed(2), subjects };
    }).sort((a, b) => parseFloat(b.pp) - parseFloat(a.pp));

    const totS = depts.reduce((a, d) => a + d.tot, 0);
    const totP = depts.reduce((a, d) => a + d.pass, 0);
    const allSubs = depts.flatMap(d => d.subjects);
    const sorted = [...allSubs].sort((a, b) => parseFloat(a.pp) - parseFloat(b.pp));
    return {
      ...sem, departments: depts,
      ov: {
        totS, totP, totF: totS - totP, pp: totS ? ((totP / totS) * 100).toFixed(1) : '0',
        best: depts[0], worst: depts[depts.length - 1], tough: sorted[0], easiest: sorted[sorted.length - 1], allSubs
      }
    };
  });
}

function normDept(s) {
  const l = s.toLowerCase().trim();
  if (l.includes('computer') || l === 'cse') return 'CSE';
  if (l.includes('electronics') || l.includes('communication') || l === 'ece') return 'ECE';
  if (l.includes('electrical') || l === 'eee') return 'EEE';
  if (l.includes('mechanical') || l === 'me') return 'ME';
  if (l.includes('civil') || l === 'ce') return 'CE';
  if (l.includes('information') || l === 'it') return 'IT';
  if (l.includes('artificial') || l.includes('machine') || l === 'aim' || l === 'ad') return 'AIM';
  const clean = s.trim().toUpperCase();
  if (['CODE', 'COUR', 'STRE', 'PROG', 'BRAN'].includes(clean.slice(0, 4))) return 'UNK';
  const m = s.match(/[A-Z]{2,4}(?=\d)/); if (m) return m[0].slice(0, 3);
  return clean.slice(0, 4) || 'UNK';
}

function deptFromRegNo(r) {
  const m = r.match(/^L?[A-Z]{3,}\d{2}([A-Z]{2,3})\d+/i);
  if (!m) return null;
  const c = m[1].toUpperCase();
  const MAP = {
    'CS': 'CSE', 'CSE': 'CSE', 'EC': 'ECE', 'ECE': 'ECE', 'EE': 'EEE', 'EEE': 'EEE',
    'ME': 'ME', 'CE': 'CE', 'IT': 'IT', 'AD': 'AIM', 'AM': 'AIM', 'AIM': 'AIM'
  };
  return MAP[c] || c;
}

function cleanGrade(raw) {
  const s = raw.replace(/[^A-Za-z+]/g, '').toUpperCase();
  const MAP = {
    'S': 'S', 'A+': 'A+', 'A': 'A', 'B+': 'B+', 'B': 'B', 'C+': 'C+', 'C': 'C', 'D': 'D', 'P': 'P',
    'F': 'F', 'FE': 'FE', 'I': 'I', 'AB': 'Ab', 'ABSENT': 'Ab', 'WITHHELD': 'Withheld'
  };
  return MAP[s] ?? s;
}

const GRADE_RE = /^(S|A\+|A|B\+|B|C\+|C|D|P|F|FE|I|Ab|AB|Absent|Withheld)$/i;
const CODE_RE = /^[A-Z]{2,6}\d{3,4}[A-Z]?$/;
const REG_RE = /[A-Z]{2,5}\d{2,4}[A-Z]{1,5}\d{2,4}/;

function parseTxt(txt) {
  const lines = txt.split(/\r?\n/);
  const students = [];
  const subNameMap = {};
  let cur = null;
  let collegeName = 'Unknown College';
  let semNumber = 'Semester';

  const collegeMatch = txt.match(/Exam\s+Centre:\s*([^\n\r]+)/i);
  if (collegeMatch) collegeName = collegeMatch[1].trim();

  const semMatch = txt.match(/B\.Tech\s+(S\d)/i) || txt.match(/(S\d)\s+Result/i);
  if (semMatch) semNumber = semMatch[1].toUpperCase();

  const pushCur = () => { if (cur && cur.subjects.length > 0) students.push(finishStu(cur)); };

  for (let li = 0; li < lines.length; li++) {
    const line = lines[li];
    const raw = line.trim();
    if (!raw) continue;

    const subDefMatch = raw.match(/([A-Z]{2,6}\d{3,4}[A-Z]?)\s+([A-Z][A-Z\d\s,.:/()-]{5,150})/i);
    if (subDefMatch && !raw.includes('(') && !REG_RE.test(raw)) {
      const code = subDefMatch[1].toUpperCase();
      const name = subDefMatch[2].trim();
      if (!subNameMap[code] || subNameMap[code] === 'Subject') subNameMap[code] = name;
    }

    const rnLabel = raw.match(/(?:Reg(?:ister)?(?:\s*No\.?|Number|\.|No)|Roll\s*No\.?|Enroll(?:ment)?\s*No\.?|Reg\.?\s*No\.?)[:\s.-]*([A-Z0-9]{8,15})/i);
    const rnBare = !rnLabel && (REG_RE.exec(raw) || raw.match(/^[A-Z]{3}\d{2}[A-Z]{2}\d{3}$/));
    const rnVal = rnLabel ? rnLabel[1] : rnBare ? rnBare[0] : null;

    if (rnVal) {
      pushCur();
      const cleanRn = rnVal.replace(/["']/g, '').trim();
      const dept = deptFromRegNo(cleanRn) || 'Unknown';
      cur = {
        regNo: cleanRn, department: dept, subjects: [], lateral: cleanRn.toUpperCase().startsWith('L'),
        admissionYear: '20' + (cleanRn.match(/[A-Z]{3,}(\d{2})/)?.[1] || '??')
      };
    }

    if (!cur) continue;
    const words = raw.split(/\s+/);

    const dpLabel = raw.match(/(?:Branch|Programme(?:me)?|Department|Course(?!\s*Code)|Stream)\s*[:\s.-]+(.{3,60}?)(?:\s{2,}|$)/i);
    if (dpLabel) { cur.department = normDept(dpLabel[1]); }

    const subWithGradeMatch = raw.match(/([A-Z]{2,6}\d{3,4}[A-Z]?)\s*[([({]\s*([A-Za-z+]{1,10})\s*[)\]})]/i);
    if (subWithGradeMatch) {
      const code = subWithGradeMatch[1].toUpperCase();
      const grade = cleanGrade(subWithGradeMatch[2]);
      if (GRADE_RE.test(grade) && !cur.subjects.find(s => s.code === code)) {
        const name = subNameMap[code] || 'Subject';
        cur.subjects.push({ code, name, grade, passed: OK.has(grade), gp: GP[grade] ?? 0 });
        let remaining = raw.replace(subWithGradeMatch[0], '');
        while (true) {
          const next = remaining.match(/([A-Z]{2,6}\d{3,4}[A-Z]?)\s*[([({]\s*([A-Za-z+]{1,10})\s*[)\]})]/i);
          if (!next) break;
          const c = next[1].toUpperCase();
          const n = subNameMap[c] || 'Subject';
          const g = cleanGrade(next[2]);
          if (GRADE_RE.test(g) && !cur.subjects.find(s => s.code === c)) cur.subjects.push({ code: c, name: n, grade: g, passed: OK.has(g), gp: GP[g] ?? 0 });
          remaining = remaining.replace(next[0], '');
        }
      }
    }

    if (CODE_RE.test(words[0])) {
      const parts = raw.split(/\s{2,}/);
      if (parts.length >= 2) {
        const code = parts[0].trim();
        const gradeToken = cleanGrade(parts[parts.length - 1]);
        const name = parts.slice(1, parts.length - 1).join(' ').trim() || 'Subject';
        if (GRADE_RE.test(gradeToken) && !cur.subjects.find(s => s.code === code)) cur.subjects.push({ code, name, grade: gradeToken, passed: OK.has(gradeToken), gp: GP[gradeToken] ?? 0 });
      }
    }

    for (let i = 0; i < words.length - 1; i++) {
      const possibleCode = words[i].toUpperCase();
      if (CODE_RE.test(possibleCode)) {
        for (let j = 1; j <= 3 && (i + j) < words.length; j++) {
          const possibleGrade = cleanGrade(words[i + j]);
          if (GRADE_RE.test(possibleGrade)) {
            if (!cur.subjects.find(s => s.code === possibleCode)) {
              cur.subjects.push({ code: possibleCode, name: subNameMap[possibleCode] || 'Subject', grade: possibleGrade, passed: OK.has(possibleGrade), gp: GP[possibleGrade] ?? 0 });
              i += j;
            }
            break;
          }
        }
      }
    }
  }

  pushCur();

  if (students.length === 0) {
    const tokens = txt.split(/\s+/);
    let fc = null;
    for (let ti = 0; ti < tokens.length; ti++) {
      const t = tokens[ti];
      if (REG_RE.test(t) && t.length >= 9) {
        if (fc && fc.subjects.length > 0) students.push(finishStu(fc));
        fc = { regNo: t, name: 'Unknown', department: deptFromRegNo(t) || 'Unknown', subjects: [] };
      } else if (fc && CODE_RE.test(t)) {
        let gi = ti + 1;
        while (gi < tokens.length && !GRADE_RE.test(cleanGrade(tokens[gi])) && gi < ti + 8) gi++;
        if (gi < tokens.length) {
          const grade = cleanGrade(tokens[gi]);
          if (GRADE_RE.test(grade) && !fc.subjects.find(s => s.code === t)) {
            const name = subNameMap[t] || 'Subject';
            fc.subjects.push({ code: t, name, grade, passed: OK.has(grade), gp: GP[grade] ?? 0 });
            ti = gi;
          }
        }
      }
    }
    if (fc && fc.subjects.length > 0) students.push(finishStu(fc));
  }

  if (!students.length) return null;
  const dm = {};
  students.forEach(s => {
    const d = s.department || 'Unknown';
    if (!dm[d]) dm[d] = { code: d, name: DEPT_MAP[d] || d, students: [] };
    dm[d].students.push(s);
  });
  return [{ id: 'uploaded', label: 'Uploaded Semester', departments: Object.values(dm), meta: { college: collegeName, semester: semNumber } }];
}

function finishStu(s) {
  const cred = s.subjects.length * 4, gp = s.subjects.reduce((a, x) => a + x.gp * 4, 0);
  return { ...s, sgpa: cred ? (gp / cred).toFixed(2) : 0, allPassed: s.subjects.every(x => x.passed) };
}

async function exportExcelBase64(sem) {
  const XL = XLSX;
  const wb = XL.utils.book_new();
  const semLabel = sem.label || 'Semester';
  const bd = (s = 'thin', c = 'CCCCCC') => ({ top: { style: s, color: { rgb: c } }, bottom: { style: s, color: { rgb: c } }, left: { style: s, color: { rgb: c } }, right: { style: s, color: { rgb: c } } });
  const S = {
    title: { font: { name: 'Calibri', bold: true, sz: 16, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1A3A6E' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' } },
    info: { font: { name: 'Calibri', sz: 10, italic: true, color: { rgb: '3A4D6B' } }, fill: { fgColor: { rgb: 'EBF3FB' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'C0D4EE') },
    hdr: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1B4F9B' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bd('medium', '1B4F9B') },
    sheetSub: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: '1A3A6E' } }, fill: { fgColor: { rgb: 'D6E4F7' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'A0B8D8') },
    kpiLbl: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '444444' } }, fill: { fgColor: { rgb: 'F2F2F2' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'DDDDDD') },
    kpiBlue: { font: { name: 'Calibri', bold: true, sz: 14, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1B4F9B' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', '1B4F9B') },
    kpiGreen: { font: { name: 'Calibri', bold: true, sz: 14, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '28A745' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', '28A745') },
    kpiRed: { font: { name: 'Calibri', bold: true, sz: 14, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: 'DC3545' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'DC3545') },
    secGreen: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '2D6A4F' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bd('thin', '2D6A4F') },
    secBlue: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1B3A6B' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bd('thin', '1B3A6B') },
    even: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd() },
    odd: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'E8F0FE' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd() },
    evenL: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bd() },
    oddL: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'E8F0FE' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bd() },
    pass: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '155724' } }, fill: { fgColor: { rgb: 'D4EDDA' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'C3E6CB') },
    fail: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '7B1D1D' } }, fill: { fgColor: { rgb: 'F8D7DA' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'F5C6CB') },
    warn: { font: { name: 'Calibri', sz: 11, color: { rgb: '856404' } }, fill: { fgColor: { rgb: 'FFF3CD' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'FFEEBA') },
    gold: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '7B5200' } }, fill: { fgColor: { rgb: 'FFF3CD' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bd('thin', 'FFEEBA') },
    footer: { font: { name: 'Calibri', italic: true, sz: 10, color: { rgb: '1A3A6E' } }, fill: { fgColor: { rgb: 'EBF3FB' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' } },
  };
  const CREDIT = 'Developed by Arun Xavier, Asst. Prof., Dept. of EEE, Vidya Academy of Science & Technology, Thrissur. (Note: SGPA might be slightly different from the official KTU result, since subjects in specific schemes have different credits)';

  function sty(ws, r, c, style) { const ref = XL.utils.encode_cell({ r, c }); if (!ws[ref]) ws[ref] = { v: '', t: 's' }; ws[ref].s = style; }
  function styRange(ws, r1, c1, r2, c2, style) { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) sty(ws, r, c, style); }
  function autoW(aoa) { const nc = Math.max(...aoa.map(r => r.length), 0), w = Array(nc).fill(10); aoa.forEach(row => row.forEach((v, ci) => { const l = String(v ?? '').length; if (l > w[ci]) w[ci] = Math.min(l + 3, 45); })); return w.map(wch => ({ wch })); }
  function mkRows(h, n) { return Array(n).fill({ hpt: h }); }
  function passPct(p) { return p >= 75 ? S.pass : p < 50 ? S.fail : S.warn; }
  function sgpaS(sg, odd) { return sg >= 8 ? S.pass : sg < 5 ? S.fail : (odd ? S.odd : S.even); }
  function rowS(odd, text) { return odd ? (text ? S.oddL : S.odd) : (text ? S.evenL : S.even); }
  function addFooter(ws, aoa, nc) {
    const cr = aoa.length + 20;
    const ref = XL.utils.encode_cell({ r: cr, c: 0 });
    ws[ref] = { v: CREDIT, t: 's', s: S.footer };
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: cr, c: 0 }, e: { r: cr, c: nc - 1 } });
    const rng = XL.utils.decode_range(ws['!ref'] || 'A1'); if (cr > rng.e.r) rng.e.r = cr; ws['!ref'] = XL.utils.encode_range(rng);
  }

  (() => {
    const NC = 8, totS = sem.departments.reduce((a, d) => a + d.tot, 0), totP = sem.departments.reduce((a, d) => a + d.pass, 0), totF = totS - totP;
    const dat = sem.departments.map((d, i) => [i + 1, d.code, d.name || DEPT_MAP[d.code] || d.code, d.tot, d.pass, d.fail, parseFloat(d.pp), parseFloat(d.avgSGPA)]);
    const aoa = [
      [`KTU RESULT ANALYSER – ${(sem.meta?.college || "Unknown College").toUpperCase()}`],
      [`Overall Department Summary – ${sem.meta?.semester || "Semester Details"}`],
      [],
      ['Total Students', 'Passed', 'Failed', 'Pass %', 'Avg SGPA', '', '', ''],
      [totS, totP, totF, (sem.ov?.pp || '0') + '%', sem.departments.reduce((a, d) => a + parseFloat(d.avgSGPA), 0) / sem.departments.length || 0, '', '', ''],
      [],
      ['Rank', 'Dept Code', 'Department Name', 'Total Students', 'Passed', 'Failed', 'Pass %', 'Avg SGPA'],
      ...dat
    ];
    const ws = XL.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } },
    ];
    ws['!cols'] = autoW(aoa); ws['!rows'] = [{ hpt: 35 }, { hpt: 25 }, { hpt: 10 }, { hpt: 20 }, { hpt: 30 }, { hpt: 12 }, { hpt: 26 }, ...mkRows(22, dat.length)];
    ws['!freeze'] = { xSplit: 0, ySplit: 7, topLeftCell: 'A8', activePane: 'bottomLeft', state: 'frozen' };
    ws['!autofilter'] = { ref: XL.utils.encode_range({ s: { r: 6, c: 0 }, e: { r: 6 + dat.length, c: NC - 1 } }) };
    
    styRange(ws, 0, 0, 0, NC - 1, S.title); 
    styRange(ws, 1, 0, 1, NC - 1, S.sheetSub);
    styRange(ws, 3, 0, 3, NC - 1, S.kpiLbl);
    // Row 5 KPIs - Simplified to 1 column per metric for reliability
    sty(ws, 4, 0, S.kpiBlue);  // Total
    sty(ws, 4, 1, S.kpiGreen); // Pass
    sty(ws, 4, 2, S.kpiRed);   // Fail
    sty(ws, 4, 3, S.kpiGreen); // Pass%
    sty(ws, 4, 4, S.kpiBlue);  // SGPA
    
    styRange(ws, 6, 0, 6, NC - 1, S.hdr);
    dat.forEach((row, ri) => {
      const r = 7 + ri, odd = ri % 2 === 1;
      row.forEach((v, ci) => {
        let st;
        if (ci === 0 && ri === 0) st = S.gold;
        else if (ci === 6) st = passPct(parseFloat(v));
        else if (ci === 7) st = sgpaS(parseFloat(v), odd);
        else if (ci === 2) st = rowS(odd, true);
        else st = rowS(odd, false);
        sty(ws, r, ci, st);
      });
    });
    addFooter(ws, aoa, NC);
    XL.utils.book_append_sheet(wb, ws, 'Departments');
  })();

  (() => {
    const NC = 6;
    const subHdrs = ['Subject Code', 'Subject Name', 'Total Appeared', 'Passed', 'Failed', 'Pass %'];
    const totalSubs = sem.departments.reduce((a, d) => a + d.subjects.length, 0);
    const aoa = []; const merges = []; const rowH = []; const rowStyles = [];
    let cr = 0;
    aoa.push([`KTU RESULT ANALYSER – ${(sem.meta?.college || "Unknown College").toUpperCase()}`]); merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: NC - 1 } }); rowH.push({ hpt: 35 }); rowStyles.push('sheetTitle'); cr++;
    aoa.push([`Subject-wise Performance Analysis – ${sem.meta?.semester || "Semester Details"}`]); merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: NC - 1 } }); rowH.push({ hpt: 25 }); rowStyles.push('sheetSub'); cr++;
    aoa.push([]); rowH.push({ hpt: 10 }); rowStyles.push('gap'); cr++;
    aoa.push(['Total Subjects', '', 'Departments', '', 'Overall Status', '']); merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: 1 } }, { s: { r: cr, c: 2 }, e: { r: cr, c: 3 } }, { s: { r: cr, c: 4 }, e: { r: cr, c: NC - 1 } }); rowH.push({ hpt: 20 }); rowStyles.push('kpiLbl'); cr++;
    aoa.push([totalSubs, '', sem.departments.length, '', (sem.ov?.pp || '0') + '% Pass', '']); merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: 1 } }, { s: { r: cr, c: 2 }, e: { r: cr, c: 3 } }, { s: { r: cr, c: 4 }, e: { r: cr, c: NC - 1 } }); rowH.push({ hpt: 30 }); rowStyles.push('kpiVal'); cr++;
    sem.departments.forEach((d, di) => {
      aoa.push([]); rowH.push({ hpt: 12 }); rowStyles.push('gap'); cr++;
      const deptLabel = `${d.code}  —  ${d.name || DEPT_MAP[d.code] || d.code}   (${d.subjects.length} subjects, Pass%: ${d.pp}%)`;
      aoa.push([deptLabel]); merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: NC - 1 } }); rowH.push({ hpt: 24 }); rowStyles.push('deptHdr'); cr++;
      aoa.push(subHdrs); rowH.push({ hpt: 22 }); rowStyles.push('colHdr'); cr++;
      d.subjects.forEach((s, ri) => { aoa.push([s.code, s.name, s.total, s.passed, s.failed, parseFloat(s.pp)]); rowH.push({ hpt: 20 }); rowStyles.push(ri % 2 === 0 ? 'even' : 'odd'); cr++; });
    });
    const ws = XL.utils.aoa_to_sheet(aoa); ws['!merges'] = merges; ws['!cols'] = autoW(aoa); ws['!rows'] = rowH;
    rowStyles.forEach((type, r) => {
      if (type === 'sheetTitle') styRange(ws, r, 0, r, NC - 1, S.title);
      else if (type === 'sheetSub') styRange(ws, r, 0, r, NC - 1, S.sheetSub);
      else if (type === 'kpiLbl') styRange(ws, r, 0, r, NC - 1, S.kpiLbl);
      else if (type === 'kpiVal') {
        styRange(ws, r, 0, r, 1, S.kpiBlue); styRange(ws, r, 2, r, 3, S.kpiBlue); styRange(ws, r, 4, r, 5, S.kpiGreen);
      }
      else if (type === 'deptHdr') styRange(ws, r, 0, r, NC - 1, S.secGreen);
      else if (type === 'colHdr') { styRange(ws, r, 0, r, NC - 1, S.hdr); }
      else if (type === 'even' || type === 'odd') {
        const odd = type === 'odd';
        aoa[r].forEach((v, ci) => {
          let st;
          if (ci === 5) st = passPct(parseFloat(v));
          else if (ci <= 1) st = rowS(odd, true);
          else st = rowS(odd, false);
          sty(ws, r, ci, st);
        });
      }
    });
    addFooter(ws, aoa, NC); XL.utils.book_append_sheet(wb, ws, 'Subjects');
  })();

  (() => {
    const NC = 5; const dat = [];
    sem.departments.forEach(d => d.students.forEach(s => dat.push([s.regNo, s.department, d.name || DEPT_MAP[s.department] || s.department, parseFloat(s.sgpa), s.allPassed ? 'PASS' : 'FAIL'])));
    const totP = dat.filter(r => r[4] === 'PASS').length; const pct = dat.length ? ((totP / dat.length) * 100).toFixed(1) : 0;
    const aoa = [
      [`KTU RESULT ANALYSER – ${(sem.meta?.college || "Unknown College").toUpperCase()}`],
      [`Complete Student Results – ${sem.meta?.semester || "Semester Details"}`],
      [],
      ['Total Students', 'Passed', 'Failed', 'Pass %', 'Avg SGPA'],
      [dat.length, totP, dat.length - totP, pct + '%', sem.ov?.avgSGPA || '0'],
      [],
      ['Register No', 'Dept Code', 'Department Name', 'SGPA', 'Status'],
      ...dat
    ];
    const ws = XL.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } }
    ];
    ws['!cols'] = autoW(aoa); ws['!rows'] = [{ hpt: 35 }, { hpt: 25 }, { hpt: 10 }, { hpt: 20 }, { hpt: 30 }, { hpt: 12 }, { hpt: 26 }, ...mkRows(21, dat.length)];
    ws['!freeze'] = { xSplit: 0, ySplit: 7, topLeftCell: 'A8', activePane: 'bottomLeft', state: 'frozen' };
    ws['!autofilter'] = { ref: XL.utils.encode_range({ s: { r: 6, c: 0 }, e: { r: 6 + dat.length, c: NC - 1 } }) };
    styRange(ws, 0, 0, 0, NC - 1, S.title); 
    styRange(ws, 1, 0, 1, NC - 1, S.sheetSub); 
    styRange(ws, 3, 0, 3, NC - 1, S.kpiLbl);
    sty(ws, 4, 0, S.kpiBlue); sty(ws, 4, 1, S.kpiGreen); sty(ws, 4, 2, S.kpiRed); sty(ws, 4, 3, S.kpiGreen); sty(ws, 4, 4, S.kpiBlue);
    styRange(ws, 6, 0, 6, NC - 1, S.hdr);
    dat.forEach((row, ri) => {
      const r = 7 + ri, odd = ri % 2 === 1;
      row.forEach((v, ci) => {
        let st;
        if (ci === 4) st = v === 'PASS' ? S.pass : S.fail;
        else if (ci === 3) st = sgpaS(parseFloat(v), odd);
        else if (ci === 0 || ci === 2) st = rowS(odd, true);
        else st = rowS(odd, false);
        sty(ws, r, ci, st);
      });
    });
    addFooter(ws, aoa, NC); XL.utils.book_append_sheet(wb, ws, 'All Students');
  })();

  sem.departments.forEach(dept => {
    const sortedStudents = [...dept.students].sort((a, b) => {
      const numA = parseInt(a.regNo.slice(-3)) || 0;
      const numB = parseInt(b.regNo.slice(-3)) || 0;
      return numA - numB;
    });

    const nSubs = dept.subjects.length; const allCodes = [...new Set(dept.students.flatMap(s => s.subjects.map(x => x.code)))];
    const nStuCols = allCodes.length + 3; const maxC = Math.max(nStuCols, 8) - 1;
    const mainHeadR = 0, collHeadR = 1, kpiLblR = 3, kpiValR = 4;
    const subSecR = 6, subHdrR = 7, subDatR = 8, stuSecR = subDatR + nSubs + 1, stuHdrR = stuSecR + 1, stuDatR = stuHdrR + 1;
    const aoa = []; const merges = []; const rowH = [];
    aoa.push([`KTU RESULT ANALYSER – ${(sem.meta?.college || "Unknown College").toUpperCase()}`]); merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: maxC } }); rowH.push({ hpt: 35 });
    aoa.push([`${(dept.name || dept.code).toUpperCase()} – ${sem.meta?.semester || "Semester Details"}`]); merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: maxC } }); rowH.push({ hpt: 25 });
    aoa.push([]); rowH.push({ hpt: 10 });
    aoa.push(['Total Students', 'Passed', 'Failed', 'Pass %', 'Avg SGPA', '', '', '']); rowH.push({ hpt: 20 });
    aoa.push([dept.tot, dept.pass, dept.fail, dept.pp + '%', dept.avgSGPA, '', '', '']); rowH.push({ hpt: 30 });
    aoa.push([]); rowH.push({ hpt: 12 });
    aoa.push(['📚  SUBJECT-WISE ANALYSIS']); merges.push({ s: { r: subSecR, c: 0 }, e: { r: subSecR, c: 5 } }); rowH.push({ hpt: 24 });
    aoa.push(['Subject Code', 'Subject Name', 'Total Appeared', 'Passed', 'Failed', 'Pass %']); rowH.push({ hpt: 22 });
    dept.subjects.forEach(s => { aoa.push([s.code, s.name, s.total, s.passed, s.failed, parseFloat(s.pp)]); rowH.push({ hpt: 20 }); });
    aoa.push([]); rowH.push({ hpt: 10 });
    aoa.push(['🎓  STUDENT-WISE RESULTS  (with Subject Grades)']); merges.push({ s: { r: stuSecR, c: 0 }, e: { r: stuSecR, c: nStuCols - 1 } }); rowH.push({ hpt: 24 });
    aoa.push(['Register No', 'SGPA', 'Status', ...allCodes]); rowH.push({ hpt: 22 });
    sortedStudents.forEach(stu => {
      const gm = {}; stu.subjects.forEach(x => gm[x.code] = x.grade);
      aoa.push([stu.regNo, parseFloat(stu.sgpa), stu.allPassed ? 'PASS' : 'FAIL', ...allCodes.map(c => gm[c] || '—')]); rowH.push({ hpt: 20 });
    });
    const ws = XL.utils.aoa_to_sheet(aoa); ws['!merges'] = merges; ws['!rows'] = rowH; ws['!cols'] = autoW(aoa);
    ws['!freeze'] = { xSplit: 0, ySplit: stuHdrR + 1, topLeftCell: `A${stuHdrR + 2}`, activePane: 'bottomLeft', state: 'frozen' };
    ws['!autofilter'] = { ref: XL.utils.encode_range({ s: { r: stuHdrR, c: 0 }, e: { r: stuDatR + dept.students.length - 1, c: nStuCols - 1 } }) };
    styRange(ws, 0, 0, 0, maxC, S.title); 
    styRange(ws, 1, 0, 1, maxC, S.sheetSub);
    styRange(ws, 3, 0, 3, maxC, S.kpiLbl);
    sty(ws, 4, 0, S.kpiBlue); sty(ws, 4, 1, S.kpiGreen); sty(ws, 4, 2, S.kpiRed); sty(ws, 4, 3, S.kpiGreen); sty(ws, 4, 4, S.kpiBlue);
    styRange(ws, subSecR, 0, subSecR, 5, S.secGreen);
    styRange(ws, subHdrR, 0, subHdrR, 5, S.hdr); 
    styRange(ws, stuSecR, 0, stuSecR, nStuCols - 1, S.secBlue); 
    styRange(ws, stuHdrR, 0, stuHdrR, nStuCols - 1, S.hdr);
    dept.subjects.forEach((s, ri) => {
      const r = subDatR + ri, odd = ri % 2 === 1;
      [s.code, s.name, s.total, s.passed, s.failed, parseFloat(s.pp)].forEach((v, ci) => {
        let st; if (ci === 5) st = passPct(parseFloat(v)); else if (ci <= 1) st = rowS(odd, true); else st = rowS(odd, false);
        sty(ws, r, ci, st);
      });
    });
    dept.students.forEach((stu, ri) => {
      const r = stuDatR + ri, odd = ri % 2 === 1;
      const row = [stu.regNo, parseFloat(stu.sgpa), stu.allPassed ? 'PASS' : 'FAIL', ...allCodes.map(c => { const x = stu.subjects.find(s => s.code === c); return x ? x.grade : '—'; })];
      row.forEach((v, ci) => {
        let st; if (ci === 2) st = v === 'PASS' ? S.pass : S.fail; else if (ci === 1) st = sgpaS(parseFloat(v), odd); else if (ci === 0) st = rowS(odd, true);
        else { const g = String(v); if (['F', 'FE', 'Ab', 'Absent', 'Withheld', 'I'].includes(g)) st = S.fail; else if (g === 'S' || g === 'A+') st = S.pass; else st = rowS(odd, false); }
        sty(ws, r, ci, st);
      });
    });
    addFooter(ws, aoa, nStuCols); XL.utils.book_append_sheet(wb, ws, (dept.code || 'DEPT').slice(0, 31));
  });

  const b64 = XL.write(wb, { type: 'base64', bookType: 'xlsx' });
  const filename = sem.meta ? `KTU ${sem.meta.semester} Results ${sem.meta.college}.xlsx` : `KTU_Results_${semLabel.replace(/\s+/g, '_')}.xlsx`;
  return { filename, b64 };
}

/* ── API HANDLER ────────────────────────────────────────────────── */
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { fileData } = req.body;
    if (!fileData) return res.status(400).json({ error: 'Missing fileData in request' });

    // Process direct base64 data (bypassing storage download)
    const buffer = Buffer.from(fileData, 'base64');

    // Parse PDF Text
    const data = await pdf(buffer);
    const textContext = data.text;

    // Analyze text
    const rawSems = parseTxt(textContext);
    if (!rawSems || !rawSems[0]?.departments?.length) throw new Error('Failed to parse any valid semantic records from the PDF.');
    
    const analysis = analyze(rawSems);
    const sem = analysis[0];

    // Export to Excel Base64
    const { filename, b64 } = await exportExcelBase64(sem);

    return res.status(200).json({
      success: true,
      analysis: analysis,
      excelFilename: filename,
      excelBase64: b64
    });

  } catch (err) {
    console.error('API Error:', err);
    return res.status(500).json({ error: err.message || 'Internal Server Error' });
  }
}
