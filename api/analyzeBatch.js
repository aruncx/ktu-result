const XLSX = require('xlsx-js-style');
const nodeFetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));

const BRANCH_MAP = {
  'CS': 'Computer Science and Engineering',
  'EE': 'Electrical and Electronics Engineering',
  'EC': 'Electronics and Communication Engineering',
  'ME': 'Mechanical Engineering',
  'CE': 'Civil Engineering',
  'MR': 'Mechatronics Engineering',
  'IT': 'Information Technology',
  'AD': 'Artificial Intelligence and Data Science',
  'AM': 'Artificial Intelligence and Machine Learning',
  'RA': 'Robotics and Automation',
  'CH': 'Chemical Engineering',
  'BT': 'Biotechnology'
};

function normalizeGrade(g) {
  g = String(g).trim().toUpperCase(); if (g === '') return 'WH';
  if (['S', 'A+', 'A', 'B+', 'B', 'C+', 'C', 'D', 'P', 'F', 'FE', 'AB', 'WH', 'PASS', 'FAIL'].includes(g)) return g;
  return g;
}

function parseSemData(label, aoa) {
  if (!aoa.length) return null;
  let hdrIdx = aoa.findIndex(row => String(row[0]).toLowerCase().includes('student'));
  if (hdrIdx < 0) hdrIdx = 0;
  const hdr = aoa[hdrIdx].map(v => String(v).trim());
  const iEarned = hdr.findIndex(h => h.toLowerCase().includes('earned'));
  const iCumul = hdr.findIndex(h => h.toLowerCase().includes('cumil') || h.toLowerCase().includes('cumul'));
  const iSGPA = hdr.findIndex(h => h.toUpperCase() === 'SGPA');
  const iCGPA = hdr.findIndex(h => h.toUpperCase() === 'CGPA');
  const end = iEarned > 0 ? iEarned : hdr.length;
  
  const subjectIdxs = []; 
  for (let i = 1; i < end; i++) if (hdr[i]) subjectIdxs.push(i);
  
  const subjectCodes = subjectIdxs.map(i => hdr[i]);
  const studentRows = [];
  
  for (let ri = hdrIdx + 1; ri < aoa.length; ri++) {
    const row = aoa[ri];
    const raw = String(row[0] || '').trim(); if (!raw) continue;
    const dashIdx = raw.indexOf('-');
    let roll = dashIdx >= 0 ? raw.slice(0, dashIdx).trim() : raw;
    let name = dashIdx >= 0 ? raw.slice(dashIdx + 1).trim() : '';
    name = name.replace(/\s+/g, ' ');
    const isLateral = roll.toUpperCase().startsWith('L') && roll.length > 1;
    const regMatch = roll.match(/^L?([A-Z]{3})(\d{2})([A-Z]{2,3})(\d{3})$/);
    let admYear = '', branchCode = '', branchName = '', collegeCode = '';
    if (regMatch) {
      collegeCode = regMatch[1];
      admYear = '20' + regMatch[2];
      branchCode = regMatch[3];
      branchName = BRANCH_MAP[branchCode] || branchCode;
    }
    const dept = branchName || roll.replace(/^L?VAS\d{2}/, '').replace(/\d+$/, '') || '';
    const rawGrades = {};
    subjectIdxs.forEach((ci, si) => { rawGrades[subjectCodes[si]] = String(row[ci] || '').trim().toUpperCase(); });
    const emptyCount = Object.values(rawGrades).filter(g => g === '').length;
    const totalSubjects = subjectIdxs.length;
    const isResultWithheld = emptyCount === totalSubjects;
    const grades = {}; let failCount = 0;
    
    subjectIdxs.forEach((ci, si) => {
      const code = subjectCodes[si]; const rawVal = rawGrades[code];
      if (isResultWithheld) { grades[code] = { raw: 'WH', passed: false, skipped: false }; failCount++; }
      else if (rawVal === '' && emptyCount <= 3) { grades[code] = { raw: '', passed: true, skipped: true }; }
      else {
        const gNorm = normalizeGrade(rawVal);
        const passed = !['F', 'FE', 'AB', 'WH'].includes(gNorm);
        if (!passed) failCount++;
        grades[code] = { raw: gNorm, passed, skipped: false };
      }
    });
    const earnedCredits = parseFloat(row[iEarned]) || 0;
    const cumulCredits = parseFloat(row[iCumul]) || 0;
    const sgpa = parseFloat(row[iSGPA]) || 0;
    const cgpa = parseFloat(row[iCGPA]) || 0;
    const semPassed = failCount === 0;
    studentRows.push({ roll, name, dept, admYear, branchCode, branchName, isLateral, grades, failCount, semPassed, earnedCredits, cumulCredits, sgpa, cgpa, subjectCodes, isResultWithheld });
  }
  
  const subStats = {};
  subjectCodes.forEach(code => {
    let fails = 0, passes = 0, abCount = 0, whCount = 0, feCount = 0, total = 0;
    studentRows.forEach(s => {
      const g = s.grades[code]; if (!g || g.skipped) return;
      total++;
      if (g.raw === 'F') fails++;
      else if (g.raw === 'FE') feCount++;
      else if (g.raw === 'AB') abCount++;
      else if (g.raw === 'WH') whCount++;
      else passes++;
    });
    const totalFail = fails + feCount + abCount + whCount;
    subStats[code] = { fails, passes, abCount, whCount, feCount, total, failPct: total ? (totalFail / total * 100).toFixed(1) : 0 };
  });
  
  const totalStu = studentRows.length;
  const passAll = studentRows.filter(s => s.semPassed).length;
  return { label, subjectCodes, subStats, studentRows, totalStu, passAll, failedSome: totalStu - passAll, totalFailures: studentRows.reduce((a, s) => a + s.failCount, 0), passRate: totalStu ? (passAll / totalStu * 100).toFixed(1) : 0 };
}

function processBatchCore(semDataArr) {
  let semData = semDataArr.filter(Boolean);
  let maxRegYear = 0;
  semData.forEach(sem => {
    sem.studentRows.forEach(sr => {
      if (!sr.isLateral && sr.admYear) {
        const yr = parseInt(sr.admYear);
        if (!isNaN(yr) && yr > maxRegYear) maxRegYear = yr;
      }
    });
  });
  if (maxRegYear === 0) {
    semData.forEach(sem => {
      sem.studentRows.forEach(sr => {
        if (sr.admYear) {
          const yr = parseInt(sr.admYear);
          if (!isNaN(yr) && yr > maxRegYear) maxRegYear = yr;
        }
      });
    });
  }

  semData.forEach(sem => {
    sem.studentRows = sem.studentRows.filter(sr => {
      if (!sr.admYear) return true;
      const yr = parseInt(sr.admYear);
      if (sr.isLateral && maxRegYear > 0) return yr === maxRegYear || yr === maxRegYear + 1;
      return yr === maxRegYear;
    });

    Object.keys(sem.subStats).forEach(code => {
      let fails = 0, passes = 0, abCount = 0, whCount = 0, feCount = 0, total = 0;
      sem.studentRows.forEach(s => {
        const g = s.grades[code]; if (!g || g.skipped) return;
        total++;
        if (g.raw === 'F') fails++;
        else if (g.raw === 'FE') feCount++;
        else if (g.raw === 'AB') abCount++;
        else if (g.raw === 'WH') whCount++;
        else passes++;
      });
      const totalFail = fails + feCount + abCount + whCount;
      sem.subStats[code] = { fails, passes, abCount, whCount, feCount, total, failPct: total ? (totalFail / total * 100).toFixed(1) : 0 };
    });
    sem.totalStu = sem.studentRows.length;
    sem.passAll = sem.studentRows.filter(s => s.semPassed).length;
    sem.failedSome = sem.totalStu - sem.passAll;
    sem.totalFailures = sem.studentRows.reduce((a, s) => a + s.failCount, 0);
    sem.passRate = sem.totalStu ? (sem.passAll / sem.totalStu * 100).toFixed(1) : 0;
  });

  const students = {};
  semData.forEach(sem => {
    sem.studentRows.forEach(sr => {
      const key = sr.roll || sr.name;
      if (!students[key]) students[key] = { roll: sr.roll, name: sr.name, dept: sr.dept, admYear: sr.admYear, branchCode: sr.branchCode, branchName: sr.branchName, isLateral: sr.isLateral, sems: {} };
      students[key].sems[sem.label] = { grades: sr.grades, failCount: sr.failCount, semPassed: sr.semPassed, earnedCredits: sr.earnedCredits, cumulCredits: sr.cumulCredits, sgpa: sr.sgpa, cgpa: sr.cgpa, subjectCodes: sr.subjectCodes };
    });
  });

  return { semData, students, branchName: semData[0]?.studentRows[0]?.branchName || 'Unknown', maxRegYear };
}

async function exportBatchExcel(students, semData, branchName, maxRegYear) {
  const XL = XLSX;
  const wb = XL.utils.book_new();
  const allStu = Object.values(students);
  
  const bdr = (s = 'thin', c = 'CCCCCC') => ({ top: { style: s, color: { rgb: c } }, bottom: { style: s, color: { rgb: c } }, left: { style: s, color: { rgb: c } }, right: { style: s, color: { rgb: c } } });
  const S = {
    title: { font: { name: 'Calibri', bold: true, sz: 16, color: { rgb: '1A3A6E' } }, fill: { fgColor: { rgb: 'D6E4F7' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'A0B8D8') },
    info: { font: { name: 'Calibri', sz: 10, italic: true, color: { rgb: '3A4D6B' } }, fill: { fgColor: { rgb: 'EBF3FB' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'C0D4EE') },
    hdr: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1B4F9B' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr('medium', '1B4F9B') },
    hdrL: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1B4F9B' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: bdr('medium', '1B4F9B') },
    secG: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '2D6A4F' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bdr('thin', '2D6A4F') },
    even: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr() },
    odd: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'E8F0FE' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr() },
    evenL: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bdr() },
    oddL: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'E8F0FE' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bdr() },
    pass: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '155724' } }, fill: { fgColor: { rgb: 'D4EDDA' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'C3E6CB') },
    fail: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '7B1D1D' } }, fill: { fgColor: { rgb: 'F8D7DA' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'F5C6CB') }
  };
  function styRange(ws, r1, c1, r2, c2, style) { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const ref = XL.utils.encode_cell({ r, c }); if (!ws[ref]) ws[ref] = { v: '', t: 's' }; ws[ref].s = style; } }
  function autoW(aoa) { const nc = Math.max(...aoa.map(r => r.length), 0); const w = Array(nc).fill(10); aoa.forEach(row => row.forEach((v, ci) => { const l = String(v ?? '').length; if (l > w[ci]) w[ci] = Math.min(l + 3, 50); })); return w.map(wch => ({ wch })); }
  function mkRows(h, n) { return Array(n).fill({ hpt: h }); }

  const getTotalArrears = s => Object.values(s.sems).reduce((a, sm) => a + sm.failCount, 0);
  const lastSemCGPA = s => { for (let i = semData.length - 1; i >= 0; i--) { const sm = s.sems[semData[i].label]; if (sm && sm.cgpa > 0) return sm.cgpa; } return 0; };
  allStu.forEach(s => { s.totalArrears = getTotalArrears(s); s.latestCgpa = lastSemCGPA(s); });
  const sortedStu = [...allStu].sort((a, b) => b.latestCgpa - a.latestCgpa);

  // SHEET 1: Master Summary
  (() => {
    let rs = [];
    rs.push([`APJ AKTU Batch Analysis — ${branchName}`]);
    rs.push([`Generated: ${new Date().toLocaleDateString('en-IN')}`]);
    rs.push([]);
    const hdrRow = ['Rank', 'Reg. No', 'Name', 'Total Arrears', 'Latest CGPA', ...semData.map(s => s.label + ' SGPA'), ...semData.map(s => s.label + ' Arrears')];
    rs.push(hdrRow);
    sortedStu.forEach((s, idx) => {
      const row = [idx + 1, s.roll, s.name, s.totalArrears, s.latestCgpa];
      semData.forEach(sem => { const sm = s.sems[sem.label]; row.push(sm ? sm.sgpa || 0 : '-'); });
      semData.forEach(sem => { const sm = s.sems[sem.label]; row.push(sm ? sm.failCount : '-'); });
      rs.push(row);
    });
    const ws = XL.utils.aoa_to_sheet(rs);
    const NC = hdrRow.length;
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: NC - 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: NC - 1 } }];
    ws['!rows'] = [{ hpt: 30 }, { hpt: 18 }, { hpt: 8 }, { hpt: 24 }, ...mkRows(21, sortedStu.length)];
    ws['!cols'] = autoW(rs);
    ws['!freeze'] = { ySplit: 4, topLeftCell: 'A5', activePane: 'bottomLeft', state: 'frozen' };
    styRange(ws, 0, 0, 0, NC - 1, S.title); styRange(ws, 1, 0, 1, NC - 1, S.info); styRange(ws, 3, 0, 3, NC - 1, S.hdr);
    for (let r = 4; r < rs.length; r++) { const odd = r % 2 === 1; for (let c = 0; c < NC; c++) { let st = c <= 2 ? (odd ? S.oddL : S.evenL) : (odd ? S.odd : S.even); styRange(ws, r, c, r, c, st); } }
    XL.utils.book_append_sheet(wb, ws, 'Master Summary');
  })();

  // Return base64 string
  const b64 = XL.write(wb, { type: 'base64', bookType: 'xlsx' });
  const filename = `APJ_AKTU_Results_${maxRegYear}_${branchName.replace(/[^A-Za-z0-9]/g, '')}_Analysis.xlsx`;

  return { filename, b64 };
}

/* ── API HANDLER ────────────────────────────────────────────────── */
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { files } = req.body; 
    if (!files || !Array.isArray(files) || files.length === 0) return res.status(400).json({ error: 'Missing or empty files array in request' });

    const semDataArr = [];
    const fetchFn = typeof fetch !== 'undefined' ? fetch : nodeFetch;

    for (let f of files) {
      if (!f.url) continue;
      const fileRes = await fetchFn(f.url);
      if (!fileRes.ok) throw new Error(`Failed to download ${f.label}`);
      
      const arrayBuffer = await fileRes.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);
      
      const wb = XLSX.read(buffer, { type: 'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      
      const semParse = parseSemData(f.label, aoa);
      if (semParse) semDataArr.push(semParse);
    }

    if (semDataArr.length === 0) throw new Error('Could not parse any valid semester data.');

    // Process Batch Core Logic
    const { semData, students, branchName, maxRegYear } = processBatchCore(semDataArr);

    // Export Excel
    const { filename, b64 } = await exportBatchExcel(students, semData, branchName, maxRegYear);

    return res.status(200).json({
      success: true,
      analysis: { semData, students },
      excelFilename: filename,
      excelBase64: b64
    });

  } catch (err) {
    console.error('API Error:', err);
    return res.status(500).json({ error: err.message || 'Internal Server Error' });
  }
}
