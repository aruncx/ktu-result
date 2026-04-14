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
  const CREDIT = 'Developed by Arun Xavier, Asst. Prof., Dept. of EEE, Vidya Academy of Science & Technology, Thrissur';

  // State helpers
  const lastSemCGPA = s => { for (let i = semData.length - 1; i >= 0; i--) { const sm = s.sems[semData[i].label]; if (sm && sm.cgpa > 0) return sm.cgpa; } return 0; };
  const latestCGPA = lastSemCGPA;

  /* --- Shared style helpers --- */
  const bdr = (s = 'thin', c = 'CCCCCC') => ({ top: { style: s, color: { rgb: c } }, bottom: { style: s, color: { rgb: c } }, left: { style: s, color: { rgb: c } }, right: { style: s, color: { rgb: c } } });
  const S = {
    title: { font: { name: 'Calibri', bold: true, sz: 16, color: { rgb: '1A3A6E' } }, fill: { fgColor: { rgb: 'D6E4F7' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'A0B8D8') },
    info: { font: { name: 'Calibri', sz: 10, italic: true, color: { rgb: '3A4D6B' } }, fill: { fgColor: { rgb: 'EBF3FB' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'C0D4EE') },
    hdr: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1B4F9B' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr('medium', '1B4F9B') },
    hdrR: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '7B1D1D' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: bdr('medium', '7B1D1D') },
    hdrL: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1B4F9B' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: bdr('medium', '1B4F9B') },
    secG: { font: { name: 'Calibri', bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '2D6A4F' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bdr('thin', '2D6A4F') },
    even: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr() },
    odd: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'E8F0FE' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr() },
    evenL: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'FFFFFF' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bdr() },
    oddL: { font: { name: 'Calibri', sz: 11 }, fill: { fgColor: { rgb: 'E8F0FE' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' }, border: bdr() },
    pass: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '155724' } }, fill: { fgColor: { rgb: 'D4EDDA' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'C3E6CB') },
    fail: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '7B1D1D' } }, fill: { fgColor: { rgb: 'F8D7DA' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'F5C6CB') },
    warn: { font: { name: 'Calibri', sz: 11, color: { rgb: '856404' } }, fill: { fgColor: { rgb: 'FFF3CD' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'FFEEBA') },
    gold: { font: { name: 'Calibri', bold: true, sz: 11, color: { rgb: '7B5200' } }, fill: { fgColor: { rgb: 'FFF3CD' }, patternType: 'solid' }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr('thin', 'FFEEBA') },
    footer: { font: { name: 'Calibri', italic: true, sz: 10, color: { rgb: '1A3A6E' } }, fill: { fgColor: { rgb: 'EBF3FB' }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' } },
  };
  const sty = (ws, r, c, st) => { const ref = XL.utils.encode_cell({ r, c }); if (!ws[ref]) ws[ref] = { v: '', t: 's' }; ws[ref].s = st; };
  const styR = (ws, r1, c1, r2, c2, st) => { for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) sty(ws, r, c, st); };
  const autoW = (aoa) => {
    const nc = Math.max(...aoa.map(r => r.length), 0), w = Array(nc).fill(10);
    aoa.forEach(row => row.forEach((v, ci) => { const l = String(v ?? '').length; if (l > w[ci]) w[ci] = Math.min(l + 3, 45); }));
    return w.map(wch => ({ wch }));
  };
  const mkRows = (h, n) => Array(n).fill({ hpt: h });
  const passPct = (p) => p >= 75 ? S.pass : p < 50 ? S.fail : S.warn;
  const rowS = (odd, txt) => odd ? (txt ? S.oddL : S.odd) : (txt ? S.evenL : S.even);
  const addFooter = (ws, aoa, nc) => {
    const cr = aoa.length + 20;
    const ref = XL.utils.encode_cell({ r: cr, c: 0 });
    ws[ref] = { v: CREDIT, t: 's', s: S.footer };
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: cr, c: 0 }, e: { r: cr, c: nc - 1 } });
    const rng = XL.utils.decode_range(ws['!ref'] || 'A1'); if (cr > rng.e.r) rng.e.r = cr; ws['!ref'] = XL.utils.encode_range(rng);
  };

  const batchLabel = (() => {
    const yrs = allStu.map(s => parseInt(s.admYear)).filter(y => !isNaN(y));
    const yr = yrs.length ? Math.max(...yrs) : '';
    const br = allStu.map(s => s.branchName).filter(Boolean)[0] || '';
    return yr ? yr + ' Admission Batch - ' + br : 'Batch Analysis';
  })();

  /* --- SHEET 1: DASHBOARD --- */
  (() => {
    const lastSem = semData[semData.length - 1];
    const allKeys = Object.keys(students);
    const allSemLabels = semData.map(s => s.label);
    const passedAll = allKeys.filter(k => allSemLabels.every(lbl => students[k].sems[lbl]?.semPassed));
    const totalFails = semData.reduce((a, s) => a + (s.totalFailures || 0), 0);
    const allClear = allStu.filter(s => Object.values(s.sems).every(sm => sm.semPassed)).length;
    const avgCGPA = allStu.length ? (allStu.reduce((a, s) => a + latestCGPA(s), 0) / allStu.length).toFixed(2) : '0.00';
    const ranked = [...allStu].sort((a, b) => lastSemCGPA(b) - lastSemCGPA(a));
    const topCGPA = ranked[0] ? lastSemCGPA(ranked[0]).toFixed(2) : '-';
    const totalSubjects = semData.reduce((a, s) => a + (s.subjectCodes ? s.subjectCodes.length : 0), 0);
    const nc = 7;

    const sumHdr = ['Semester', 'Students', 'All Clear', 'Has Arrears', 'F Grades', 'Pass Rate %'];
    const sumDat = semData.map(s => [s.label, s.totalStu, s.passAll, s.failedSome, s.totalFailures, parseFloat(s.passRate)]);

    const ppHdr = ['Semester', 'Students (Sem)', 'Full Pass (Sem)', 'Pass % (Sem)', 'Upto Semester', 'Full Pass (Cumul)', 'Pass % (Cumul)'];
    const ppDat = semData.map((sem, si) => {
      const labelsUpTo = semData.slice(0, si + 1).map(s => s.label);
      const cumulPass = Object.keys(students).filter(k => labelsUpTo.every(lbl => students[k].sems[lbl]?.semPassed)).length;
      const cumulPct = sem.totalStu ? (cumulPass / sem.totalStu * 100).toFixed(2) : '0.00';
      return [sem.label, sem.totalStu, sem.passAll, parseFloat(sem.passRate), sem.label, cumulPass, parseFloat(cumulPct)];
    });

    const topHdr = ['Rank', 'Register No.', 'Name', 'Dept', 'Latest SGPA', 'Latest CGPA', 'Status'];
    const topDat = ranked.slice(0, 10).map((s, i) => {
      let sgpa = 0;
      for (let j = semData.length - 1; j >= 0; j--) { const sm = s.sems[semData[j].label]; if (sm && sm.sgpa > 0) { sgpa = sm.sgpa; break; } }
      const totF = Object.values(s.sems).reduce((a, sm) => a + sm.failCount, 0);
      return [i + 1, s.roll, s.name, s.branchName || s.dept || '', sgpa, lastSemCGPA(s).toFixed(2), totF === 0 ? 'ALL CLEAR' : 'HAS ARREARS'];
    });

    const getTotArr = s => Object.values(s.sems).reduce((a, sm) => a + sm.failCount, 0);
    const csHdr = ['Category', 'No. of Students'];
    const csDat = [
      ['Total No. of Students', allStu.length],
      ['Full Pass (No Arrears)', allStu.filter(s => getTotArr(s) === 0).length],
      ['1 Arrear', allStu.filter(s => getTotArr(s) === 1).length],
      ['2 Arrears', allStu.filter(s => getTotArr(s) === 2).length],
      ['3 Arrears', allStu.filter(s => getTotArr(s) === 3).length],
      ['4 Arrears', allStu.filter(s => getTotArr(s) === 4).length],
      ['5 Arrears', allStu.filter(s => getTotArr(s) === 5).length],
      ['6 Arrears', allStu.filter(s => getTotArr(s) === 6).length],
      ['More than 6 Arrears', allStu.filter(s => getTotArr(s) > 6).length],
    ];

    const aoa = []; const merges = []; const rowH = []; const rowMeta = []; let cr = 0;
    const push = (row, hpt, meta) => { aoa.push(row); rowH.push({ hpt }); rowMeta.push(meta || 'data'); cr++; };
    const pushMerge = (row, hpt, nc2, meta) => { merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: nc2 - 1 } }); push(row, hpt, meta); };

    pushMerge([batchLabel + ' - Dashboard Summary'], 32, nc, 'sheetTitle');
    pushMerge(['Semesters: ' + semData.length + ' | Students: ' + allStu.length + ' | Avg CGPA: ' + avgCGPA + ' | High: ' + topCGPA + ' | Generated: ' + new Date().toLocaleDateString('en-IN')], 18, nc, 'info');
    push([], 10, 'gap');

    push(['KPI', 'Total Students', 'Total Subjects', 'Cumul Pass %', 'Total F Grades', 'All Sems Clear'], 22, 'kpiHdr');
    push(['Values', allStu.length, totalSubjects, parseFloat(passedAll.length / Math.max(lastSem?.totalStu || 1, 1) * 100).toFixed(1) + '%', totalFails, allClear], 20, 'kpiRow');
    push([], 10, 'gap');

    pushMerge(['# SEMESTER SUMMARY'], 24, nc, 'secHdr');
    push(sumHdr, 22, 'hdr');
    sumDat.forEach((row, ri) => { push(row, 20, ri % 2 === 0 ? 'even' : 'odd'); });
    push([], 10, 'gap');

    pushMerge(['# PASS PERCENTAGE REPORT'], 24, nc, 'secHdr');
    push(ppHdr, 22, 'hdr');
    ppDat.forEach((row, ri) => { push(row, 20, ri % 2 === 0 ? 'even' : 'odd'); });
    push([], 10, 'gap');

    pushMerge(['# CGPA TOPPERS (Top 10)'], 24, nc, 'secHdr');
    push(topHdr, 22, 'hdr');
    topDat.forEach((row, ri) => { push(row, 20, ri % 2 === 0 ? 'even' : 'odd'); });
    push([], 10, 'gap');

    pushMerge(['# CUMULATIVE STATUS'], 24, nc, 'secHdr');
    push(csHdr, 22, 'hdr');
    csDat.forEach((row, ri) => { push(row, 20, ri % 2 === 0 ? 'even' : 'odd'); });

    const ws = XL.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges; ws['!rows'] = rowH; ws['!cols'] = autoW(aoa);
    ws['!freeze'] = { xSplit: 0, ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' };

    rowMeta.forEach((meta, r) => {
      if (meta === 'sheetTitle') styR(ws, r, 0, r, nc - 1, S.title);
      else if (meta === 'info') styR(ws, r, 0, r, nc - 1, S.info);
      else if (meta === 'secHdr') styR(ws, r, 0, r, nc - 1, S.secG);
      else if (meta === 'kpiHdr') styR(ws, r, 0, r, 5, S.hdr);
      else if (meta === 'kpiRow') styR(ws, r, 0, r, 5, S.gold);
      else if (meta === 'hdr') styR(ws, r, 0, r, nc - 1, S.hdr);
      else if (meta === 'even' || meta === 'odd') {
        const odd = (meta === 'odd');
        aoa[r].forEach((v, ci) => {
          let st; const vStr = String(v);
          if (vStr === 'ALL CLEAR') st = S.pass;
          else if (vStr === 'HAS ARREARS') st = S.fail;
          else if (typeof v === 'number' && ci >= 3 && ci <= 4) st = passPct(parseFloat(v));
          else if (ci === 0) st = rowS(odd, true);
          else st = rowS(odd, false);
          sty(ws, r, ci, st);
        });
      }
    });
    addFooter(ws, aoa, nc); XL.utils.book_append_sheet(wb, ws, 'Dashboard');
  })();

  /* --- SHEET: ACADEMIC SUMMARY --- */
  (() => {
    const semLabels = semData.map(s => s.label);
    const nc = 1 + (semLabels.length * 2) + 2;
    const aoa = []; const merges = []; const rowH = []; const rowMeta = []; let cr = 0;
    const push = (row, hpt, meta) => { aoa.push(row); rowH.push({ hpt }); rowMeta.push(meta || 'data'); cr++; };
    const pushM = (row, hpt, nc2, meta) => { merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: nc2 - 1 } }); push(row, hpt, meta); };

    pushM([batchLabel + ' - Academic Summary'], 32, nc, 'sheetTitle');
    pushM(['Overview of academic performance across all uploaded semesters'], 18, nc, 'info');
    push([], 10, 'gap');

    const hdr = ['REGISTER NUMBER']; semLabels.forEach(lbl => { hdr.push(lbl + ' SGPA'); hdr.push(lbl + ' BACKLOG'); });
    hdr.push('CGPA'); hdr.push('TOTAL BACKLOGS'); push(hdr, 22, 'hdr');

    [...allStu].sort((a, b) => {
      const n1 = parseInt(String(a.roll).slice(-3)) || 0;
      const n2 = parseInt(String(b.roll).slice(-3)) || 0;
      return n1 !== n2 ? n1 - n2 : String(a.roll).localeCompare(String(b.roll));
    }).forEach((s, ri) => {
      const row = [s.roll]; let totB = 0;
      semLabels.forEach(lbl => { const sm = s.sems[lbl]; if (sm) { row.push(sm.sgpa || 0); row.push(sm.failCount || 0); totB += (sm.failCount || 0); } else { row.push(''); row.push(''); } });
      row.push(parseFloat(lastSemCGPA(s).toFixed(2))); row.push(totB); push(row, 20, ri % 2 === 0 ? 'even' : 'odd');
    });

    const ws = XL.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges; ws['!rows'] = rowH; ws['!cols'] = autoW(aoa);
    ws['!freeze'] = { xSplit: 1, ySplit: 3, topLeftCell: 'B4', activePane: 'bottomRight', state: 'frozen' };

    rowMeta.forEach((meta, r) => {
      if (meta === 'sheetTitle') styR(ws, r, 0, r, nc - 1, S.title);
      else if (meta === 'info') styR(ws, r, 0, r, nc - 1, S.info);
      else if (meta === 'hdr') styR(ws, r, 0, r, nc - 1, S.hdr);
      else if (meta === 'even' || meta === 'odd') {
        const odd = (meta === 'odd');
        aoa[r].forEach((v, ci) => {
          let st;
          if (ci === 0) st = rowS(odd, true);
          else if (ci === nc - 1) st = (v > 0) ? S.fail : S.pass;
          else if (ci === nc - 2) { const cg = parseFloat(v); st = cg >= 8 ? S.pass : cg >= 6 ? rowS(odd, false) : cg >= 5 ? S.warn : S.fail; }
          else if (ci > 0 && ci < nc - 2) { if (ci % 2 === 0) st = (v > 0) ? S.fail : rowS(odd, false); else st = rowS(odd, false); }
          else st = rowS(odd, false);
          sty(ws, r, ci, st);
        });
      }
    });
    addFooter(ws, aoa, nc); XL.utils.book_append_sheet(wb, ws, 'Academic Summary');
  })();

  /* --- SHEETS: PER SEMESTER --- */
  semData.forEach(sem => {
    const subCodes = sem.subjectCodes || [];
    const nc = Math.max(subCodes.length + 8, 10);
    const aoa = []; const merges = []; const rowH = []; const rowMeta = []; let cr = 0;
    const push = (row, hpt, meta) => { aoa.push(row); rowH.push({ hpt }); rowMeta.push(meta || 'data'); cr++; };
    const pushM = (row, hpt, nc2, meta) => { merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: nc2 - 1 } }); push(row, hpt, meta); };

    pushM([batchLabel + ' - ' + sem.label + ' Analysis'], 32, nc, 'sheetTitle');
    pushM(['Students: ' + sem.totalStu + ' | All Clear: ' + sem.passAll + ' | Pass: ' + sem.passRate + '% | Subjects: ' + subCodes.length], 18, nc, 'info');
    push([], 10, 'gap');

    push(['All Clear', 'Has Arrears', 'Pass Rate %', 'Subjects', 'F Grades'], 22, 'kpiHdr');
    push([sem.passAll, sem.failedSome, parseFloat(sem.passRate), subCodes.length, sem.totalFailures], 20, 'kpiRow');
    push([], 10, 'gap');

    pushM(['# SUBJECT-WISE ANALYSIS'], 24, 8, 'secHdr');
    push(['Subject Code', 'Total', 'Passed', 'F', 'FE', 'AB', 'WH', 'Pass %'], 22, 'subHdr');
    subCodes.forEach((code, ri) => {
      const st = sem.subStats[code] || { total: 0, passes: 0, fails: 0, abCount: 0, whCount: 0 }; 
      const pp = st.total ? (st.passes / st.total * 100).toFixed(1) : '0.0';
      push([code, st.total, st.passes, st.fails, st.feCount || 0, st.abCount, st.whCount || 0, parseFloat(pp)], 20, ri % 2 === 0 ? 'subEven' : 'subOdd');
    });
    push([], 10, 'gap');

    pushM(['# TOP 5 BY SGPA'], 24, 6, 'secHdr');
    push(['Rank', 'Reg. No.', 'Name', 'SGPA', 'CGPA', 'Arrears'], 22, 'hdr');
    [...sem.studentRows].sort((a, b) => b.sgpa - a.sgpa).slice(0, 5).forEach((sr, i) => { push([i + 1, sr.roll, sr.name, sr.sgpa, sr.cgpa, sr.failCount], 20, i % 2 === 0 ? 'even' : 'odd'); });
    push([], 10, 'gap');

    pushM(['# NEEDS ATTENTION (Bottom 5 BY SGPA)'], 24, 6, 'secHdrR');
    push(['Rank', 'Reg. No.', 'Name', 'SGPA', 'CGPA', 'Arrears'], 22, 'hdrR');
    [...sem.studentRows].filter(s => s.sgpa > 0).sort((a, b) => a.sgpa - b.sgpa).slice(0, 5).forEach((sr, i) => { push([i + 1, sr.roll, sr.name, sr.sgpa, sr.cgpa, sr.failCount], 20, i % 2 === 0 ? 'attnEven' : 'attnOdd'); });

    pushM(['# PASS/FAIL SUMMARY'], 24, 2, 'secHdr');
    push(['PASS/FAIL', 'NUMBER'], 22, 'hdr');
    const srs = sem.studentRows; const fP = srs.filter(s => s.semPassed && !s.isResultWithheld).length; const wH = srs.filter(s => s.isResultWithheld).length;
    const fC = [1, 2, 3, 4, 5, 6].map(x => srs.filter(s => !s.isResultWithheld && s.failCount === x).length);
    const sRows = [['No. of full Pass', fP], ['No: of students whose results are withheld', wH], ['No. of students failed in 1 subject', fC[0]], ['No. of students failed in 2 subjects', fC[1]], ['No. of students failed in 3 subjects', fC[2]], ['No. of students failed in 4 subjects', fC[3]], ['No. of students failed in 5 Subjects', fC[4]], ['No. of students failed in 6 subjects', fC[5]]];
    sRows.forEach((row, ri) => { push(row, 20, ri % 2 === 0 ? 'sumEven' : 'sumOdd'); });
    push([], 10, 'gap');

    pushM(['# ALL STUDENTS'], 24, nc, 'secHdr');
    push(['#', 'Reg. No.', 'Name', ...subCodes, 'Earned Cr.', 'Cumul Cr.', 'SGPA', 'CGPA', 'Arrears', 'Status'], 22, 'hdr');
    [...sem.studentRows].sort((a, b) => {
      const n1 = parseInt(String(a.roll).slice(-3)) || 0;
      const n2 = parseInt(String(b.roll).slice(-3)) || 0;
      return n1 !== n2 ? n1 - n2 : String(a.roll).localeCompare(String(b.roll));
    }).forEach((sr, ri) => {
      const row = [ri + 1, sr.roll, sr.name, ...subCodes.map(code => sr.grades[code]?.skipped ? '—' : (sr.grades[code]?.raw || '')), sr.earnedCredits, sr.cumulCredits, sr.sgpa, sr.cgpa, sr.failCount, sr.semPassed ? 'CLEAR' : 'ARREAR'];
      push(row, 20, ri % 2 === 0 ? 'even' : 'odd');
    });

    const ws = XL.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges; ws['!rows'] = rowH; ws['!cols'] = autoW(aoa);
    ws['!freeze'] = { xSplit: 3, ySplit: 2, topLeftCell: 'D3', activePane: 'bottomRight', state: 'frozen' };

    rowMeta.forEach((meta, r) => {
      if (meta === 'sheetTitle') styR(ws, r, 0, r, nc - 1, S.title);
      else if (meta === 'info') styR(ws, r, 0, r, nc - 1, S.info);
      else if (meta === 'secHdr') styR(ws, r, 0, r, nc - 1, S.secG);
      else if (meta === 'secHdrR') styR(ws, r, 0, r, nc - 1, S.hdrR); 
      else if (meta === 'kpiHdr') styR(ws, r, 0, r, 4, S.hdr);
      else if (meta === 'kpiRow') aoa[r].forEach((v, ci) => { if(ci < 5) sty(ws, r, ci, passPct(typeof v === 'number' ? v : 100)) });
      else if (meta === 'subHdr') styR(ws, r, 0, r, 7, S.hdr);
      else if (meta === 'hdr' || meta === 'hdrR') styR(ws, r, 0, r, nc - 1, S.hdr);
      else if (meta === 'sumEven' || meta === 'sumOdd') { const odd = (meta === 'sumOdd'); sty(ws, r, 0, rowS(odd, true)); sty(ws, r, 1, rowS(odd, false)); }
      else if (meta === 'subEven' || meta === 'subOdd') { const odd = (meta === 'subOdd'); aoa[r].forEach((v, ci) => sty(ws, r, ci, ci === 7 ? passPct(parseFloat(v)) : (ci === 0 ? rowS(odd, true) : rowS(odd, false)))); }
      else if (meta === 'even' || meta === 'odd') { const odd = (meta === 'odd'); const lC = aoa[r].length - 1; aoa[r].forEach((v, ci) => { let st; if (ci === lC) st = String(v) === 'CLEAR' ? S.pass : S.fail; else if (ci === 1 || ci === 2) st = rowS(odd, true); else st = rowS(odd, false); sty(ws, r, ci, st); }); }
      else if (meta === 'attnEven' || meta === 'attnOdd') { const odd = (meta === 'attnOdd'); aoa[r].forEach((v, ci) => sty(ws, r, ci, ci === 2 ? rowS(odd, true) : rowS(odd, false))); }
    });
    addFooter(ws, aoa, nc); XL.utils.book_append_sheet(wb, ws, sem.label);
  });

  /* --- SHEET: RANKING --- */
  (() => {
    const nc = 9;
    const ranked = [...allStu].map(s => {
      const sems = Object.values(s.sems); const cgpa = lastSemCGPA(s); let sgpa = 0;
      for (let i = semData.length - 1; i >= 0; i--) { const sm = s.sems[semData[i].label]; if (sm && sm.sgpa > 0) { sgpa = sm.sgpa; break; } }
      return { roll: s.roll, name: s.name, dept: s.branchName || s.dept || '', isLateral: s.isLateral, cgpa, sgpa, totF: sems.reduce((a, sm) => a + sm.failCount, 0), clr: sems.every(sm => sm.semPassed) };
    }).sort((a, b) => b.cgpa - a.cgpa);
    const hdr = ['Rank', 'Reg. No.', 'Name', 'Dept', 'Type', 'Latest SGPA', 'Latest CGPA', 'Total Arrears', 'Status'];
    const dat = ranked.map((s, i) => [i + 1, s.roll, s.name, s.dept, s.isLateral ? 'Lateral' : 'Regular', s.sgpa, parseFloat(s.cgpa.toFixed(2)), s.totF, s.clr ? 'ALL CLEAR' : 'HAS ARREARS']);
    const aoa = [[batchLabel + ' - Leaderboard'], ['Total: ' + ranked.length], [], hdr, ...dat];
    const ws = XL.utils.aoa_to_sheet(aoa); ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: nc - 1 } }]; ws['!cols'] = autoW(aoa);
    ws['!rows'] = [{ hpt: 32 }, { hpt: 18 }, { hpt: 8 }, { hpt: 24 }, ...mkRows(20, dat.length)];
    ws['!freeze'] = { xSplit: 0, ySplit: 4, topLeftCell: 'A5', activePane: 'bottomLeft', state: 'frozen' };
    styR(ws, 0, 0, 0, nc - 1, S.title); styR(ws, 1, 0, 1, nc - 1, S.info); styR(ws, 3, 0, 3, nc - 1, S.hdr);
    dat.forEach((row, ri) => {
      const r = 4 + ri, odd = ri % 2 === 1;
      row.forEach((v, ci) => {
        let st; const vS = String(v);
        if (ci === 8) st = (vS === 'ALL CLEAR') ? S.pass : S.fail;
        else if (ci === 4) st = (vS === 'Lateral') ? S.warn : rowS(odd, false);
        else if (ci === 6) { const cg = parseFloat(v); st = cg >= 8 ? S.pass : cg >= 6 ? rowS(odd, false) : cg >= 5 ? S.warn : S.fail; }
        else if (ci === 0) st = (ri < 3) ? S.gold : rowS(odd, false);
        else if (ci === 1 || ci === 2 || ci === 3) st = rowS(odd, true);
        else st = rowS(odd, false);
        sty(ws, r, ci, st);
      });
    });
    addFooter(ws, aoa, nc);
    XL.utils.book_append_sheet(wb, ws, 'Ranking');
  })();

  /* --- SHEET: FAILURE ANALYSIS --- */
  (() => {
    const subMap = {};
    semData.forEach(sem => {
      if (sem.subStats) {
        Object.entries(sem.subStats).forEach(([code, st]) => {
          if (!subMap[code]) subMap[code] = { passes: 0, fails: 0, feCount: 0, abCount: 0, whCount: 0, total: 0, sems: [] };
          subMap[code].passes += st.passes;
          subMap[code].fails += st.fails;
          subMap[code].feCount += (st.feCount || 0);
          subMap[code].abCount += (st.abCount || 0);
          subMap[code].whCount += (st.whCount || 0);
          subMap[code].total += st.total;
          subMap[code].sems.push(sem.label);
        });
      }
    });

    const subjects = Object.entries(subMap).sort((a, b) => b[1].fails - a[1].fails);
    const stuArr = allStu.map(s => {
      const sems = Object.values(s.sems);
      const totF = sems.reduce((a, sm) => a + sm.failCount, 0);
      return { name: s.name, roll: s.roll, totF, latCGPA: latestCGPA(s), semCount: sems.length };
    }).filter(s => s.totF > 0).sort((a, b) => b.totF - a.totF);

    const nc = 10;
    const aoa = []; const merges = []; const rowH = []; let cr = 0;
    const push = (row, hpt) => { aoa.push(row); rowH.push({ hpt }); cr++; };
    const pushM = (row, hpt, nc2) => { merges.push({ s: { r: cr, c: 0 }, e: { r: cr, c: nc2 - 1 } }); push(row, hpt); };

    pushM([`${batchLabel} — Failure Analysis`], 32, nc);
    pushM([`Subjects with Failures: ${subjects.filter(([, s]) => s.fails > 0).length} | Students with Arrears: ${stuArr.length} | Total F Grades: ${semData.reduce((a, s) => a + (s.totalFailures || 0), 0)}`], 18, nc);
    push([], 10);

    // Students with Highest Arrears
    pushM(['⚠️ STUDENTS WITH HIGHEST ARREARS'], 24, nc);
    const saHdr = ['Rank', 'Reg. No.', 'Name', 'Semesters', 'Total Arrears', 'Latest CGPA'];
    push(saHdr, 22);
    stuArr.forEach((s, ri) => push([ri + 1, s.roll, s.name, s.semCount, s.totF, parseFloat(s.latCGPA.toFixed(2))], 20));
    const saDatEnd = cr;

    push([], 10);

    // Subject Failure Details
    pushM(['📊 SUBJECT FAILURE DETAILS'], 24, nc);
    const sfHdr = ['#', 'Subject Code', 'Semesters', 'Total', 'Passed', 'F', 'FE', 'AB', 'WH', 'Fail %'];
    push(sfHdr, 22);
    const sfDat = subjects.map(([code, st], i) => {
      const totalFail = st.fails + (st.feCount || 0) + st.abCount + (st.whCount || 0);
      const fp = st.total ? (totalFail / st.total * 100).toFixed(1) : 0;
      return [i + 1, code, st.sems.join(', '), st.total, st.passes, st.fails, st.feCount || 0, st.abCount, st.whCount || 0, parseFloat(fp)];
    });
    sfDat.forEach((row, ri) => push(row, 20));
    const sfDatEnd = cr;

    const ws = XL.utils.aoa_to_sheet(aoa);
    ws['!merges'] = merges;
    ws['!rows'] = rowH;
    ws['!cols'] = autoW(aoa);
    ws['!freeze'] = { xSplit: 0, ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' };

    styR(ws, 0, 0, 0, nc - 1, S.title); styR(ws, 1, 0, 1, nc - 1, S.info);

    const saHdrRow = 3;
    const saColHdrRow = 4;
    const saDatStartRow = 5;
    styR(ws, saHdrRow, 0, saHdrRow, nc - 1, S.hdrR);
    styR(ws, saColHdrRow, 0, saColHdrRow, nc - 1, S.hdr);
    stuArr.forEach((s, ri) => {
      const r = saDatStartRow + ri, odd = ri % 2 === 1;
      [ri + 1, s.roll, s.name, s.semCount, s.totF, s.latCGPA.toFixed(2)].forEach((v, ci) => {
        let st;
        if (ci === 4) st = s.totF >= 5 ? S.fail : s.totF >= 3 ? S.warn : rowS(odd, false);
        else if (ci === 1 || ci === 2) st = rowS(odd, true);
        else st = rowS(odd, false);
        sty(ws, r, ci, st);
      });
    });

    const sfHdrRow = saDatEnd + 1;
    const sfColHdrRow = sfHdrRow + 1;
    const sfDatStartRow = sfColHdrRow + 1;
    styR(ws, sfHdrRow, 0, sfHdrRow, nc - 1, S.secG);
    styR(ws, sfColHdrRow, 0, sfColHdrRow, nc - 1, S.hdr);
    sfDat.forEach((row, ri) => {
      const r = sfDatStartRow + ri, odd = ri % 2 === 1;
      row.forEach((v, ci) => {
        let st;
        if (ci === 9) st = parseFloat(v) >= 50 ? S.fail : parseFloat(v) >= 25 ? S.warn : rowS(odd, false);
        else if (ci <= 2) st = rowS(odd, true);
        else st = rowS(odd, false);
        sty(ws, r, ci, st);
      });
    });

    addFooter(ws, aoa, nc);
    XL.utils.book_append_sheet(wb, ws, 'Failure Analysis');
  })();

  const b64 = XL.write(wb, { type: 'base64', bookType: 'xlsx' });
  const allYrs = allStu.map(s => parseInt(s.admYear)).filter(y => !isNaN(y));
  const yr = allYrs.length ? Math.max(...allYrs) : 'Batch';
  const allBrs = allStu.map(s => s.branchCode).filter(Boolean);
  const br = allBrs.length ? allBrs[0] : '';
  const filename = `APJ_AKTU_Results_${yr}_${br}_Analysis.xlsx`;
  return { filename, b64 };
}

/* ── API HANDLER ────────────────────────────────────────────────── */
module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { files } = req.body; 
    if (!files || !Array.isArray(files) || files.length === 0) return res.status(400).json({ error: 'Missing or empty files array in request' });

    const semDataArr = [];
    for (let f of files) {
      if (!f.base64) continue;
      const buffer = Buffer.from(f.base64, 'base64');
      
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
