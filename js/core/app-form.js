/*
 * 教學住診／教學門診「補助費申請單」：比對與轉換（不含畫面）
 *
 * 規則（使用者確認 2026-11-10）：以各單位送來的申請單（PDF 掃描）為計費基準；
 * rawdata 用來比對：取得卡號、確認系統有紀錄。申請單有、系統沒有 → 仍計入但紅字標示。
 *
 * 申請單列：{ date: 'YYYY-MM-DD', teacher, student, unit }
 * rawdata 列：parseRawRows 產生 { tCard, tName, sName, date, done, unit }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./inpatient-calc.js'));
  else root.AppForm = factory(root.InpatientCalc);
})(typeof self !== 'undefined' ? self : this, (Base) => {
  const isPUA = (ch) => { const c = ch.codePointAt(0); return c >= 0xE000 && c <= 0xF8FF; };
  /** 姓名正規化：NFKC、只留中日韓文字與英數 */
  const normName = (s) => String(s ?? '').normalize('NFKC').replace(/[^\p{Script=Han}A-Za-z0-9-]/gu, '');

  /** 姓名相似度 0～1（共同字數 / 較長者字數；造字視為相符） */
  function sim(a, b) {
    const x = [...normName(a)]; const y = [...normName(b)];
    if (!x.length || !y.length) return 0;
    const pool = [...y];
    let hit = 0;
    x.forEach((ch, i) => {
      let j = pool.indexOf(ch);
      if (j < 0 && (isPUA(ch) || (y[i] && isPUA(y[i])))) j = pool.indexOf(y[i]) >= 0 ? pool.indexOf(y[i]) : -1;
      if (j >= 0) { hit += 1; pool.splice(j, 1); }
    });
    return hit / Math.max(x.length, y.length);
  }

  /** OCR 日期文字 → ISO；年月有誤時以目標年月補正 */
  function fixDate(text, ym) {
    const d = String(text ?? '').normalize('NFKC').replace(/[Oo]/g, '0').replace(/[Il|]/g, '1').replace(/\D/g, '');
    const ty = Number(ym.slice(0, 4)); const tm = Number(ym.slice(5, 7));
    let y; let m; let day;
    if (d.length >= 8) { y = Number(d.slice(0, 4)); m = Number(d.slice(4, 6)); day = Number(d.slice(6, 8)); } else if (d.length >= 3) { day = Number(d.slice(-2)); }
    if (!day || day > 31) return '';
    // 申請單一份只申報一個月：年、月與目標月份不同時，視為辨識錯誤，以目標年月為準
    y = ty;
    m = tm;
    const dt = new Date(y, m - 1, day);
    if (dt.getMonth() !== m - 1) return '';
    return `${y}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  /** rawdata「課程rawdata」物件陣列 → 比對用列（不限月份、不限是否評核完成） */
  function parseRawRows(rows) {
    const C = Base.COL;
    return rows.map((r) => ({
      tCard: Base.normCard(r[C.card]),
      tName: String(r[C.name] || '').trim(),
      sName: String(r['姓名'] || '').trim(),
      date: Base.normDate(r[C.date]),
      done: String(r[C.done] || '').trim().toUpperCase() === 'Y',
      unit: String(r[C.unit] || '').trim(),
      rowNum: typeof r.__rowNum__ === 'number' ? r.__rowNum__ + 1 : null,   // 原始 Excel 列號（SheetJS 提供，0 起算）
    })).filter((r) => r.date && r.tCard);
  }

  /**
   * 申請單列 ↔ rawdata 比對（同一筆 rawdata 只配對一次，分數高者優先）
   * 回傳每列：{ ...row, match: rawRow|null, score, status: 'ok'|'not-done'|'no-system'|'unsure' }
   */
  function matchRows(appRows, rawRows, ym) {
    const byDate = new Map();
    rawRows.forEach((r) => { if (!byDate.has(r.date)) byDate.set(r.date, []); byDate.get(r.date).push(r); });
    const monthRows = rawRows.filter((r) => r.date.startsWith(ym));
    const cands = [];
    appRows.forEach((a, i) => {
      const pool = a.date && byDate.has(a.date) ? byDate.get(a.date) : (a.date ? [] : monthRows);
      const scored = pool.map((r) => {
        const st = sim(a.teacher, r.tName); const ss = sim(a.student, r.sName);
        return { r, st, ss, score: st + ss + (a.date && r.date === a.date ? 0.2 : 0) };
      }).sort((x, y) => y.score - x.score);
      scored.forEach(({ r, st, ss, score }, k) => {
        let ok = a.date
          // 同日期：兩個名字都有部分相符，或其中一個名字高度相符
          ? (st >= 0.34 && ss >= 0.34 && score >= 1.0) || ((st >= 0.66 || ss >= 0.66) && st + ss >= 0.66)
          : st >= 0.5 && ss >= 0.5 && score >= 1.4;
        // 同日期候選很少：第一名明顯領先（≥0.3）且有一點相似度，也採用（標為不確定）
        if (!ok && a.date && k === 0 && st + ss >= 0.3 && score - (scored[1]?.score || 0) >= 0.3) ok = true;
        if (ok) cands.push({ i, r, score });
      });
    });
    cands.sort((x, y) => y.score - x.score);
    const usedRaw = new Set(); const result = appRows.map((a) => ({ ...a, match: null, score: 0, status: 'no-system' }));
    cands.forEach(({ i, r, score }) => {
      if (result[i].match || usedRaw.has(r)) return;
      usedRaw.add(r);
      result[i].match = r; result[i].score = score;
      result[i].status = r.done ? (score >= 1.4 ? 'ok' : 'unsure') : 'not-done';
    });

    // 第二輪（同一頁＝同一單位）：老師姓名辨識失敗但日期正確 → 在本頁已比對到的老師中找當天紀錄
    const pageTeachers = new Map();
    result.forEach((m) => {
      if (!m.page || !m.match) return;
      if (!pageTeachers.has(m.page)) pageTeachers.set(m.page, new Set());
      pageTeachers.get(m.page).add(m.match.tCard);
    });
    result.forEach((m) => {
      if (m.match || !m.date || !m.page || !pageTeachers.has(m.page)) return;
      const tset = pageTeachers.get(m.page);
      const pool = (byDate.get(m.date) || []).filter((r) => tset.has(r.tCard) && !usedRaw.has(r))
        .map((r) => ({ r, s: sim(m.student, r.sName) + 0.5 * sim(m.teacher, r.tName) }))
        .sort((x, y) => y.s - x.s);
      if (!pool.length) return;
      if (pool.length === 1 || pool[0].s - pool[1].s >= 0.3) {
        const r = pool[0].r;
        usedRaw.add(r);
        m.match = r; m.score = pool[0].s;
        m.status = r.done ? 'unsure' : 'not-done';
      }
    });
    return result;
  }

  /** 由姓名找卡號（rawdata 中的評核者），找不到回傳 '' */
  function teacherCard(rawRows, name) {
    let best = null; let bestS = 0;
    const seen = new Map();
    rawRows.forEach((r) => { if (!seen.has(r.tCard)) seen.set(r.tCard, r); });
    seen.forEach((r) => { const s = sim(name, r.tName); if (s > bestS) { bestS = s; best = r; } });
    return bestS >= 0.66 ? best : null;
  }

  /**
   * 比對結果 → 計費紀錄
   * 有系統紀錄者用系統的卡號／姓名；沒有者以姓名找卡號（或使用者手動填的 card），標記 inSystem=false。
   */
  function toRecords(matched, rawRows) {
    return matched.filter((m) => m.include !== false).map((m) => {
      if (m.match) {
        return { card: m.match.tCard, name: m.match.tName, unit: m.unit || m.match.unit, date: m.match.date,
          student: normName(m.match.sName), source: '申請單', appUnit: m.unit, inSystem: m.match.done };
      }
      const t = m.card ? { tCard: Base.normCard(m.card), tName: m.teacher } : teacherCard(rawRows, m.teacher);
      // 系統無紀錄：直接計入、紅字；找不到卡號時卡號留空、姓名用辨識文字（下載後人工核對）
      return { card: t?.tCard || '', name: t?.tName || normName(m.teacher) || `（無法辨識：${m.page?.file || ''} p${m.page?.page || ''}）`,
        unit: m.unit, date: m.date, student: normName(m.student), source: '申請單', appUnit: m.unit, inSystem: false, unresolved: !t };
    }).filter((r) => r.date);
  }

  return { normName, sim, fixDate, parseRawRows, matchRows, teacherCard, toRecords };
});
