/*
 * 教學門診指導費：計算與比對邏輯（不含畫面，可在瀏覽器與 Node 測試共用）
 *
 * 規則（使用者確認 2026-11-09；申請單說明：一位學員 1 份紀錄 500 元，每次上限 1500 元）：
 *  - 依「評核/訓練開始日(自填)」判斷教學日，只取目標年月；只算老師已評核完成 = Y。
 *  - 每位老師每個日期：學員數 × 500，上限 3 位（1,500 元）。同一學員同日只算一次。
 *  - 各單位的「教學門診補助費申請單」(.doc) 與系統資料比對；
 *    申請單有、系統沒有的學員，可選擇是否計入。
 *  - 姓名比對：先做 Unicode NFKC 正規化（Word 常用相容字，如 沈 U+F972）；
 *    醫院造字（私用區 U+E000–U+F8FF）視為可與任何字相符。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./inpatient-calc.js'));
  else root.OutpatientCalc = factory(root.InpatientCalc);
})(typeof self !== 'undefined' ? self : this, (Base) => {
  const RATE_PER_STUDENT = 500;
  const CAP_STUDENTS_PER_DAY = 3;
  // 這些單位另外排在名冊最後（使用者指示：一般內科情況特殊）
  const LAST_UNITS = ['一般內科'];
  // 內科部門診表的科別代碼 → 單位（使用者提供，2026-11-17）
  const CODE_UNIT = {
    GM: '一般內科', CV: '心臟內科', CM: '胸腔內科', HEMA: '血液科', ONC: '腫瘤科', INF: '感染科', META: '新陳代謝科',
    NEPH: '腎臟科', NEUR: '神經內科', AIR: '過敏免疫風濕科', GI: '胃腸科', GER: '高齡醫學中心',
  };
  // 單位名稱轉成正式名稱（申請單／系統上的簡稱 → 名冊用的名稱）
  const UNIT_ALIAS = {
    復健科: '復健醫學部', 復健部: '復健醫學部', 婦產部: '婦女醫學部', 婦產科: '婦女醫學部', 耳鼻喉部: '耳鼻喉頭頸部', 耳鼻喉科: '耳鼻喉頭頸部',
    兒科: '兒童醫學部', 兒醫部: '兒童醫學部', 高齡醫學: '高齡醫學中心', 腸胃科: '胃腸科', 風濕免疫科: '過敏免疫風濕科',
  };
  const officialUnit = (u) => UNIT_ALIAS[String(u || '').trim()] || String(u || '').trim();

  const normName = (s) => String(s ?? '').normalize('NFKC').replace(/\s+/g, '');
  const isPUA = (ch) => { const c = ch.codePointAt(0); return c >= 0xE000 && c <= 0xF8FF; };

  /** 兩個姓名是否為同一人（容許醫院造字） */
  function sameName(a, b) {
    const x = [...normName(a)]; const y = [...normName(b)];
    if (x.length !== y.length || !x.length) return false;
    return x.every((ch, i) => ch === y[i] || isPUA(ch) || isPUA(y[i]));
  }

  /** 去除重複學員（容許造字），保留第一次出現的寫法；系統寫法優先 */
  function uniqueStudents(list) {
    const out = [];
    for (const s of list) {
      const hit = out.find((o) => sameName(o.name, s.name));
      if (!hit) out.push({ ...s, sources: new Set([s.source]) });
      else hit.sources.add(s.source);
    }
    return out;
  }

  /** 解析 rawdata「課程rawdata」→ 系統紀錄 + 評核者通訊錄（姓名→卡號） */
  function parseRaw(rows, year, month) {
    const C = Base.COL;
    const prefix = `${year}-${String(month).padStart(2, '0')}`;
    const records = [];
    const directory = new Map();
    const skipped = { notDone: 0, otherMonth: 0, noCard: 0 };
    for (const r of rows) {
      const card = Base.normCard(r[C.card]);
      const name = normName(r[C.name]);
      const unit = String(r[C.unit] || '').trim();
      if (name && card) {
        if (!directory.has(name)) directory.set(name, new Map());
        const e = directory.get(name).get(card) || { count: 0, unit };
        e.count += 1;
        directory.get(name).set(card, e);
      }
      const date = Base.normDate(r[C.date]);
      if (!date.startsWith(prefix)) { skipped.otherMonth += 1; continue; }
      if (String(r[C.done] || '').trim().toUpperCase() !== 'Y') { skipped.notDone += 1; continue; }
      if (!card) { skipped.noCard += 1; continue; }
      records.push({
        card, name, unit, date,
        student: normName(r['姓名']), source: '系統',
      });
    }
    return { records, directory, skipped };
  }

  /** 申請單日期：'1150908'、'115/09/08'、'115.9.8'、'09/08' → ISO */
  function parseAppDate(s, year) {
    const t = String(s ?? '').normalize('NFKC').trim();
    let m = t.match(/^(\d{3})(\d{2})(\d{2})$/) || t.match(/^(\d{2,4})[/.\-年](\d{1,2})[/.\-月](\d{1,2})日?$/);
    if (m) {
      let y = Number(m[1]); if (y < 1911) y += 1911;
      return Base.normDate(`${y}-${m[2]}-${m[3]}`);
    }
    m = t.match(/^(\d{1,2})[/.\-月](\d{1,2})日?$/);
    if (m && year) return Base.normDate(`${year}-${m[1]}-${m[2]}`);
    return '';
  }

  /**
   * 解析申請單（DocReader.parse 的結果）
   * 回傳 { unit, applicant, rocYear, month, entries: [{ teacher, students[], date, raw }], errors[] }
   */
  function parseApplication(doc, fallbackYear) {
    const head = doc.paragraphs.join('\n').normalize('NFKC');
    const unit = (head.match(/單位[:：]\s*(\S+)/) || [])[1] || '';
    const applicant = (head.match(/申請人[:：]\s*(\S+)/) || [])[1] || '';
    const ym = head.match(/申請月份[:：]\s*(\d{2,4})\s*年\s*(\d{1,2})\s*月/);
    const rocYear = ym ? Number(ym[1]) : null;
    const month = ym ? Number(ym[2]) : null;
    const year = rocYear ? rocYear + 1911 : fallbackYear;

    const entries = [];
    const errors = [];
    const [header, ...rows] = doc.table;
    if (!header) return { unit, applicant, rocYear, month, entries, errors: ['找不到表格'] };
    const h = header.map((x) => normName(x));
    const ci = {
      teacher: h.findIndex((x) => /醫師|老師|教師/.test(x)),
      students: h.findIndex((x) => /學員|學生/.test(x)),
      date: h.findIndex((x) => /日期/.test(x)),
    };
    if (Object.values(ci).some((i) => i < 0)) {
      return { unit, applicant, rocYear, month, entries, errors: ['表頭需含「主治醫師」「學員」「日期」'] };
    }
    rows.forEach((r, i) => {
      const teacher = normName(r[ci.teacher]);
      const studentsText = String(r[ci.students] || '');
      const dateText = String(r[ci.date] || '');
      if (!teacher && !studentsText.trim() && !dateText.trim()) return;   // 空白列
      const date = parseAppDate(dateText, year);
      const students = studentsText.split(/[、,，;；/\s\n]+/).map(normName).filter(Boolean);
      if (!teacher || !date) errors.push(`第 ${i + 2} 列資料不完整：${r.join(' | ')}`);
      entries.push({ teacher, students, date, row: i + 2 });
    });
    return { unit, applicant, rocYear, month, entries, errors };
  }

  /** 申請單 → 紀錄（需卡號對照；cardOverride: 姓名→卡號 手動指定） */
  function applicationRecords(apps, directory, prefix, cardOverride = {}) {
    const recs = [];
    for (const app of apps) {
      for (const e of app.entries) {
        if (!e.teacher || !e.date || !e.date.startsWith(prefix)) continue;
        const hit = Base.lookupCard(directory, e.teacher);
        const card = Base.normCard(cardOverride[e.teacher] || hit.card);
        const unit = hit.unit || app.unit;
        e.students.forEach((s) => recs.push({
          card, name: e.teacher, unit, date: e.date, student: s, source: '申請單', appUnit: app.unit,
        }));
      }
    }
    return recs;
  }

  const teacherKey = (r) => r.card || `name:${r.name}`;

  /**
   * 比對：以「老師 + 日期」為單位，列出申請單與系統的學員差異。
   * 只比對申請單上出現的老師（申請單是各單位自行提出）。
   */
  function compare(sysRecords, appRecords) {
    const group = new Map();
    for (const r of [...sysRecords, ...appRecords]) {
      const k = `${teacherKey(r)}|${r.date}`;
      if (!group.has(k)) group.set(k, { card: r.card, name: r.name, unit: r.appUnit || r.unit, date: r.date, sys: [], app: [] });
      const g = group.get(k);
      (r.source === '系統' ? g.sys : g.app).push(r.student);
      if (r.appUnit) g.unit = r.appUnit;
    }
    const appTeachers = new Set(appRecords.map(teacherKey));
    const rows = [];
    for (const g of group.values()) {
      const tk = g.card || `name:${g.name}`;
      if (!appTeachers.has(tk)) continue;
      const onlyApp = g.app.filter((a) => !g.sys.some((s) => sameName(a, s)));
      const onlySys = g.sys.filter((s) => !g.app.some((a) => sameName(a, s)));
      let status = '一致';
      if (!g.app.length) status = '申請單未列';
      else if (!g.sys.length) status = '系統無紀錄';
      else if (onlyApp.length || onlySys.length) status = '有差異';
      rows.push({ ...g, sysCount: uniqueStudents(g.sys.map((n) => ({ name: n }))).length, appCount: g.app.length, onlyApp, onlySys, status });
    }
    rows.sort((a, b) => a.unit.localeCompare(b.unit, 'zh-Hant') || a.name.localeCompare(b.name, 'zh-Hant') || a.date.localeCompare(b.date));
    return rows;
  }

  /** 彙整每位老師：每日學員數（上限 3）× 500 */
  function summarize(records, { rate = RATE_PER_STUDENT, cap = CAP_STUDENTS_PER_DAY } = {}) {
    const byTeacher = new Map();
    for (const r of records) {
      const k = teacherKey(r);
      if (!byTeacher.has(k)) byTeacher.set(k, []);
      byTeacher.get(k).push(r);
    }
    const result = [];
    for (const list of byTeacher.values()) {
      const byDate = new Map();
      list.forEach((r) => {
        if (!byDate.has(r.date)) byDate.set(r.date, []);
        byDate.get(r.date).push({ name: r.student, source: r.source });
      });
      const days = [...byDate.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([date, studs]) => {
        // 系統紀錄排前面，讓去重時保留系統寫法
        const students = uniqueStudents(studs.sort((a, b) => (a.source === '系統' ? -1 : 1) - (b.source === '系統' ? -1 : 1)));
        const counted = Math.min(students.length, cap);
        return { date, students, count: students.length, counted, paid: counted * rate };
      });
      const sysNames = list.filter((r) => r.source === '系統').map((r) => r.name);
      const pick = (arr) => {
        const c = new Map(); arr.filter(Boolean).forEach((v) => c.set(v, (c.get(v) || 0) + 1));
        return [...c.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || '';
      };
      // 單位：老師有出現在申請單 → 用申請單的單位；否則用系統的訓練單位
      const appUnit = pick(list.map((r) => r.appUnit));
      result.push({
        card: list[0].card,
        name: pick(sysNames) || pick(list.map((r) => r.name)),
        unit: appUnit || pick(list.filter((r) => r.source === '系統').map((r) => r.unit)) || pick(list.map((r) => r.unit)),
        days,
        daysText: days.map((d) => `${d.date.slice(5, 7)}/${d.date.slice(8, 10)}`).join(', '),
        amount: days.reduce((a, d) => a + d.paid, 0),
        fromApp: days.reduce((a, d) => a + d.students.filter((s) => !s.sources.has('系統')).length, 0),
        overCap: days.filter((d) => d.count > cap).length,
      });
    }
    const lastRank = (u) => { const i = LAST_UNITS.indexOf(u); return i < 0 ? -1 : i; };
    result.sort((a, b) => lastRank(a.unit) - lastRank(b.unit)
      || a.unit.localeCompare(b.unit, 'zh-Hant') || a.card.localeCompare(b.card));
    return result;
  }

  /**
   * 標記「因申請單而增加金額」的老師：與只用系統資料的結果比較。
   * 會在每位老師加上 sysAmount、appExtra（增加的金額）、highlight（是否紅字）。
   */
  function markApplicationEffect(summaryAll, summarySys) {
    const sys = new Map(summarySys.map((s) => [s.card || `name:${s.name}`, s.amount]));
    summaryAll.forEach((s) => {
      s.sysAmount = sys.get(s.card || `name:${s.name}`) || 0;
      s.appExtra = s.amount - s.sysAmount;
      s.highlight = s.appExtra > 0;
    });
    return summaryAll;
  }

  /** 顯示用：'09/08(3), 09/15(5→3)' */
  function formatDays(days, cap = CAP_STUDENTS_PER_DAY) {
    return days.map((d) => `${d.date.slice(5, 7)}/${d.date.slice(8, 10)}(${d.count > cap ? `${d.count}→${cap}` : d.count})`).join(', ');
  }

  return {
    RATE_PER_STUDENT, CAP_STUDENTS_PER_DAY, LAST_UNITS, CODE_UNIT, UNIT_ALIAS, officialUnit,
    normName, sameName, parseRaw, parseAppDate, parseApplication, applicationRecords, compare, summarize,
    markApplicationEffect, formatDays,
  };
});
