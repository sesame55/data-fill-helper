/*
 * 教學住診指導費：計算邏輯（不含畫面，可在瀏覽器與 Node 測試共用）
 *
 * 規則（使用者確認 2026-11-09）：
 *  - 依「評核/訓練開始日(自填)」判斷教學日，只取目標年月。
 *  - 只算「老師是否評核完成」= Y 的紀錄。
 *  - 依教學日計算：同一老師同一天只算一次。
 *  - 每週 1,000 元；同一週教學多天仍只算 1,000 元。
 *  - 手寫紀錄（PDF）同樣算入，與系統紀錄合併後去除重複日期。
 *  - 不使用 VS 名單，老師資料取自 rawdata 的評核者。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.InpatientCalc = factory();
})(typeof self !== 'undefined' ? self : this, () => {
  const RATE_PER_WEEK = 1000;
  // 這些單位另外排在名冊最後（使用者指示：一般內科情況特殊）
  const LAST_UNITS = ['一般內科'];

  const COL = {
    card: '評核者卡號',
    name: '評核者',
    date: '評核/訓練開始日(自填)',
    done: '老師是否評核完成',
    unit: '訓練單位名稱',
  };

  /** 卡號統一成至少 4 碼字串（12 → 0012，'A123' 不變） */
  function normCard(v) {
    if (v === null || v === undefined) return '';
    const s = String(v).trim();
    return s ? s.padStart(4, '0') : '';
  }

  /** 各種日期格式 → 'YYYY-MM-DD'；無效回傳 '' */
  function normDate(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date && !isNaN(v)) {
      return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
    }
    const m = String(v).trim().match(/^(\d{3,4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (!m) return '';
    let y = Number(m[1]);
    if (y < 1911) y += 1911;              // 民國年
    return `${y}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  }

  /** 該日期所屬週的起始日（weekStart: 1=週一, 0=週日），作為「週」的識別 */
  function weekKey(isoDate, weekStart = 1) {
    const [y, m, d] = isoDate.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    const diff = (dt.getUTCDay() - weekStart + 7) % 7;
    dt.setUTCDate(dt.getUTCDate() - diff);
    return dt.toISOString().slice(0, 10);
  }

  /**
   * 解析 rawdata「課程rawdata」工作表（物件陣列，key 為表頭）
   * 回傳 { records, directory, skipped }
   *  - records：符合年月且老師已評核完成的教學紀錄
   *  - directory：全部評核者 姓名→卡號 對照（給手寫紀錄找卡號用）
   */
  function parseRaw(rows, year, month) {
    const prefix = `${year}-${String(month).padStart(2, '0')}`;
    const records = [];
    const directory = new Map();     // name -> Map(card -> {count, unit})
    const skipped = { notDone: 0, otherMonth: 0, noCard: 0 };

    for (const r of rows) {
      const card = normCard(r[COL.card]);
      const name = String(r[COL.name] || '').trim();
      const unit = String(r[COL.unit] || '').trim();
      if (name && card) {
        if (!directory.has(name)) directory.set(name, new Map());
        const cards = directory.get(name);
        const e = cards.get(card) || { count: 0, unit };
        e.count += 1;
        cards.set(card, e);
      }
      const date = normDate(r[COL.date]);
      if (!date.startsWith(prefix)) { skipped.otherMonth += 1; continue; }
      if (String(r[COL.done] || '').trim().toUpperCase() !== 'Y') { skipped.notDone += 1; continue; }
      if (!card) { skipped.noCard += 1; continue; }
      records.push({ card, name, unit, date, source: '系統' });
    }
    return { records, directory, skipped };
  }

  /** 由姓名找卡號：唯一時回傳，否則回傳候選清單 */
  function lookupCard(directory, name) {
    const cards = directory.get(String(name || '').trim());
    if (!cards) return { card: '', unit: '', candidates: [] };
    const list = [...cards.entries()].sort((a, b) => b[1].count - a[1].count);
    return {
      card: list.length === 1 ? list[0][0] : '',
      unit: list.length === 1 ? list[0][1].unit : '',
      candidates: list.map(([card, e]) => ({ card, unit: e.unit })),
    };
  }

  const mostCommon = (arr) => {
    const c = new Map();
    arr.filter(Boolean).forEach((v) => c.set(v, (c.get(v) || 0) + 1));
    return [...c.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || '';
  };

  /**
   * 彙整：每位老師的教學日、週數、金額
   * records 可包含系統與手寫紀錄；同一老師以卡號歸戶（無卡號時以姓名）
   */
  function summarize(records, { weekStart = 1, rate = RATE_PER_WEEK } = {}) {
    const byTeacher = new Map();
    for (const r of records) {
      const key = r.card || `name:${r.name}`;
      if (!byTeacher.has(key)) byTeacher.set(key, []);
      byTeacher.get(key).push(r);
    }

    const result = [];
    for (const [, list] of byTeacher) {
      const days = [...new Set(list.map((r) => r.date))].sort();
      const weeks = [...new Set(days.map((d) => weekKey(d, weekStart)))].sort();
      const sysName = mostCommon(list.filter((r) => r.source === '系統').map((r) => r.name));
      result.push({
        card: list[0].card,
        name: sysName || mostCommon(list.map((r) => r.name)),
        unit: mostCommon(list.map((r) => r.unit)),
        days,
        weeks,
        sessions: list.length,
        handwritten: list.filter((r) => r.source === '手寫').length,
        amount: weeks.length * rate,
      });
    }
    const lastRank = (u) => { const i = LAST_UNITS.indexOf(u); return i < 0 ? -1 : i; };
    result.sort((a, b) => lastRank(a.unit) - lastRank(b.unit)
      || a.unit.localeCompare(b.unit, 'zh-Hant') || a.card.localeCompare(b.card));
    return result;
  }

  /** 教學日清單：'09/10, 09/20' */
  function formatDays(days) {
    return days.map((d) => `${d.slice(5, 7)}/${d.slice(8, 10)}`).join(', ');
  }

  return { RATE_PER_WEEK, LAST_UNITS, COL, normCard, normDate, weekKey, parseRaw, lookupCard, summarize, formatDays };
});
