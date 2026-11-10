/*
 * 教學住診指導費：把彙整結果填入使用者提供的範本（ExcelJS 活頁簿）
 * 範本格式：第一個工作表，表頭列含「卡號」，下面為資料列，最後有「合計」列。
 * 資料超過範本列數時，自動複製最後一列格式往下插入，並更新合計公式。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./inpatient-calc.js'));
  else root.InpatientFill = factory(root.InpatientCalc);
})(typeof self !== 'undefined' ? self : this, (Calc) => {
  const cellText = (c) => {
    const v = c.value;
    if (v && typeof v === 'object') return String(v.result ?? v.text ?? (v.richText ? v.richText.map((t) => t.text).join('') : ''));
    return v === null || v === undefined ? '' : String(v);
  };

  /** 找出範本結構：表頭列、欄位位置、資料起始列、合計列 */
  function analyzeTemplate(ws) {
    let headerRow = 0;
    let totalRow = 0;
    const cols = {};
    ws.eachRow({ includeEmpty: false }, (row, n) => {
      if (!headerRow) {
        row.eachCell((cell) => { if (cellText(cell).replace(/\s/g, '') === '卡號') headerRow = n; });
      }
      row.eachCell((cell, c) => {
        const t = cellText(cell).replace(/\s/g, '');
        if (headerRow === n) {
          if (t === '年') cols.year = c;
          if (t === '月') cols.month = c;
          if (t === '單位') cols.unit = c;
          if (t === '卡號') cols.card = c;
          if (t === '姓名') cols.name = c;
          if (t === '金額') cols.amount = c;
          if (t === '備註') cols.note = c;
          if (t === '教學日') cols.days = c;
        }
        if (headerRow && n > headerRow && !totalRow && t === '合計') totalRow = n;
      });
    });
    if (!headerRow || !totalRow || !cols.card || !cols.amount) {
      throw new Error('範本格式無法辨識：找不到「卡號」表頭或「合計」列');
    }
    const startRow = headerRow + 1;
    const yearCell = cols.year ? Number(cellText(ws.getRow(startRow).getCell(cols.year))) : NaN;
    const monthCell = cols.month ? Number(cellText(ws.getRow(startRow).getCell(cols.month))) : NaN;
    return {
      headerRow, startRow, totalRow, cols,
      capacity: totalRow - startRow,
      rocYear: yearCell || null,
      month: monthCell || null,
    };
  }

  /** 填寫；回傳填入筆數與是否有插入列 */
  function fillTemplate(wb, summary, { today = new Date() } = {}) {
    const ws = wb.worksheets[0];
    const t = analyzeTemplate(ws);
    const { cols, startRow } = t;
    let totalRow = t.totalRow;

    const extra = summary.length - t.capacity;
    if (extra > 0) {
      ws.duplicateRow(totalRow - 1, extra, true);   // 複製最後一列（含格式）往下插入
      totalRow += extra;
    }
    const lastRow = totalRow - 1;

    for (let i = 0; i < summary.length; i += 1) {
      const s = summary[i];
      const row = ws.getRow(startRow + i);
      if (cols.year && t.rocYear) row.getCell(cols.year).value = t.rocYear;
      if (cols.month && t.month) row.getCell(cols.month).value = t.month;
      if (cols.unit) row.getCell(cols.unit).value = s.unit;
      row.getCell(cols.card).value = s.card;                 // 文字，保留前導 0
      if (cols.name) row.getCell(cols.name).value = s.name;
      row.getCell(cols.amount).value = s.amount;
      if (cols.days) row.getCell(cols.days).value = s.daysText ?? Calc.formatDays(s.days);
      // 紅字標記（例如：申請單有、系統無紀錄或無法辨識）。
      // 每一格都給獨立的 style 物件：ExcelJS 複製列（duplicateRow）會共用 style，只改一格會讓其他列一起變紅。
      ['unit', 'card', 'name', 'amount', 'days'].forEach((k) => {
        if (!cols[k]) return;
        const cell = row.getCell(cols[k]);
        const font = { ...((cell.style && cell.style.font) || {}) };
        if (s.highlight) font.color = { argb: 'FFFF0000' };
        else if (font.color && font.color.argb === 'FFFF0000') delete font.color;
        cell.style = { ...(cell.style || {}), font };
      });
      row.commit();
    }
    // 多出的範本列：清空可能殘留的資料欄（保留年、月、備註原樣）
    for (let r = startRow + summary.length; r <= lastRow; r += 1) {
      const row = ws.getRow(r);
      ['unit', 'card', 'name', 'amount', 'days'].forEach((k) => { if (cols[k]) row.getCell(cols[k]).value = null; });
    }

    // 合計公式與製表日期
    const amountLetter = ws.getColumn(cols.amount).letter;
    const totalCell = ws.getRow(totalRow).getCell(cols.amount);
    const total = summary.reduce((a, s) => a + s.amount, 0);
    totalCell.value = { formula: `SUM(${amountLetter}${startRow}:${amountLetter}${lastRow})`, result: total };
    ws.getRow(totalRow).eachCell((cell) => {
      const txt = cellText(cell);
      if (txt.startsWith('製表日期')) {
        const d = today;
        cell.value = `製表日期：${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
      }
    });

    return { filled: summary.length, inserted: Math.max(0, extra), total };
  }

  return { analyzeTemplate, fillTemplate };
});
