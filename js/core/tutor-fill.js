/*
 * 導生座談：把結果填入原始「導談完成率統計」活頁簿（直接改 xlsx 內的 XML，原檔格式完整保留）
 *
 * ① rawdata（使用者確認 2026-11-09）：
 *  - 依每筆紀錄的「座談日期」月份，填入該月欄位：狀態「V」＋ 座談日期（同月多筆取最早）。
 *  - 只填空白格；已有內容（院內／院外／外訓…）不覆寫，列入報告。
 * ② 簽到簿：
 *  - 有簽到 → 「院內」（空白或原為 V 時寫入；院外／外訓等手動內容不覆寫）。
 *  - 簽到但 rawdata 當月無紀錄 → 「院內」紅底白字。
 *  - 臨培（醫事人員導談補助申請紀錄）每月只有一欄，只寫「院內」。
 *
 * 找月份欄：表頭中導師欄右邊符合「N月」或「民國年.月（115.10）」的欄位；
 *   若有多個同名欄位，以該欄現有日期的年月確認；都沒有日期時，只接受導師欄後「第一個區塊」內的欄位，
 *   避免寫到舊年度或其他區塊（例如「餐費申請」）的同名欄。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./tutor-calc.js'));
  else root.TutorFill = factory(root.TutorCalc);
})(typeof self !== 'undefined' ? self : this, (T) => {
  const text = (v) => {
    if (v && typeof v === 'object' && !(v instanceof Date)) {
      if (v.richText) return v.richText.map((t) => t.text).join('');
      if ('result' in v) return text(v.result);
      if (v.text) return String(v.text);
    }
    if (typeof v === 'number') return String(v);
    return v === null || v === undefined ? '' : String(v).trim();
  };
  const isoOf = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;

  const SHEET_CAT = (name) => {
    if (/刪除/.test(name)) return null;
    if (/^R檢核/.test(name)) return 'R';
    if (/^PGY\(/.test(name)) return 'PGY';
    if (/^C\(/.test(name)) return 'UGY';
    if (/醫事人員/.test(name)) return '臨培';
    return null;
  };

  /** 表頭文字是否代表目標月份 */
  function isMonthHeader(raw, ym) {
    const y = Number(ym.slice(0, 4)); const m = Number(ym.slice(5, 7));
    if (typeof raw === 'number') return Math.abs(raw - ((y - 1911) + m / 100)) < 0.001;
    const s = text(raw).replace(/\s/g, '');
    if (s === `${m}月`) return true;
    const roc = s.match(/^(\d{3})\.(\d{1,2})$/);
    return !!roc && Number(roc[1]) === y - 1911 && Number(roc[2]) === m;
  }

  /** 工作表結構：表頭列、卡號欄、姓名欄、導師欄 */
  function analyzeSheet(ws) {
    let headerRow = 0;
    for (let r = 1; r <= 3 && !headerRow; r += 1) {
      ws.getRow(r).eachCell((c) => { if (/卡號/.test(text(c.value))) headerRow = r; });
    }
    if (!headerRow) return null;
    let cardCol = 0; let tutorCol = 0; let nameCol = 0;
    ws.getRow(headerRow).eachCell((c, col) => {
      const t = text(c.value);
      if (!cardCol && /卡號$/.test(t)) cardCol = col;
      if (!nameCol && /^(學員)?姓名$/.test(t)) nameCol = col;
      if (!tutorCol && /^導師(姓名)?$/.test(t)) tutorCol = col;
    });
    return { headerRow, cardCol, nameCol, tutorCol, monthCache: {} };
  }

  /** 找某月的狀態欄（日期欄 = 狀態欄 + 1；hasDate=false 時只有狀態欄） */
  function findMonthCol(ws, a, ym, hasDate = true) {
    if (a.monthCache[ym] !== undefined) return a.monthCache[ym];
    const hr = ws.getRow(a.headerRow);
    const headers = [];
    hr.eachCell((c, col) => { if (col > a.tutorCol && text(c.value)) headers.push({ col, v: c.value }); });
    // 導師欄後第一個區塊：表頭連續（間隔 ≤ 3 欄）
    let blockEnd = a.tutorCol;
    for (const h of headers) { if (h.col - blockEnd > 3) break; blockEnd = h.col; }
    const cands = headers.filter((h) => isMonthHeader(h.v, ym)).map((h) => h.col);
    let found = 0;
    if (hasDate) {
      // 該欄現有日期的年月分布（偶有打錯年份，所以要「多數」符合才算）
      const statsOf = (col) => {
        let total = 0; let hit = 0;
        for (let r = a.headerRow + 1; r <= ws.rowCount; r += 1) {
          const v = ws.getRow(r).getCell(col + 1).value;
          if (v instanceof Date) { total += 1; if (isoOf(v).slice(0, 7) === ym) hit += 1; }
        }
        return { col, total, hit };
      };
      const info = cands.map(statsOf);
      found = info.find((x) => x.total && x.hit / x.total >= 0.5)?.col
        || info.find((x) => x.total === 0 && x.col <= blockEnd)?.col || 0;
    } else {
      found = cands.find((col) => col <= blockEnd) || 0;
    }
    a.monthCache[ym] = found;
    return found;
  }

  /** 各類所有工作表（新到舊），例如 PGY 有 115 與 114 學年度兩張 */
  function sheetsByCat(wb) {
    const all = {};
    wb.worksheets.forEach((ws) => {
      const cat = SHEET_CAT(ws.name);
      if (!cat) return;
      const start = Number((ws.name.match(/\((\d+(?:\.\d+)?)/) || [])[1] || 0);
      (all[cat] = all[cat] || []).push({ ws, start, a: analyzeSheet(ws) });
    });
    Object.values(all).forEach((l) => l.sort((x, y) => y.start - x.start));
    return all;
  }

  /** 每類取一個工作表（PGY 取起始年最大者） */
  function pickSheets(wb) {
    const pick = {};
    wb.worksheets.forEach((ws) => {
      const cat = SHEET_CAT(ws.name);
      if (!cat) return;
      const start = Number((ws.name.match(/\((\d+(?:\.\d+)?)/) || [])[1] || 0);
      if (!pick[cat] || start > pick[cat].start) pick[cat] = { ws, start, a: analyzeSheet(ws) };
    });
    return pick;
  }

  const rocYm = (ym) => `${Number(ym.slice(0, 4)) - 1911}/${ym.slice(5, 7)}`;

  /**
   * ① rawdata → 規劃要填的儲存格（ExcelJS 只用來讀）
   * 依每筆紀錄的座談日期月份填入；months 省略時處理 rawdata 中所有月份。
   * @returns { months, sheets[], filled[], kept[], notInRoster[], missing[], writes[] }
   */
  function planFill(wb, rawAll, months) {
    const byMonth = new Map();   // ym → Map(card → 最早紀錄)
    [...rawAll].sort((a, b) => a.date.localeCompare(b.date)).forEach((r) => {
      const ym = r.date.slice(0, 7);
      if (months && !months.includes(ym)) return;
      if (!byMonth.has(ym)) byMonth.set(ym, new Map());
      const m = byMonth.get(ym);
      if (!m.has(r.card)) m.set(r.card, { ...r, count: 1 });
      else m.get(r.card).count += 1;
    });
    const yms = [...byMonth.keys()].sort();
    const byCat = sheetsByCat(wb);
    const report = { months: yms, sheets: [], filled: [], kept: [], notInRoster: [], missing: [], writes: [] };

    // 名冊：卡號 → 所有出現的位置（新工作表在前；同類可能有新舊學年度兩張）
    const rosterRows = new Map();
    for (const cat of ['R', 'PGY', 'UGY']) {
      const list = byCat[cat] || [];
      if (!list.length) { report.missing.push(`${T.CAT_LABEL[cat]}：找不到工作表`); continue; }
      list.forEach((p) => {
        if (!p.a?.cardCol) { report.missing.push(`${p.ws.name}：找不到「卡號」表頭`); return; }
        for (let r = p.a.headerRow + 1; r <= p.ws.rowCount; r += 1) {
          const card = T.normCard(text(p.ws.getRow(r).getCell(p.a.cardCol).value));
          if (!card) continue;
          if (!rosterRows.has(card)) rosterRows.set(card, []);
          rosterRows.get(card).push({ ...p, cat, row: r });
        }
      });
    }

    const sheetInfo = new Map();
    for (const ym of yms) {
      byMonth.get(ym).forEach((rec, card) => {
        const locs = rosterRows.get(card);
        if (!locs) { report.notInRoster.push({ ym, cat: rec.cat, card, name: rec.name, tutor: rec.tutor, date: rec.date }); return; }
        // 取第一個「有這個月份欄」的工作表；都沒有時記在最新那張
        const loc = locs.find((l) => findMonthCol(l.ws, l.a, ym, true)) || locs[0];
        const { ws, a, cat, row: r } = loc;
        const key = `${ws.name}|${ym}`;
        const statusCol = findMonthCol(ws, a, ym, true);
        if (!sheetInfo.has(key)) {
          sheetInfo.set(key, { name: ws.name, cat, ym, statusCol: statusCol ? ws.getColumn(statusCol).letter : '', filled: 0, kept: 0, skipped: 0 });
        }
        const info = sheetInfo.get(key);
        const row = ws.getRow(r);
        const name = text(row.getCell(a.nameCol || 1).value);
        const base = { ym, sheet: ws.name, row: r, cat, card, name, date: rec.date, count: rec.count };
        if (!statusCol) { info.skipped += 1; return; }
        const sCell = row.getCell(statusCol); const dCell = row.getCell(statusCol + 1);
        const existing = text(sCell.value);
        const dTxt = dCell.value instanceof Date ? isoOf(dCell.value) : text(dCell.value);
        // 「算完成*」（人工註記）可被覆蓋；其他已有內容不覆寫
        const doneMark = (v) => /算完成/.test(v);
        const canS = !existing || doneMark(existing);
        const canD = !(dCell.value instanceof Date) && (!dTxt || doneMark(dTxt));
        if (!(canS && canD) && !(doneMark(existing) || doneMark(dTxt))) {
          info.kept += 1;
          report.kept.push({ ...base, existing, existingDate: dTxt, same: existing === 'V' && dTxt === rec.date });
          return;
        }
        if (canS) report.writes.push({ sheet: ws.name, row: r, col: statusCol, kind: 'status', value: 'V' });
        if (canD) report.writes.push({ sheet: ws.name, row: r, col: statusCol + 1, kind: 'date', value: rec.date, headerRow: a.headerRow });
        info.filled += 1;
        report.filled.push({ ...base, replaced: doneMark(existing) || doneMark(dTxt) ? '算完成' : '' });
      });
    }
    report.sheets = [...sheetInfo.values()].sort((x, y) => x.ym.localeCompare(y.ym) || x.name.localeCompare(y.name));
    report.sheets.filter((s) => !s.statusCol).forEach((s) => report.missing.push(`${s.name}：找不到 ${rocYm(s.ym)} 的月份欄（${s.skipped} 人未填）`));
    return report;
  }

  /**
   * ② 簽到結果 → 規劃寫入「院內」（TutorCalc.compute 的 rows）
   * 空白或 V → 寫入院內；其他手動內容不覆寫。red=true 的格子用紅底白字。
   */
  function planSignFill(wb, rows, ym) {
    const pick = pickSheets(wb);
    const report = { ym, written: [], kept: [], notInRoster: [], missing: [], writes: [] };
    const signed = rows.filter((x) => x.status === '院內');
    const missingMonth = new Set();
    for (const x of signed) {
      const p = pick[x.cat];
      if (!p?.a?.cardCol) { report.notInRoster.push(x); continue; }
      let r = 0;
      for (let i = p.a.headerRow + 1; i <= p.ws.rowCount && !r; i += 1) {
        if (T.normCard(text(p.ws.getRow(i).getCell(p.a.cardCol).value)) === x.card) r = i;
      }
      if (!r) { report.notInRoster.push(x); continue; }
      const hasDate = x.cat !== '臨培';
      const col = findMonthCol(p.ws, p.a, ym, hasDate);
      if (!col) { missingMonth.add(p.ws.name); continue; }
      const row = p.ws.getRow(r);
      const existing = text(row.getCell(col).value);
      const base = { sheet: p.ws.name, row: r, cat: x.cat, card: x.card, name: x.name, red: x.red, date: x.date };
      if (existing && existing !== 'V' && existing !== '院內' && !/算完成/.test(existing)) { report.kept.push({ ...base, existing }); continue; }
      report.writes.push({ sheet: p.ws.name, row: r, col, kind: 'status', value: '院內', red: x.red });
      if (hasDate && x.date && !(row.getCell(col + 1).value instanceof Date)) {
        report.writes.push({ sheet: p.ws.name, row: r, col: col + 1, kind: 'date', value: x.date, headerRow: p.a.headerRow });
      }
      report.written.push({ ...base, before: existing });
    }
    missingMonth.forEach((n) => report.missing.push(`${n}：找不到 ${rocYm(ym)} 的月份欄`));
    return report;
  }

  // ---------- 直接修改 xlsx 內的 XML（保留原檔所有格式、條件式格式、欄寬…） ----------
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const colLetter = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colNumber = (ref) => ref.replace(/\d+/g, '').split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const excelSerial = (iso) => { const [y, m, d] = iso.split('-').map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };

  /** styles.xml → 判斷某個樣式索引是否為日期格式 */
  function dateStyleChecker(stylesDoc) {
    const custom = {};
    Array.from(stylesDoc.getElementsByTagNameNS(NS, 'numFmt')).forEach((n) => { custom[n.getAttribute('numFmtId')] = n.getAttribute('formatCode'); });
    const xfsParent = stylesDoc.getElementsByTagNameNS(NS, 'cellXfs')[0];
    const xfs = xfsParent ? Array.from(xfsParent.getElementsByTagNameNS(NS, 'xf')) : [];
    return (s) => {
      const xf = xfs[Number(s || 0)];
      if (!xf) return false;
      const id = Number(xf.getAttribute('numFmtId') || 0);
      if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47) || (id >= 27 && id <= 36) || (id >= 50 && id <= 58)) return true;
      const code = (custom[id] || '').replace(/"[^"]*"|\[[^\]]*\]/g, '');
      return /[ymd]/i.test(code) && !/[#0]/.test(code.replace(/[ymd]/gi, ''));
    };
  }

  /**
   * 把 writes 套用到原始 xlsx（ArrayBuffer）→ 回傳新的 Uint8Array
   * deps: { JSZip, DOMParser, XMLSerializer }
   */
  async function patchXlsx(buffer, writes, { JSZip, DOMParser, XMLSerializer }) {
    const zip = await JSZip.loadAsync(buffer);
    const parse = async (path) => new DOMParser().parseFromString(await zip.file(path).async('string'), 'application/xml');
    const wbDoc = await parse('xl/workbook.xml');
    const rels = await parse('xl/_rels/workbook.xml.rels');
    const relTarget = {};
    Array.from(rels.getElementsByTagName('Relationship')).forEach((r) => { relTarget[r.getAttribute('Id')] = r.getAttribute('Target'); });
    const sheetPath = {};
    Array.from(wbDoc.getElementsByTagNameNS(NS, 'sheet')).forEach((s) => {
      const t = relTarget[s.getAttributeNS(NS_R, 'id') || s.getAttribute('r:id')];
      if (t) sheetPath[s.getAttribute('name')] = t.startsWith('/') ? t.slice(1) : `xl/${t.replace(/^\.\//, '')}`;
    });
    const stylesDoc = await parse('xl/styles.xml');
    const isDateStyle = dateStyleChecker(stylesDoc);

    // 換底色／字色：複製原儲存格樣式，換成指定填滿色（與字色）（新增到 styles.xml）
    const styleCache = {};
    let stylesChanged = false;
    const colorStyle = (s, fillRgb, fontRgb = null, bold = false) => {
      const base = Number(s || 0);
      const key = `${base}|${fillRgb}|${fontRgb}|${bold}`;
      if (styleCache[key] !== undefined) return styleCache[key];
      const tag = (name) => stylesDoc.getElementsByTagNameNS(NS, name)[0];
      const fills = tag('fills'); const fonts = tag('fonts'); const cellXfs = tag('cellXfs');
      const xfs = Array.from(cellXfs.getElementsByTagNameNS(NS, 'xf'));
      const baseXf = xfs[base] || xfs[0];
      const xf = baseXf.cloneNode(true);
      // 填滿
      const fill = stylesDoc.createElementNS(NS, 'fill');
      const pf = stylesDoc.createElementNS(NS, 'patternFill'); pf.setAttribute('patternType', 'solid');
      const fg = stylesDoc.createElementNS(NS, 'fgColor'); fg.setAttribute('rgb', fillRgb);
      const bg = stylesDoc.createElementNS(NS, 'bgColor'); bg.setAttribute('indexed', '64');
      pf.appendChild(fg); pf.appendChild(bg); fill.appendChild(pf); fills.appendChild(fill);
      const fillId = Array.from(fills.getElementsByTagNameNS(NS, 'fill')).length - 1;
      fills.setAttribute('count', String(fillId + 1));
      xf.setAttribute('fillId', String(fillId)); xf.setAttribute('applyFill', '1');
      // 字型：複製原字型，改顏色（與粗體）
      if (fontRgb || bold) {
        const fontList = Array.from(fonts.getElementsByTagNameNS(NS, 'font'));
        const font = (fontList[Number(baseXf.getAttribute('fontId') || 0)] || fontList[0]).cloneNode(true);
        if (bold && !font.getElementsByTagNameNS(NS, 'b')[0]) font.insertBefore(stylesDoc.createElementNS(NS, 'b'), font.firstChild);
        if (fontRgb) {
          Array.from(font.getElementsByTagNameNS(NS, 'color')).forEach((n) => font.removeChild(n));
          const color = stylesDoc.createElementNS(NS, 'color'); color.setAttribute('rgb', fontRgb);
          const sz = font.getElementsByTagNameNS(NS, 'sz')[0];
          font.insertBefore(color, sz ? sz.nextSibling : null);
        }
        fonts.appendChild(font);
        fonts.setAttribute('count', String(fontList.length + 1));
        xf.setAttribute('fontId', String(fontList.length)); xf.setAttribute('applyFont', '1');
      }
      cellXfs.appendChild(xf);
      cellXfs.setAttribute('count', String(xfs.length + 1));
      styleCache[key] = xfs.length;
      stylesChanged = true;
      return styleCache[key];
    };
    const redStyle = (s) => colorStyle(s, 'FFFF0000', 'FFFFFFFF', true);

    const bySheet = {};
    writes.forEach((w) => { (bySheet[w.sheet] = bySheet[w.sheet] || []).push(w); });
    for (const [name, list] of Object.entries(bySheet)) {
      const path = sheetPath[name];
      if (!path || !zip.file(path)) throw new Error(`找不到工作表 XML：${name}`);
      const doc = await parse(path);
      const sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
      const rowMap = new Map(Array.from(sheetData.getElementsByTagNameNS(NS, 'row')).map((r) => [Number(r.getAttribute('r')), r]));
      const cellAt = (rowNum, col) => {
        let row = rowMap.get(rowNum);
        if (!row) {
          row = doc.createElementNS(NS, 'row'); row.setAttribute('r', String(rowNum));
          const after = [...rowMap.keys()].filter((k) => k > rowNum).sort((a, b) => a - b)[0];
          sheetData.insertBefore(row, after ? rowMap.get(after) : null);
          rowMap.set(rowNum, row);
        }
        const ref = `${colLetter(col)}${rowNum}`;
        const cells = Array.from(row.getElementsByTagNameNS(NS, 'c'));
        let c = cells.find((x) => x.getAttribute('r') === ref);
        if (!c) {
          c = doc.createElementNS(NS, 'c'); c.setAttribute('r', ref);
          const next = cells.find((x) => colNumber(x.getAttribute('r')) > col);
          row.insertBefore(c, next || null);
        }
        return c;
      };
      const clear = (c) => { while (c.firstChild) c.removeChild(c.firstChild); c.removeAttribute('t'); };
      // 同欄附近已有日期的樣式（原格若不是日期格式時沿用）
      const nearbyDateStyle = (col, rowNum, headerRow) => {
        const keys = [...rowMap.keys()].sort((a, b) => Math.abs(a - rowNum) - Math.abs(b - rowNum));
        for (const k of keys) {
          if (k <= headerRow) continue;
          const c = Array.from(rowMap.get(k).getElementsByTagNameNS(NS, 'c')).find((x) => x.getAttribute('r') === `${colLetter(col)}${k}`);
          if (c && isDateStyle(c.getAttribute('s')) && c.getElementsByTagNameNS(NS, 'v')[0]) return c.getAttribute('s');
        }
        return null;
      };

      for (const w of list) {
        const c = cellAt(w.row, w.col);
        // 只改顏色（不動內容），例如 rawdata 比對成功標綠色
        if (w.kind === 'style') { c.setAttribute('s', String(colorStyle(c.getAttribute('s'), w.fill, w.font || null, !!w.bold))); continue; }
        clear(c);
        if (w.kind === 'status') {
          c.setAttribute('t', 'inlineStr');
          const is = doc.createElementNS(NS, 'is'); const t = doc.createElementNS(NS, 't');
          t.appendChild(doc.createTextNode(String(w.value))); is.appendChild(t); c.appendChild(is);
          if (w.red) c.setAttribute('s', String(redStyle(c.getAttribute('s'))));
        } else {
          if (!isDateStyle(c.getAttribute('s'))) {
            const s = nearbyDateStyle(w.col, w.row, w.headerRow || 1);
            if (s) c.setAttribute('s', s);
          }
          const v = doc.createElementNS(NS, 'v'); v.appendChild(doc.createTextNode(String(excelSerial(w.value)))); c.appendChild(v);
        }
      }
      zip.file(path, new XMLSerializer().serializeToString(doc));
    }
    if (stylesChanged) zip.file('xl/styles.xml', new XMLSerializer().serializeToString(stylesDoc));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  return { isMonthHeader, analyzeSheet, findMonthCol, planFill, planSignFill, patchXlsx, excelSerial };
});
