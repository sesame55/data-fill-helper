/*
 * 簽到簿表格偵測（純影像處理，不需網路）
 * 輸入：ImageData（RGBA）。輸出：表格的列（水平線之間）與欄（垂直線）位置、哪些列有手寫內容。
 * 簽到簿欄位順序：編號｜日期｜導生姓名｜卡號｜GSM｜輔導老師
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SheetDetect = factory();
})(typeof self !== 'undefined' ? self : this, () => {
  const COLS = ['no', 'date', 'name', 'card', 'gsm', 'tutor'];
  const DATA_ROWS = 20;   // 簽到簿每頁固定 20 列

  function toDark(img, threshold = 150) {
    const { width: w, height: h, data } = img;
    const dark = new Uint8Array(w * h);
    for (let i = 0, p = 0; i < dark.length; i += 1, p += 4) {
      const g = 0.299 * data[p] + 0.587 * data[p + 1] + 0.114 * data[p + 2];
      dark[i] = g < threshold ? 1 : 0;
    }
    return dark;
  }

  /** 把相鄰的線位置合併成一條（取中心） */
  function merge(positions, gap = 3) {
    const out = [];
    positions.forEach((p) => {
      const last = out[out.length - 1];
      if (last && p - last.end <= gap) last.end = p;
      else out.push({ start: p, end: p });
    });
    return out.map((l) => Math.round((l.start + l.end) / 2));
  }

  /**
   * 偵測表格
   * @returns { rows: [{ top, bottom }], cols: [x...], left, right, dataRows: [{ index, top, bottom, ink }] } 或 null
   */
  function detect(img) {
    const { width: w, height: h } = img;
    const lines = toDark(img, 200);   // 格線常是淺灰：用寬鬆門檻
    const dark = toDark(img, 160);    // 手寫墨水：較嚴門檻
    // 1) 水平線：每一列（上下各容許 2px 傾斜）暗點數 ≥ 頁寬 40%
    const hPos = [];
    for (let y = 2; y < h - 2; y += 1) {
      let n = 0;
      for (let x = 0; x < w; x += 1) {
        if (lines[(y - 2) * w + x] || lines[(y - 1) * w + x] || lines[y * w + x] || lines[(y + 1) * w + x] || lines[(y + 2) * w + x]) n += 1;
      }
      if (n >= w * 0.4) hPos.push(y);
    }
    const hAll = merge(hPos, 4);
    if (hAll.length < 4) return null;
    // 取「間距相近」最長的一串線（表格資料列），再把上方較高的表頭列併入
    const gaps = hAll.slice(1).map((y, i) => y - hAll[i]).filter((g) => g > h * 0.012 && g < h * 0.07).sort((a, b) => a - b);
    if (!gaps.length) return null;
    const med = gaps[Math.floor(gaps.length / 2)];
    let best = []; let cur = [hAll[0]];
    for (let i = 1; i < hAll.length; i += 1) {
      const g = hAll[i] - hAll[i - 1];
      if (Math.abs(g - med) <= med * 0.35) cur.push(hAll[i]);
      else if (Math.abs(g - 2 * med) <= med * 0.35) cur.push(Math.round(hAll[i - 1] + g / 2), hAll[i]);   // 漏掉一條線
      else { if (cur.length > best.length) best = cur; cur = [hAll[i]]; }
    }
    if (cur.length > best.length) best = cur;
    // 表頭列比資料列高：往上找一條，使上方那列高度在 1.2～2.6 倍之間
    const fi = hAll.indexOf(best[0]);
    if (fi > 0) {
      const g = best[0] - hAll[fi - 1];
      if (g >= med * 1.2 && g <= med * 2.6) best = [hAll[fi - 1], ...best];
    }
    if (best.length < 4) return null;
    // 以「表格底線」為基準往上推 20 列資料列（簽到簿固定 20 列）：
    // 先把底線延伸到最後一條符合「整數倍列高」的線，再依列高推算每條線，貼齊最近的實際線。
    let bottomLine = best[best.length - 1];
    hAll.filter((y) => y > bottomLine).forEach((y) => {
      const k = Math.round((y - bottomLine) / med);
      if (k >= 1 && k <= 5 && Math.abs(y - bottomLine - k * med) <= med * 0.3) bottomLine = y;
    });
    const snap = (y) => {
      const near = hAll.reduce((a, b) => (Math.abs(b - y) < Math.abs(a - y) ? b : a), hAll[0]);
      return Math.abs(near - y) <= med * 0.3 ? near : Math.round(y);
    };
    // 依實際底線與往上 20 列處的線重新估計列高
    const topGuess = snap(bottomLine - DATA_ROWS * med);
    const rowH = (bottomLine - topGuess) / DATA_ROWS;
    const dataLines = [];
    for (let k = 0; k <= DATA_ROWS; k += 1) dataLines.push(snap(bottomLine - (DATA_ROWS - k) * rowH));
    if (dataLines[0] < 0) return null;
    // 表頭：資料區上方最近的一條線（列高較高）
    const above = hAll.filter((y) => y < dataLines[0] - rowH * 0.5);
    const headerTop = above.length ? above[above.length - 1] : Math.max(0, Math.round(dataLines[0] - rowH * 1.6));
    const hLines = [headerTop, ...dataLines];
    const top = hLines[0]; const bottom = hLines[hLines.length - 1];

    // 2) 垂直線：在表格高度內（左右容許 2px）暗點 ≥ 80%
    const vPos = [];
    const span = bottom - top;
    for (let x = 2; x < w - 2; x += 1) {
      let n = 0;
      for (let y = top; y <= bottom; y += 1) {
        const o = y * w + x;
        if (lines[o - 2] || lines[o - 1] || lines[o] || lines[o + 1] || lines[o + 2]) n += 1;
      }
      if (n >= span * 0.8) vPos.push(x);
    }
    const cols = merge(vPos, 5);
    if (cols.length < 2) return null;
    const left = cols[0]; const right = cols[cols.length - 1];

    const rows = [];
    for (let i = 0; i < hLines.length - 1; i += 1) rows.push({ top: hLines[i], bottom: hLines[i + 1] });

    // 3) 每列（跳過表頭）計算手寫墨水量：排除「編號」欄、格線附近
    const inkIn = (x0, x1, y0, y1) => {
      let n = 0; let total = 0;
      for (let y = y0; y < y1; y += 1) for (let x = x0; x < x1; x += 1) { total += 1; n += dark[y * w + x]; }
      return total ? n / total : 0;
    };
    const pad = 6;
    const dataRows = rows.slice(1).map((r, i) => {
      let ink = 0;
      for (let c = 1; c < cols.length - 1; c += 1) {
        // 每格內縮避開格線，再減去格線殘影
        ink = Math.max(ink, inkIn(cols[c] + pad, cols[c + 1] - pad, r.top + pad, r.bottom - pad));
      }
      return { index: i + 1, top: r.top, bottom: r.bottom, ink };
    });
    return { rows, cols, left, right, dataRows };
  }

  /**
   * 一般表格（列數不固定，例如補助費申請單）：找出表格的列與欄
   * @returns { rows: [{ top, bottom }], cols: [x...] } 或 null（rows[0] 為表頭）
   */
  function detectGrid(img, maxGap = 0.08) {
    const { width: w, height: h } = img;
    const lines = toDark(img, 200);
    const hPos = [];
    for (let y = 2; y < h - 2; y += 1) {
      let n = 0;
      for (let x = 0; x < w; x += 1) {
        if (lines[(y - 1) * w + x] || lines[y * w + x] || lines[(y + 1) * w + x]) n += 1;
      }
      if (n >= w * 0.4) hPos.push(y);
    }
    const hAll = merge(hPos, 4);
    if (hAll.length < 3) return null;
    // 表格本體：相鄰間距 < maxGap（頁高比例）的最長一串線；Word 版申請單列很高，呼叫端可放寬到 0.15
    let best = []; let cur = [hAll[0]];
    for (let i = 1; i < hAll.length; i += 1) {
      if (hAll[i] - hAll[i - 1] < h * maxGap) cur.push(hAll[i]);
      else { if (cur.length > best.length) best = cur; cur = [hAll[i]]; }
    }
    if (cur.length > best.length) best = cur;
    if (best.length < 3) return null;
    const top = best[0]; const bottom = best[best.length - 1];
    const span = bottom - top;
    const vPos = [];
    for (let x = 2; x < w - 2; x += 1) {
      let n = 0;
      for (let y = top; y <= bottom; y += 1) {
        const o = y * w + x;
        if (lines[o - 1] || lines[o] || lines[o + 1]) n += 1;
      }
      if (n >= span * 0.8) vPos.push(x);
    }
    const cols = merge(vPos, 5);
    const rows = [];
    for (let i = 0; i < best.length - 1; i += 1) {
      if (best[i + 1] - best[i] >= h * 0.01) rows.push({ top: best[i], bottom: best[i + 1] });
    }
    return rows.length >= 2 && cols.length >= 4 ? { rows, cols } : null;
  }

  /**
   * 整列刪除線：在列的中段（上下 12%～88%）找一條近乎水平、橫跨 x0～x1 至少 minRatio（預設 40%）的連續暗線
   * （容許上下 ±3px 的歪斜、≤15px 的斷點（筆跡較淡會斷續）；格線的垂直線會自然接上）。手寫字不會有這麼長的連續橫線。
   */
  function isStruck(img, x0, x1, top, bottom, minRatio = 0.4) {
    const { width: w, data } = img;
    // 暗點：一般墨跡灰階 < 160；藍筆筆跡較淡，偏藍（B−R > 40）且灰階 < 200 也算
    const dark = (x, y) => {
      const p = (y * w + x) * 4; const g = 0.299 * data[p] + 0.587 * data[p + 1] + 0.114 * data[p + 2];
      return g < 160 || (g < 200 && data[p + 2] - data[p] > 40);
    };
    // 檢查列高 12%～88%；追蹤範圍不碰上下格線（各留 5px），避免把格線當成刪除線
    const y0 = Math.round(top + (bottom - top) * 0.12); const y1 = Math.round(bottom - (bottom - top) * 0.12);
    const lo = top + 5; const hi = bottom - 5;
    const need = (x1 - x0) * minRatio;
    for (let y = y0; y <= y1; y += 2) {
      let run = 0; let gap = 0; let cy = y;
      for (let x = x0; x < x1; x += 1) {
        // 跟著線走：在目前高度上下 3px 內找暗點
        let hit = -1;
        for (let dy = 0; dy <= 3 && hit < 0; dy += 1) {
          if (cy + dy <= hi && dark(x, cy + dy)) hit = cy + dy;
          else if (dy && cy - dy >= lo && dark(x, cy - dy)) hit = cy - dy;
        }
        if (hit >= 0) { run += 1 + gap; gap = 0; cy = hit; } else if (run && gap < 15) gap += 1; else { run = 0; gap = 0; cy = y; }
        if (run >= need) return true;
      }
    }
    return false;
  }

  /** 判斷有內容的列（墨水比例門檻） */
  const filledRows = (det, threshold = 0.012) => det.dataRows.filter((r) => r.ink >= threshold);

  /** 欄位對應：6 欄時依「編號｜日期｜姓名｜卡號｜GSM｜老師」 */
  function columnRects(det) {
    if (det.cols.length - 1 !== COLS.length) return null;
    const out = {};
    COLS.forEach((k, i) => { out[k] = { x0: det.cols[i], x1: det.cols[i + 1] }; });
    return out;
  }

  return { detect, detectGrid, isStruck, filledRows, columnRects, COLS };
});
